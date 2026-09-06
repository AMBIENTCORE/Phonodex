"""Offline regressions; run with python -m unittest discover -s tests -v."""

import importlib.util
import io
import os
import subprocess
import sys
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from PIL import Image

from utils import image_handling
from utils.file_operations import handle_drop
from utils.runtime import default_folder_format, maximize_window, user_data_directory

ROOT = Path(__file__).resolve().parents[1]


class RuntimePathsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.env = patch.dict(os.environ, {
            "HOME": str(self.root), "USERPROFILE": str(self.root),
        }, clear=True)
        self.env.start()
        self.addCleanup(self.env.stop)

    def test_windows_source_keeps_existing_location(self):
        self.assertEqual(user_data_directory(self.root, system="Windows"), self.root)

    def test_windows_frozen_keeps_appdata(self):
        os.environ["LOCALAPPDATA"] = str(self.root)
        self.assertEqual(user_data_directory(ROOT, True, "Windows"), self.root / "Phonodex")

    def test_linux_uses_xdg(self):
        os.environ["XDG_CONFIG_HOME"] = str(self.root)
        self.assertEqual(user_data_directory(ROOT, system="Linux"), self.root / "Phonodex")

    def test_linux_default_stays_outside_source(self):
        self.assertEqual(user_data_directory(ROOT, system="Linux"),
                         Path.home() / ".config" / "Phonodex")

    def test_config_override_on_both_platforms(self):
        os.environ["PHONODEX_CONFIG_DIR"] = str(self.root / "config")
        for system in ("Linux", "Windows"):
            self.assertEqual(user_data_directory(ROOT, True, system), self.root / "config")

    def test_relative_config_rejected(self):
        os.environ["PHONODEX_CONFIG_DIR"] = "relative/config"
        with self.assertRaisesRegex(ValueError, "absolute"):
            user_data_directory(ROOT)

    def test_windows_default_export_unchanged(self):
        self.assertTrue(default_folder_format("Windows").startswith("D:\\Music\\Collection\\"))

    def test_linux_export_is_native_absolute_path(self):
        value = default_folder_format("Linux")
        self.assertTrue(Path(value).is_absolute())
        self.assertIn("%albumartist%", value)
        self.assertNotIn("D:", value)

    def test_library_override(self):
        os.environ["PHONODEX_LIBRARY_DIR"] = str(self.root / "music")
        self.assertTrue(default_folder_format().startswith(str(self.root / "music")))

    def test_relative_library_rejected(self):
        os.environ["PHONODEX_LIBRARY_DIR"] = "relative/music"
        with self.assertRaisesRegex(ValueError, "absolute"):
            default_folder_format()

    def test_settings_survive_reload_and_preserve_skin(self):
        os.environ["PHONODEX_CONFIG_DIR"] = str(self.root / "config")

        def load():
            spec = importlib.util.spec_from_file_location("test_config", ROOT / "config.py")
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module

        first = load()
        first.Config.API_KEY_FILE.write_text("fixture-token", encoding="utf-8")
        first.Config.GENRES_FILE.write_text("Fixture Genre\n", encoding="utf-8")
        self.assertTrue(first.save_skin_name("Slime"))
        first.save_folder_format(str(self.root / "music" / "%title%"))
        second = load()
        self.assertEqual(second.ACTIVE_SKIN, "Slime")
        self.assertEqual(second.folder_format, first.folder_format)
        self.assertEqual(second.Config.API_KEY_FILE.read_text(), "fixture-token")
        self.assertEqual(second.Config.GENRES_FILE.read_text(), "Fixture Genre\n")

    def test_unwritable_config_path_fails_explicitly(self):
        path = self.root / "not-a-directory"
        path.write_text("fixture")
        os.environ["PHONODEX_CONFIG_DIR"] = str(path)
        spec = importlib.util.spec_from_file_location("bad_config", ROOT / "config.py")
        module = importlib.util.module_from_spec(spec)
        with self.assertRaises(FileExistsError):
            spec.loader.exec_module(module)


class WindowTests(unittest.TestCase):
    def test_x11_maximization(self):
        window = Mock()
        window.tk.call.return_value = "x11"
        maximize_window(window)
        window.attributes.assert_called_once_with("-zoomed", True)
        window.state.assert_not_called()

    def test_windows_maximization(self):
        window = Mock()
        window.tk.call.return_value = "win32"
        maximize_window(window)
        window.state.assert_called_once_with("zoomed")

    def test_unsupported_maximization_warns(self):
        window = Mock()
        window.tk.call.return_value = "x11"
        window.attributes.side_effect = tk.TclError("unsupported")
        with patch("utils.runtime.log_message") as log:
            maximize_window(window)
        self.assertIn("WARNING", log.call_args.args[0])


class DropTests(unittest.TestCase):
    def test_tcl_paths_with_spaces_braces_unicode_and_apostrophes(self):
        with tempfile.TemporaryDirectory() as temp:
            paths = [Path(temp) / name for name in (
                "simple.mp3", "an album's track.mp3", "track {live}.flac", "caf\u00e9.mp3")]
            for path in paths:
                path.touch()
            data = tk._join([path.as_posix() for path in paths])
            result = []
            handle_drop(data, file_list_var=result, supported_extensions=[".mp3", ".flac"])
            self.assertEqual([Path(p) for p in result], paths)

    def test_windows_and_posix_payloads_use_same_parser(self):
        interpreter = tk.Tcl()
        for data, expected in (
            ("{C:/Music/Album A/track.mp3}", "C:/Music/Album A/track.mp3"),
            ("/music/track.mp3", "/music/track.mp3"),
        ):
            with patch("utils.file_operations.tk.Tcl", return_value=interpreter), \
                 patch("utils.file_operations.os.path.exists", return_value=True), \
                 patch("utils.file_operations.os.path.isdir", return_value=False), \
                 patch("utils.file_operations.os.path.isfile", return_value=True):
                files = []
                handle_drop(data, file_list_var=files, supported_extensions=[".mp3"])
            self.assertEqual(files, [expected])

    def test_invalid_payload_reports_error(self):
        with patch("utils.file_operations.log_message") as log:
            self.assertEqual(handle_drop("{unclosed", supported_extensions=[".mp3"]), [])
        self.assertIn("Invalid dropped", log.call_args.args[0])


class ClipboardTests(unittest.TestCase):
    def setUp(self):
        stream = io.BytesIO()
        Image.new("RGB", (4, 4), "blue").save(stream, "PNG")
        self.png = stream.getvalue()
        image_handling._original_image_data = None
        self.addCleanup(setattr, image_handling, "_original_image_data", None)

    def test_linux_copy_without_xclip_keeps_lossless_internal_paste(self):
        with patch.object(sys, "platform", "linux"), \
             patch.object(image_handling.shutil, "which", return_value=None):
            self.assertTrue(image_handling.copy_image_to_clipboard(self.png))
            self.assertEqual(image_handling.paste_image_from_clipboard(), self.png)

    def test_linux_system_copy_sends_png_without_shell(self):
        with patch.object(sys, "platform", "linux"), \
             patch.object(image_handling.shutil, "which", return_value="/usr/bin/xclip"), \
             patch.object(image_handling.subprocess, "run") as run:
            self.assertTrue(image_handling.copy_image_to_clipboard(self.png))
        self.assertEqual(run.call_args.args[0][0], "xclip")
        self.assertTrue(run.call_args.kwargs["input"].startswith(b"\x89PNG"))
        self.assertNotIn("shell", run.call_args.kwargs)

    def test_linux_clipboard_failure_preserves_internal_copy(self):
        with patch.object(sys, "platform", "linux"), \
             patch.object(image_handling.shutil, "which", return_value="xclip"), \
             patch.object(image_handling.subprocess, "run",
                          side_effect=subprocess.TimeoutExpired("xclip", 5)):
            self.assertTrue(image_handling.copy_image_to_clipboard(self.png))
            self.assertEqual(image_handling.paste_image_from_clipboard(), self.png)

    def test_linux_external_paste(self):
        for mode in ("RGB", "RGBA", "P", "LA", "L"):
            with self.subTest(mode=mode):
                stream = io.BytesIO()
                Image.new(mode, (4, 4)).save(stream, "PNG")
                with patch.object(sys, "platform", "linux"), \
                     patch.object(image_handling.shutil, "which", return_value="xclip"), \
                     patch.object(image_handling.subprocess, "run",
                                  return_value=Mock(stdout=stream.getvalue())):
                    data = image_handling.paste_image_from_clipboard()
                self.assertIsNotNone(data)
                self.assertEqual(Image.open(io.BytesIO(data)).size, (4, 4))

    @unittest.skipUnless(sys.platform == "win32", "Windows clipboard branch")
    def test_windows_clipboard_api_is_preserved(self):
        with patch.object(image_handling, "win32clipboard") as clipboard:
            self.assertTrue(image_handling.copy_image_to_clipboard(self.png))
        clipboard.OpenClipboard.assert_called_once()
        clipboard.SetClipboardData.assert_called_once()
        clipboard.CloseClipboard.assert_called_once()


if __name__ == "__main__":
    unittest.main()
