"""Offline GUI/tag/export smoke test. Requires a display (or xvfb-run on Linux)."""

import io
import os
import runpy
import sys
import tempfile
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    with tempfile.TemporaryDirectory(prefix="phonodex-smoke-") as temp:
        work = Path(temp)
        with patch.dict(os.environ, {
            "PHONODEX_CONFIG_DIR": str(work / "config"),
            "PHONODEX_LIBRARY_DIR": str(work / "music"),
        }), patch.object(tk.Tk, "mainloop"), \
             patch("requests.get", side_effect=AssertionError("Smoke test must remain offline")):
            namespace = runpy.run_path(str(ROOT / "main.py"), run_name="__main__")
        app = namespace["app"]
        from utils.logging import logger
        try:
            app.update()
            assert app.winfo_exists()
            assert app.title() == "Phonodex"

            # Synthetic MPEG frames: no real music, API credentials, or network calls.
            frames = (b"\xff\xfb\x90\x00" + bytes(413)) * 10
            track = work / "Synthetic track.mp3"
            track.write_bytes(frames)
            from mutagen.mp3 import MP3
            from mutagen.id3 import TIT2
            initial = MP3(track)
            initial.add_tags()
            initial.tags.add(TIT2(encoding=3, text=["Initial Title"]))
            initial.save()
            namespace["file_list"].append(str(track))
            namespace["apply_filter"]()
            table = namespace["file_table"]
            items = table.get_children()
            assert len(items) == 1, "Import did not populate the GUI"
            table.selection_set(items[0])
            app.update()
            metadata = {
                "Artist": "Fixture Artist", "Title": "Fixture Title",
                "Album": "Fixture Album", "Album Artist": "Fixture Artist",
                "Track": "1", "Year": "2026", "Genre": "Fixture",
                "Catalog Number": "TEST-001", "Comment": "Offline smoke fixture",
            }
            for field, value in metadata.items():
                namespace["basic_field_vars"][field].set(value)

            from PIL import Image
            image = io.BytesIO()
            Image.new("RGB", (16, 16), "blue").save(image, "PNG")
            globals_ = namespace["apply_basic_fields"].__globals__
            globals_["pending_album_art"] = image.getvalue()
            namespace["apply_basic_fields"]()
            audio = MP3(track)
            assert str(audio.tags["TIT2"]) == "Fixture Title"
            assert audio.tags.getall("APIC")[0].data == image.getvalue()
            assert track.read_bytes().endswith(frames), "Audio frames changed during tagging"

            import config
            config.save_skin_name("Slime")
            export_format = str(work / "music" / "%albumartist%" / "%album%" / "%title%")
            config.save_folder_format(export_format)
            assert config.load_skin_name() == "Slime"

            def approve(_parent, _font, files, skipped, on_confirm):
                assert len(files) == 1 and not skipped
                on_confirm()

            with patch.dict(globals_, {"show_move_confirmation_dialog": approve}), \
                 patch("tkinter.messagebox.showinfo"):
                namespace["organize_files_with_format"]()
            exported = work / "music" / "Fixture Artist" / "Fixture Album" / "Fixture Title.mp3"
            assert exported.is_file(), "Export did not create the expected path"
            assert not track.exists(), "Export is expected to move the source file"
            assert str(MP3(exported).tags["TIT2"]) == "Fixture Title"
            assert exported.read_bytes().endswith(frames)
            print("PASS: GUI startup, synthetic MP3 import, tags/artwork, settings, and export")
        finally:
            logger.set_debug_widget(None)
            app.destroy()


if __name__ == "__main__":
    main()
