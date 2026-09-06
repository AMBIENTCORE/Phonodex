"""Fast standard-library checks; no Docker daemon or GUI dependencies required."""

import io
import re
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import entrypoint
import fail_fast
import healthcheck
import wait_for_x


class EntrypointTests(unittest.TestCase):
    def test_supported_screens(self):
        for value in ("1600x900x24", "1920x1080x24", "3840x2160x24"):
            with self.subTest(value=value):
                self.assertEqual(entrypoint.validate_screen(value), value)

    def test_invalid_screens(self):
        for value in ("800x600x24", "1920x800x24", "1920x1080x16", "9999x9999x24", "1920x1080x24 -ac"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                entrypoint.validate_screen(value)

    def test_rejects_root_before_starting(self):
        with patch.object(entrypoint.os, "geteuid", return_value=0, create=True), \
                patch.object(entrypoint.os, "execv") as execute, \
                patch.object(entrypoint.sys, "stderr", io.StringIO()):
            self.assertEqual(entrypoint.main(), 1)
            execute.assert_not_called()

    def test_read_only_config_is_actionable(self):
        with patch.object(entrypoint.os, "geteuid", return_value=1000, create=True), \
                patch.object(entrypoint.os, "umask"), \
                patch.object(entrypoint, "writable_directory", side_effect=PermissionError("read-only")), \
                patch.object(entrypoint.sys, "stderr", io.StringIO()) as errors:
            self.assertEqual(entrypoint.main(), 1)
            self.assertIn("mounted music is never chowned", errors.getvalue())


class ReadinessTests(unittest.TestCase):
    def test_waits_for_a_successful_x_connection(self):
        results = [SimpleNamespace(returncode=1), SimpleNamespace(returncode=0)]
        with patch.object(wait_for_x.subprocess, "run", side_effect=results) as run, \
                patch.object(wait_for_x.time, "sleep") as sleep:
            wait_for_x.wait_for_x()
            self.assertEqual(run.call_count, 2)
            sleep.assert_called_once()

    def test_missing_display_times_out(self):
        with patch.object(wait_for_x.time, "monotonic", side_effect=[0, 31]), \
                self.assertRaises(RuntimeError):
            wait_for_x.wait_for_x()


class HealthTests(unittest.TestCase):
    def setUp(self):
        self.patches = [
            patch.object(healthcheck, "supervisor"),
            patch.object(healthcheck.Path, "read_bytes", return_value=b"python\0/opt/phonodex/main.py\0"),
            patch.object(healthcheck.subprocess, "run", return_value=SimpleNamespace(stdout='0x400001 "Phonodex": ("tk" "Tk")')),
            patch.object(healthcheck.urllib.request, "build_opener"),
            patch.object(healthcheck.socket, "create_connection"),
        ]
        server, self.command, self.tree, opener, connection = [item.start() for item in self.patches]
        for item in self.patches:
            self.addCleanup(item.stop)
        self.processes = [
            {"name": name, "statename": "RUNNING", "pid": 42}
            for name in healthcheck.REQUIRED
        ]
        server.return_value.__enter__.return_value.supervisor.getAllProcessInfo.return_value = self.processes
        response = opener.return_value.open.return_value.__enter__.return_value
        response.status = 200
        response.read.return_value = b"<title>noVNC</title>"
        self.vnc = connection.return_value.__enter__.return_value
        self.vnc.recv.return_value = b"RFB 003.008\n"

    def test_full_stack_is_healthy(self):
        healthcheck.check_health()

    def test_every_required_process_is_checked(self):
        for process in self.processes:
            with self.subTest(process=process["name"]):
                process["statename"] = "FATAL"
                with self.assertRaisesRegex(RuntimeError, re.escape(process["name"])):
                    healthcheck.check_health()
                process["statename"] = "RUNNING"

    def test_readiness_wrapper_is_not_the_app(self):
        self.command.return_value = b"python\0wait_for_x.py\0"
        with self.assertRaisesRegex(RuntimeError, "has not started"):
            healthcheck.check_health()

    def test_missing_window_is_unhealthy(self):
        self.tree.return_value.stdout = '0x400001 "Other": ("tk" "Tk")'
        with self.assertRaisesRegex(RuntimeError, "window is missing"):
            healthcheck.check_health()

    def test_closed_vnc_connection_is_unhealthy(self):
        self.vnc.recv.return_value = b""
        with self.assertRaisesRegex(RuntimeError, "Connection closed"):
            healthcheck.check_health()

    def test_invalid_vnc_greeting_is_unhealthy(self):
        self.vnc.recv.return_value = b"not vnc data"
        with self.assertRaisesRegex(RuntimeError, "valid RFB"):
            healthcheck.check_health()

    def test_greeting_can_be_split_across_packets(self):
        connection = MagicMock()
        connection.recv.side_effect = [b"RFB ", b"003.", b"008\n"]
        self.assertEqual(healthcheck.read_exact(connection, 12), b"RFB 003.008\n")


class FailFastTests(unittest.TestCase):
    def test_exit_and_fatal_shut_down_supervisor(self):
        for event in ("PROCESS_STATE_EXITED", "PROCESS_STATE_FATAL"):
            with self.subTest(event=event):
                payload = "processname:phonodex groupname:phonodex"
                data = f"eventname:{event} len:{len(payload)}\n{payload}"
                with patch.object(fail_fast.sys, "stdin", io.StringIO(data)), \
                        patch.object(fail_fast.sys, "stdout", io.StringIO()) as output, \
                        patch.object(fail_fast.sys, "stderr", io.StringIO()), \
                        patch.object(fail_fast.os, "getppid", return_value=123), \
                        patch.object(fail_fast.os, "kill") as kill:
                    fail_fast.main()
                    kill.assert_called_once_with(123, fail_fast.signal.SIGTERM)
                    self.assertEqual(output.getvalue(), "READY\nRESULT 2\nOK")


if __name__ == "__main__":
    unittest.main()
