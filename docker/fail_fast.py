"""Shut down the entire stack when a required process exits or cannot start."""

import os
import signal
import sys


def main():
    while True:
        # stdout is exclusively Supervisor's event-listener wire protocol.
        sys.stdout.write("READY\n")
        sys.stdout.flush()
        line = sys.stdin.readline()
        if not line:
            return
        headers = dict(field.split(":", 1) for field in line.split())
        payload = sys.stdin.read(int(headers["len"]))
        sys.stdout.write("RESULT 2\nOK")
        sys.stdout.flush()
        if headers["eventname"] in {"PROCESS_STATE_EXITED", "PROCESS_STATE_FATAL"}:
            print(f"Required process failed/exited: {payload}", file=sys.stderr, flush=True)
            os.kill(os.getppid(), signal.SIGTERM)
            return


if __name__ == "__main__":
    main()
