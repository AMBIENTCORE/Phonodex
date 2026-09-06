"""Wait for a real X connection, then exec without changing the supervised PID."""

import os
import subprocess
import sys
import time


def wait_for_x(timeout=30):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            result = subprocess.run(
                ["/usr/bin/xdpyinfo"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=2,
                check=False,
            )
            if result.returncode == 0:
                return
        except subprocess.TimeoutExpired:
            pass
        time.sleep(0.2)
    raise RuntimeError("Xvfb did not become ready within 30 seconds")


def main():
    try:
        if len(sys.argv) < 2:
            raise ValueError("A program to execute is required")
        wait_for_x()
        os.execv(sys.argv[1], sys.argv[1:])
    except (OSError, RuntimeError, ValueError) as error:
        print(str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
