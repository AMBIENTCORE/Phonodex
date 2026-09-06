"""Exercise a built image using disposable Docker containers and one test volume.

Run from the repository: python docker/smoke.py --image phonodex:local
Requires only Python's standard library, Docker, and an already-built image.
Never uses real music or credentials; all created resources are removed on exit.
"""

import argparse
import base64
import hashlib
import json
import os
import socket
import struct
import subprocess
import sys
import time
import urllib.request
import uuid

from healthcheck import read_exact

HEALTHCHECK = "/opt/phonodex/docker/healthcheck.py"
SUPERVISOR = "/opt/phonodex/docker/supervisord.conf"


def docker(*arguments, check=True, timeout=60):
    result = subprocess.run(
        ["docker", *arguments], capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout, check=False,
    )
    if check and result.returncode:
        raise RuntimeError(f"docker {' '.join(arguments)}:\n{result.stdout}{result.stderr}")
    return result


def state(name):
    return json.loads(docker("inspect", "--format", "{{json .State}}", name).stdout)


def wait_healthy(name, timeout):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        current = state(name)
        if not current["Running"]:
            raise RuntimeError(f"Container exited before becoming healthy: {current}")
        if current.get("Health", {}).get("Status") == "healthy":
            # Reject a stale Docker health result after intentionally stopping the app.
            if docker("exec", name, "/usr/bin/python3", HEALTHCHECK, check=False).returncode == 0:
                return
        time.sleep(1)
    raise RuntimeError("Timed out waiting for the actual Phonodex window and browser stack")


def browser_check(port):
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(f"http://127.0.0.1:{port}/vnc.html", timeout=5) as response:
        if response.status != 200 or b"noVNC" not in response.read():
            raise RuntimeError("Published HTTP endpoint is not noVNC")
    key = base64.b64encode(os.urandom(16)).decode("ascii")
    request = (
        f"GET /websockify HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\n"
        "Upgrade: websocket\r\nConnection: Upgrade\r\n"
        f"Sec-WebSocket-Key: {key}\r\nSec-WebSocket-Version: 13\r\n"
        "Sec-WebSocket-Protocol: binary\r\n\r\n"
    )
    with socket.create_connection(("127.0.0.1", port), timeout=5) as connection:
        connection.sendall(request.encode("ascii"))
        header = bytearray()
        while not header.endswith(b"\r\n\r\n"):
            header.extend(read_exact(connection, 1))
            if len(header) > 16384:
                raise RuntimeError("WebSocket upgrade headers are too large")
        lines = header.decode("iso-8859-1").split("\r\n")
        if lines[0].split()[1] != "101":
            raise RuntimeError(f"WebSocket upgrade failed: {lines[0]}")
        fields = {
            name.lower(): value.strip()
            for name, value in (line.split(":", 1) for line in lines[1:] if ":" in line)
        }
        expected = base64.b64encode(
            hashlib.sha1((key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11").encode("ascii")).digest()
        ).decode("ascii")
        if fields.get("sec-websocket-accept") != expected:
            raise RuntimeError("Invalid WebSocket handshake")
        greeting = bytearray()
        while len(greeting) < 12:
            first, second = read_exact(connection, 2)
            if second & 0x80:
                raise RuntimeError("The WebSocket server unexpectedly masked its frame")
            length = second & 0x7F
            if length == 126:
                length = struct.unpack("!H", read_exact(connection, 2))[0]
            elif length == 127:
                length = struct.unpack("!Q", read_exact(connection, 8))[0]
            if length > 65536 or first & 0x0F not in (0, 2):
                raise RuntimeError("Unexpected frame instead of the VNC greeting")
            greeting.extend(read_exact(connection, length))
        if not greeting.startswith(b"RFB 003."):
            raise RuntimeError("The WebSocket endpoint did not relay the VNC handshake")


def run_container(name, volume, image):
    docker(
        "run", "--detach", "--name", name,
        "--publish", "127.0.0.1::6080",
        "--mount", f"type=volume,source={volume},target=/config",
        "--cap-drop", "ALL", "--security-opt", "no-new-privileges:true",
        "--shm-size", "256m", image,
    )


def check_isolation(name):
    script = """
import os
from pathlib import Path
assert os.geteuid() != 0, "Container is root"
for status in Path("/proc").glob("[0-9]*/status"):
    try:
        lines = status.read_text().splitlines()
    except FileNotFoundError:
        continue
    uid = next(line.split()[2] for line in lines if line.startswith("Uid:"))
    assert int(uid) == os.geteuid(), f"Unexpected process UID in {status}: {uid}"
listeners = []
for table in ("/proc/net/tcp", "/proc/net/tcp6"):
    if not Path(table).exists():
        continue
    for row in Path(table).read_text().splitlines()[1:]:
        columns = row.split()
        address, port = columns[1].split(":")
        if port == "170C" and columns[3] == "0A":
            listeners.append(address)
assert listeners, "No VNC listener"
assert all(address in ("0100007F", "00000000000000000000000001000000")
           for address in listeners), f"VNC exposed beyond loopback: {listeners}"
"""
    docker("exec", name, "/usr/bin/python3", "-c", script)


def run(args):
    suffix = uuid.uuid4().hex[:12]
    name, volume = f"phonodex-smoke-{suffix}", f"phonodex-smoke-config-{suffix}"
    docker("image", "inspect", args.image)
    docker("volume", "create", volume)
    try:
        run_container(name, volume, args.image)
        wait_healthy(name, args.timeout)
        port = int(docker("port", name, "6080/tcp").stdout.strip().rsplit(":", 1)[1])
        browser_check(port)
        check_isolation(name)
        print("PASS: ready app window, HTTP/WebSocket/VNC and non-root isolation")

        marker = uuid.uuid4().hex
        write = (
            "import os; from pathlib import Path; "
            f"(Path(os.environ['PHONODEX_CONFIG_DIR']) / '.smoke-persistence').write_text('{marker}')"
        )
        docker("exec", name, "/usr/bin/python3", "-c", write)
        docker("stop", "--time", "25", name)
        if state(name)["ExitCode"] != 0:
            raise RuntimeError("Supervisor did not shut down cleanly on SIGTERM")
        docker("rm", name)
        run_container(name, volume, args.image)
        wait_healthy(name, args.timeout)
        read = (
            "import os; from pathlib import Path; "
            f"assert (Path(os.environ['PHONODEX_CONFIG_DIR']) / '.smoke-persistence').read_text() == '{marker}'"
        )
        docker("exec", name, "/usr/bin/python3", "-c", read)
        print("PASS: clean signal shutdown and config persistence across recreation")

        docker("exec", name, "/usr/bin/supervisorctl", "-c", SUPERVISOR, "stop", "phonodex")
        result = docker("exec", name, "/usr/bin/python3", HEALTHCHECK, check=False)
        if result.returncode == 0 or "phonodex" not in result.stderr:
            raise RuntimeError("Healthcheck accepted a stopped application")
        port = int(docker("port", name, "6080/tcp").stdout.strip().rsplit(":", 1)[1])
        browser_check(port)
        print("PASS: a live noVNC page cannot hide a stopped application")
        docker("exec", name, "/usr/bin/supervisorctl", "-c", SUPERVISOR, "start", "phonodex")
        wait_healthy(name, args.timeout)

        pid = int(docker(
            "exec", name, "/usr/bin/supervisorctl", "-c", SUPERVISOR, "pid", "xvfb",
        ).stdout.strip())
        docker("exec", name, "/usr/bin/python3", "-c", f"import os, signal; os.kill({pid}, signal.SIGKILL)")
        deadline = time.monotonic() + 30
        while state(name)["Running"]:
            if time.monotonic() >= deadline:
                raise RuntimeError("A crashed display server left the container running")
            time.sleep(0.5)
        print("PASS: a required-process crash shuts down the complete stack")

        for options, label in (
            (["--env", "PHONODEX_SCREEN=800x600x24"], "invalid resolution"),
            (["--mount", f"type=volume,source={volume},target=/config,readonly"], "read-only config"),
        ):
            # Reuse our known name so even a timeout is covered by finally cleanup.
            docker("rm", name)
            result = docker("run", "--name", name, *options, args.image, check=False, timeout=40)
            if result.returncode == 0 or "Phonodex startup failed:" not in result.stderr:
                raise RuntimeError(f"Startup did not reject {label}: {result.stdout}{result.stderr}")
            print(f"PASS: startup rejects {label}")
    except Exception:
        logs = docker("logs", "--tail", "80", name, check=False)
        print(logs.stdout + logs.stderr, file=sys.stderr)
        raise
    finally:
        docker("rm", "--force", name, check=False)
        docker("volume", "rm", volume, check=False)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--image", default="phonodex:local")
    parser.add_argument("--timeout", type=int, default=180, help="Startup timeout in seconds")
    args = parser.parse_args()
    try:
        run(args)
    except (OSError, RuntimeError, subprocess.TimeoutExpired) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1
    print("All container smoke checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
