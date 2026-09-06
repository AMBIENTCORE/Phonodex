"""Check the supervised app, its X window, noVNC HTTP and the VNC handshake."""

import http.client
import os
from pathlib import Path
import re
import socket
import subprocess
import sys
import urllib.request
import xmlrpc.client

REQUIRED = {"fail-fast", "xvfb", "dbus", "openbox", "vnc", "novnc", "phonodex"}


class UnixConnection(http.client.HTTPConnection):
    def __init__(self, path):
        super().__init__("localhost", timeout=3)
        self.path = path

    def connect(self):
        self.sock = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        self.sock.settimeout(self.timeout)
        self.sock.connect(self.path)


class UnixTransport(xmlrpc.client.Transport):
    def __init__(self, path):
        super().__init__()
        self.path = path

    def make_connection(self, host):
        return UnixConnection(self.path)


def supervisor():
    # Docker exec/healthcheck do not inherit variables set by the entrypoint.
    config = Path(os.environ.get("PHONODEX_CONFIG_DIR", "/config"))
    return xmlrpc.client.ServerProxy(
        "http://localhost",
        transport=UnixTransport(str(config / ".runtime" / "supervisor.sock")),
    )


def read_exact(connection, length):
    data = bytearray()
    while len(data) < length:
        chunk = connection.recv(length - len(data))
        if not chunk:
            raise RuntimeError("Connection closed before the handshake completed")
        data.extend(chunk)
    return bytes(data)


def check_health():
    with supervisor() as server:
        processes = {item["name"]: item for item in server.supervisor.getAllProcessInfo()}
    for name in sorted(REQUIRED):
        if name not in processes or processes[name]["statename"] != "RUNNING":
            raise RuntimeError(f"Required process is not RUNNING: {name}")
    pid = processes["phonodex"]["pid"]
    command = Path(f"/proc/{pid}/cmdline").read_bytes().split(b"\0")
    if b"/opt/phonodex/main.py" not in command:
        raise RuntimeError("The supervised Phonodex process has not started the app")
    tree = subprocess.run(
        ["/usr/bin/xwininfo", "-display", ":99", "-root", "-tree"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        check=True, timeout=3,
    ).stdout
    if not re.search(r'0x[0-9a-fA-F]+\s+"Phonodex":', tree):
        raise RuntimeError("The Phonodex X window is missing")
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open("http://127.0.0.1:6080/vnc.html", timeout=3) as response:
        if response.status != 200 or b"noVNC" not in response.read(65536):
            raise RuntimeError("noVNC is not serving its browser client")
    with socket.create_connection(("127.0.0.1", 5900), timeout=3) as connection:
        if not re.fullmatch(rb"RFB \d{3}\.\d{3}\n", read_exact(connection, 12)):
            raise RuntimeError("VNC did not provide a valid RFB greeting")


def main():
    try:
        check_health()
    except Exception as error:
        print(f"unhealthy: {error}", file=sys.stderr)
        return 1
    print("healthy: Phonodex window/process, X, noVNC and VNC")
    return 0


if __name__ == "__main__":
    sys.exit(main())
