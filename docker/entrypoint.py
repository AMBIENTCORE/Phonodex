"""Validate writable state and replace this process with the supervisor."""

import os
from pathlib import Path
import re
import sys
import uuid

SUPERVISOR_CONFIG = "/opt/phonodex/docker/supervisord.conf"


def validate_screen(value):
    match = re.fullmatch(r"(\d{4})x(\d{3,4})x24", value)
    if not match:
        raise ValueError("PHONODEX_SCREEN must be WIDTHxHEIGHTx24, e.g. 1920x1080x24")
    width, height = map(int, match.groups())
    if not (1600 <= width <= 7680 and 900 <= height <= 4320):
        raise ValueError("PHONODEX_SCREEN must be at least 1600x900 and at most 7680x4320")
    return value


def writable_directory(value):
    path = Path(value)
    if not path.is_absolute():
        raise ValueError(f"Container data directories must be absolute: {value}")
    path.mkdir(parents=True, exist_ok=True)
    probe = path / f".phonodex-write-check-{uuid.uuid4().hex}"
    try:
        with probe.open("x", encoding="utf-8") as stream:
            stream.write("writable\n")
    finally:
        probe.unlink(missing_ok=True)
    return str(path)


def main():
    try:
        if os.geteuid() == 0:
            raise ValueError("Run Phonodex as a non-root numeric UID:GID")
        os.umask(0o077)
        validate_screen(os.environ.get("PHONODEX_SCREEN", "1920x1080x24"))
        config = writable_directory(os.environ.get("PHONODEX_CONFIG_DIR", "/config"))
        library = writable_directory(os.environ.get("PHONODEX_LIBRARY_DIR", "/music"))
        writable_directory("/work")
        runtime = writable_directory(str(Path(config) / ".runtime"))
        os.chmod(runtime, 0o700)
        os.environ.update(
            HOME=config,
            DISPLAY=":99",
            PHONODEX_CONFIG_DIR=config,
            PHONODEX_LIBRARY_DIR=library,
            PHONODEX_RUNTIME_DIR=runtime,
            DBUS_SESSION_BUS_ADDRESS=f"unix:path={runtime}/bus",
            PHONODEX_SCREEN=os.environ.get("PHONODEX_SCREEN", "1920x1080x24"),
            XDG_RUNTIME_DIR=runtime,
            XDG_CONFIG_HOME=str(Path(config) / ".config"),
            XDG_CACHE_HOME=str(Path(config) / ".cache"),
            XDG_DATA_HOME=str(Path(config) / ".local" / "share"),
        )
        os.execv("/usr/bin/supervisord", ["supervisord", "-c", SUPERVISOR_CONFIG])
    except (OSError, ValueError) as error:
        print(
            f"Phonodex startup failed: {error}\n"
            "Config, library and /work must be writable by the selected UID:GID. "
            "Prepare mount ownership yourself; mounted music is never chowned.",
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
