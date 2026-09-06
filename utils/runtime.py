"""Platform-specific paths and window behavior."""

import os
import platform
import tkinter as tk
from pathlib import Path

from utils.logging import log_message


def user_data_directory(project_root, frozen=False, system=None):
    system = system or platform.system()
    override = os.environ.get("PHONODEX_CONFIG_DIR")
    if override:
        directory = Path(override).expanduser()
        if not directory.is_absolute():
            raise ValueError("PHONODEX_CONFIG_DIR must be an absolute path")
        return directory
    if system == "Windows":
        if frozen:
            return Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "Phonodex"
        return Path(project_root)
    return Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config") / "Phonodex"


def default_folder_format(system=None):
    system = system or platform.system()
    library = os.environ.get("PHONODEX_LIBRARY_DIR")
    if library:
        root = Path(library).expanduser()
        if not root.is_absolute():
            raise ValueError("PHONODEX_LIBRARY_DIR must be an absolute path")
    elif system == "Windows":
        return r"D:\Music\Collection\%genre%\%year%\[%catalognumber%] %albumartist% - %album%\%artist% - %title%"
    else:
        root = Path.home() / "Music" / "Collection"
    return str(root / "%genre%" / "%year%" /
               "[%catalognumber%] %albumartist% - %album%" / "%artist% - %title%")


def maximize_window(window):
    try:
        if window.tk.call("tk", "windowingsystem") == "x11":
            window.attributes("-zoomed", True)
        else:
            window.state("zoomed")
    except tk.TclError as error:
        log_message(f"[WARNING] Window manager cannot maximize Phonodex: {error}")
