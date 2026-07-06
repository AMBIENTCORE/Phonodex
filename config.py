"""
Configuration management for Phonodex application.
"""

from pathlib import Path
import json
import os
import sys

# Project root (directory containing this config.py).
# Used for bundled read-only assets when running as a script.
_PROJECT_ROOT = Path(__file__).resolve().parent

# User-data root: where api_key.txt, folder_format_settings.json, etc. live.
# When frozen by PyInstaller, store user data in the standard per-user
# location (%LOCALAPPDATA%\Phonodex on Windows). This survives uninstalls,
# works regardless of whether the app is installed to Program Files
# (read-only) or AppData, and matches how other Windows apps behave.
# When running as a plain script, keep files next to the project for
# convenience during development.
if getattr(sys, 'frozen', False):
    _local_appdata = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
    _USER_DATA_ROOT = Path(_local_appdata) / "Phonodex"
    _USER_DATA_ROOT.mkdir(parents=True, exist_ok=True)
    # One-time migration: if an old api_key.txt or settings file exists
    # next to the .exe (from previous app versions), move it into the
    # new location so the user doesn't have to re-enter their key.
    _legacy_root = Path(sys.executable).resolve().parent
    for _name in ("api_key.txt", "folder_format_settings.json"):
        _legacy = _legacy_root / _name
        _new = _USER_DATA_ROOT / _name
        if _legacy.exists() and not _new.exists():
            try:
                _new.write_bytes(_legacy.read_bytes())
            except Exception:
                pass
else:
    _USER_DATA_ROOT = _PROJECT_ROOT

# Default folder structure format
DEFAULT_FOLDER_FORMAT = "D:\\Music\\Collection\\%genre%\\%year%\\[%catalognumber%] %albumartist% - %album%\\%artist% - %title%"
folder_format = DEFAULT_FOLDER_FORMAT

# Settings file path (anchored next to the exe when frozen, otherwise next to config.py)
SETTINGS_FILE = str(_USER_DATA_ROOT / "folder_format_settings.json")

def load_settings():
    """Load settings from file."""
    global folder_format
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                folder_format = settings.get('folder_format', DEFAULT_FOLDER_FORMAT)
    except Exception as e:
        print(f"Error loading settings: {e}")
        folder_format = DEFAULT_FOLDER_FORMAT

# Load settings at startup
load_settings()

class Config:
    """Main configuration class."""

    # App version -- keep in sync with AppVersion in Phonodex.iss.
    # Used for display (window title / About dialog if added later) and
    # for the "check for updates" comparison against GitHub Releases.
    # Bump both this and Phonodex.iss (plus OutputBaseFilename) on every
    # release.
    APP_VERSION = "1.1"

    # API Configuration
    DISCOGS_SEARCH_URL = "https://api.discogs.com/database/search"
    API_KEY_FILE = _USER_DATA_ROOT / "api_key.txt"
    # Genre autocomplete dictionary (one genre per line; lines starting with # are comments).
    # Location: %LOCALAPPDATA%\Phonodex\genres.txt when frozen, next to the project when run as script.
    GENRES_FILE = _USER_DATA_ROOT / "genres.txt"
    API = {
        "RATE_LIMIT": 60,
        "TIMEOUT": 10,
        "RATE_LIMIT_WAIT": 60,  # seconds
        "USAGE_THRESHOLDS": {
            "WARNING": 0.7,  # 70% - switch to orange
            "CRITICAL": 0.9  # 90% - switch to red
        }
    }
    MAX_API_CALLS_PER_MINUTE = 60
    API_RATE_LIMIT_WAIT = 60  # seconds
    
    # Window Configuration
    WINDOW_TITLE = "Phonodex"
    WINDOW_SIZE = "1600x900"
    MIN_WINDOW_SIZE = (1600, 900)
    
    # GUI Element Sizes
    API_BUTTON_WIDTH = 3
    TABLE_HEIGHT = 20
    DEBUG_LOG_HEIGHT = 12
    PROCESSING_LOG_HEIGHT = 6
    
    # File Types
    SUPPORTED_AUDIO_EXTENSIONS = [".mp3", ".flac", ".m4a", ".mp4", ".wma", ".ogg", ".wav"]
    FILE_TYPE_DESCRIPTION = "Audio Files"
    
    # Album Art Configuration
    ALBUM_ART = {
        "COVER_SIZE": 240,  # Default size for album art display
        "DEFAULT_IMAGE": "assets/no_cover.png",  # Path to default "no cover" image
        "PREFER_ITUNES": True  # Try iTunes Search API for album art before falling back to Discogs
    }
    
    # GUI Layout
    PADDING = {
        "DEFAULT": 10,
        "SMALL": 5
    }
    
    # Status Messages
    MESSAGES = {
        "API_KEY_MISSING": "API Key is required!",
        "API_KEY_SAVED": "API Key saved successfully!",
        "NO_FILES_SELECTED": "No files selected!",
        "API_RESUMING": "API rate limit reset, resuming operations..."
    }
    
    # Colors
    COLORS = {
        "SUCCESS": "#4CAF50",  # Material Design Green
        "ERROR": "#F44336",    # Material Design Red
        "BACKGROUND": "#1e1e1e",  # Dark background
        "SECONDARY_BACKGROUND": "#252526",  # Slightly lighter dark background
        "BORDER": "#3e3e42",  # Column divider / outline (VS Code separator gray)
        "TEXT": "#ffffff",  # White text
        "SECONDARY_TEXT": "#cccccc",  # Light gray text
        "VALID_ENTRY": "#1b3a1b",    # Dark green
        "INVALID_ENTRY": "#3a1b1b",  # Dark red
        "UPDATED_ROW": "#1b3a1b",    # Dark green
        "FAILED_ROW": "#3a1b1b",     # Dark red
        "PROGRESSBAR": {
            "GREEN": "#4CAF50",   # Material Design Green
            "ORANGE": "#FF9800",  # Material Design Orange
            "RED": "#F44336",     # Material Design Red
            "TROUGH": "#2d2d2d"   # Dark gray
        }
    }

    # UI Dimensions
    DIMENSIONS = {
        "API_ENTRY_WIDTH": 40,
        "SAVE_BUTTON_WIDTH": 3,
        "PROGRESS_BAR_LENGTH": 100,
        "TABLE_HEIGHT": 20,
        # Debug log shrunk from 12 to 6 rows so the file table can claim
        # the reclaimed vertical space on the right side.
        "DEBUG_LOG_HEIGHT": 6,
        "DEBUG_LOG_WIDTH": 80,
        "PROCESSING_LOG_HEIGHT": 6,
        "PROCESSING_LOG_WIDTH": 80,
        "TABLE_COLUMN_WIDTH": 200
    }

    # Font Configuration
    FONTS = {
        "DEFAULT_SIZE": 10,
        "TABLE_SIZE": 8,
        "TABLE_HEADING_SIZE": 8,
        "LOG_SIZE": 9
    }

    # Style Configuration
    STYLES = {
        "THEME": "clam",
        "FRAME": {
            "LEFT_PANEL": "LeftPanel.TFrame",
            "SECTION": "Section.TLabelframe",
            "SECTION_LABEL": "Section.TLabelframe.Label"
        },
        "CUSTOM_FONT": {
            "FAMILY": "I pixel u",
            "FILE": "assets/I-pixel-u.ttf"
        },
        # Internal button padding (padx/pady on tk.Button). Lowered from 5
        # to 2 to fit more buttons vertically in the left panel now that
        # PROCESS/EXPORT/STOP/checkboxes/REFRESH/REMOVE SELECTED were moved
        # over from the right side. Buttons stay readable at this size.
        "WIDGET_PADDING": 2,
        "PANEL_WEIGHTS": {
            "LEFT": 1,
            "RIGHT": 3
        }
    }

    # Folder Structure Settings
    FOLDER_STRUCTURE = {
        "DEFAULT_FORMAT": DEFAULT_FOLDER_FORMAT,
        "SETTINGS_FILE": SETTINGS_FILE
    }


# ---------------------------------------------------------------------------
# Skins
# ---------------------------------------------------------------------------
# A skin overrides the aesthetic colors, the custom "chrome" font (the pixel
# font), and the default "no cover" album-art image. The data table / logs
# keep their monospace (Consolas) font. Status colors are shared via the
# Default palette so green=updated / red=failed stays meaningful unless a
# skin deliberately overrides them.
#
# The selected skin is saved in the settings JSON under "skin" and applied at
# startup BEFORE any styles are built. Switching skins auto-restarts the app.

import copy as _copy

# Snapshot the built-in look as the "Default" skin before any overlay.
_DEFAULT_SKIN = {
    "display_name": "DANGER",
    # Aesthetic accent for previews/highlights (its "Pantone"). This is the
    # app's red -- the LEAVE / STOP danger red -- NOT SUCCESS, which stays a
    # fixed semantic green. Each new skin supplies its own accent.
    "accent": "#990000",
    "colors": _copy.deepcopy(Config.COLORS),
    "font": dict(Config.STYLES["CUSTOM_FONT"]),
    "no_cover": Config.ALBUM_ART["DEFAULT_IMAGE"],
}

# Registry of available skins. New skins are added here.
SKINS = {
    "Default": _DEFAULT_SKIN,
    "Slime": {
        "display_name": "SLIME",
        "accent": "#aaff00",
        "colors": {
            "BACKGROUND": "#141a0d",
            "SECONDARY_BACKGROUND": "#1f2913",
            "BORDER": "#7cb518",
            "TEXT": "#c6ff4d",
            "SECONDARY_TEXT": "#8a9a3c",
            "PROGRESSBAR": {"TROUGH": "#1f2913"},
        },
        "no_cover": "assets/no_cover_slime.png",
    },
    "Ye Olde Tagger": {
        "display_name": "YE OLDE TAGGER",
        "accent": "#a86b2e",
        "labels": {"STOP": "HALT", "LEAVE": "BEGONE"},
        "colors": {
            "BACKGROUND": "#e8d8b0",
            "SECONDARY_BACKGROUND": "#d8c290",
            "BORDER": "#7a5230",
            "TEXT": "#3b2a17",
            "SECONDARY_TEXT": "#6b533a",
            "VALID_ENTRY": "#b9cf8f",
            "INVALID_ENTRY": "#dda28f",
            "UPDATED_ROW": "#b9cf8f",
            "FAILED_ROW": "#dda28f",
            "PROGRESSBAR": {"TROUGH": "#d8c290"},
        },
        "no_cover": "assets/no_cover_yeoldetagger.png",
    },
    "Amber Terminal": {
        "display_name": "AMBER TERMINAL",
        "accent": "#ffcf4d",
        "labels": {"STOP": "KILL", "LEAVE": "EXIT"},
        "colors": {
            "BACKGROUND": "#0a0a08",
            "SECONDARY_BACKGROUND": "#14120c",
            "BORDER": "#5a4410",
            "TEXT": "#ffb000",
            "SECONDARY_TEXT": "#a9760a",
            "PROGRESSBAR": {"TROUGH": "#14120c"},
        },
        "no_cover": "assets/no_cover_amberterminal.png",
    },
    # ------------------------------------------------------------------
    # Extra skins. Light themes override the status-row/entry tints so
    # OK/NOK stay readable on a light background. Each carries a signature
    # accent (used on selected rows) and themed STOP/LEAVE labels.
    # ------------------------------------------------------------------
    "Sakura": {
        "display_name": "SAKURA",
        "accent": "#ff6f91",
        "colors": {
            "BACKGROUND": "#fff0f5",
            "SECONDARY_BACKGROUND": "#ffe0ec",
            "BORDER": "#e6a8c3",
            "TEXT": "#5c3a4a",
            "SECONDARY_TEXT": "#a86f86",
            "VALID_ENTRY": "#d6ecc9",
            "INVALID_ENTRY": "#f7c5c5",
            "UPDATED_ROW": "#d6ecc9",
            "FAILED_ROW": "#f7c5c5",
            "PROGRESSBAR": {"TROUGH": "#ffe0ec"},
        },
        "no_cover": "assets/no_cover_sakura.png",
    },
    "E-Ink": {
        "display_name": "E-INK",
        "accent": "#2b2b2b",
        "labels": {"STOP": "STOP", "LEAVE": "CLOSE"},
        "colors": {
            "BACKGROUND": "#f4f3ee",
            "SECONDARY_BACKGROUND": "#e7e5dd",
            "BORDER": "#4a4a4a",
            "TEXT": "#1a1a1a",
            "SECONDARY_TEXT": "#6b6b6b",
            "VALID_ENTRY": "#dfe6df",
            "INVALID_ENTRY": "#efd9d9",
            "UPDATED_ROW": "#dfe6df",
            "FAILED_ROW": "#efd9d9",
            "PROGRESSBAR": {"TROUGH": "#e7e5dd"},
        },
        "no_cover": "assets/no_cover_eink.png",
    },
}

SKIN_SETTINGS_KEY = "skin"


def _read_settings_dict():
    """Read the whole settings JSON as a dict (empty dict on any failure)."""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def load_skin_name():
    """Return the saved skin name, defaulting to 'Default'."""
    name = _read_settings_dict().get(SKIN_SETTINGS_KEY, "Default")
    return name if name in SKINS else "Default"


def save_skin_name(name):
    """Persist the chosen skin name into settings JSON (preserving other keys)."""
    data = _read_settings_dict()
    data[SKIN_SETTINGS_KEY] = name
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving skin setting: {e}")
        return False


def _contrast_text(hex_color):
    """Return black-ish or white (whichever is more readable) over hex_color."""
    try:
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        return "#1a1a1a" if luminance > 140 else "#ffffff"
    except Exception:
        return "#ffffff"


def apply_skin(name):
    """Overlay a skin's colors/font/no-cover onto Config (in place).

    Starts from a clean Default palette so no stale keys survive a switch;
    then overlays the chosen skin's values. Unspecified keys fall back to
    Default. Called once at startup.
    """
    skin = SKINS.get(name, _DEFAULT_SKIN)
    merged = _copy.deepcopy(_DEFAULT_SKIN["colors"])
    for k, v in skin.get("colors", {}).items():
        if isinstance(v, dict) and isinstance(merged.get(k), dict):
            merged[k].update(v)
        else:
            merged[k] = v
    # Expose the skin's accent (its Pantone) + a readable text color on it so
    # ttk styles can actually USE it (e.g. selected rows) -- not just previews.
    accent = skin.get("accent") or _DEFAULT_SKIN.get("accent") or "#990000"
    merged["ACCENT"] = accent
    merged["ACCENT_TEXT"] = skin.get("accent_text") or _contrast_text(accent)
    Config.COLORS = merged
    Config.STYLES["CUSTOM_FONT"] = dict(skin.get("font", _DEFAULT_SKIN["font"]))
    Config.ALBUM_ART["DEFAULT_IMAGE"] = skin.get("no_cover", _DEFAULT_SKIN["no_cover"])
    # Optional per-skin button label overrides (e.g. STOP -> HALT).
    Config.LABELS = skin.get("labels", {})


def skin_label(text):
    """Return the active skin's label override for `text`, or `text` unchanged."""
    return getattr(Config, "LABELS", {}).get(text, text)


# Currently active skin name; apply it immediately so downstream style
# building (in main.py) reads the skinned values.
ACTIVE_SKIN = load_skin_name()
apply_skin(ACTIVE_SKIN)


