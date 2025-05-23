"""
Configuration management for Phonodex application.
"""

from pathlib import Path
import json
import os

# Default folder structure format
DEFAULT_FOLDER_FORMAT = "D:\\Music\\Collection\\%genre%\\%year%\\[%catalognumber%] %albumartist% - %album%\\%artist% - %title%"
folder_format = DEFAULT_FOLDER_FORMAT

# Settings file path
SETTINGS_FILE = "folder_format_settings.json"

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

def save_settings():
    """Save settings to file."""
    try:
        with open(SETTINGS_FILE, 'w') as f:
            settings = {
                'folder_format': folder_format
            }
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"Error saving settings: {e}")

# Load settings at startup
load_settings()

class Config:
    """Main configuration class."""
    
    # API Configuration
    DISCOGS_SEARCH_URL = "https://api.discogs.com/database/search"
    API_KEY_FILE = Path("api_key.txt")
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
        "DEFAULT_IMAGE": "assets/no_cover.png"  # Path to default "no cover" image
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
        "DEBUG_LOG_HEIGHT": 12,
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
        "WIDGET_PADDING": 5,
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


