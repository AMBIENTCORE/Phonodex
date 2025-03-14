import tkinter as tk
from tkinter import filedialog, ttk, StringVar, IntVar, font, BooleanVar, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import requests
import os
import time
import sys
import shutil
from collections import Counter
import mutagen
from mutagen.easyid3 import EasyID3
from mutagen.mp3 import MP3
from mutagen.flac import FLAC, Picture
from mutagen.mp4 import MP4
from mutagen.oggvorbis import OggVorbis
from mutagen.asf import ASF
from mutagen.wave import WAVE
from mutagen.id3 import ID3, APIC, TPE1, TIT2, TALB, TPE2, TXXX, TDRC, TRCK, TCON
import threading
from config import Config
import json
import win32clipboard
import hashlib
import win32con
from array import array

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# Cache and API rate limiter
album_catalog_cache = {}
failed_search_cache = set()  # Cache for artist-album combinations that returned no results
cache_lock = threading.Lock()  # Lock for thread-safe cache access
processed_lock = threading.Lock()  # Lock for thread-safe processed files access
file_metadata_cache = {}  # Cache for file metadata

# Track selected folders for refresh functionality
selected_folders = set()  # Store paths of selected folders

# Default folder structure format
DEFAULT_FOLDER_FORMAT = "D:\\Music\\Collection\\%genre%\\%year%\\[%catalognumber%] %albumartist% - %album%\\%artist% - %title%"
folder_format = DEFAULT_FOLDER_FORMAT

# Settings file
SETTINGS_FILE = "folder_format_settings.json"

# Load settings from file
def load_settings():
    global folder_format
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                folder_format = settings.get('folder_format', DEFAULT_FOLDER_FORMAT)
    except Exception as e:
        print(f"Error loading settings: {e}")
        folder_format = DEFAULT_FOLDER_FORMAT

# Save settings to file
def save_settings():
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

# Sorting variables
sort_column = None  # Track which column we're sorting by
sort_reverse = False  # Track sort direction

def treeview_sort_column(tv, col, reverse):
    """Sort treeview content when a column header is clicked."""
    global sort_column, sort_reverse
    
    # Get all items in the table
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    
    try:
        # Try to sort as numbers if possible
        l.sort(key=lambda t: float(t[0]), reverse=reverse)
    except ValueError:
        # Fall back to string sort if not numbers
        l.sort(key=lambda t: t[0].lower(), reverse=reverse)  # Case-insensitive sort

    # Rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    # Reverse sort next time
    sort_column = col
    sort_reverse = not reverse
    
    # Update header arrows
    for column in columns:  # columns is your list of column names
        if column == col:
            tv.heading(column, text=f"{column} {'↓' if reverse else '↑'}")
        else:
            tv.heading(column, text=column)

# API rate limiting
rate_limit_total = 60  # Default to authenticated limit
rate_limit_used = 0
rate_limit_remaining = 60
first_request_time = 0  # Track when the first request was made in the current window

# Load saved API Key
if os.path.exists(Config.API_KEY_FILE):
    with open(Config.API_KEY_FILE, "r") as f:
        DISCOGS_API_TOKEN = f.read().strip()
else:
    DISCOGS_API_TOKEN = ""

# ---------------- GUI SETUP ---------------- #

app = TkinterDnD.Tk()
app.title(Config.WINDOW_TITLE)
app.geometry(Config.WINDOW_SIZE)
app.minsize(*Config.MIN_WINDOW_SIZE)

# Set ttk style to clam
style = ttk.Style()

def configure_styles(style, custom_font):
    """Configure all ttk styles for the application."""
    # Set theme
    style.theme_use(Config.STYLES["THEME"])
    
    # Configure dark theme styles
    style.configure('Dark.TPanedwindow', background=Config.COLORS["BACKGROUND"], sashwidth=0)  # Set sashwidth to 0 to remove separator
    style.configure('TFrame', background=Config.COLORS["BACKGROUND"])
    style.configure('TButton', padding=Config.STYLES["WIDGET_PADDING"], font=custom_font, background=Config.COLORS["SECONDARY_BACKGROUND"], foreground=Config.COLORS["TEXT"], relief="solid", borderwidth=1)
    style.map('TButton',
        relief=[('pressed', 'sunken'), ('!pressed', 'solid')],
        borderwidth=[('pressed', 1), ('!pressed', 1)])
    style.configure('TEntry', padding=Config.STYLES["WIDGET_PADDING"], fieldbackground=Config.COLORS["SECONDARY_BACKGROUND"], foreground=Config.COLORS["TEXT"])
    style.configure('TLabel', background=Config.COLORS["BACKGROUND"], foreground=Config.COLORS["TEXT"], font=custom_font)
    style.configure('TText', padding=Config.STYLES["WIDGET_PADDING"], background=Config.COLORS["SECONDARY_BACKGROUND"], foreground=Config.COLORS["TEXT"])
    
    # Table styles with dark theme
    style.configure('Treeview', 
                   rowheight=15,
                   font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                   background=Config.COLORS["SECONDARY_BACKGROUND"],
                   foreground=Config.COLORS["TEXT"],
                   fieldbackground=Config.COLORS["SECONDARY_BACKGROUND"])
    style.configure('Treeview.Heading', 
                   font=('Consolas', Config.FONTS["TABLE_HEADING_SIZE"], 'bold'),
                   background=Config.COLORS["BACKGROUND"],
                   foreground=Config.COLORS["TEXT"])

    # Left panel styles with dark theme
    style.configure('LeftPanel.TFrame', background=Config.COLORS["BACKGROUND"])
    style.configure('Section.TLabelframe', background=Config.COLORS["BACKGROUND"], foreground=Config.COLORS["TEXT"], padding=0)
    style.configure('Section.TLabelframe.Label', 
                   font=custom_font, 
                   background=Config.COLORS["BACKGROUND"],
                   foreground=Config.COLORS["TEXT"])

    # Entry validation styles with dark theme
    style.configure('Valid.TEntry', fieldbackground=Config.COLORS["VALID_ENTRY"], foreground=Config.COLORS["TEXT"])
    style.configure('Invalid.TEntry', fieldbackground=Config.COLORS["INVALID_ENTRY"], foreground=Config.COLORS["TEXT"])

    # Custom checkbutton style with dark theme
    style.configure('Custom.TCheckbutton', 
                   font=(Config.STYLES["CUSTOM_FONT"]["FAMILY"], Config.FONTS["DEFAULT_SIZE"]),
                   background=Config.COLORS["BACKGROUND"],
                   foreground=Config.COLORS["TEXT"])

    # Progress bar styles
    style.configure("API.Horizontal.TProgressbar",
                   background=Config.COLORS["PROGRESSBAR"]["GREEN"],
                   troughcolor=Config.COLORS["PROGRESSBAR"]["TROUGH"])
    style.configure("Gradient.Horizontal.TProgressbar",
                   background=Config.COLORS["SUCCESS"],
                   troughcolor=Config.COLORS["PROGRESSBAR"]["TROUGH"])

    # Table row status styles with dark theme
    style.configure("updated", background=Config.COLORS["UPDATED_ROW"])
    style.configure("failed", background=Config.COLORS["FAILED_ROW"])

# Load custom font
custom_font_path = resource_path(Config.STYLES["CUSTOM_FONT"]["FILE"])
if not os.path.exists(custom_font_path):
    raise FileNotFoundError(f"Required font file '{custom_font_path}' not found! Please ensure it's in the same directory as the script.")

# Create font configuration
custom_font = font.Font(family=Config.STYLES["CUSTOM_FONT"]["FAMILY"], size=Config.FONTS["DEFAULT_SIZE"])

# Configure all styles
configure_styles(style, custom_font)

# Variables
api_key_var = StringVar(value=DISCOGS_API_TOKEN)
save_art_var = BooleanVar(value=True)  # Default to True
save_year_var = BooleanVar(value=True)  # Default to True
save_catalog_var = BooleanVar(value=True)  # Default to True
stop_processing = False  # Flag to control processing state
file_list = []
processed_count = 0
processed_files = set()
updated_files = set()  # Track which files have been updated with catalog numbers
editing_item = None  # Track which item is being edited
editing_column = None  # Track which column is being edited
editing_entry = None  # Reference to the editing entry widget

# Create main horizontal split
main_paned = ttk.PanedWindow(app, orient=tk.HORIZONTAL)
main_paned.pack(fill="both", expand=True)

# Set the fixed size for the album cover
album_cover_size = 240  # This matches the typical width of the notes frame
left_panel_width = album_cover_size + 100  # Add padding for the frame borders

# Create left panel with fixed width
left_panel = ttk.Frame(main_paned, style=Config.STYLES["FRAME"]["LEFT_PANEL"], width=left_panel_width)
left_panel.pack_propagate(False)  # Prevent the frame from shrinking
main_paned.add(left_panel, weight=0)  # Set weight to 0 to maintain fixed size

# Create right panel
right_panel = ttk.Frame(main_paned)
main_paned.add(right_panel, weight=1)  # Set weight to 1 to allow expansion

# Buttons directly in left panel
buttons_subframe = ttk.Frame(left_panel)
buttons_subframe.pack(fill="x", pady=(15, 5), padx=10)

for button_text, command in [
    ("FILES", lambda: select_files()),
    ("FOLDER", lambda: select_folder()),
    ("LEAVE", app.quit)
]:
    tk.Button(buttons_subframe, text=button_text,
              command=command,
              font=custom_font,
              bg=Config.COLORS["SECONDARY_BACKGROUND"],
              fg=Config.COLORS["TEXT"] if button_text != "LEAVE" else "#990000",
              padx=Config.STYLES["WIDGET_PADDING"],
              pady=Config.STYLES["WIDGET_PADDING"]).pack(side="left", padx=Config.PADDING["SMALL"], expand=True, fill="x")

# API Key Entry with Save Button - directly in left panel
api_subframe = ttk.Frame(left_panel)
api_subframe.pack(fill="x", pady=(5, 5), padx=10)

api_entry = tk.Entry(api_subframe, 
                    textvariable=api_key_var, 
                    width=Config.DIMENSIONS["API_ENTRY_WIDTH"], 
                    justify="center", 
                    font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                    bg=Config.COLORS["SECONDARY_BACKGROUND"],
                    fg=Config.COLORS["TEXT"],
                    insertbackground=Config.COLORS["TEXT"])
api_entry.pack(side="left", fill="x", expand=True)

# Update the validation style
def update_api_entry_style(is_valid):
    api_entry.configure(bg=Config.COLORS["VALID_ENTRY"] if is_valid else Config.COLORS["INVALID_ENTRY"])

# Initial style based on API token
update_api_entry_style(bool(DISCOGS_API_TOKEN))

save_button = tk.Button(api_subframe, text="💾", 
                       width=Config.DIMENSIONS["SAVE_BUTTON_WIDTH"],
                       command=lambda: save_api_key(),
                       font=custom_font,
                       bg=Config.COLORS["SECONDARY_BACKGROUND"],
                       fg=Config.COLORS["TEXT"],
                       padx=Config.STYLES["WIDGET_PADDING"],
                       pady=Config.STYLES["WIDGET_PADDING"])
save_button.pack(side="left", padx=(Config.PADDING["SMALL"], 0))

# Basic Fields Section (Now as a direct child of left_panel)
basic_fields_subframe = ttk.Frame(left_panel, style='TFrame', borderwidth=0)
basic_fields_subframe.pack(fill="x", pady=(5, 5), padx=10)

# Create variables for the fields
basic_field_vars = {
    "Artist": StringVar(),
    "Title": StringVar(),
    "Album": StringVar(),
    "Album Artist": StringVar(),
    "Catalog Number": StringVar(),
    "Year": StringVar(),
    "Track": StringVar(),
    "Genre": StringVar()
}


# Create text boxes for basic fields - use the original field names for variable lookup
for field in ["Artist", "Title", "Album", "Album Artist", "Catalog Number", "Year", "Track", "Genre"]:
    field_frame = ttk.Frame(basic_fields_subframe, style='TFrame', borderwidth=0)
    field_frame.pack(fill="x", padx=5, pady=2)
    
    # Get the current custom font details to create a smaller version
    current_font = font.nametofont(custom_font.name)
    current_size = current_font.cget("size")
    smaller_size = current_size - 1  # Decrease by 1pt
    
    # Use capitalized field name directly
    tk.Label(field_frame, 
             text=field.upper() + ":", 
             font=(current_font.cget("family"), smaller_size),  # Smaller font
             bg=Config.COLORS["BACKGROUND"],
             fg=Config.COLORS["TEXT"],
             bd=0).pack(fill="x")  # No border on label
    
    # Keep using the original field name to access the variable
    tk.Entry(field_frame,
             textvariable=basic_field_vars[field],
             font=('Consolas', Config.FONTS["TABLE_SIZE"]),
             bg=Config.COLORS["SECONDARY_BACKGROUND"],
             fg=Config.COLORS["TEXT"],
             insertbackground=Config.COLORS["TEXT"],
             bd=1).pack(fill="x", pady=(2, 0))  # Minimal border on entry

# Add a frame for album cover as a direct child of left_panel
album_cover_subframe = ttk.Frame(left_panel, style='TFrame', borderwidth=0)
album_cover_subframe.pack(fill="none", expand=False, pady=(5, 5), padx=10)

# Create a fixed-size frame to contain the album art
album_art_container = ttk.Frame(album_cover_subframe, 
                               style=Config.STYLES["FRAME"]["SECTION"],
                               relief="solid",
                               borderwidth=1)
album_art_container.pack(padx=5, pady=5)

# Create a label for the album cover
album_cover_label = ttk.Label(album_art_container,
                             background=Config.COLORS["SECONDARY_BACKGROUND"],
                             relief="flat",
                             borderwidth=0,
                             padding=0)  # Remove any padding
album_cover_label.pack(padx=0, pady=0, fill="both", expand=True)  # Make label fill the container

# Set the fixed size for the album cover
album_art_container.configure(width=album_cover_size, height=album_cover_size)
# Force the container to keep its size
album_art_container.pack_propagate(False)

# Create a variable to store the current album art image reference
current_album_art = None
# Create a variable to store pending image data (for pasting before saving)
pending_album_art = None

# Create a context menu for the album art
def show_album_art_context_menu(event):
    """Display the context menu when right-clicking on album art."""
    album_art_context_menu.tk_popup(event.x_root, event.y_root)

# Function to paste image from clipboard
def paste_image_from_clipboard():
    """Paste image from clipboard to album art display."""
    global pending_album_art
    
    try:
        from PIL import ImageGrab, Image
        import io
        
        # Get image from clipboard
        img = ImageGrab.grabclipboard()
        if img is None:
            log_message("[COVER] No image found in clipboard", log_type="processing")
            return
        
        if not isinstance(img, Image.Image):
            log_message("[COVER] Clipboard content is not an image", log_type="processing")
            return
        
        # Convert image to bytes
        img_buffer = io.BytesIO()
        img.save(img_buffer, format='PNG')
        image_data = img_buffer.getvalue()
        
        # Store the image data for later saving
        pending_album_art = image_data
        
        # Display the image
        update_album_art_display(image_data)
        log_message("[COVER] Image pasted from clipboard (not saved until 'SAVE METADATA' is clicked)", log_type="processing")
    except Exception as e:
        log_message(f"[COVER] Error pasting image from clipboard: {e}", log_type="processing")

# Function to remove the album art
def remove_album_art():
    """Remove the album art and set to default."""
    global pending_album_art
    
    # Set pending_album_art to None to indicate removal
    pending_album_art = "REMOVE"
    
    # Load default image
    load_default_album_art()
    log_message("[COVER] Album art removed (not applied until 'SAVE METADATA' is clicked)", log_type="processing")

def copy_album_art_to_clipboard():
    """Copy the currently displayed album art to clipboard."""
    global current_album_art, pending_album_art
    
    try:
        from PIL import Image
        import io
        import win32clipboard
        import win32con
        
        image_data = None
        
        if pending_album_art and pending_album_art != "REMOVE":
            # If there's pending art, use that
            image_data = pending_album_art
        elif current_album_art:
            # Get all selected items
            selected_items = file_table.selection()
            if not selected_items:
                log_message("[COVER] No files selected", log_type="processing")
                return
                
            # If multiple items are selected, verify they all have the same album art
            if len(selected_items) > 1:
                art_hashes = set()
                
                # Collect art data from all selected files
                for selected_item in selected_items:
                    values = file_table.item(selected_item)['values']
                    
                    # Get the metadata we'll use to match
                    selected_artist = str(values[0]).strip()
                    selected_title = str(values[1]).strip()
                    selected_album = str(values[2]).strip()
                    selected_albumartist = str(values[4]).strip()
                    
                    # Find the matching file
                    for file_path in file_list:
                        if file_path in file_metadata_cache:
                            metadata = file_metadata_cache[file_path]
                            if (str(metadata.get("artist", "")).strip() == selected_artist and
                                str(metadata.get("title", "")).strip() == selected_title and
                                str(metadata.get("album", "")).strip() == selected_album and
                                str(metadata.get("albumartist", "")).strip() == selected_albumartist):
                                
                                # Found the matching file, now get its album art
                                audio = get_audio_file(file_path)
                                if audio:
                                    temp_data = None
                                    if isinstance(audio, MP3) and audio.tags:
                                        apic_frames = audio.tags.getall("APIC")
                                        if apic_frames:
                                            temp_data = apic_frames[0].data
                                    elif isinstance(audio, FLAC) and audio.pictures:
                                        temp_data = audio.pictures[0].data
                                    elif isinstance(audio, MP4) and "covr" in audio:
                                        temp_data = audio["covr"][0]
                                    
                                    if temp_data:
                                        # Hash the image data
                                        art_hash = hashlib.md5(temp_data).hexdigest()
                                        art_hashes.add(art_hash)
                                        
                                        # Store the first image data we find
                                        if not image_data:
                                            image_data = temp_data
                                break
                
                # If we found different art hashes or some files don't have art
                if len(art_hashes) != 1:
                    log_message("[COVER] Selected files have different album art or some files have no art", log_type="processing")
                    return
            else:
                # Single file selected, get its art
                selected_item = selected_items[0]
                values = file_table.item(selected_item)['values']
                
                # Get the metadata we'll use to match
                selected_artist = str(values[0]).strip()
                selected_title = str(values[1]).strip()
                selected_album = str(values[2]).strip()
                selected_albumartist = str(values[4]).strip()
                
                # Find the matching file
                for file_path in file_list:
                    if file_path in file_metadata_cache:
                        metadata = file_metadata_cache[file_path]
                        if (str(metadata.get("artist", "")).strip() == selected_artist and
                            str(metadata.get("title", "")).strip() == selected_title and
                            str(metadata.get("album", "")).strip() == selected_album and
                            str(metadata.get("albumartist", "")).strip() == selected_albumartist):
                            
                            # Found the matching file, now get its album art
                            audio = get_audio_file(file_path)
                            if audio:
                                if isinstance(audio, MP3) and audio.tags:
                                    apic_frames = audio.tags.getall("APIC")
                                    if apic_frames:
                                        image_data = apic_frames[0].data
                                        break
                                elif isinstance(audio, FLAC) and audio.pictures:
                                    image_data = audio.pictures[0].data
                                    break
                                elif isinstance(audio, MP4) and "covr" in audio:
                                    image_data = audio["covr"][0]
                                    break
        
        if image_data:
            # Convert bytes to PIL Image
            img = Image.open(io.BytesIO(image_data))
            
            # Get original format and size
            original_format = img.format
            original_size = len(image_data)
            
            # Convert to RGB if needed
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Save as BMP in memory
            output = io.BytesIO()
            img.save(output, 'BMP')
            data = output.getvalue()
            output.close()
            
            # Copy to clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_DIB, data[14:])  # Skip BMP header
            win32clipboard.CloseClipboard()
            
            log_message(f"[COVER] Album art copied to clipboard (Format: {original_format}, Size: {original_size:,} bytes)", log_type="processing")
        else:
            log_message("[COVER] No album art to copy", log_type="processing")
    except Exception as e:
        log_message(f"[COVER] Error copying album art to clipboard: {e}", log_type="processing")
        try:
            win32clipboard.CloseClipboard()
        except:
            pass

# Create the album art context menu
album_art_context_menu = tk.Menu(album_cover_label, tearoff=0, bg=Config.COLORS["SECONDARY_BACKGROUND"], fg=Config.COLORS["TEXT"])
album_art_context_menu.add_command(label="Copy Image", command=copy_album_art_to_clipboard)
album_art_context_menu.add_command(label="Paste Image", command=paste_image_from_clipboard)
album_art_context_menu.add_command(label="Remove Image", command=remove_album_art)

# Bind the context menu to the right mouse button on the album art label
album_cover_label.bind("<Button-3>", show_album_art_context_menu)

# TOP
# (Only Progress Bars & Status Remain)

top_frame = ttk.Frame(right_panel)
top_frame.pack(fill="x", padx=Config.PADDING["DEFAULT"], pady=Config.PADDING["DEFAULT"])

# Remove the original API frame and buttons frame from top_frame

# API Calls Counter & Status (Now centered)
status_frame = ttk.Frame(top_frame)
status_frame.pack(fill="x")

# API calls progress bar
api_bar_container = ttk.Frame(status_frame)
api_bar_container.pack(fill="x", pady=(Config.PADDING["SMALL"], 0))

api_progress_var = IntVar()

api_start_label = ttk.Label(api_bar_container, text="0", font=custom_font, width=1)
api_start_label.pack(side="left", padx=(0, 2))

api_progress_bar = ttk.Progressbar(api_bar_container, 
                                 variable=api_progress_var, 
                                 maximum=Config.API["RATE_LIMIT"],
                                 style="API.Horizontal.TProgressbar",
                                 length=Config.DIMENSIONS["PROGRESS_BAR_LENGTH"])
api_progress_bar.pack(side="left", fill="x", expand=True, padx=Config.PADDING["SMALL"])

api_end_label = ttk.Label(api_bar_container, text=str(Config.MAX_API_CALLS_PER_MINUTE), font=custom_font, width=7)
api_end_label.pack(side="left", padx=(2, 0))

# File processing progress bar
file_bar_container = ttk.Frame(status_frame)
file_bar_container.pack(fill="x", pady=Config.PADDING["SMALL"])

progress_var = IntVar()

file_start_label = ttk.Label(file_bar_container, text="0", font=custom_font, width=1)
file_start_label.pack(side="left", padx=(0, 2))

progress_bar = ttk.Progressbar(file_bar_container, 
                             variable=progress_var, 
                             maximum=100,
                             style="Gradient.Horizontal.TProgressbar",
                             length=Config.DIMENSIONS["PROGRESS_BAR_LENGTH"])
progress_bar.pack(side="left", fill="x", expand=True, padx=Config.PADDING["SMALL"])

file_count_var = StringVar(value="0/0")  # Changed to show selected/total format
file_end_label = ttk.Label(file_bar_container, textvariable=file_count_var, font=custom_font, width=7)  # Increased width
file_end_label.pack(side="left", padx=(2, 0))

# MIDDLE
# (Table for Files & Metadata)

middle_frame = ttk.Frame(right_panel)
middle_frame.pack(fill="both", expand=True, padx=Config.PADDING["DEFAULT"], pady=Config.PADDING["SMALL"])

# Create a filter frame above the table
filter_frame = ttk.Frame(middle_frame)
filter_frame.pack(fill="x", pady=(0, Config.PADDING["SMALL"]))

# Add filter label and entry
ttk.Label(filter_frame, text="FILTER:", font=custom_font).pack(side="left", padx=(0, Config.PADDING["SMALL"]))
filter_var = StringVar()
filter_entry = tk.Entry(filter_frame, 
                       textvariable=filter_var, 
                       width=40, 
                       font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                       bg=Config.COLORS["SECONDARY_BACKGROUND"],
                       fg=Config.COLORS["TEXT"],
                       insertbackground=Config.COLORS["TEXT"])
filter_entry.pack(side="left", fill="x", expand=True)

# Create a container for the table and its scrollbar
table_container = ttk.Frame(middle_frame)
table_container.pack(fill="both", expand=True)

columns = ("Artist", "Title", "Album", "Catalog Number", "Album Artist", "Year", "Track", "Genre")

# Create a frame with a border for the table
table_border_frame = ttk.Frame(table_container, relief="solid", borderwidth=1)  # Use ttk.Frame with system border style
table_border_frame.pack(fill="both", expand=True)

# Clear existing packing/layout first
for widget in table_border_frame.winfo_children():
    widget.pack_forget()

# Pack the table inside the border frame using grid
file_table = ttk.Treeview(table_border_frame, columns=columns, show="headings", height=Config.DIMENSIONS["TABLE_HEIGHT"])
file_table.grid(row=0, column=0, sticky="nsew")

# Enable drag and drop
file_table.drop_target_register(DND_FILES)
file_table.dnd_bind('<<Drop>>', lambda e: handle_drop(e.data))

def handle_drop(files):
    """Handle dropped files and add them to the file list."""
    global file_list, processed_files, updated_files, selected_folders, file_metadata_cache
    
    # Clear all data structures first
    file_list = []
    processed_files.clear()
    updated_files.clear()
    selected_folders.clear()
    file_metadata_cache.clear()
    
    # Clear the table
    file_table.delete(*file_table.get_children())
    
    # Handle Windows paths from drag and drop - might come in different formats
    dropped_paths = []
    if files:
        # Method 1: Standard format {path1} {path2} {path3}
        if files.startswith('{') and '}' in files:
            log_message(f"[DEBUG] Processing drag with standard format")
            # Split by } { and clean up each path
            paths = files.split('} {')
            for path in paths:
                # Remove outer braces and clean up the path
                clean_path = path.strip('{}')
                if clean_path:
                    dropped_paths.append(clean_path)
        # Method 2: Alternative format with multiple paths separated by spaces
        else:
            log_message(f"[DEBUG] Processing drag with alternative format")
            import re
            # Try to extract valid paths using a regex pattern for Windows paths
            # Look for drive letter patterns (C:/, D:/, etc.) or UNC paths (\\server\)
            potential_paths = re.findall(r'[A-Za-z]:[/\\][^\s]*|\\\\[^\s]*', files)
            for path in potential_paths:
                path = path.strip()
                if path and (os.path.exists(path) or os.path.exists(path.strip('"').strip("'"))):
                    # Clean up any quotes
                    clean_path = path.strip('"').strip("'")
                    dropped_paths.append(clean_path)
    
    log_message(f"[DEBUG] Found {len(dropped_paths)} paths to process")
    
    # Process each dropped path (could be file or folder)
    all_files = []
    for path in dropped_paths:
        try:
            if os.path.isdir(path):
                # It's a folder - add it to selected folders
                selected_folders.add(path)
                # Find all audio files in the folder and subfolders
                folder_files = []
                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(tuple(Config.SUPPORTED_AUDIO_EXTENSIONS)):
                            full_path = os.path.join(root, f)
                            folder_files.append(full_path)
                
                all_files.extend(folder_files)
                log_message(f"[INFO] Added folder: {path} ({len(folder_files)} files)")
            elif os.path.isfile(path) and path.lower().endswith(tuple(Config.SUPPORTED_AUDIO_EXTENSIONS)):
                # It's an individual audio file
                all_files.append(path)
        except Exception as e:
            log_message(f"[ERROR] Failed to process path: {path} - {str(e)}")
    
    # Update file list with all valid files
    if all_files:
        file_list.extend(all_files)
        update_table()
        log_message(f"[INFO] Added {len(all_files)} files via drag and drop")
    else:
        log_message("[WARNING] No valid audio files found in dropped items")
        # Reset the file count display
        file_count_var.set("0/0")

# Configure table borders and remove extra spacing
style.layout("Treeview", [
    ('Treeview.treearea', {'sticky': 'nswe'})
])

# Add scrollbar to table with autohide
table_scrollbar = ttk.Scrollbar(table_border_frame, orient="vertical", command=file_table.yview)
file_table.configure(yscrollcommand=lambda f, l: autohide_scrollbar(table_scrollbar, f, l))
# Initially hide the scrollbar
table_scrollbar.grid_remove()  # Hide initially instead of showing

# Configure grid weights and sizes
table_border_frame.columnconfigure(0, weight=1)
table_border_frame.columnconfigure(1, weight=0, minsize=0)  # Set minsize to 0 to prevent reserved space
table_border_frame.rowconfigure(0, weight=1)

style.layout("Treeview.Heading", [
    ("Treeview.Heading.cell", {'sticky': 'nswe'}),
    ("Treeview.Heading.padding", {'sticky': 'nswe', 'children': [
        ("Treeview.Heading.image", {'side': 'right', 'sticky': ''}),
        ("Treeview.Heading.text", {'sticky': 'we'})
    ]})
])

# Configure table colors and styles
style.configure("Treeview",
               background=Config.COLORS["SECONDARY_BACKGROUND"],
               foreground=Config.COLORS["TEXT"],
               fieldbackground=Config.COLORS["SECONDARY_BACKGROUND"])

style.configure("Treeview.Heading",
               background=Config.COLORS["BACKGROUND"],
               foreground=Config.COLORS["TEXT"],
               relief="flat")  # Remove header border

# Configure alternating row colors
file_table.tag_configure('oddrow', background=Config.COLORS["BACKGROUND"])
file_table.tag_configure('evenrow', background=Config.COLORS["SECONDARY_BACKGROUND"])

# Configure style for hidden rows
file_table.tag_configure('hidden', foreground='#FFFFFF', background='#FFFFFF')  # Make text invisible

# Define column widths
column_widths = {
    "Artist": 200,
    "Title": 200,
    "Album": 200,
    "Catalog Number": 130,
    "Album Artist": 200,
    "Year": 80,
    "Track": 80,
    "Genre": 200
}

for col in columns:
    file_table.heading(col, text=col,
                      command=lambda c=col: treeview_sort_column(file_table, c, sort_reverse))
    file_table.column(col, anchor="center", width=column_widths[col])

def apply_filter(*args):
    """Filter table contents based on filter text."""
    filter_text = filter_entry.get().lower()  # Convert filter text to lowercase
    
    # Clear the current table
    file_table.delete(*file_table.get_children())
    
    # Repopulate with filtered items in the same order as file_list
    for idx, file_path in enumerate(file_list):
        # Use cached metadata if available, otherwise read from file
        if file_path not in file_metadata_cache:
            audio = get_audio_file(file_path)
            if audio:
                file_metadata_cache[file_path] = {
                    "artist": get_tag_value(audio, "artist"),
                    "title": get_tag_value(audio, "title"),
                    "album": get_tag_value(audio, "album"),
                    "albumartist": get_tag_value(audio, "albumartist"),
                    "catalognumber": get_tag_value(audio, "catalognumber"),
                    "date": get_tag_value(audio, "date"),
                    "tracknumber": get_tag_value(audio, "tracknumber"),
                    "genre": get_tag_value(audio, "genre")
                }
        
        metadata = file_metadata_cache.get(file_path)
        if metadata:
            # Get all metadata values with safe access using .get()
            data = [
                metadata.get("artist", ""),
                metadata.get("title", ""),
                metadata.get("album", ""),
                metadata.get("catalognumber", ""),
                metadata.get("albumartist", ""),
                metadata.get("date", ""),
                metadata.get("tracknumber", ""),
                metadata.get("genre", "")
            ]
            
            # Check if any value matches the filter (case-insensitive)
            if not filter_text or any(filter_text in str(value).lower() for value in data):
                item = file_table.insert("", "end", values=data)
                
                # Apply alternating row colors
                if idx % 2 == 0:
                    file_table.item(item, tags=('evenrow',))
                else:
                    file_table.item(item, tags=('oddrow',))
                
                # Normalize the file path for comparison
                normalized_path = os.path.normpath(file_path)
                
                # Apply appropriate tags based on file status
                if normalized_path in updated_files:
                    file_table.tag_configure("updated", background=Config.COLORS["UPDATED_ROW"])
                    file_table.item(item, tags=("updated",))
                elif normalized_path in processed_files:
                    file_table.tag_configure("failed", background=Config.COLORS["FAILED_ROW"])
                    file_table.item(item, tags=("failed",))
        else:
            # Only show error items if they match the filter or if there's no filter
            if not filter_text or "error" in filter_text.lower():
                item = file_table.insert("", "end", values=["Error", "", "", "", "", "", "", ""])
                file_table.tag_configure("failed", background=Config.COLORS["FAILED_ROW"])
                file_table.item(item, tags=("failed",))
    
    # Update file count label
    selected_count = len(file_table.selection())
    total_count = len(file_table.get_children())  # Count actual visible items
    file_count_var.set(f"{selected_count}/{total_count}")
    
    # Auto-adjust column widths after filtering
    auto_adjust_column_widths()
    
    # Force UI update
    app.update_idletasks()

# Bind filter entry to any key event
filter_entry.bind('<Key>', apply_filter)
filter_entry.bind('<KeyRelease>', apply_filter)

# BOTTOM
# (Logs & Buttons)

bottom_frame = ttk.Frame(right_panel)
bottom_frame.pack(fill="x", padx=Config.PADDING["DEFAULT"], pady=Config.PADDING["DEFAULT"])

# Logs Section
log_frame = ttk.Frame(bottom_frame)
log_frame.pack(side="left", fill="both", expand=True, padx=Config.PADDING["SMALL"])

# Debug log container (only log section in bottom frame now)
debug_container = ttk.Frame(log_frame)
debug_container.pack(fill="both", expand=True)

debug_logbox = tk.Text(debug_container, 
                     height=Config.DIMENSIONS["DEBUG_LOG_HEIGHT"], 
                     width=Config.DIMENSIONS["DEBUG_LOG_WIDTH"],
                     state="disabled",
                     wrap="word",
                     font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                     bg=Config.COLORS["SECONDARY_BACKGROUND"],
                     fg=Config.COLORS["TEXT"],
                     insertbackground=Config.COLORS["TEXT"])
debug_logbox.pack(side="left", fill="both", expand=True)

# Add scrollbar to debug logbox with autohide - using default style
debug_scrollbar = ttk.Scrollbar(debug_container, orient="vertical", command=debug_logbox.yview)
debug_logbox.configure(yscrollcommand=lambda f, l: autohide_scrollbar(debug_scrollbar, f, l))
debug_scrollbar.pack(side="right", fill="y")

# Clear Buttons
button_frame = ttk.Frame(bottom_frame)
button_frame.pack(side="right", padx=Config.PADDING["SMALL"])

# Process button at the top of button frame
tk.Button(button_frame, text="PROCESS",
          command=lambda: start_processing(),
          font=custom_font,
          bg=Config.COLORS["SECONDARY_BACKGROUND"],
          fg=Config.COLORS["TEXT"],
          padx=Config.STYLES["WIDGET_PADDING"],
          pady=Config.STYLES["WIDGET_PADDING"]).pack(fill="x", pady=Config.PADDING["SMALL"])

# Ship button below process button
tk.Button(button_frame, text="EXPORT",
          command=lambda: organize_to_collection(),
          font=custom_font,
          bg=Config.COLORS["SECONDARY_BACKGROUND"],
          fg=Config.COLORS["TEXT"],
          padx=Config.STYLES["WIDGET_PADDING"],
          pady=Config.STYLES["WIDGET_PADDING"]).pack(fill="x", pady=Config.PADDING["SMALL"])

# Stop button below ship button
tk.Button(button_frame, text="STOP",
          command=lambda: stop_processing_files(),
          font=custom_font,
          bg=Config.COLORS["SECONDARY_BACKGROUND"],
          fg="#990000",
          padx=Config.STYLES["WIDGET_PADDING"],
          pady=Config.STYLES["WIDGET_PADDING"]).pack(fill="x", pady=Config.PADDING["SMALL"])

# Create a frame for checkboxes to be on the same line
checkbox_frame = ttk.Frame(button_frame)
checkbox_frame.pack(fill="x", pady=Config.PADDING["SMALL"])

# Metadata save options in horizontal layout
for checkbox_text, variable in [
    ("art", save_art_var),
    ("year", save_year_var),
    ("catalog", save_catalog_var),
]:
    tk.Checkbutton(checkbox_frame, text=checkbox_text,
                   variable=variable,
                   font=custom_font,
                   bg=Config.COLORS["BACKGROUND"],
                   fg=Config.COLORS["TEXT"],
                   selectcolor=Config.COLORS["SECONDARY_BACKGROUND"],
                   activebackground=Config.COLORS["BACKGROUND"],
                   activeforeground=Config.COLORS["TEXT"]).pack(side="left", padx=2)

# Add new control buttons
for button_text, command in [
    ("REFRESH", lambda: refresh_file_list()),
    ("REMOVE SELECTED", lambda: remove_selected_items()),
]:
    tk.Button(button_frame, text=button_text,
              command=command,
              font=custom_font,
              bg=Config.COLORS["SECONDARY_BACKGROUND"],
              fg=Config.COLORS["TEXT"],
              padx=Config.STYLES["WIDGET_PADDING"],
              pady=Config.STYLES["WIDGET_PADDING"]).pack(fill="x", pady=Config.PADDING["SMALL"])

# Add Delete key binding to the table
file_table.bind('<Delete>', lambda e: remove_selected_items())

def autohide_scrollbar(scrollbar, first, last):
    """Hide scrollbar if not needed, show if needed."""
    try:
        # Convert to float for comparison
        first = float(first)
        last = float(last)
        
        # Check if scrollbar still exists
        if not scrollbar.winfo_exists():
            return
            
        # Get current manager (grid or pack)
        manager = scrollbar.winfo_manager()
        
        # If no manager, initialize based on parent widget's existing children
        if not manager:
            parent_children = scrollbar.master.pack_slaves()
            if parent_children:  # If parent already has packed children
                scrollbar.pack(side="right", fill="y")
                manager = 'pack'
            else:  # Default to grid for table
                scrollbar.grid(row=0, column=1, sticky="ns")
                manager = 'grid'
        
        # Only hide if we're absolutely sure we don't need it
        if first <= 0.0 and last >= 1.0:
            # Hide the scrollbar
            if manager == 'grid':
                scrollbar.grid_remove()
            elif manager == 'pack':
                scrollbar.pack_forget()
        else:
            # Show the scrollbar with correct parameters
            if manager == 'grid' and not scrollbar.grid_info():
                scrollbar.grid(row=0, column=1, sticky="ns")
            elif manager == 'pack' and not scrollbar.pack_info():
                scrollbar.pack(side="right", fill="y")
        
        # Update scrollbar position without triggering another update
        scrollbar.set(first, last)
        
    except Exception as e:
        # Log error but don't raise it to prevent breaking the UI
        print(f"[ERROR] Scrollbar error: {str(e)}")  # Use print instead of log_message to avoid potential recursion

# ---------------- FUNCTIONS ---------------- #

def log_message(message, log_type="debug"):
    """Log messages in the appropriate text box based on type.
    
    Args:
        message: The message text to log
        log_type: Either "debug" (for technical messages) or "processing" (for operation results)
            - "debug" messages will appear in the debug_logbox (technical information)
            - "processing" messages will appear in the processing_listbox (success/failure results)
    """
    # Filter out debug messages and cover art logs
    if message.startswith("[DEBUG]") or message.startswith("[COVER]"):
        return  # Skip these messages entirely
    
    # Handle the case when UI elements aren't defined yet (early startup)
    global debug_logbox, processing_listbox
    if 'debug_logbox' not in globals() or debug_logbox is None:
        print(f"Early log: {message}")
        return
        
    # Use debug_logbox as fallback if processing_listbox isn't defined yet
    if log_type == "processing" and ('processing_listbox' not in globals() or processing_listbox is None):
        target_box = debug_logbox
    else:
        # Normal operation - send to appropriate box
        target_box = debug_logbox if log_type == "debug" else processing_listbox
        
    target_box.configure(state="normal")
    
    # Special handling for OK/NOK tags in processing messages
    if log_type == "processing" and message.startswith("[OK]") and target_box == processing_listbox:
        target_box.insert("end", "[OK] ", "ok")
        target_box.insert("end", message[4:] + "\n")
    elif log_type == "processing" and message.startswith("[NOK]") and target_box == processing_listbox:
        target_box.insert("end", "[NOK] ", "nok")
        target_box.insert("end", message[5:] + "\n")
    else:
        target_box.insert("end", message + "\n")
        
    target_box.configure(state="disabled")
    target_box.see("end")  # Auto-scroll to the latest message

def update_progress_bar(progress, bar_type="file"):
    """Update progress bar value and color based on type.
    
    Args:
        progress: For file progress, 0-100 percentage. For API, number of used calls.
        bar_type: Either "file" or "api"
    """
    if bar_type == "file":
        progress_var.set(progress)
        # Convert progress (0-100) to color (green->yellow->red)
        if progress < 50:
            # Green to Yellow (mix more yellow as progress increases)
            green = 255
            red = int((progress / 50) * 255)
        else:
            # Yellow to Red (reduce green as progress increases)
            red = 255
            green = int(((100 - progress) / 50) * 255)
        
        color = f'#{red:02x}{green:02x}00'
        style.configure("Gradient.Horizontal.TProgressbar", background=color)
    else:  # API progress bar
        api_progress_var.set(progress)
        # Calculate usage ratio for API bar
        usage_ratio = progress / Config.API["RATE_LIMIT"]
        if usage_ratio < Config.API["USAGE_THRESHOLDS"]["WARNING"]:  # Below 70% - Green
            color = Config.COLORS["PROGRESSBAR"]["GREEN"]
        elif usage_ratio < Config.API["USAGE_THRESHOLDS"]["CRITICAL"]:  # Between 70% and 90% - Orange
            color = Config.COLORS["PROGRESSBAR"]["ORANGE"]
        else:  # Above 90% - Red
            color = Config.COLORS["PROGRESSBAR"]["RED"]
        style.configure("API.Horizontal.TProgressbar", background=color)

def update_api_progress():
    """Update API progress bar based on rate limit headers"""
    global rate_limit_remaining
    
    # If no requests in 60 seconds, reset the window
    current_time = time.time()
    if current_time - first_request_time > Config.API_RATE_LIMIT_WAIT:
        rate_limit_remaining = rate_limit_total
    
    # Update progress bar to show used requests
    update_progress_bar(rate_limit_total - rate_limit_remaining, "api")

def enforce_api_limit():
    """Ensure we do not exceed API rate limit."""
    global rate_limit_remaining, first_request_time, rate_limit_total, rate_limit_used
    
    current_time = time.time()
    time_since_first_request = current_time - first_request_time if first_request_time > 0 else float('inf')
    
    # If this is the first request or we're in a new window
    if first_request_time == 0 or time_since_first_request >= Config.API_RATE_LIMIT_WAIT:
        first_request_time = current_time
        rate_limit_remaining = rate_limit_total
        rate_limit_used = 0
        log_message("[INFO] Starting new rate limit window.", log_type="debug")
    # If we're within the current window and out of requests
    elif rate_limit_remaining <= 0:
        wait_time = Config.API_RATE_LIMIT_WAIT - time_since_first_request
        log_message(f"[INFO] API limit reached. Waiting {wait_time:.1f} seconds...", log_type="debug")
        app.update()
        time.sleep(wait_time)
        # Reset counters after waiting
        first_request_time = time.time()
        rate_limit_remaining = rate_limit_total
        rate_limit_used = 0
        log_message(Config.MESSAGES["API_RESUMING"], log_type="debug")
    
    app.update()
    update_api_progress()

def save_api_key():
    """Save API Key to file and update visual state."""
    global DISCOGS_API_TOKEN
    new_api_key = api_key_var.get().strip()

    if not new_api_key:
        update_api_entry_style(False)
        log_message("[ERROR] API Key cannot be empty", log_type="processing")
        return

    # Test the API key with a simple request
    test_response = requests.get(
        Config.DISCOGS_SEARCH_URL,
        params={"token": new_api_key, "q": "test", "per_page": 1},
        timeout=10
    )

    if test_response.status_code != 200:
        update_api_entry_style(False)
        log_message("[ERROR] Invalid API Key - Authentication failed", log_type="processing")
        return

    DISCOGS_API_TOKEN = new_api_key
    with open(Config.API_KEY_FILE, "w") as f:
        f.write(DISCOGS_API_TOKEN)
    
    update_api_entry_style(True)
    log_message("[SUCCESS] API Key validated and saved", log_type="processing")

# Also validate the API key on startup
if DISCOGS_API_TOKEN:
    try:
        test_response = requests.get(
            Config.DISCOGS_SEARCH_URL,
            params={"token": DISCOGS_API_TOKEN, "q": "test", "per_page": 1},
            timeout=10
        )
        if test_response.status_code != 200:
            DISCOGS_API_TOKEN = ""
            update_api_entry_style(False)
            log_message("[ERROR] Saved API Key is invalid", log_type="processing")
    except:
        DISCOGS_API_TOKEN = ""
        update_api_entry_style(False)
        log_message("[ERROR] Could not validate saved API Key", log_type="processing")

def select_files():
    """Open file dialog to select MP3 files."""
    global file_list
    files = filedialog.askopenfilenames(filetypes=[(Config.FILE_TYPE_DESCRIPTION, "*" + Config.SUPPORTED_AUDIO_EXTENSIONS[0])])
    if files:
        file_list.extend(files)
        file_count_var.set(f"{len(file_list)}/{len(file_list)}")  # Update counter immediately
        update_table()

def select_folder():
    """Open a dialog to select a folder and add all MP3 files inside it, including subfolders."""
    global file_list, file_metadata_cache
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        # Clear existing data structures
        file_list = []
        file_metadata_cache.clear()
        processed_files.clear()
        updated_files.clear()
        
        # Clear the table
        file_table.delete(*file_table.get_children())
        
        # Add to selected folders set
        selected_folders.add(folder_selected)
        
        # Find all audio files in the folder and subfolders
        mp3_files = []
        for root, _, files in os.walk(folder_selected):
            for f in files:
                if f.lower().endswith(tuple(Config.SUPPORTED_AUDIO_EXTENSIONS)):
                    full_path = os.path.join(root, f)
                    mp3_files.append(full_path)
                    
        if mp3_files:
            file_list.extend(mp3_files)
            log_message(f"[INFO] Added {len(mp3_files)} files from folder: {folder_selected}")
            update_table()

def get_audio_file(file_path):
    """Helper function to safely get an audio file object with appropriate tag handling."""
    try:
        # Get the file extension
        ext = os.path.splitext(file_path)[1].lower()
        
        # Use appropriate handler based on file type
        if ext == '.mp3':
            # For MP3, first try to load with MP3
            audio = MP3(file_path)
            
            # Ensure ID3 tags exist
            if audio.tags is None:
                try:
                    audio.add_tags()
                except mutagen.MutagenError:
                    pass  # Tags already exist
            
            return audio
            
        elif ext == '.flac':
            return FLAC(file_path)
        elif ext in ['.m4a', '.mp4']:
            return MP4(file_path)
        elif ext == '.ogg':
            return OggVorbis(file_path)
        elif ext == '.wma':
            return ASF(file_path)
        elif ext == '.wav':
            return WAVE(file_path)
        else:
            # Fallback to basic File
            audio = mutagen.File(file_path)
        
        if not audio:
            log_message(f"[ERROR] Could not open file: {file_path}")
            return None
            
        return audio
        
    except Exception as e:
        log_message(f"[ERROR] Failed to read file {os.path.basename(file_path)}: {str(e)}")
        return None

def get_tag_value(audio, tag_name, default=""):
    """Helper function to get tag value across different audio formats."""
    try:
        # MP3
        if isinstance(audio, MP3):
            if not audio.tags:
                return default
                
            id3_mapping = {
                "artist": "TPE1",
                "title": "TIT2",
                "album": "TALB",
                "albumartist": "TPE2",
                "catalognumber": "TXXX:CATALOGNUMBER",
                "date": "TDRC",  # Year/Date
                "tracknumber": "TRCK",  # Track number
                "genre": "TCON"  # Genre
            }
            mapped_tag = id3_mapping.get(tag_name)
            
            if not mapped_tag:
                return default
                
            # Handle TXXX frames specially
            if mapped_tag.startswith("TXXX:"):
                desc = mapped_tag.split(":")[1]
                for tag in audio.tags.getall("TXXX"):
                    if tag.desc == desc:
                        return str(tag.text[0])
            # Handle regular ID3 frames
            elif mapped_tag in audio.tags:
                return str(audio.tags[mapped_tag].text[0])
                
            return default
            
        # FLAC
        elif isinstance(audio, FLAC):
            flac_mapping = {
                "date": "date",
                "tracknumber": "tracknumber",
                "genre": "genre"
            }
            mapped_tag = flac_mapping.get(tag_name, tag_name)
            return audio.get(mapped_tag, [default])[0]
        # MP4/M4A
        elif isinstance(audio, MP4):
            mp4_mapping = {
                "artist": "©ART",
                "title": "©nam",
                "album": "©alb",
                "albumartist": "aART",
                "catalognumber": "----:com.apple.iTunes:CATALOGNUMBER",
                "date": "©day",
                "tracknumber": "trkn",
                "genre": "©gen"
            }
            mapped_tag = mp4_mapping.get(tag_name)
            if mapped_tag and mapped_tag in audio:
                # Special handling for track number in MP4
                if mapped_tag == "trkn" and audio.get(mapped_tag):
                    return str(audio[mapped_tag][0][0])  # Track numbers are stored as tuples
                # Special handling for custom iTunes tags (bytes data)
                elif mapped_tag.startswith("----"):
                    # Custom iTunes tags are stored as bytes and need to be decoded
                    try:
                        byte_value = audio[mapped_tag][0]
                        if isinstance(byte_value, bytes):
                            return byte_value.decode('utf-8')
                        return str(byte_value)
                    except Exception as e:
                        log_message(f"[ERROR] Failed to decode MP4 custom tag {mapped_tag}: {e}")
                        return default
                return str(audio[mapped_tag][0])
        # OGG
        elif isinstance(audio, OggVorbis):
            ogg_mapping = {
                "date": "date",
                "tracknumber": "tracknumber",
                "genre": "genre"
            }
            mapped_tag = ogg_mapping.get(tag_name, tag_name)
            return audio.get(mapped_tag, [default])[0]
        # WMA
        elif isinstance(audio, ASF):
            asf_mapping = {
                "artist": "Author",
                "title": "Title",
                "album": "WM/AlbumTitle",
                "albumartist": "WM/AlbumArtist",
                "catalognumber": "WM/CatalogNo",
                "date": "WM/Year",
                "tracknumber": "WM/TrackNumber",
                "genre": "WM/Genre"
            }
            mapped_tag = asf_mapping.get(tag_name)
            if mapped_tag and mapped_tag in audio:
                return str(audio[mapped_tag][0])
            
        return default
    except Exception as e:
        log_message(f"[ERROR] Failed to get tag {tag_name}: {str(e)}")
        return default

def set_tag_value(audio, tag_name, value):
    """Helper function to set tag value across different audio formats."""
    try:
        # Import necessary modules
        from mutagen.flac import FLAC
        from mutagen.id3 import TPE1, TIT2, TALB, TPE2, TXXX, TDRC, TRCK, TCON
        
        # FLAC
        if isinstance(audio, FLAC):
            # FLAC tags need to be lists
            audio[tag_name] = [value]
        # MP4/M4A
        elif isinstance(audio, mutagen.mp4.MP4):
            mp4_mapping = {
                "artist": "©ART",
                "title": "©nam",
                "album": "©alb",
                "albumartist": "aART",
                "catalognumber": "----:com.apple.iTunes:CATALOGNUMBER",
                "date": "©day",  # Year/date
                "tracknumber": "trkn",  # Track number
                "genre": "©gen"  # Genre
            }
            mapped_tag = mp4_mapping.get(tag_name)
            if mapped_tag:
                if mapped_tag == "trkn":
                    # Special handling for track numbers in MP4
                    audio[mapped_tag] = [(int(value), 0)]  # Format as (track_number, total_tracks)
                elif mapped_tag.startswith("----"):
                    # Special handling for custom iTunes tags (like CATALOGNUMBER)
                    try:
                        # Custom iTunes tags need to be encoded as bytes with a special format
                        tag_parts = mapped_tag.split(":")
                        namespace = tag_parts[1]  # e.g., "com.apple.iTunes"
                        name = tag_parts[2]       # e.g., "CATALOGNUMBER"
                        
                        # Create a properly formatted custom tag
                        log_message(f"[DEBUG] Setting iTunes custom tag: {namespace}:{name}={value}")
                        audio[mapped_tag] = [value.encode("utf-8")]
                    except Exception as e:
                        log_message(f"[ERROR] Failed to set custom MP4 tag {mapped_tag}: {e}")
                else:
                    # Regular MP4 tags
                    audio[mapped_tag] = [value]
        # OGG
        elif isinstance(audio, mutagen.oggvorbis.OggVorbis):
            audio[tag_name] = [value]
        # WMA
        elif isinstance(audio, mutagen.asf.ASF):
            asf_mapping = {
                "artist": "Author",
                "title": "Title",
                "album": "WM/AlbumTitle",
                "albumartist": "WM/AlbumArtist",
                "catalognumber": "WM/CatalogNo",
                "date": "WM/Year",  # Year/date
                "tracknumber": "WM/TrackNumber",  # Track number
                "genre": "WM/Genre"  # Genre
            }
            mapped_tag = asf_mapping.get(tag_name)
            if mapped_tag:
                audio[mapped_tag] = [value]
        # MP3
        elif isinstance(audio, mutagen.mp3.MP3):
            id3_mapping = {
                "artist": (TPE1, lambda v: TPE1(encoding=3, text=[v])),
                "title": (TIT2, lambda v: TIT2(encoding=3, text=[v])),
                "album": (TALB, lambda v: TALB(encoding=3, text=[v])),
                "albumartist": (TPE2, lambda v: TPE2(encoding=3, text=[v])),
                "catalognumber": (TXXX, lambda v: TXXX(encoding=3, desc="CATALOGNUMBER", text=[v])),
                "date": (TDRC, lambda v: TDRC(encoding=3, text=[v])),
                "tracknumber": (TRCK, lambda v: TRCK(encoding=3, text=[v])),
                "genre": (TCON, lambda v: TCON(encoding=3, text=[v]))
            }
            if tag_name in id3_mapping:
                frame_class, frame_creator = id3_mapping[tag_name]
                if audio.tags is None:
                    audio.add_tags()
                
                # Remove existing frame before adding new one
                if tag_name == "catalognumber":
                    # Special handling for TXXX frames
                    for tag in list(audio.tags.getall("TXXX")):
                        if tag.desc == "CATALOGNUMBER":
                            audio.tags.delall("TXXX:" + tag.desc)
                else:
                    # Regular ID3 frames
                    frame_name = frame_class.__name__
                    if frame_name in audio.tags:
                        audio.tags.delall(frame_name)
                
                # Always add the frame with the new value (even if empty)
                audio.tags.add(frame_creator(value))
                
        audio.save()
        return True
    except Exception as e:
        log_message(f"[ERROR] Failed to set tag {tag_name}: {str(e)}")
        return False

def update_table():
    """Update the table with current file list and metadata."""
    # Clear the current table
    file_table.delete(*file_table.get_children())
    
    # Apply the current filter to show the correct items
    apply_filter()
    
    # Update file count label to show total files - use actual table items
    selected_count = len(file_table.selection())
    total_count = len(file_table.get_children())  # Count actual visible items
    file_count_var.set(f"{selected_count}/{total_count}")
    
    # Auto-adjust column widths after updating the table
    auto_adjust_column_widths()
    
    # Force UI update
    app.update_idletasks()

def auto_adjust_column_widths():
    """Calculate and set optimal column widths based on content."""
    # Get all items in the table
    items = file_table.get_children()
    if not items:
        return
        
    # Initialize column widths with header text lengths (using 6 pixels per character instead of 7)
    column_widths = {col: len(col) * 6 for col in columns}
    
    # Calculate maximum width for each column based on content
    for item in items:
        values = file_table.item(item)['values']
        for col_idx, value in enumerate(values):
            if value:  # Only consider non-empty values
                # Calculate width based on content length (using 6 pixels per character instead of 7)
                content_width = len(str(value)) * 6
                # Add smaller padding (5 pixels instead of 10)
                content_width += 5
                # Update maximum width if this content is wider
                column_widths[columns[col_idx]] = max(column_widths[columns[col_idx]], content_width)
    
    # Set minimum widths for specific columns
    column_widths["Track"] = max(60, min(80, column_widths["Track"]))        # Between 60-80 pixels
    column_widths["Year"] = max(60, min(80, column_widths["Year"]))          # Between 60-80 pixels
    column_widths["Catalog Number"] = max(80, min(120, column_widths["Catalog Number"]))  # Between 80-120 pixels
    
    # Set maximum widths for Artist and Title columns to prevent them from taking too much space
    column_widths["Artist"] = min(180, column_widths["Artist"])              # Maximum 180 pixels
    column_widths["Title"] = min(180, column_widths["Title"])                # Maximum 180 pixels
    column_widths["Album"] = min(180, column_widths["Album"])                # Maximum 180 pixels
    column_widths["Album Artist"] = min(180, column_widths["Album Artist"])  # Maximum 180 pixels
    column_widths["Genre"] = min(160, column_widths["Genre"])                # Maximum 160 pixels
    
    # Apply the calculated widths
    for col in columns:
        file_table.column(col, width=column_widths[col])

def clear_file_list():
    """Clear all file-related data structures and update the UI."""
    global file_list
    
    # Clear all data structures
    file_list = []
    processed_files.clear()
    updated_files.clear()
    selected_folders.clear()
    file_metadata_cache.clear()
    
    # Clear the table
    file_table.delete(*file_table.get_children())
    
    # Reset the file count
    file_count_var.set("0/0")
    
    # Force UI update
    app.update_idletasks()
    
    log_message("[INFO] File list and all related data cleared.")

def clear_logs():
    """Clear both log boxes and properly reset their scrollbars."""
    # Clear processing log box
    processing_listbox.configure(state="normal")
    processing_listbox.delete("1.0", "end")
    processing_listbox.configure(state="disabled")
    # Force scrollbar update for processing log
    processing_listbox.yview_moveto(0)
    autohide_scrollbar(processing_scrollbar, 0, 1)
    
    # Clear debug log box
    debug_logbox.configure(state="normal")
    debug_logbox.delete("1.0", "end")
    debug_logbox.configure(state="disabled")
    # Force scrollbar update for debug log
    debug_logbox.yview_moveto(0)
    autohide_scrollbar(debug_scrollbar, 0, 1)
    
    # Update the UI
    app.update_idletasks()

def fetch_metadata(artist, album, title=None):
    """Fetch the most common catalog number and essential metadata for an album.
    
    Args:
        artist: The artist name
        album: The album name
        title: The track title (optional) - used as fallback search
    """
    if not album:
        log_message("[WARNING] No album metadata found, skipping.")
        return None
        
    # Clean up artist and album names for better search results
    artist = artist.strip()
    album = album.strip()
    if title:
        title = title.strip()
    
    # Use consistent cache key that includes both artist and album
    cache_key = f"{artist.lower()}|{album.lower()}"
    
    # Thread-safe cache access
    with cache_lock:
        if cache_key in album_catalog_cache:
            log_message(f"[INFO] Using cached catalog number for '{artist} - {album}'.")
            return album_catalog_cache[cache_key]
        if cache_key in failed_search_cache:
            log_message(f"[INFO] Skipping known failed search for '{artist} - {album}'.")
            return None
    
    log_message(f"[API CALL] Requesting Discogs for: Artist='{artist}', Album='{album}'")
    
    # Try first with exact search but use q parameter instead of separate fields
    response_data = make_api_request(
        Config.DISCOGS_SEARCH_URL,
        {
            "q": f'"{artist}" "{album}"',  # Quote the terms for exact matching
            "token": DISCOGS_API_TOKEN,
            "type": "release"  # Ensure we're only getting releases
        }
    )
    
    if not response_data or not response_data.get("results"):
        # If no results, try a more lenient search
        log_message(f"[INFO] No exact matches found, trying broader search...")
        response_data = make_api_request(
            Config.DISCOGS_SEARCH_URL,
            {
                "q": f"{artist} {album}",  # Search all fields
                "token": DISCOGS_API_TOKEN,
                "type": "release"
            }
        )
        
        # If still no results and we have a title that's different from the album name
        if (not response_data or not response_data.get("results")) and title and title.lower() != album.lower():
            log_message(f"[INFO] No matches found with album name, trying with title: {title}")
            response_data = make_api_request(
                Config.DISCOGS_SEARCH_URL,
                {
                    "q": f"{artist} {title}",  # Search using title instead of album
                    "token": DISCOGS_API_TOKEN,
                    "type": "release"
                }
            )
    
    if not response_data or not response_data.get("results"):
        # Cache the failed search
        with cache_lock:
            failed_search_cache.add(cache_key)
            log_message(f"[INFO] Caching failed search for '{artist} - {album}'")
        return None
        
    releases = response_data.get("results", [])
    
    # Enhanced logging to show all matches for debugging
    total_results = response_data.get("pagination", {}).get("items", len(releases))
    per_page = response_data.get("pagination", {}).get("per_page", len(releases))
    current_page = response_data.get("pagination", {}).get("page", 1)
    log_message(f"[INFO] Discogs reports {total_results} total matches, showing page {current_page} with {per_page} per page")
    log_message(f"[INFO] Looking through {len(releases)} releases received from Discogs:")
    for idx, release in enumerate(releases[:10], 1):  # Show first 10 for debugging
        log_message(f"[INFO] Match {idx}: '{release.get('title', '')}' ({release.get('year', 'Unknown')}), Catalog: '{release.get('catno', 'Unknown')}'")
    
    # CRITICAL CHANGE: First filter for EXACT album matches, not just artist
    exact_album_matches = []
    for release in releases:
        release_title = release.get('title', '').lower()
        # Split on " - " to separate artist and album title
        if ' - ' in release_title:
            parts = release_title.split(' - ')
            release_artist = parts[0].strip()
            # The album title is everything after the first " - "
            release_album = ' - '.join(parts[1:]).strip()
            
            # Check if both artist and album match (using fuzzy matching to accommodate minor variations)
            artist_match = release_artist.lower() == artist.lower() or release_artist.lower() in artist.lower() or artist.lower() in release_artist.lower()
            album_match = release_album.lower() == album.lower() or release_album.lower() in album.lower() or album.lower() in release_album.lower()
            
            if artist_match and album_match:
                # Verify catalog number is preserved
                catno = release.get("catno", "").strip()
                if catno:
                    log_message(f"[DEBUG] Found exact album match with catalog {catno}: {release.get('title')}")
                else:
                    log_message(f"[DEBUG] Found exact album match WITHOUT catalog: {release.get('title')}")
                exact_album_matches.append(release)
        elif not ' - ' in release_title:
            # Some releases might not follow the "Artist - Title" format
            # Try fuzzy matching on the whole title
            title_match = album.lower() in release_title or release_title in album.lower()
            if title_match:
                log_message(f"[DEBUG] Found title-only match: {release.get('title')}")
                exact_album_matches.append(release)
    
    # If we have exact album matches, use ONLY those instead of all releases
    if exact_album_matches:
        log_message(f"[DEBUG] Using {len(exact_album_matches)} exact album matches instead of all {len(releases)} search results")
        # CRITICAL DEBUG: Verify catalog numbers are preserved in exact matches
        exact_catalogs = [r.get("catno", "") for r in exact_album_matches if r.get("catno", "").strip()]
        log_message(f"[DEBUG] Catalog numbers in exact matches: {exact_catalogs}")
        target_releases = exact_album_matches
    else:
        log_message(f"[WARNING] No exact album matches found. Results may be less accurate.")
        # Even with no exact matches, still try to find any artist matches at least
        exact_artist_matches = []
        for release in releases:
            release_title = release.get('title', '').lower()
            if ' - ' in release_title:
                release_artist = release_title.split(' - ')[0].strip()
                if release_artist.lower() == artist.lower() or release_artist.lower() in artist.lower() or artist.lower() in release_artist.lower():
                    exact_artist_matches.append(release)
                    log_message(f"[DEBUG] Found artist-only match: {release.get('title')}")
        
        target_releases = exact_artist_matches if exact_artist_matches else releases
        if exact_artist_matches:
            log_message(f"[DEBUG] Using {len(exact_artist_matches)} artist-only matches as fallback")
            # CRITICAL DEBUG: Verify catalog numbers are preserved in artist matches
            artist_catalogs = [r.get("catno", "") for r in exact_artist_matches if r.get("catno", "").strip()]
            log_message(f"[DEBUG] Catalog numbers in artist matches: {artist_catalogs}")
    
    # NEW STEP: Check for exact album title matches BEFORE filtering by catalog number
    log_message(f"[DEBUG] Looking for exact album title matches for: '{album}'")
    exact_album_title_matches = []
    
    for release in target_releases:
        release_title = release.get('title', '').lower()
        log_message(f"[DEBUG] Checking release: '{release_title}'")
        
        # Extract just the album part from "Artist - Album"
        if ' - ' in release_title:
            release_album = ' - '.join(release_title.split(' - ')[1:]).strip()
            log_message(f"[DEBUG] Extracted album part: '{release_album}'")
        else:
            release_album = release_title
            log_message(f"[DEBUG] No album part extraction possible, using whole title: '{release_album}'")
        
        # Check for exact album name match
        if release_album.lower() == album.lower():
            log_message(f"[INFO] Found exact album title match: '{release.get('title')}'")
            exact_album_title_matches.append(release)
    
    # If we found exact album title matches, use only those regardless of catalog number
    if exact_album_title_matches:
        log_message(f"[INFO] Using {len(exact_album_title_matches)} exact album title matches - these take priority over catalog number")
        filtered_releases = exact_album_title_matches
        
        # Check if ANY of these exact title matches have non-NONE catalogs
        non_none_exact_matches = [r for r in exact_album_title_matches if r.get("catno", "").strip().upper() != "NONE" and r.get("catno", "").strip() != ""]
        if non_none_exact_matches:
            log_message(f"[INFO] Found {len(non_none_exact_matches)} exact album title matches with valid catalog numbers")
            filtered_releases = non_none_exact_matches
        else:
            log_message(f"[INFO] No exact album title matches have valid catalog numbers, keeping all exact album title matches anyway")
    else:
        log_message(f"[WARNING] No exact album title matches found, proceeding with standard catalog number filtering")
        # Only now apply the NONE catalog filter if we didn't find exact album title matches
        non_none_releases = [r for r in target_releases if r.get("catno", "").strip().upper() != "NONE" and r.get("catno", "").strip() != ""]
        
        # Use non-none releases if available, otherwise fall back to all target releases
        if non_none_releases:
            log_message(f"[DEBUG] Found {len(non_none_releases)} releases with non-NONE catalog numbers")
            filtered_releases = non_none_releases
            # CRITICAL DEBUG: Verify catalog numbers are preserved after NONE filtering
            filtered_catalogs = [r.get("catno", "") for r in filtered_releases]
            log_message(f"[DEBUG] Catalog numbers after filtering: {filtered_catalogs}")
        else:
            log_message(f"[WARNING] All releases have NONE or empty catalog numbers, using all target releases")
            filtered_releases = target_releases
    
    # If no valid releases, return None
    if not filtered_releases:
        log_message(f"[WARNING] No valid releases found for '{album}'.")
        return None
    
    # Sort releases by year (oldest first) if year information is available
    releases_with_year = [r for r in filtered_releases if r.get("year") and str(r.get("year")).isdigit()]
    if releases_with_year:
        log_message(f"[DEBUG] --------------------------------------------------")
        log_message(f"[DEBUG] Sorting {len(releases_with_year)} releases by year (oldest first)")
        releases_with_year.sort(key=lambda r: int(r.get("year", 9999)))
        
        # Use the oldest release that has a valid catalog number AND matches artist
        oldest_release = None
        for release in releases_with_year:
            # Re-verify artist match to ensure it's not an unrelated old album
            release_title = release.get('title', '').lower()
            if ' - ' in release_title:
                release_artist = release_title.split(' - ')[0].strip()
                artist_match = (release_artist.lower() == artist.lower() or 
                               release_artist.lower() in artist.lower() or 
                               artist.lower() in release_artist.lower())
                
                if artist_match:
                    oldest_release = release
                    break
            else:
                # If no artist info in title, only use if we had exact album matches initially
                if exact_album_matches and release in exact_album_matches:
                    oldest_release = release
                    break
        
        # If no valid artist match found, fall back to the first release but log warning
        if not oldest_release and releases_with_year:
            oldest_release = releases_with_year[0]
            log_message(f"[WARNING] Oldest release may not match artist '{artist}'. Title: '{oldest_release.get('title')}'")
        
        # Continue with the rest of your existing logic for catalog number validation
        if oldest_release:
            catno = oldest_release.get("catno", "").strip()
            
            # CRITICAL FIX: Check if the catalog number exists and is not "NONE" before using it
            log_message(f"[DEBUG] Oldest release from year {oldest_release.get('year')} has catalog: '{catno}'")
            
            if catno and catno.upper() != "NONE":
                log_message(f"[INFO] Selected oldest release from year {oldest_release.get('year')} with catalog number: {catno}")
                selected_release = oldest_release
                normalized_catalog = catno.replace(" ", "").upper()
            else:
                # Even if the catalog is NONE, if this is an exact album title match, still use it
                if exact_album_title_matches and oldest_release in exact_album_title_matches:
                    log_message(f"[INFO] Selected oldest release with NONE catalog because it's an exact album title match")
                    selected_release = oldest_release
                    normalized_catalog = "NONE"  # Use a standardized placeholder
                else:
                    # Fall back to frequency-based selection if the oldest doesn't have a valid catalog
                    log_message(f"[INFO] Oldest release doesn't have a valid catalog number, falling back to frequency-based selection")
                    selected_release, normalized_catalog = select_by_frequency(filtered_releases)
    else:
        # If no year information, fall back to frequency-based selection
        log_message(f"[DEBUG] --------------------------------------------------")
        log_message(f"[DEBUG] No year information available, using frequency-based selection")
        selected_release, normalized_catalog = select_by_frequency(filtered_releases)
    
    if not selected_release or not normalized_catalog:
        log_message(f"[WARNING] Could not select a valid catalog number for '{album}'.")
        # Last resort: just use the first release with any catalog number
        for release in filtered_releases:
            catno = release.get("catno", "").strip()
            if catno and catno.upper() != "NONE":
                log_message(f"[INFO] Last resort: using first available catalog number: {catno}")
                selected_release = release
                normalized_catalog = catno.replace(" ", "").upper()
                break
        
        # If still no catalog, return None
        if not selected_release or not normalized_catalog:
            return None
    
    # Extract essential metadata
    metadata = {
        "catalog_number": normalized_catalog,  # Use normalized version
        "year": selected_release.get("year", ""),
        "album": selected_release.get("title", album),  # Use API's title if available, otherwise use original
        "cover_image": selected_release.get("cover_image", ""),
        "thumb": selected_release.get("thumb", "")
    }
    
    # Thread-safe cache update
    with cache_lock:
        album_catalog_cache[cache_key] = metadata
        
    log_message(f"[INFO] Found metadata for '{artist} - {album}': {metadata}")
    return metadata

def select_by_frequency(releases):
    """Helper function to select a release based on catalog number frequency."""
    # First pass: collect all catalog numbers
    log_message(f"[DEBUG] --- Processing all {len(releases)} releases to find catalog numbers ---")
    all_catalog_numbers = []
    
    # CRITICAL FIX: Debug raw catalog values before filtering
    raw_catalogs = [release.get("catno", "MISSING") for release in releases]
    log_message(f"[DEBUG] Raw catalog values: {raw_catalogs}")
    
    for release in releases:
        catno = release.get("catno", "").strip()
        if catno and catno.upper() != "NONE":  # Explicitly exclude NONE values
            all_catalog_numbers.append(catno.upper())
            log_message(f"[DEBUG] Found catalog number: {catno}")
    
    if not all_catalog_numbers:
        log_message(f"[WARNING] No valid catalog numbers found in the filtered releases.")
        
        # CRITICAL FIX: Pick first release with ANY catalog value, even if it's "NONE"
        for release in releases:
            catno = release.get("catno", "").strip()
            if catno:  # Any non-empty catalog, even "NONE"
                log_message(f"[DEBUG] Falling back to using catalog: {catno}")
                normalized_catalog = catno.replace(" ", "").upper()
                return release, normalized_catalog
                
        # If still nothing, just use the first release and assign a placeholder
        if releases:
            log_message(f"[DEBUG] Last resort: using first release with placeholder catalog")
            release = releases[0]
            return release, "UNKNOWN"
            
        return None, None
        
    # Count occurrences of all catalog numbers
    catalog_counts = Counter(all_catalog_numbers)
    log_message(f"[DEBUG] --- Analyzing frequency of {len(catalog_counts)} unique catalog numbers ---")
    
    # Get the top 2 most common catalog numbers
    most_common = catalog_counts.most_common(2)
    
    # Start with the most common
    most_common_catalog = most_common[0][0]
    normalized_catalog = most_common_catalog.replace(" ", "").upper()
    log_message(f"[DEBUG] Most common catalog: '{most_common_catalog}' (occurs {most_common[0][1]} times)")
    
    # If the second most common is not "NONE" and the first one is very similar to "NONE",
    # use the second one instead
    if len(most_common) > 1 and normalized_catalog == "NONE" or not normalized_catalog:
        second_common = most_common[1][0]
        second_normalized = second_common.replace(" ", "").upper()
        log_message(f"[DEBUG] First catalog is '{normalized_catalog}', trying second most common: '{second_normalized}' (occurs {most_common[1][1]} times)")
        normalized_catalog = second_normalized
    
    log_message(f"[DEBUG] Selected catalog number by frequency: {normalized_catalog}")
    
    # Find the release with this catalog number
    matching_release = next((release for release in releases 
                           if release.get("catno", "").strip().upper().replace(" ", "") == normalized_catalog), None)
    
    # If no matching release found (shouldn't happen but just in case)
    if not matching_release and releases:
        log_message(f"[DEBUG] No release found with catalog {normalized_catalog}, using first release")
        matching_release = releases[0]
    
    return matching_release, normalized_catalog

def update_file_metadata(file_path, metadata):
    """Update the MP3 file's metadata based on checkbox selections."""
    try:
        audio = get_audio_file(file_path)
        if not audio:
            return False

        updated = False  # Track if any updates were made
        normalized_path = os.path.normpath(file_path)

        # Update catalog number if selected
        if save_catalog_var.get() and metadata.get("catalog_number"):
            try:
                set_tag_value(audio, "catalognumber", metadata["catalog_number"])
                updated = True
                log_message(f"[SUCCESS] Updated catalog number for {os.path.basename(file_path)}")
            except Exception as e:
                log_message(f"[ERROR] Failed to update catalog number: {e}")

        # Update year if selected
        if save_year_var.get() and metadata.get("year"):
            try:
                set_tag_value(audio, "date", str(metadata["year"]))
                updated = True
                log_message(f"[SUCCESS] Updated year to {metadata['year']} for {os.path.basename(file_path)}")
            except Exception as e:
                log_message(f"[ERROR] Failed to update year: {e}")

        # Save changes if any were made
        if updated:
            audio.save()
            updated_files.add(normalized_path)  # Add normalized path to updated files

        # Update album art if selected
        if save_art_var.get() and (metadata.get("cover_image") or metadata.get("thumb")):
            try:
                cover_url = metadata.get("cover_image") or metadata.get("thumb")
                headers = {
                    'User-Agent': 'Phonodex/1.0',
                    'Referer': 'https://www.discogs.com/',
                    'Authorization': f'Discogs token={DISCOGS_API_TOKEN}'
                }
                
                response = requests.get(cover_url, headers=headers, timeout=10)
                if response.status_code == 200:
                    # Handle FLAC files
                    if isinstance(audio, FLAC):
                        # Clear existing pictures
                        audio.clear_pictures()
                        
                        # Create new picture
                        picture = Picture()
                        picture.type = 3  # Front cover
                        picture.mime = response.headers.get('content-type', 'image/jpeg')
                        picture.desc = 'Cover'
                        picture.data = response.content
                        
                        # Add picture to FLAC file
                        audio.add_picture(picture)
                        audio.save()
                        updated = True
                        log_message(f"[SUCCESS] Updated cover art for {os.path.basename(file_path)}")
                    
                    # Handle MP3 files
                    elif isinstance(audio, MP3):
                        log_message(f"[COVER] Updating cover art for MP3 file")
                        # Need to use regular mutagen.File for APIC frame
                        if audio.tags is None:
                            from mutagen.id3 import ID3
                            audio.add_tags()
                            log_message(f"[COVER] Added new ID3 tags to file")
                        
                        # Check for existing cover art
                        existing_apic = audio.tags.getall("APIC")
                        if existing_apic:
                            log_message(f"[COVER] Found {len(existing_apic)} existing APIC frames, removing them")
                            audio.tags.delall("APIC")
                        
                        # Add new cover art
                        try:
                            from mutagen.id3 import APIC
                            mime_type = response.headers.get('content-type', 'image/jpeg')
                            log_message(f"[COVER] Adding new cover art: {len(response.content)} bytes, mime: {mime_type}")
                            
                            # Always use type 3 (front cover) for new cover art
                            audio.tags.add(
                                APIC(
                                    encoding=3,
                                    mime=mime_type,
                                    type=3,  # Front cover
                                    desc='Front Cover',
                                    data=response.content
                                )
                            )
                            log_message(f"[COVER] Successfully added front cover APIC frame")
                            audio.save()
                            updated = True
                            log_message(f"[SUCCESS] Updated cover art for {os.path.basename(file_path)}")
                        except Exception as e:
                            log_message(f"[COVER] Error adding APIC frame: {e}")
                    # Handle MP4/M4A files
                    elif isinstance(audio, MP4):
                        log_message(f"[COVER] Updating cover art for MP4/M4A file")
                        
                        try:
                            # Get the image data and content type
                            image_data = response.content
                            mime_type = response.headers.get('content-type', 'image/jpeg')
                            log_message(f"[COVER] Adding cover art: {len(image_data)} bytes, mime: {mime_type}")
                            
                            # Set the cover art ('covr' atom in MP4)
                            # MP4 requires cover art to be in a specific format
                            from mutagen.mp4 import MP4Cover
                            
                            # Determine correct cover format based on mime type
                            if mime_type.endswith('png'):
                                cover_format = MP4Cover.FORMAT_PNG
                            else:
                                cover_format = MP4Cover.FORMAT_JPEG
                                
                            # Create MP4Cover object and set it
                            cover = MP4Cover(image_data, cover_format)
                            audio['covr'] = [cover]
                            
                            # Save the file
                            audio.save()
                            updated = True
                            log_message(f"[SUCCESS] Updated cover art for {os.path.basename(file_path)}")
                        except Exception as e:
                            log_message(f"[COVER] Error updating MP4 cover art: {e}")
                    else:
                        log_message(f"[COVER] Album art update not supported for this file type: {type(audio).__name__}")
                else:
                    log_message(f"[ERROR] Failed to download cover image (Status {response.status_code})")
            except Exception as e:
                log_message(f"[ERROR] Failed to update cover art: {str(e)}")

        if updated:
            updated_files.add(normalized_path)  # Add normalized path to updated files
            processed_files.add(normalized_path)  # Also mark as processed
        return updated

    except Exception as e:
        log_message(f"[ERROR] Failed to update metadata for {os.path.basename(file_path)}: {str(e)}")
        return False

def stop_processing_files():
    """Stop the file processing."""
    global stop_processing
    stop_processing = True
    log_message("[INFO] Stopping file processing...", log_type="processing")

def start_processing():
    global stop_processing
    log_message("[DEBUG] Start Processing Clicked!")
    
    if not DISCOGS_API_TOKEN:
        api_entry.configure(style='Invalid.TEntry')
        log_message("[ERROR] API Key is required", log_type="processing")
        return
        
    if not file_table.selection():
        log_message("[ERROR] No files selected for processing", log_type="processing")
        return
        
    stop_processing = False  # Reset stop flag
    log_message("[DEBUG] Creating background processing thread...")
    try:
        processing_thread = threading.Thread(target=process_files, daemon=True)
        processing_thread.start()
        log_message("[DEBUG] Background thread started!")
    except Exception as e:
        log_message(f"[ERROR] Failed to start thread: {e}")

def make_api_request(url, params):
    """Helper function to make API requests with proper error handling and rate limiting."""
    global rate_limit_total, rate_limit_used, rate_limit_remaining, first_request_time
    
    # Check rate limit before making request
    enforce_api_limit()
    
    try:
        response = requests.get(url, params=params, timeout=10)
        
        # Update rate limit info from headers
        rate_limit_total = int(response.headers.get('X-Discogs-Ratelimit', rate_limit_total))
        rate_limit_used = int(response.headers.get('X-Discogs-Ratelimit-Used', rate_limit_used + 1))
        rate_limit_remaining = int(response.headers.get('X-Discogs-Ratelimit-Remaining', rate_limit_total - rate_limit_used))
        
        # If this is the first request in a new window
        if first_request_time == 0:
            first_request_time = time.time()
        
        # Update progress bar
        update_api_progress()
        log_message(f"[INFO] API Calls: {rate_limit_used}/{rate_limit_total} (Remaining: {rate_limit_remaining})")
        
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 429:  # Too Many Requests
            log_message("[ERROR] Rate limit exceeded, waiting for reset...", log_type="debug")
            rate_limit_remaining = 0  # Force a wait
            enforce_api_limit()  # This will wait for the window to reset
            # Retry the request once after waiting
            return make_api_request(url, params)
        else:
            log_message(f"[ERROR] API response {response.status_code}: {response.text}")
            return None
            
    except requests.Timeout:
        log_message("[ERROR] API request timed out")
        return None
    except requests.RequestException as e:
        log_message(f"[ERROR] API request failed: {e}")
        return None

def process_files():
    log_message("[DEBUG] Entered process_files()...", log_type="debug")
    global processed_count, stop_processing, file_metadata_cache
    
    # Get selected items
    selected_items = file_table.selection()
    if not selected_items:
        log_message("[ERROR] No files selected for processing", log_type="processing")
        return
    
    # Update the file count display with selected/total format
    file_count_var.set(f"{len(selected_items)}/{len(file_list)}")
    
    # Create a cache of file metadata to avoid repeated file reads
    file_metadata_cache.clear()  # Clear existing cache before populating
    for file_path in file_list:
        audio = get_audio_file(file_path)
        if audio:
            file_metadata_cache[file_path] = {
                "artist": get_tag_value(audio, "artist"),
                "title": get_tag_value(audio, "title"),
                "album": get_tag_value(audio, "album"),
                "albumartist": get_tag_value(audio, "albumartist")
            }
    
    # Get the file paths for selected items using the cache
    selected_files = []
    for item in selected_items:
        values = file_table.item(item)['values']
        table_metadata = [values[0], values[1], values[2], values[4]]  # Artist, Title, Album, Album Artist
        
        # Find matching file using cached metadata
        for file_path, metadata in file_metadata_cache.items():
            current_metadata = [
                metadata.get("artist", ""),
                metadata.get("title", ""),
                metadata.get("album", ""),
                metadata.get("albumartist", "")
            ]
            # Use string comparison instead of exact comparison
            if all(str(a).strip() == str(b).strip() for a, b in zip(current_metadata, table_metadata)):
                selected_files.append(file_path)
                break
    
    # Thread-safe access to unprocessed files
    with processed_lock:
        unprocessed_files = []
        for file_path in selected_files:
            if os.path.normpath(file_path) not in processed_files:
                unprocessed_files.append(file_path)
                log_message(f"[DEBUG] Found unprocessed file: {file_path}")
    
    if not unprocessed_files:
        log_message("[SKIP] No unprocessed files found among selected items.", log_type="debug")
        update_progress_bar(0, "file")  # Reset progress bar
        return
        
    total_files = len(unprocessed_files)
    for idx, file_path in enumerate(unprocessed_files, 1):
        if stop_processing:
            log_message("[INFO] Processing stopped by user.", log_type="processing")
            update_progress_bar(0, "file")  # Reset progress bar
            return
            
        # Update progress bar
        progress = int((idx / total_files) * 100)
        update_progress_bar(progress, "file")
        app.update_idletasks()  # Update UI without blocking
        
        log_message(f"[INFO] Processing file: {file_path}", log_type="debug")
        
        # Use cached metadata instead of reading file again
        metadata = file_metadata_cache.get(file_path)
        if not metadata:
            log_message(f"[ERROR] Could not process file: {file_path}", log_type="processing")
            continue
            
        try:
            artist = metadata["artist"]
            title = metadata["title"]
            album = metadata["album"]
            log_message(f"[INFO] Extracted Metadata: Artist={artist}, Album={album}", log_type="debug")
            
            metadata = fetch_metadata(artist, album, title)
            
            if metadata:
                # Update all selected metadata in one go
                if update_file_metadata(file_path, metadata):
                    # Use log_message function for consistency
                    log_message(f"[OK] {artist} - {title} [{album}]", log_type="processing")
                else:
                    # Use log_message function for consistency
                    log_message(f"[NOK] {artist} - {title}", log_type="processing")
            
            # Thread-safe update of processed files
            with processed_lock:
                processed_files.add(os.path.normpath(file_path))
                processed_count += 1
                
        except Exception as e:
            log_message(f"[ERROR] Failed to process metadata for {os.path.basename(file_path)}: {e}", log_type="processing")
            continue
    
    # Update visual state using cached metadata
    for item in file_table.get_children():
        values = file_table.item(item)['values']
        table_metadata = [values[0], values[1], values[2], values[4]]
        
        # Find matching file using cached metadata
        for file_path, metadata in file_metadata_cache.items():
            current_metadata = [
                metadata.get("artist", ""),
                metadata.get("title", ""),
                metadata.get("album", ""),
                metadata.get("albumartist", "")
            ]
            if current_metadata == table_metadata:
                normalized_path = os.path.normpath(file_path)
                if normalized_path in updated_files:
                    file_table.tag_configure("updated", background=Config.COLORS["UPDATED_ROW"])
                    file_table.item(item, tags=("updated",))
                elif normalized_path in processed_files:
                    file_table.tag_configure("failed", background=Config.COLORS["FAILED_ROW"])
                    file_table.item(item, tags=("failed",))
                break
    
    log_message("[DEBUG] Finished processing selected files.", log_type="debug")

def start_editing(event):
    """Start editing a cell when double-clicked."""
    global editing_item, editing_column, editing_entry
    
    # Get the region that was clicked
    region = file_table.identify_region(event.x, event.y)
    if region != "cell":
        return
        
    # Get the item and column that was clicked
    editing_item = file_table.identify_row(event.y)
    editing_column = file_table.identify_column(event.x)
    
    if not editing_item or not editing_column:
        return
        
    # Get the column number
    column_num = int(editing_column[1]) - 1
    
    # Get the current value
    current_value = file_table.item(editing_item)['values'][column_num]
    
    # Get the cell's bounding box
    x, y, w, h = file_table.bbox(editing_item, editing_column)
    
    # Create and place the entry widget
    editing_entry = tk.Entry(table_container, bg='white', fg='black')  # Use tk.Entry instead of ttk.Entry for better color control
    editing_entry.insert(0, current_value)
    editing_entry.place(x=x, y=y-2, width=w, height=h+4)  # Slightly taller and shifted up
    editing_entry.select_range(0, tk.END)  # Select all text
    editing_entry.focus_set()
    editing_entry.tkraise()
    
    # Bind events
    editing_entry.bind('<Return>', finish_editing)
    editing_entry.bind('<Escape>', cancel_editing)
    
    # Bind click outside event to root window
    app.bind('<Button-1>', check_click_outside)

def check_click_outside(event):
    """Check if the click was outside the editing area."""
    global editing_entry
    
    if editing_entry:
        # Get the entry widget's geometry
        entry_x = editing_entry.winfo_rootx()
        entry_y = editing_entry.winfo_rooty()
        entry_w = editing_entry.winfo_width()
        entry_h = editing_entry.winfo_height()
        
        # Check if click was outside the entry widget
        if (event.x_root < entry_x or 
            event.x_root > entry_x + entry_w or 
            event.y_root < entry_y or 
            event.y_root > entry_y + entry_h):
            cancel_editing(None)

def cancel_editing(event):
    """Cancel the editing operation."""
    global editing_entry, editing_item, editing_column
    if editing_entry:
        # Unbind the click outside event
        app.unbind('<Button-1>')
        editing_entry.destroy()
        editing_entry = None
        editing_item = None
        editing_column = None

def finish_editing(event):
    """Finish editing and save the changes."""
    global editing_entry, editing_item, editing_column
    
    if not editing_entry or not editing_item or not editing_column:
        return
        
    # Get the new value and ensure it's a clean string
    try:
        new_value = editing_entry.get().strip()
    except tk.TclError:
        new_value = ""
    
    # Get the column number
    column_num = int(editing_column[1]) - 1
    
    # Get the current values BEFORE updating them
    try:
        current_values = list(file_table.item(editing_item)['values'])
        # Store original metadata for matching BEFORE updating the value
        original_metadata = [current_values[0], current_values[1], current_values[2], current_values[4]]  # Artist, Title, Album, Album Artist
        
        # Now update the value in the table
        current_values[column_num] = new_value
        file_table.item(editing_item, values=current_values)
    except Exception as e:
        log_message(f"[ERROR] Failed to update table: {e}")
        return
    
    # Find matching file using the ORIGINAL metadata
    matching_file = None
    for file_path, metadata in file_metadata_cache.items():
        current_metadata = [
            metadata.get("artist", ""),
            metadata.get("title", ""),
            metadata.get("album", ""),
            metadata.get("albumartist", "")
        ]
        # Use string comparison instead of exact comparison
        if all(str(a).strip() == str(b).strip() for a, b in zip(current_metadata, original_metadata)):
            matching_file = file_path
            break
    
    if matching_file:
        # Update the MP3 file
        update_mp3_metadata(matching_file, column_num, new_value)
        # Update the cache with the new value
        if matching_file in file_metadata_cache:
            if column_num == 0:  # Artist
                file_metadata_cache[matching_file]["artist"] = new_value
            elif column_num == 1:  # Title
                file_metadata_cache[matching_file]["title"] = new_value
            elif column_num == 2:  # Album
                file_metadata_cache[matching_file]["album"] = new_value
            elif column_num == 4:  # Album Artist
                file_metadata_cache[matching_file]["albumartist"] = new_value
    else:
        log_message("[ERROR] Could not find matching file to update metadata")
    
    # Clean up
    app.unbind('<Button-1>')  # Unbind the click outside event
    editing_entry.destroy()
    editing_entry = None
    editing_item = None
    editing_column = None
    
    # Auto-adjust column widths after editing
    auto_adjust_column_widths()

def update_mp3_metadata(file_path, column_num, new_value):
    """Update the MP3 file's metadata based on the edited column."""
    audio = get_audio_file(file_path)
    if not audio:
        return
        
    try:
        # Map column numbers to ID3 tags
        tag_mapping = {
            0: "artist",
            1: "title",
            2: "album",
            3: "catalognumber",
            4: "albumartist",
            5: "date",
            6: "tracknumber",
            7: "genre",
        }
        
        if column_num in tag_mapping:
            tag = tag_mapping[column_num]
            set_tag_value(audio, tag, new_value)
            updated_files.add(file_path)  # Mark file as updated
            log_message(f"[SUCCESS] Updated {os.path.basename(file_path)} {tag}: {new_value}")
    except Exception as e:
        log_message(f"[ERROR] Failed to update {tag} for {os.path.basename(file_path)}: {e}")

def remove_selected_items():
    """Remove selected items from the file list."""
    global file_list
    selected_items = file_table.selection()
    if not selected_items:
        return
    
    # Store the values of selected items before deleting them
    items_to_remove = []
    for item in selected_items:
        try:
            values = file_table.item(item, 'values')
            if values:
                items_to_remove.append(values)
        except Exception as e:
            log_message(f"[WARNING] Could not get values for item {item}: {e}")
            continue
    
    # Delete the items from the table
    file_table.delete(*selected_items)
    
    # Update the file count based on actual table items
    total_count = len(file_table.get_children())
    file_count_var.set(f"0/{total_count}")
    
    # Now clean up the backend data structures using the cache
    for values in items_to_remove:
        # Find and remove matching files using the cache
        for file_path in list(file_metadata_cache.keys()):  # Create a copy of keys to iterate
            metadata = file_metadata_cache[file_path]
            current_metadata = [
                metadata["artist"],
                metadata["title"],
                metadata["album"],
                metadata["albumartist"]
            ]
            if current_metadata == [values[0], values[1], values[2], values[4]]:
                if file_path in file_list:
                    file_list.remove(file_path)
                processed_files.discard(file_path)
                updated_files.discard(file_path)
                file_metadata_cache.pop(file_path, None)
    
    # Force UI update
    app.update_idletasks()
    
    log_message(f"[INFO] Removed {len(items_to_remove)} items from the list")

def refresh_file_list():
    """Refresh the file list by re-scanning selected folders and keeping individual files."""
    global file_list, processed_files, updated_files, file_metadata_cache
    
    # Keep track of individual files (not from folders)
    individual_files = [f for f in file_list if os.path.dirname(f) not in selected_folders]
    
    # Re-scan all selected folders
    folder_files = []
    for folder in selected_folders:
        if os.path.exists(folder):  # Check if folder still exists
            new_files = [os.path.join(root, f) 
                        for root, _, files in os.walk(folder) 
                        for f in files if f.lower().endswith(tuple(Config.SUPPORTED_AUDIO_EXTENSIONS))]
            folder_files.extend(new_files)
        else:
            log_message(f"[WARNING] Folder no longer exists: {folder}")
            selected_folders.remove(folder)
    
    # Create new file list while preserving order and removing duplicates
    seen = set()
    new_file_list = []
    
    # First add individual files while preserving their order
    for f in individual_files:
        if f not in seen:
            seen.add(f)
            new_file_list.append(f)
    
    # Then add folder files while preserving their order
    for f in folder_files:
        if f not in seen:
            seen.add(f)
            new_file_list.append(f)
    
    # Update the file list
    file_list = new_file_list
    
    # Clear processed and updated files sets
    processed_files.clear()
    updated_files.clear()
    file_metadata_cache.clear()  # Clear the metadata cache
    
    # Clear and update the table
    file_table.delete(*file_table.get_children())
    update_table()
    log_message(f"[INFO] Refreshed file list. Total files: {len(file_list)}")

def file_table_selection_callback(event):
    """Update the file count when selection changes."""
    selected_count = len(file_table.selection())
    total_count = len(file_table.get_children())  # Use actual table items count instead of file_list
    file_count_var.set(f"{selected_count}/{total_count}")

def select_all_visible(event):
    """Select all visible items in the table when CTRL+A is pressed."""
    # Only proceed if the event originated from the table or its children
    if not str(event.widget).startswith(str(file_table)):
        return
        
    # Get all visible items
    visible_items = file_table.get_children()
    if visible_items:
        # Select all visible items
        file_table.selection_set(visible_items)
        # Update the file count display
        file_table_selection_callback(event)

def update_basic_fields(event=None):
    """Update the basic fields based on table selection."""
    global current_album_art, file_metadata_cache, pending_album_art
    
    # Reset pending album art when switching files
    pending_album_art = None
    
    selected_items = file_table.selection()
    
    # If no items are selected, clear all fields and show default album art
    if not selected_items:
        for var in basic_field_vars.values():
            var.set("")
        load_default_album_art()
        return
    
    # Get values for all selected items
    values_by_field = {field: [] for field in basic_field_vars.keys()}
    
    # Get the first selected item for album art
    first_item = selected_items[0]
    first_values = file_table.item(first_item)['values']
    
    # Check if multiple albums are selected
    albums = set()
    artists = set()
    for item in selected_items:
        values = file_table.item(item)['values']
        if values[2]:  # Album
            albums.add(values[2])
        if values[0]:  # Artist
            artists.add(values[0])
    
    # Initialize variables for art comparison
    art_hashes = set()
    found_album_art = False
    first_art_data = None
    
    # Process each selected item for album art
    for item in selected_items:
        values = file_table.item(item)['values']
        selected_artist = str(values[0]).strip()
        selected_title = str(values[1]).strip()
        selected_album = str(values[2]).strip()
        selected_albumartist = str(values[4]).strip()
        
        # Find the matching file and get its art
        for file_path in file_list:
            if file_path in file_metadata_cache:
                metadata = file_metadata_cache[file_path]
                if (str(metadata.get("artist", "")).strip() == selected_artist and
                    str(metadata.get("title", "")).strip() == selected_title and
                    str(metadata.get("album", "")).strip() == selected_album and
                    str(metadata.get("albumartist", "")).strip() == selected_albumartist):
                    
                    # Found the matching file, get its album art
                    audio = get_audio_file(file_path)
                    if audio:
                        art_data = None
                        if isinstance(audio, MP3) and audio.tags:
                            apic_frames = audio.tags.getall("APIC")
                            if apic_frames:
                                art_data = apic_frames[0].data
                        elif isinstance(audio, FLAC) and audio.pictures:
                            art_data = audio.pictures[0].data
                        elif isinstance(audio, MP4) and "covr" in audio:
                            art_data = audio["covr"][0]
                        
                        if art_data:
                            # Hash the image data
                            art_hash = hashlib.md5(art_data).hexdigest()
                            art_hashes.add(art_hash)
                            
                            # Store the first art data we find
                            if not first_art_data:
                                first_art_data = art_data
                                found_album_art = True
                    break
    
    # Display album art based on hash comparison
    if found_album_art:
        if len(art_hashes) == 1:  # All selected files have the same art
            update_album_art_display(first_art_data)
        else:  # Different art found
            load_default_album_art()
            log_message("[COVER] Selected files have different album art", log_type="processing")
    else:
        load_default_album_art()
    
    # Process metadata fields
    for item in selected_items:
        values = file_table.item(item)['values']
        field_mapping = {
            "Artist": values[0],
            "Title": values[1],
            "Album": values[2],
            "Catalog Number": values[3],
            "Album Artist": values[4],
            "Year": values[5],
            "Track": values[6],
            "Genre": values[7]
        }
        
        # Add values to their respective lists
        for field, value in field_mapping.items():
            values_by_field[field].append(value)
    
    # Set values in all fields
    for field, var in basic_field_vars.items():
        # Remove any empty strings or None values
        values = [v for v in values_by_field[field] if v]
        
        if not values:
            var.set("")
        elif len(set(values)) == 1:
            # All values are the same
            var.set(values[0])
        else:
            # Different values
            var.set("<different values>")

def update_album_art_display(image_data):
    """Update the album art display with the provided image data."""
    global current_album_art
    try:
        log_message(f"[COVER] Processing image data: {len(image_data)} bytes")
        from PIL import Image, ImageTk
        import io
        
        # Open the image data
        img_buffer = io.BytesIO(image_data)
        img = Image.open(img_buffer)
        log_message(f"[COVER] Image opened successfully: {img.format}, {img.size}, {img.mode}")
        
        # Get the size of our container
        cover_size = Config.ALBUM_ART["COVER_SIZE"]
        
        # Instead of thumbnail which may leave empty space, we'll resize with padding
        # to ensure the image fills the entire space while maintaining aspect ratio
        
        # Calculate the scaling factor to fill the container
        width, height = img.size
        width_ratio = cover_size / width
        height_ratio = cover_size / height
        
        # Use the larger ratio to ensure the image fills the space
        ratio = max(width_ratio, height_ratio)
        
        # Calculate new dimensions
        new_width = round(width * ratio)
        new_height = round(height * ratio)
        
        # Resize the image (will be larger than container in one dimension)
        img = img.resize((new_width, new_height), Image.LANCZOS)
        
        # If image is larger than container, crop it to center
        if new_width > cover_size or new_height > cover_size:
            left = (new_width - cover_size) // 2
            top = (new_height - cover_size) // 2
            right = left + cover_size
            bottom = top + cover_size
            img = img.crop((left, top, right, bottom))
        
        log_message(f"[COVER] Image resized and cropped to fill {cover_size}x{cover_size}")
        
        # Create a PhotoImage object
        try:
            photo = ImageTk.PhotoImage(img)
            log_message(f"[COVER] PhotoImage created successfully")
            
            # Update the label
            album_cover_label.configure(image=photo)
            log_message(f"[COVER] Album cover label updated with new image")
            
            # Keep a reference to prevent garbage collection
            current_album_art = photo
            return True
        except Exception as e:
            log_message(f"[COVER] Failed to create or apply PhotoImage: {e}")
            return False
        
    except Exception as e:
        log_message(f"[COVER] Failed to update album art display: {e}")
        # Load default image if we can't display the provided image
        default_result = load_default_album_art()
        log_message(f"[COVER] Loaded default album art: {default_result}")
        return default_result

def apply_basic_fields():
    """Apply metadata from basic fields to selected files."""
    global pending_album_art
    
    selected_items = file_table.selection()
    if not selected_items:
        log_message("[ERROR] No files selected for updating", log_type="processing")
        return
    
    # Get values from basic fields
    new_metadata = {field: var.get() for field, var in basic_field_vars.items()}
    
    # Only skip fields with "<different values>" but allow empty strings
    new_metadata = {k: v for k, v in new_metadata.items() if v != "<different values>"}
    
    if not new_metadata and pending_album_art is None:
        log_message("[ERROR] No valid metadata to apply", log_type="processing")
        return
    
    # Map field names to tag names
    field_to_tag = {
        "Artist": "artist",
        "Title": "title",
        "Album": "album",
        "Album Artist": "albumartist",
        "Catalog Number": "catalognumber",
        "Year": "date",
        "Track": "tracknumber",
        "Genre": "genre"
    }
    
    # Process each selected item
    updated_count = 0
    for item in selected_items:
        values = file_table.item(item)['values']
        table_metadata = [values[0], values[1], values[2], values[4]]  # Artist, Title, Album, Album Artist
        
        # Find matching file using cached metadata
        matching_file = None
        for file_path, metadata in file_metadata_cache.items():
            current_metadata = [
                metadata.get("artist", ""),
                metadata.get("title", ""),
                metadata.get("album", ""),
                metadata.get("albumartist", "")
            ]
            if all(str(a).strip() == str(b).strip() for a, b in zip(current_metadata, table_metadata)):
                matching_file = file_path
                break
        
        if matching_file:
            try:
                audio = get_audio_file(matching_file)
                if not audio:
                    continue
                
                # Apply each metadata field
                updated = False
                for field, value in new_metadata.items():
                    tag = field_to_tag[field]
                    # Even if value is empty, it should be set (to clear existing value)
                    if set_tag_value(audio, tag, value):
                        updated = True
                
                # Handle album art if there's a pending change
                if pending_album_art is not None:
                    if pending_album_art == "REMOVE":
                        # Remove the album art
                        if isinstance(audio, mutagen.mp3.MP3):
                            if audio.tags:
                                audio.tags.delall("APIC")
                                updated = True
                                log_message(f"[SUCCESS] Removed album art from {os.path.basename(matching_file)}")
                        elif isinstance(audio, mutagen.flac.FLAC):
                            audio.clear_pictures()
                            updated = True
                            log_message(f"[SUCCESS] Removed album art from {os.path.basename(matching_file)}")
                        elif isinstance(audio, mutagen.mp4.MP4):
                            if "covr" in audio:
                                del audio["covr"]
                                updated = True
                                log_message(f"[SUCCESS] Removed album art from {os.path.basename(matching_file)}")
                    elif isinstance(pending_album_art, bytes):
                        # Add the new album art
                        try:
                            from PIL import Image
                            import io
                            
                            # Determine the image format
                            img_buffer = io.BytesIO(pending_album_art)
                            img = Image.open(img_buffer)
                            mime_type = f"image/{img.format.lower()}"
                            
                            # Apply based on file type
                            if isinstance(audio, mutagen.mp3.MP3):
                                # Remove existing album art
                                if audio.tags:
                                    audio.tags.delall("APIC")
                                else:
                                    audio.tags = mutagen.id3.ID3()
                                
                                # Add new album art
                                audio.tags.add(
                                    mutagen.id3.APIC(
                                        encoding=3,  # UTF-8
                                        mime=mime_type,
                                        type=3,  # Front cover
                                        desc='Cover',
                                        data=pending_album_art
                                    )
                                )
                                updated = True
                                log_message(f"[SUCCESS] Updated album art for {os.path.basename(matching_file)}")
                            elif isinstance(audio, mutagen.flac.FLAC):
                                # Clear existing pictures
                                audio.clear_pictures()
                                
                                # Create new picture
                                picture = mutagen.flac.Picture()
                                picture.type = 3  # Front cover
                                picture.mime = mime_type
                                picture.desc = 'Cover'
                                picture.data = pending_album_art
                                
                                # Add picture to FLAC file
                                audio.add_picture(picture)
                                updated = True
                                log_message(f"[SUCCESS] Updated album art for {os.path.basename(matching_file)}")
                            elif isinstance(audio, mutagen.mp4.MP4):
                                # MP4 requires special handling
                                from mutagen.mp4 import MP4Cover
                                
                                # Determine correct cover format based on mime type
                                if mime_type.endswith('png'):
                                    cover_format = MP4Cover.FORMAT_PNG
                                else:
                                    cover_format = MP4Cover.FORMAT_JPEG
                                    
                                # Create MP4Cover object and set it
                                cover = MP4Cover(pending_album_art, cover_format)
                                audio['covr'] = [cover]
                                updated = True
                                log_message(f"[SUCCESS] Updated album art for {os.path.basename(matching_file)}")
                            else:
                                log_message(f"[WARNING] Album art update not supported for this file type: {type(audio).__name__}")
                        except Exception as e:
                            log_message(f"[ERROR] Failed to update album art: {str(e)}")
                
                if updated:
                    # Save the file
                    audio.save()
                    
                    # Update cache
                    for field, value in new_metadata.items():
                        if field in ["Artist", "Title", "Album", "Album Artist"]:
                            file_metadata_cache[matching_file][field_to_tag[field]] = value
                    
                    # Update table display
                    current_values = list(values)
                    for field, value in new_metadata.items():
                        col_idx = list(columns).index(field)
                        current_values[col_idx] = value
                    file_table.item(item, values=current_values)
                    
                    # Mark as updated
                    normalized_path = os.path.normpath(matching_file)
                    updated_files.add(normalized_path)
                    file_table.tag_configure("updated", background=Config.COLORS["UPDATED_ROW"])
                    file_table.item(item, tags=("updated",))
                    updated_count += 1
                    
                    log_message(f"[SUCCESS] Updated metadata for {os.path.basename(matching_file)}")
            except Exception as e:
                log_message(f"[ERROR] Failed to update {os.path.basename(matching_file)}: {str(e)}")
    
    # Reset pending album art after applying
    pending_album_art = None
    
    if updated_count > 0:
        log_message(f"[INFO] Successfully updated {updated_count} files")
    else:
        log_message("[WARNING] No files were updated")

# Add save button below album art
save_metadata_button = tk.Button(left_panel, text="SAVE METADATA",
                               command=apply_basic_fields,
                               font=custom_font,
                               bg=Config.COLORS["SECONDARY_BACKGROUND"],
                               fg=Config.COLORS["TEXT"],
                               padx=Config.STYLES["WIDGET_PADDING"],
                               pady=Config.STYLES["WIDGET_PADDING"])
save_metadata_button.pack(fill="x", padx=10, pady=5)

# Add Clear buttons at the bottom of left panel - CREATE THESE FIRST
try:
    # Create a container for the Clear buttons - side by side
    buttons_container = ttk.Frame(left_panel)
    buttons_container.pack(side="bottom", fill="x", padx=10, pady=5)
    
    # Create two half-width buttons side by side
    tk.Button(buttons_container, text="CLEAR FILES",
              command=lambda: clear_file_list(),
              font=custom_font,
              bg=Config.COLORS["SECONDARY_BACKGROUND"],
              fg=Config.COLORS["TEXT"],
              padx=Config.STYLES["WIDGET_PADDING"],
              pady=Config.STYLES["WIDGET_PADDING"]).pack(side="left", fill="x", expand=True, padx=(0, 2))
    
    tk.Button(buttons_container, text="CLEAR LOGS",
              command=lambda: clear_logs(),
              font=custom_font,
              bg=Config.COLORS["SECONDARY_BACKGROUND"],
              fg=Config.COLORS["TEXT"],
              padx=Config.STYLES["WIDGET_PADDING"],
              pady=Config.STYLES["WIDGET_PADDING"]).pack(side="left", fill="x", expand=True, padx=(2, 0))
    
except Exception as e:
    log_message(f"[ERROR] Failed to add clear buttons: {str(e)}")

# Add processing log to left panel - AFTER the buttons are created
processing_container = ttk.Frame(left_panel)
processing_container.pack(fill="both", expand=True, padx=10, pady=5)

processing_listbox = tk.Text(processing_container, 
                           width=25,
                           state="disabled",
                           wrap="word",
                           font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                           bg=Config.COLORS["SECONDARY_BACKGROUND"],
                           fg=Config.COLORS["TEXT"],
                           insertbackground=Config.COLORS["TEXT"])
processing_listbox.pack(side="left", fill="both", expand=True)

# Configure tags for colored text
processing_listbox.tag_config("ok", foreground="#006400")  # Dark green
processing_listbox.tag_config("nok", foreground="#8B0000")  # Dark red

# Add scrollbar to processing listbox with autohide - using default style
processing_scrollbar = ttk.Scrollbar(processing_container, orient="vertical", command=processing_listbox.yview)
processing_listbox.configure(yscrollcommand=lambda f, l: autohide_scrollbar(processing_scrollbar, f, l))
processing_scrollbar.pack(side="right", fill="y")

def load_default_album_art():
    """Load the default album art image."""
    global current_album_art
    try:
        # Try to load the placeholder image from resources
        placeholder_path = resource_path(Config.ALBUM_ART["DEFAULT_IMAGE"])
        if os.path.exists(placeholder_path):
            from PIL import Image, ImageTk
            img = Image.open(placeholder_path)
            img = img.resize((240, 240), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            current_album_art = photo
            album_cover_label.configure(image=photo)
            album_cover_label.image = photo  # Keep a reference!
        else:
            log_message(f"[WARNING] Default album art not found at {placeholder_path}")
            album_cover_label.configure(image='')
    except Exception as e:
        log_message(f"[ERROR] Failed to load default album art: {str(e)}")
        album_cover_label.configure(image='')

def sanitize_filename(filename):
    """Sanitize filename by removing or replacing invalid characters."""
    # Replace characters that are invalid in filenames
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename

def show_folder_format_dialog():
    """Show a dialog to edit the folder structure format."""
    global folder_format
    
    format_dialog = tk.Toplevel(app)
    format_dialog.title("Configure Folder Structure")
    format_dialog.geometry("600x400")  # Increased height from 300 to 400
    format_dialog.configure(bg=Config.COLORS["BACKGROUND"])
    format_dialog.grab_set()  # Make the dialog modal
    
    # Add a label explaining the dialog
    tk.Label(format_dialog, 
            text="CHOOSE A FOLDER STRUCTURE FORMAT:",
            font=custom_font,
            bg=Config.COLORS["BACKGROUND"],
            fg=Config.COLORS["TEXT"]).pack(pady=(10, 5), padx=10)
    
    # Add explanation text
    explanation_text = """
Available placeholders:
%genre% - The genre tag
%year% - The release year
%catalognumber% - The catalog number
%albumartist% - The album artist
%album% - The album name
%artist% - The track artist
%title% - The track title

Default format:
D:\\Music\\Collection\\%genre%\\%year%\\[%catalognumber%] %albumartist% - %album%\\%artist% - %title%
    """
    
    # Add explanation text widget
    explanation = tk.Text(format_dialog, 
                       height=10, 
                       width=80,
                       font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                       bg=Config.COLORS["SECONDARY_BACKGROUND"],
                       fg=Config.COLORS["TEXT"])
    explanation.pack(padx=10, pady=5, fill="x")
    explanation.insert(tk.END, explanation_text)
    explanation.configure(state="disabled")
    
    # Add entry field with current format
    tk.Label(format_dialog, 
            text="FOLDER FORMAT:",
            font=custom_font,
            bg=Config.COLORS["BACKGROUND"],
            fg=Config.COLORS["TEXT"]).pack(pady=(5, 0), padx=10, anchor="w")
    
    format_var = StringVar(value=folder_format)
    format_entry = tk.Entry(format_dialog, 
                          textvariable=format_var,
                          width=80,
                          font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                          bg=Config.COLORS["SECONDARY_BACKGROUND"],
                          fg=Config.COLORS["TEXT"])
    format_entry.pack(padx=10, pady=5, fill="x")
    
    # Add reset to default button
    def reset_to_default():
        format_var.set(DEFAULT_FOLDER_FORMAT)
    
    reset_button = tk.Button(format_dialog, 
                           text="RESET TO DEFAULT",
                           command=reset_to_default,
                           font=custom_font,
                           bg=Config.COLORS["SECONDARY_BACKGROUND"],
                           fg=Config.COLORS["TEXT"])
    reset_button.pack(pady=5, padx=10, anchor="e")
    
    # Add button frame at the bottom
    button_frame = ttk.Frame(format_dialog)
    button_frame.pack(fill="x", pady=10, padx=10, side="bottom")
    
    # Function to save and continue
    def save_and_continue():
        global folder_format
        # Save the new format
        folder_format = format_var.get()
        # Save to persistent storage
        save_settings()
        # Close the dialog
        format_dialog.destroy()
        # Continue with organization
        organize_files_with_format()
    
    # Function to cancel
    def cancel_operation():
        format_dialog.destroy()
    
    # Cancel button (left)
    tk.Button(button_frame, text="CANCEL",
             command=cancel_operation,
             font=custom_font,
             bg=Config.COLORS["SECONDARY_BACKGROUND"],
             fg="#990000",
             padx=Config.STYLES["WIDGET_PADDING"],
             pady=Config.STYLES["WIDGET_PADDING"]).pack(side="left", fill="x", expand=True, padx=(0, 5))
    
    # Continue button (right)
    tk.Button(button_frame, text="CONTINUE",
             command=save_and_continue,
             font=custom_font,
             bg=Config.COLORS["SECONDARY_BACKGROUND"],
             fg="#006400",
             padx=Config.STYLES["WIDGET_PADDING"],
             pady=Config.STYLES["WIDGET_PADDING"]).pack(side="right", fill="x", expand=True, padx=(5, 0))
    
    # Wait for the dialog to close
    app.wait_window(format_dialog)
    return False  # Dialog was canceled

def organize_to_collection():
    """
    Entry point for organizing files to collection.
    Shows the folder format dialog first, then continues with organization.
    """
    # First check if any files are selected
    selected_items = file_table.selection()
    if not selected_items:
        log_message("[ERROR] No files selected for organizing", log_type="processing")
        messagebox.showinfo("No Files Selected", "Please select files to export first.")
        return
        
    # Show the folder format dialog
    show_folder_format_dialog()

def organize_files_with_format():
    """
    Organize selected files to collection folder using metadata.
    Uses the configured folder format.
    """
    # Get selected items
    selected_items = file_table.selection()
    if not selected_items:
        log_message("[ERROR] No files selected for organizing", log_type="processing")
        return
    
    # Create a list to store files that will be moved with their destinations
    files_to_move = []
    skipped_files = []
    
    # Check each selected file for required metadata
    for item in selected_items:
        values = file_table.item(item)['values']
        table_metadata = [values[0], values[1], values[2], values[4]]  # Artist, Title, Album, Album Artist
        
        # Find matching file using cached metadata
        matching_file = None
        for file_path, metadata in file_metadata_cache.items():
            current_metadata = [
                metadata.get("artist", ""),
                metadata.get("title", ""),
                metadata.get("album", ""),
                metadata.get("albumartist", "")
            ]
            if all(str(a).strip() == str(b).strip() for a, b in zip(current_metadata, table_metadata)):
                matching_file = file_path
                break
        
        if not matching_file:
            log_message(f"[ERROR] Could not find file for {values[0]} - {values[1]}")
            skipped_files.append(f"{values[0]} - {values[1]}")
            continue
        
        # Get required metadata
        audio = get_audio_file(matching_file)
        if not audio:
            skipped_files.append(os.path.basename(matching_file))
            continue
            
        # Get all required tags
        artist = get_tag_value(audio, "artist", "")
        title = get_tag_value(audio, "title", "")
        album = get_tag_value(audio, "album", "")
        albumartist = get_tag_value(audio, "albumartist", "")
        genre = get_tag_value(audio, "genre", "")
        year = get_tag_value(audio, "date", "")
        catalognumber = get_tag_value(audio, "catalognumber", "")
        
        # Check if all required fields are present
        required_fields = {"artist": artist, "title": title, "album": album, 
                          "albumartist": albumartist, "genre": genre, 
                          "year": year, "catalognumber": catalognumber}
        
        missing_fields = [field for field, value in required_fields.items() if not value.strip()]
        
        if missing_fields:
            log_message(f"[ERROR] Skipping {os.path.basename(matching_file)} - Missing required fields: {', '.join(missing_fields)}")
            skipped_files.append(f"{os.path.basename(matching_file)} (missing: {', '.join(missing_fields)})")
            continue
        
        # Handle genre with backslash separator - take only the part before first backslash
        if "\\" in genre:
            genre = genre.split("\\")[0].strip()
            log_message(f"[INFO] Using first genre component: {genre}")
            
        # Sanitize values for use in paths
        safe_genre = sanitize_filename(genre)
        safe_year = sanitize_filename(year)
        safe_catalognumber = sanitize_filename(catalognumber)
        safe_albumartist = sanitize_filename(albumartist)
        safe_album = sanitize_filename(album)
        safe_artist = sanitize_filename(artist)
        safe_title = sanitize_filename(title)
        
        # Get file extension
        _, ext = os.path.splitext(matching_file)
        
        # Build the destination path using the configured format
        # Replace placeholders with actual values
        destination_path = folder_format
        destination_path = destination_path.replace("%genre%", safe_genre)
        destination_path = destination_path.replace("%year%", safe_year)
        destination_path = destination_path.replace("%catalognumber%", safe_catalognumber)
        destination_path = destination_path.replace("%albumartist%", safe_albumartist)
        destination_path = destination_path.replace("%album%", safe_album)
        destination_path = destination_path.replace("%artist%", safe_artist)
        destination_path = destination_path.replace("%title%", safe_title)
        
        # Add file extension if not already in the format
        if not destination_path.endswith(ext):
            destination_path += ext
        
        # Check if the destination path exists and is different from source
        if matching_file != destination_path:
            files_to_move.append((matching_file, destination_path))
        else:
            log_message(f"[SKIP] File is already in correct location: {os.path.basename(matching_file)}")
    
    # If no valid files found, exit
    if not files_to_move:
        messagebox.showinfo("No Valid Files", f"No files to move. Either all files are missing required metadata or already in the correct location.\n\nSkipped files: {len(skipped_files)}")
        return
    
    # Create confirmation dialog
    confirmation_dialog = tk.Toplevel(app)
    confirmation_dialog.title("Confirm File Organization")
    confirmation_dialog.geometry("600x500")  # Increased height from 400 to 500
    confirmation_dialog.configure(bg=Config.COLORS["BACKGROUND"])
    confirmation_dialog.grab_set()  # Make the dialog modal
    
    # Add a label explaining what will happen
    tk.Label(confirmation_dialog, 
            text=f"The following {len(files_to_move)} files will be moved to the Collection folder structure:",
            font=custom_font,
            bg=Config.COLORS["BACKGROUND"],
            fg=Config.COLORS["TEXT"]).pack(pady=10, padx=10)
    
    # Create a scrollable list to display file moves
    file_frame = ttk.Frame(confirmation_dialog)
    file_frame.pack(fill="both", expand=True, padx=10, pady=5)
    
    # Add scrollbar
    scrollbar = ttk.Scrollbar(file_frame)
    scrollbar.pack(side="right", fill="y")
    
    # Text widget to show file paths
    file_list_text = tk.Text(file_frame,
                          font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                          bg=Config.COLORS["SECONDARY_BACKGROUND"],
                          fg=Config.COLORS["TEXT"],
                          yscrollcommand=scrollbar.set)
    file_list_text.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=file_list_text.yview)
    
    # Add all files to the text widget
    for src, dest in files_to_move:
        file_list_text.insert("end", f"From: {src}\nTo: {os.path.normpath(dest)}\n\n")
    file_list_text.configure(state="disabled")  # Make read-only
    
    # Show skipped files if any
    if skipped_files:
        tk.Label(confirmation_dialog, 
                text=f"The following {len(skipped_files)} files will be skipped (missing required metadata):",
                font=custom_font,
                bg=Config.COLORS["BACKGROUND"],
                fg="#990000").pack(pady=(10, 0), padx=10)
                
        # Create another scrollable list for skipped files
        skipped_frame = ttk.Frame(confirmation_dialog)
        skipped_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Add scrollbar
        skipped_scrollbar = ttk.Scrollbar(skipped_frame)
        skipped_scrollbar.pack(side="right", fill="y")
        
        # Text widget to show skipped files
        skipped_list_text = tk.Text(skipped_frame, height=5,
                                  font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                                  bg=Config.COLORS["SECONDARY_BACKGROUND"],
                                  fg="#990000",
                                  yscrollcommand=skipped_scrollbar.set)
        skipped_list_text.pack(side="left", fill="both", expand=True)
        skipped_scrollbar.config(command=skipped_list_text.yview)
        
        # Add all skipped files to the text widget
        for file in skipped_files:
            skipped_list_text.insert("end", f"{file}\n")
        skipped_list_text.configure(state="disabled")  # Make read-only
    
    # Button frame at the bottom
    button_frame = ttk.Frame(confirmation_dialog)
    button_frame.pack(fill="x", pady=10, padx=10)
    
    # Function to execute the move
    def execute_move():
        confirmation_dialog.destroy()
        
        moved_count = 0
        errors = 0
        moved_file_paths = []  # Track which files were successfully moved
        
        for src, dest in files_to_move:
            try:
                # Create destination directory if it doesn't exist
                dest_dir = os.path.dirname(dest)
                os.makedirs(dest_dir, exist_ok=True)
                
                # Move the file
                if os.path.exists(dest):
                    log_message(f"[WARNING] Destination file already exists, overwriting: {dest}")
                shutil.move(src, dest)
                moved_count += 1
                log_message(f"[SUCCESS] Moved file to: {dest}")
                moved_file_paths.append(src)  # Track the successfully moved file
            except Exception as e:
                errors += 1
                log_message(f"[ERROR] Failed to move {src}: {str(e)}")
        
        # Show summary
        if moved_count > 0:
            messagebox.showinfo("Organization Complete", f"Successfully moved {moved_count} files to Collection structure.\n{errors} errors occurred.")
            
            # Remove the moved files from the file_list and related data structures
            global file_list
            for path in moved_file_paths:
                if path in file_list:
                    file_list.remove(path)
                if path in file_metadata_cache:
                    file_metadata_cache.pop(path)
                processed_files.discard(path)
                updated_files.discard(path)
                
            # Remove the moved files from the table
            items_to_remove = []
            for item in selected_items:
                values = file_table.item(item)['values']
                table_metadata = [values[0], values[1], values[2], values[4]]  # Artist, Title, Album, Album Artist
                
                # Find the corresponding file path
                for file_path in moved_file_paths:
                    if file_path in file_metadata_cache:
                        continue  # Already removed from cache
                        
                    # Find if this item's metadata matches any moved file
                    for src, _ in files_to_move:
                        if os.path.normpath(src) == os.path.normpath(file_path):
                            items_to_remove.append(item)
                            break
            
            # Remove items from table
            if items_to_remove:
                file_table.delete(*items_to_remove)
                
            # Update file count
            file_count_var.set(f"{len(file_table.selection())}/{len(file_table.get_children())}")
            
            # Force UI update
            app.update_idletasks()
            
            # Remove any references to folders that no longer exist
            folder_list = list(selected_folders)
            for folder in folder_list:
                if not os.path.exists(folder):
                    selected_folders.remove(folder)
        else:
            messagebox.showerror("Organization Failed", "No files were moved. Check logs for details.")
    
    # Cancel button (left)
    tk.Button(button_frame, text="CANCEL",
             command=confirmation_dialog.destroy,
             font=custom_font,
             bg=Config.COLORS["SECONDARY_BACKGROUND"],
             fg="#990000",
             padx=Config.STYLES["WIDGET_PADDING"],
             pady=Config.STYLES["WIDGET_PADDING"]).pack(side="left", fill="x", expand=True, padx=(0, 5))
    
    # Confirm button (right)
    tk.Button(button_frame, text="CONFIRM",
             command=execute_move,
             font=custom_font,
             bg=Config.COLORS["SECONDARY_BACKGROUND"],
             fg="#006400",
             padx=Config.STYLES["WIDGET_PADDING"],
             pady=Config.STYLES["WIDGET_PADDING"]).pack(side="right", fill="x", expand=True, padx=(5, 0))

# Bind table events after all functions are defined
file_table.bind('<Double-1>', start_editing)
file_table.bind('<Escape>', cancel_editing)
file_table.bind('<Return>', finish_editing)
file_table.bind('<<TreeviewSelect>>', lambda e: (file_table_selection_callback(e), update_basic_fields(e)))
file_table.bind('<Control-a>', select_all_visible)  # Add CTRL+A binding
file_table.bind('<Control-A>', select_all_visible)  # Also bind capital A for caps lock cases

# Configure text widgets with dark theme
processing_listbox.configure(bg=Config.COLORS["SECONDARY_BACKGROUND"], fg=Config.COLORS["TEXT"])
debug_logbox.configure(bg=Config.COLORS["SECONDARY_BACKGROUND"], fg=Config.COLORS["TEXT"])

# Configure the root window background
app.configure(bg=Config.COLORS["BACKGROUND"])

# Apply style to main paned window
main_paned.configure(style='Dark.TPanedwindow')

# Initialize the album art display with the default image
load_default_album_art()

# Set the window to start maximized (windowed fullscreen)
app.state('zoomed')  # For Windows systems

app.mainloop()
