from tkinter import filedialog, ttk, StringVar, IntVar, font, BooleanVar, messagebox
import os
import shutil
import mutagen
import requests
import tkinter as tk
import subprocess
import platform
from tkinterdnd2 import DND_FILES, TkinterDnD
from collections import Counter
from mutagen.easyid3 import EasyID3
from mutagen.mp3 import MP3
from mutagen.flac import FLAC, Picture
from mutagen.mp4 import MP4
from mutagen.oggvorbis import OggVorbis
from mutagen.asf import ASF
from mutagen.wave import WAVE
from mutagen.id3 import ID3, APIC, TPE1, TIT2, TALB, TPE2, TXXX, TDRC, TRCK, TCON
import threading
from config import Config, load_settings, save_settings, folder_format, DEFAULT_FOLDER_FORMAT
from utils.logging import logger, log_message, autohide_scrollbar
from utils.file_operations import (resource_path, select_files as file_ops_select_files, 
                                 select_folder as file_ops_select_folder, handle_drop as file_ops_handle_drop, 
                                 get_audio_file, sanitize_filename,
                                 move_file_to_destination, copy_file_to_destination)
from utils.image_handling import (get_image_from_clipboard, copy_image_to_clipboard, 
                                resize_image, create_photo_image, 
                                load_default_album_art as image_load_default_album_art,
                                update_album_art_display as image_update_album_art_display,
                                paste_image_from_clipboard as image_paste_from_clipboard,
                                extract_album_art_from_file)
from utils.metadata import (
    get_tag_value, set_tag_value, select_by_frequency,
    fetch_metadata as metadata_fetch_metadata, update_album_metadata, update_tag_by_column,
    album_catalog_cache, cache_lock, update_mp3_metadata as metadata_update_mp3_metadata
)
from services.api_client import (
    make_api_request, update_api_entry_style, save_api_key,
    update_api_progress, enforce_api_limit, update_rate_limits_from_headers,
    rate_limit_total, rate_limit_used, rate_limit_remaining, first_request_time
)
import hashlib
from array import array
from ui.dialogs import show_folder_format_dialog, show_move_confirmation_dialog
from utils.table_operations import (
    auto_adjust_column_widths, 
    treeview_sort_column, 
    select_all_visible,
    file_table_selection_callback,
    update_table as table_ops_update_table,  # This matches the actual function name
    apply_filter as table_apply_filter,
    remove_selected_items as table_ops_remove_items  # Add this import
)
from ui.styles import (configure_styles, style_button, style_entry, style_label, style_checkbutton, configure_context_menu,
                      update_progress_bar_style, set_api_entry_style, configure_text_tags,
                      configure_table_columns, configure_table_tags, create_styled_button,
                      create_styled_entry, create_styled_text, create_button_pair)


# Cache and API rate limiter
# Removed duplicate cache variables: album_catalog_cache, failed_search_cache, cache_lock
# (now imported from utils.metadata)
processed_lock = threading.Lock()  # Lock for thread-safe processed files access
file_metadata_cache = {}  # Cache for file metadata
shared_album_art_files = set()  # Set of files that share the currently pending album art

# Track selected folders for refresh functionality
selected_folders = set()  # Store paths of selected folders

# Sorting variables
sort_column = None  # Track which column we're sorting by
sort_reverse = False  # Track sort direction

def sort_table(col):
    """Wrapper function to handle sorting state."""
    global sort_column, sort_reverse
    sort_reverse = treeview_sort_column(file_table, col, sort_reverse, columns)
    sort_column = col

# Load saved API Key
if os.path.exists(Config.API_KEY_FILE):
    with open(Config.API_KEY_FILE, "r") as f:
        DISCOGS_API_TOKEN = f.read().strip()
else:
    DISCOGS_API_TOKEN = ""

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
            # We can't update the UI style yet as it's not created
            log_message("[ERROR] Saved API Key is invalid", log_type="processing")
    except:
        DISCOGS_API_TOKEN = ""
        log_message("[ERROR] Could not validate saved API Key", log_type="processing")

# Function to update global DISCOGS_API_TOKEN
def update_global_token(token):
    global DISCOGS_API_TOKEN
    DISCOGS_API_TOKEN = token

# ---------------- GUI SETUP ---------------- #

app = TkinterDnD.Tk()
app.title(Config.WINDOW_TITLE)
app.geometry(Config.WINDOW_SIZE)
app.minsize(*Config.MIN_WINDOW_SIZE)

# Set ttk style to clam
style = ttk.Style()

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

# Global variables
current_song = None
current_album_art = None
current_album_art_bytes = None  # Store the raw bytes of the album art
pending_album_art = None

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
    ("FILES", lambda: file_ops_select_files(
        Config.FILE_TYPE_DESCRIPTION, 
        Config.SUPPORTED_AUDIO_EXTENSIONS, 
        file_list_var=file_list, 
        count_var=file_count_var, 
        update_table_func=update_table
    )),
    ("FOLDER", lambda: file_ops_select_folder(
        update_table_func=update_table,
        file_list_var=file_list,
        metadata_cache=file_metadata_cache,
        processed_files=processed_files,
        updated_files=updated_files,
        selected_folders_var=selected_folders,
        supported_extensions=Config.SUPPORTED_AUDIO_EXTENSIONS,
        count_var=file_count_var
    )),
    ("LEAVE", app.quit)
]:
    button = create_styled_button(
        buttons_subframe, 
        button_text, 
        command, 
        is_danger=(button_text == "LEAVE")
    )
    button.pack(side="left", padx=Config.PADDING["SMALL"], expand=True, fill="x")

# API Key Entry with Save Button - directly in left panel
api_subframe = ttk.Frame(left_panel)
api_subframe.pack(fill="x", pady=(5, 5), padx=10)

api_entry = create_styled_entry(
    api_subframe,
    textvariable=api_key_var,
    width=Config.DIMENSIONS["API_ENTRY_WIDTH"],
    justify="center"
)
api_entry.pack(side="left", fill="x", expand=True)

# Update the validation style - using the imported function as a wrapper
def update_api_entry_style(is_valid):
    """Update the API entry styling based on validity."""
    set_api_entry_style(api_entry, is_valid)

# Initial style based on API token
update_api_entry_style(bool(DISCOGS_API_TOKEN))

save_button = tk.Button(api_subframe, text="💾", 
                       width=Config.DIMENSIONS["SAVE_BUTTON_WIDTH"],
                       command=lambda: save_api_key(api_key_var, api_entry, update_global_token))
style_button(save_button)
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
    label = tk.Label(field_frame, text=field.upper() + ":")
    style_label(label, use_smaller_font=True)
    label.pack(fill="x")
    
    # Keep using the original field name to access the variable
    entry = create_styled_entry(field_frame, textvariable=basic_field_vars[field])
    entry.pack(fill="x", pady=(2, 0))

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
    """Display the context menu when right-clicking on album art.
    
    Note: This function must remain in main.py as it:
    - Relies on the global album_art_context_menu variable
    - Is part of the album art management system
    - Uses functions that manipulate the application state (paste, copy, remove)
    - Is directly bound to the album cover label UI element
    
    Args:
        event: The mouse event that triggered the context menu
    """
    album_art_context_menu.tk_popup(event.x_root, event.y_root)

# Function to paste image from clipboard
def paste_image_from_clipboard():
    """Paste image from clipboard to album art display."""
    global pending_album_art, shared_album_art_files, current_album_art, current_album_art_bytes
    
    # Use the function from utils.image_handling
    image_data = image_paste_from_clipboard()
    
    if image_data:
        # Store the image data for later saving
        pending_album_art = image_data
        
        # Record which files share this album art (all selected files)
        selected_items = file_table.selection()
        shared_album_art_files.clear()  # Clear previous sharing
        
        for item in selected_items:
            values = file_table.item(item)['values']
            file_path = values[-1]  # Last column is file path
            if file_path:
                # Normalize the path to ensure consistent comparison
                normalized_path = os.path.normpath(file_path)
                shared_album_art_files.add(normalized_path)
                log_message(f"[COVER] Marked file as sharing pending album art: {os.path.basename(file_path)}", log_type="debug")
        
        # Display the image
        current_album_art_bytes = image_data
        photo = image_update_album_art_display(
            image_data,
            label=album_cover_label,
            size=Config.ALBUM_ART["COVER_SIZE"],
            load_default_func=load_default_album_art
        )
        current_album_art = photo
        log_message("[COVER] Image pasted from clipboard (not saved until 'SAVE METADATA' is clicked)", log_type="processing")
        log_message(f"[COVER] Marked {len(shared_album_art_files)} files as sharing the same album art", log_type="debug")

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
    global current_album_art, current_album_art_bytes, pending_album_art
    
    try:
        image_data = None
        
        if pending_album_art and pending_album_art != "REMOVE":
            # If there's pending art, use that
            log_message("[COVER] Using pending album art for clipboard", log_type="processing")
            image_data = pending_album_art
        elif current_album_art_bytes:
            # If there's current album art displayed, use the raw bytes data
            log_message("[COVER] Using current album art for clipboard", log_type="processing")
            image_data = current_album_art_bytes
        else:
            # Get all selected items
            selected_items = file_table.selection()
            if not selected_items:
                log_message("[COVER] No files selected", log_type="processing")
                return
                
            # If multiple items are selected, verify they all have the same album art
            if len(selected_items) > 1:
                art_hashes = set()
                for item in selected_items:
                    # Get the file path from the values array
                    values = file_table.item(item)['values']
                    
                    # Check if the values array has enough elements
                    if len(values) < 9:
                        log_message(f"[ERROR] Invalid table values for copy: {values}", log_type="debug")
                        continue
                        
                    # Get the file path from the values array (last element)
                    file_path = values[8]  # File path is now in position 8 (9th element, 0-indexed)
                    
                    if not file_path:
                        log_message("[ERROR] Missing file path in table item", log_type="debug")
                        continue
                    
                    # Only process audio files
                    ext = os.path.splitext(file_path)[1].lower()
                    if ext in ['.mp3', '.flac', '.m4a', '.mp4', '.ogg', '.wma', '.wav']:
                                audio = get_audio_file(file_path)
                                if audio:
                                    art_data = extract_album_art_from_file(file_path, audio)
                    if art_data:
                        art_hash = hashlib.md5(art_data).hexdigest()
                        art_hashes.add(art_hash)
                                        
                if len(art_hashes) > 1:
                    log_message("[COVER] Selected files have different album art", log_type="processing")
                    return
                elif len(art_hashes) == 0:
                    log_message("[COVER] No album art found in selected files", log_type="processing")
                    return
            
            # Get the first selected item
            # Get the file path from the values array
            values = file_table.item(selected_items[0])['values']
            
            # Check if the values array has enough elements
            if len(values) < 9:
                log_message(f"[ERROR] Invalid table values for copy: {values}", log_type="debug")
                return
                
            # Get the file path from the values array (last element)
            file_path = values[8]  # File path is now in position 8 (9th element, 0-indexed)
            
            if not file_path:
                log_message("[ERROR] Missing file path in table item", log_type="debug")
                return
                
            # Check if the file is a supported audio format
            ext = os.path.splitext(file_path)[1].lower()
            if ext not in ['.mp3', '.flac', '.m4a', '.mp4', '.ogg', '.wma', '.wav']:
                log_message(f"[COVER] File type not supported for album art: {ext}", log_type="processing")
                return
            
            audio = get_audio_file(file_path)
            if audio:
                image_data = extract_album_art_from_file(file_path, audio)
        
        if not image_data:
            log_message("[COVER] No album art to copy", log_type="processing")
            return
        
        # Use the function from utils.image_handling
        if copy_image_to_clipboard(image_data):
            log_message("[COVER] Album art copied to clipboard", log_type="processing")
        else:
            log_message("[COVER] Failed to copy album art to clipboard", log_type="processing")
    
    except Exception as e:
        log_message(f"[ERROR] Error in copy_album_art_to_clipboard: {str(e)}", log_type="error")

# Create the album art context menu
album_art_context_menu = tk.Menu(album_cover_label)
configure_context_menu(album_art_context_menu)
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
filter_entry = create_styled_entry(filter_frame, textvariable=filter_var, width=40)
filter_entry.pack(side="left", fill="x", expand=True)

# Create a container for the table and its scrollbar
table_container = ttk.Frame(middle_frame)
table_container.pack(fill="both", expand=True)

columns = ("Artist", "Title", "Album", "Catalog Number", "Album Artist", "Year", "Track", "Genre", "File Path")

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
file_table.dnd_bind('<<Drop>>', lambda e: file_ops_handle_drop(
    e.data,
    file_list_var=file_list,
    processed_files=processed_files,
    updated_files=updated_files,
    selected_folders_var=selected_folders,
    metadata_cache=file_metadata_cache,
    table=file_table,
    supported_extensions=Config.SUPPORTED_AUDIO_EXTENSIONS,
    count_var=file_count_var,
    update_table_func=update_table
))

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

# Configure table tags
configure_table_tags(file_table)

# Define column widths and configure
column_widths = {
    "Artist": 200,
    "Title": 200,
    "Album": 200,
    "Catalog Number": 130,
    "Album Artist": 200,
    "Year": 80,
    "Track": 80,
    "Genre": 200,
    "File Path": 0  # Set width to 0 to hide it
}

for col in columns:
    file_table.heading(col, text=col,
                      command=lambda c=col: sort_table(c))

# Use the utility function to configure columns
configure_table_columns(file_table, columns, column_widths, hide_columns=["File Path"])

def apply_filter(*args):
    """Filter table contents based on filter text."""
    filter_text = filter_entry.get().lower()  # Convert filter text to lowercase
    table_apply_filter(
        file_table, 
        filter_text, 
        file_list, 
        file_metadata_cache, 
        get_audio_file, 
        get_tag_value, 
        updated_files, 
        processed_files, 
        file_count_var, 
        columns
    )
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

debug_logbox = create_styled_text(
    debug_container,
    height=Config.DIMENSIONS["DEBUG_LOG_HEIGHT"],
    width=Config.DIMENSIONS["DEBUG_LOG_WIDTH"],
    state="disabled",
    wrap="word"
)
debug_logbox.pack(side="left", fill="both", expand=True)

# Add scrollbar to debug logbox with autohide - using default style
debug_scrollbar = ttk.Scrollbar(debug_container, orient="vertical", command=debug_logbox.yview)
debug_logbox.configure(yscrollcommand=lambda f, l: autohide_scrollbar(debug_scrollbar, f, l))
debug_scrollbar.pack(side="right", fill="y")

# Configure the logger with the text widgets
logger.set_debug_widget(debug_logbox)

# Clear Buttons
button_frame = ttk.Frame(bottom_frame)
button_frame.pack(side="right", padx=Config.PADDING["SMALL"])

# Process button at the top of button frame
process_button = create_styled_button(button_frame, "PROCESS", lambda: start_processing())
process_button.pack(fill="x", pady=Config.PADDING["SMALL"])

# Ship button below process button
process_button = create_styled_button(button_frame, "EXPORT", lambda: organize_to_collection())
process_button.pack(fill="x", pady=Config.PADDING["SMALL"])

# Stop button below ship button
process_button = create_styled_button(button_frame, "STOP", lambda: stop_processing_files(), is_danger=True)
process_button.pack(fill="x", pady=Config.PADDING["SMALL"])

# Create a frame for checkboxes to be on the same line
checkbox_frame = ttk.Frame(button_frame)
checkbox_frame.pack(fill="x", pady=Config.PADDING["SMALL"])

# Metadata save options in horizontal layout
for checkbox_text, variable in [
    ("art", save_art_var),
    ("year", save_year_var),
    ("catalog", save_catalog_var),
]:
    checkbox = tk.Checkbutton(checkbox_frame, text=checkbox_text, variable=variable)
    style_checkbutton(checkbox)
    checkbox.pack(side="left", padx=2)

# Add new control buttons
for button_text, command in [
    ("REFRESH", lambda: refresh_file_list()),
    ("REMOVE SELECTED", lambda: remove_selected_items()),
]:
    process_button = create_styled_button(button_frame, button_text, command)
    process_button.pack(fill="x", pady=Config.PADDING["SMALL"])
    

# Add Delete key binding to the table
file_table.bind('<Delete>', lambda e: remove_selected_items())

# Create the file table context menu
file_table_context_menu = tk.Menu(file_table)
configure_context_menu(file_table_context_menu)
file_table_context_menu.add_command(label="Play Selected", command=lambda: play_selected_files())
# Add the Show in Explorer option
file_table_context_menu.add_command(label="Show in Explorer", command=lambda: show_in_explorer(), state="disabled")

# Function to show file table context menu - update to check explorer menu state
def show_file_table_context_menu(event):
    """Display the context menu when right-clicking on the file table."""
    # Only show if there are items selected
    if file_table.selection():
        # Update the state of the Show in Explorer menu item
        update_explorer_menu_state()
        file_table_context_menu.tk_popup(event.x_root, event.y_root)

# Bind the context menu to the right mouse button on the file table
file_table.bind("<Button-3>", show_file_table_context_menu)

# Add the function to play selected files
def play_selected_files():
    """Play selected files in the default music player."""
    selected_items = file_table.selection()
    if not selected_items:
        log_message("[ERROR] No files selected for playback", log_type="processing")
        return
        
    # Get all table items in their display order
    all_table_items = file_table.get_children()
    
    # Filter to only selected items, but maintain table order
    files_to_play = []
    for item in all_table_items:
        if item in selected_items:  # Only process selected items
            values = file_table.item(item)['values']
            file_path = values[8]  # File path is in position 8 (9th element, 0-indexed)
            
            if not file_path:
                log_message("[ERROR] Missing file path for playback", log_type="processing")
                continue
                
            # Check if the file exists
            if os.path.exists(file_path):
                files_to_play.append(file_path)
            else:
                log_message(f"[ERROR] File does not exist: {file_path}", log_type="processing")
    
    if not files_to_play:
        log_message("[ERROR] No valid files to play", log_type="processing")
        return
        
    try:
        # Use the appropriate command based on the operating system
        if platform.system() == 'Windows':
            for file in files_to_play:
                os.startfile(file)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open'] + files_to_play)
        else:  # Linux and others
            subprocess.call(['xdg-open'] + files_to_play)
        
        log_message(f"[SUCCESS] Playing {len(files_to_play)} files in table order", log_type="processing")
    except Exception as e:
        log_message(f"[ERROR] Failed to play files: {str(e)}", log_type="processing")

# ---------------- FUNCTIONS ---------------- #

def update_progress_bar(progress, bar_type="file", verbose=False):  # Changed default to False
    """Update progress bar value and color based on type."""
    # Debug print (only if verbose)
    if verbose:
        print(f"DEBUG: update_progress_bar called with progress={progress}, bar_type={bar_type}")
    
    if bar_type == "file":
        progress_var.set(progress)
    else:
        api_progress_var.set(progress)
    
    # Use the utility function for styling
    update_progress_bar_style(style, progress, bar_type)

def update_api_progress(state=None, verbose=False):  # Changed default to False
    """Update API progress bar based on rate limit headers
    
    Args:
        state: Optional state parameter ("start", "complete", or None)
        verbose: Whether to output detailed debug messages
    """
    # Use the imported function but pass our local update_progress_bar as the callback
    from services.api_client import update_api_progress as api_progress
    api_progress(state, verbose, update_progress_bar)

def enforce_api_limit():
    """Wrapper for API client's enforce_api_limit function."""
    from services.api_client import enforce_api_limit as api_enforce_limit
    return api_enforce_limit(app.update)

def update_rate_limits_from_headers(headers, update_progress=True, verbose=False):
    """Wrapper for API client's update_rate_limits_from_headers function."""
    from services.api_client import update_rate_limits_from_headers as api_update_rate_limits
    return api_update_rate_limits(headers, update_progress, verbose, update_progress_bar)

def update_file_metadata(file_path, metadata):
    """Update the MP3 file's metadata based on checkbox selections."""
    # Prepare the options based on UI checkboxes
    options = {
        'catalog': save_catalog_var.get(),
        'year': save_year_var.get(),
        'art': save_art_var.get()
    }
    
    # Prepare callbacks to maintain application state
    callbacks = {
        'log_message': log_message,
        'mark_updated': lambda path: updated_files.add(path),
        'mark_processed': lambda path: processed_files.add(path)
    }
    
    # Add API token to metadata if needed for cover art
    if 'api_token' not in metadata and DISCOGS_API_TOKEN:
        metadata['api_token'] = DISCOGS_API_TOKEN
    
    # Call the utility function
    return update_album_metadata(file_path, metadata, options=options, callbacks=callbacks)

def stop_processing_files():
    """Stop the file processing thread if it's running."""
    global stop_processing
    stop_processing = True
    log_message("Processing stopped by user", log_type="info")

def start_processing():
    """Start processing files in a separate thread."""
    global processing_thread, stop_processing
    
    stop_processing = False
    processing_thread = threading.Thread(target=process_files)
    processing_thread.daemon = True
    processing_thread.start()

def process_files():
    """Process the selected files and fetch metadata from the API."""
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
    
    # First, collect all file paths from selected items
    selected_files = []
    for item in selected_items:
        values = file_table.item(item)['values']
        if len(values) >= 9:  # Ensure there's a file path
            file_path = values[8]  # File path is in position 8
            if file_path and os.path.exists(file_path):
                selected_files.append(file_path)
                
                # Build the cache directly from files
                audio = get_audio_file(file_path)
                if audio:
                    file_metadata_cache[file_path] = {
                        "artist": get_tag_value(audio, "artist"),
                        "title": get_tag_value(audio, "title"),
                        "album": get_tag_value(audio, "album"),
                        "albumartist": get_tag_value(audio, "albumartist")
                    }
    
    # Also cache remaining files that aren't selected
    for file_path in file_list:
        if file_path not in selected_files and os.path.exists(file_path):
            audio = get_audio_file(file_path)
            if audio:
                file_metadata_cache[file_path] = {
                    "artist": get_tag_value(audio, "artist"),
                    "title": get_tag_value(audio, "title"),
                    "album": get_tag_value(audio, "album"),
                    "albumartist": get_tag_value(audio, "albumartist")
                }
    
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
            albumartist = metadata.get("albumartist", "")  # Add this line to get album artist
            log_message(f"[INFO] Extracted Metadata: Artist={artist}, Album={album}", log_type="debug")
            
            # Create cache key to check if we have this metadata already
            cache_key = f"{artist.lower()}|{album.lower()}"
            
            # Check if we already have cached metadata for this album
            cached_metadata = None
            with cache_lock:
                if cache_key in album_catalog_cache:
                    cached_metadata = album_catalog_cache[cache_key]
                    log_message(f"[INFO] Using cached metadata for '{artist} - {album}'", log_type="debug")
                # Add fallback check using albumartist+album if artist check fails
                elif albumartist and album:
                    albumartist_cache_key = f"{albumartist.lower()}|{album.lower()}"
                    if albumartist_cache_key in album_catalog_cache:
                        cached_metadata = album_catalog_cache[albumartist_cache_key]
                        log_message(f"[INFO] Using cached metadata via album artist match for '{albumartist} - {album}'", log_type="debug")
            
            # If we have cached metadata, use it without making an API call
            if cached_metadata:
                metadata = cached_metadata
                log_message(f"[INFO] Using cached metadata for '{artist} - {album}' - No API call needed", log_type="debug")
                # We don't touch the API progress bar at all when using cached metadata
            else:
                log_message(f"[INFO] No cached metadata found for '{artist} - {album}' - Making API call", log_type="debug")
                
                # Only enforce API limits and update progress if we're actually making an API call
                if not enforce_api_limit():
                    log_message("[WARNING] API rate limit reached. Pausing processing.", log_type="processing")
                    break
                    
                # Update API progress before call (to show we're about to make a call)
                # Use verbose=False to reduce debug output
                update_api_progress("start", verbose=False)
                
                # Make the actual API call to fetch metadata
                metadata_result = metadata_fetch_metadata(artist, album, title, api_token=DISCOGS_API_TOKEN, search_url=Config.DISCOGS_SEARCH_URL)
                
                # Handle the new return format (metadata, headers)
                if isinstance(metadata_result, tuple):
                    metadata, response_headers = metadata_result
                    # Update rate limits from the headers if available 
                    # Pass update_progress=True to let it handle the completion update
                    if response_headers:
                        update_rate_limits_from_headers(response_headers, update_progress=True, verbose=False)
                else:
                    # Backwards compatibility with older versions
                    metadata = metadata_result
                    # In this case, we still need to update the API progress manually
                    update_api_progress("complete", verbose=False)
                
                # No need for a separate update_api_progress call here since it's handled above
            
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
        table_metadata = [values[0], values[1], values[2], values[4]]  # Artist, Title, Album, Album Artist
        
        # Find matching file using cached metadata - with improved numeric matching
        for file_path, metadata in file_metadata_cache.items():
            current_metadata = [
                metadata.get("artist", ""),
                metadata.get("title", ""),
                metadata.get("album", ""),
                metadata.get("albumartist", "")
            ]
            
            # Check if all values match, with special handling for numeric values
            is_match = True
            for a, b in zip(current_metadata, table_metadata):
                a_str = str(a).strip()
                b_str = str(b).strip()
                
                # If strings are equal, they match
                if a_str == b_str:
                    continue
                
                # Try numeric comparison if both can be converted to numbers
                try:
                    if a_str and b_str and float(a_str) == float(b_str):
                        continue
                except ValueError:
                    pass
                
                # If we reach here, values don't match
                is_match = False
                break
            
            if is_match:
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
    
    # Get values from the table row, including the file path
    values = file_table.item(editing_item)['values']
    file_path = values[8]  # File path is in the last column
    
    # Map column indices to tag names
    column_to_tag = {
        0: "artist",       # Artist
        1: "title",        # Title
        2: "album",        # Album
        3: "catalognumber", # Catalog Number
        4: "albumartist",  # Album Artist
        5: "date",         # Year
        6: "tracknumber",  # Track
        7: "genre",        # Genre
    }
    
    # Get value directly from the audio file if possible
    current_value = ""
    if file_path and os.path.exists(file_path) and column_num in column_to_tag:
        tag_name = column_to_tag[column_num]
        audio = get_audio_file(file_path)
        if audio:
            # Get the value directly from the file to preserve leading zeros
            current_value = get_tag_value(audio, tag_name, "")
        else:
            # Fallback to table value
            current_value = values[column_num]
            if current_value is not None:
                current_value = str(current_value)
            else:
                current_value = ""
    else:
        # Fallback to table value
        current_value = values[column_num]
        if current_value is not None:
            current_value = str(current_value)
        else:
            current_value = ""

    # Get the cell's bounding box
    x, y, w, h = file_table.bbox(editing_item, editing_column)
    
    # Create and place the entry widget
    editing_entry = create_styled_entry(table_container, textvariable=tk.StringVar(value=current_value), width=w)
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
    
    # Find matching file using the ORIGINAL metadata - with improved numeric matching
    matching_file = None
    for file_path, metadata in file_metadata_cache.items():
        current_metadata = [
            metadata.get("artist", ""),
            metadata.get("title", ""),
            metadata.get("album", ""),
            metadata.get("albumartist", "")
        ]
        
        # Check if all values match, with special handling for numeric values
        is_match = True
        for a, b in zip(current_metadata, original_metadata):
            a_str = str(a).strip()
            b_str = str(b).strip()
            
            # If strings are equal, they match
            if a_str == b_str:
                continue
            
            # Try numeric comparison if both can be converted to numbers
            try:
                if a_str and b_str and float(a_str) == float(b_str):
                    continue
            except ValueError:
                pass
            
            # If we reach here, values don't match
            is_match = False
            break
        
        if is_match:
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
    auto_adjust_column_widths(file_table, columns)

def update_mp3_metadata(file_path, column_num, new_value):
    """Update the audio file's metadata based on the edited column."""
    # Define callbacks to use with the utility function
    callbacks = {
        'log_message': log_message,
        'mark_updated': lambda path: updated_files.add(path)
    }
    
    # Use the utility function from utils/metadata.py
    return metadata_update_mp3_metadata(file_path, column_num, new_value, callbacks=callbacks)

def remove_selected_items():
    """Remove selected items from the file list."""
    table_ops_remove_items(
        file_table,
        file_list,
        file_metadata_cache,
        processed_files,
        updated_files,
        file_count_var,
        log_message
    )
    # Force UI update
    app.update_idletasks()

def refresh_file_list():
    """Refresh the file list by re-scanning selected folders and keeping individual files."""
    global file_list, processed_files, updated_files, file_metadata_cache
    
    log_message(f"[DEBUG] Starting refresh. Current selected folders: {list(selected_folders)}")
    log_message(f"[DEBUG] Current file list has {len(file_list)} files")
    
    # Save current table order before refreshing
    current_table_order = []
    for item in file_table.get_children():
        values = file_table.item(item)['values']
        if len(values) >= 9:  # Ensure we have the file path
            file_path = values[8]  # File path is in position 8
            if file_path:
                current_table_order.append(file_path)
    
    log_message(f"[DEBUG] Saved table order with {len(current_table_order)} files")
    
    # Keep track of individual files (not from folders)
    individual_files = [f for f in file_list if os.path.dirname(f) not in selected_folders]
    
    # Make sure individual files still exist
    individual_files = [f for f in individual_files if os.path.exists(f)]
    
    log_message(f"[DEBUG] Found {len(individual_files)} individual files to preserve")
    
    # Re-scan all selected folders - but only scan the exact folder, not the entire tree
    folder_files = []
    for folder in selected_folders:
        if os.path.exists(folder):  # Check if folder still exists
            log_message(f"[DEBUG] Scanning folder: {folder}")
            # Only scan the selected folder itself, not recursively through subfolders
            # This prevents loading the entire collection when refreshing a subfolder
            try:
                files_in_folder = os.listdir(folder)
                new_files = []
                for file in files_in_folder:
                    file_path = os.path.join(folder, file)
                    if os.path.isfile(file_path) and file.lower().endswith(tuple(Config.SUPPORTED_AUDIO_EXTENSIONS)):
                        new_files.append(file_path)
                    elif os.path.isdir(file_path):
                        # If it's a subdirectory, only scan it if it was explicitly selected
                        # This maintains the current behavior for explicitly selected subfolders
                        if file_path in selected_folders:
                            log_message(f"[DEBUG] Found explicitly selected subfolder: {file_path}")
                            subfolder_files = [os.path.join(root, f) 
                                            for root, _, files in os.walk(file_path) 
                                            for f in files if f.lower().endswith(tuple(Config.SUPPORTED_AUDIO_EXTENSIONS))]
                            new_files.extend(subfolder_files)
                            log_message(f"[DEBUG] Added {len(subfolder_files)} files from subfolder")
                folder_files.extend(new_files)
                log_message(f"[DEBUG] Added {len(new_files)} files from folder {folder}")
            except PermissionError:
                log_message(f"[WARNING] Permission denied accessing folder: {folder}")
                continue
        else:
            log_message(f"[WARNING] Folder no longer exists: {folder}")
            selected_folders.remove(folder)
    
    log_message(f"[DEBUG] Total folder files found: {len(folder_files)}")
    
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
    
    log_message(f"[DEBUG] New file list created with {len(new_file_list)} files (removed {len(folder_files) + len(individual_files) - len(new_file_list)} duplicates)")
    
    # Update the file list
    file_list = new_file_list
    
    # Clear processed and updated files sets
    processed_files.clear()
    updated_files.clear()
    file_metadata_cache.clear()  # Clear the metadata cache
    
    # Clear and update the table
    file_table.delete(*file_table.get_children())
    update_table()
    
    # Restore the table order if we had a previous order
    if current_table_order:
        # Create a mapping of file paths to their desired positions
        order_map = {path: idx for idx, path in enumerate(current_table_order)}
        
        # Get all current table items
        current_items = file_table.get_children()
        
        # Sort the items based on the saved order
        def get_sort_key(item):
            values = file_table.item(item)['values']
            if len(values) >= 9:
                file_path = values[8]
                return order_map.get(file_path, len(current_table_order))  # Put unknown items at the end
            return len(current_table_order)
        
        # Sort items by their position in the saved order
        sorted_items = sorted(current_items, key=get_sort_key)
        
        # Reorder the items in the table
        for idx, item in enumerate(sorted_items):
            file_table.move(item, '', idx)
        
        log_message(f"[DEBUG] Restored table order for {len(sorted_items)} items")
    
    log_message(f"[INFO] Refreshed file list. Total files: {len(file_list)}")


def update_basic_fields(event=None):
    """Update the basic fields based on table selection."""
    global current_album_art, current_album_art_bytes, file_metadata_cache, pending_album_art, shared_album_art_files
    
    # Get the selected items
    selected_items = file_table.selection()
    
    # If no items are selected, clear all fields and show default album art
    if not selected_items:
        for var in basic_field_vars.values():
            var.set("")
        load_default_album_art()
        pending_album_art = None
        current_album_art_bytes = None
        return
    
    # Get values for all selected items
    values_by_field = {field: [] for field in basic_field_vars.keys()}
    
    # Check if multiple albums are selected
    albums = set()
    artists = set()
    for item in selected_items:
        values = file_table.item(item)['values']
        if values[2]:  # Album
            albums.add(values[2])
        if values[0]:  # Artist
            artists.add(values[0])
    
    # Keep pending album art if available
    if pending_album_art and pending_album_art != "REMOVE":
        log_message("[COVER] Using pending album art for display", log_type="debug")
        current_album_art_bytes = pending_album_art
        photo = image_update_album_art_display(
            pending_album_art,
            label=album_cover_label,
            size=Config.ALBUM_ART["COVER_SIZE"],
            load_default_func=load_default_album_art
        )
        current_album_art = photo
        
        # Process metadata fields
        process_metadata_fields(selected_items, values_by_field)
        return
        
    # For album art, we need to check if all files have the same art
    art_data = None
    found_album_art = False
    different_art = False
    first_file_path = None
    
    log_message(f"[DEBUG] Checking album art for {len(selected_items)} selected items", log_type="debug")
    
    # Check for album art in selected files
    for item in selected_items:
        values = file_table.item(item)['values']
        
        # Check if the values array has enough elements
        if len(values) < 9:
            log_message(f"[ERROR] Invalid table values: {values}", log_type="debug")
            continue
            
        # Get the file path from the values array (last element)
        file_path = values[8]  # File path is now in position 8 (9th element, 0-indexed)
        
        if not file_path:
            log_message("[ERROR] Missing file path in table item", log_type="debug")
            continue
            
        # Make sure this is a string
        file_path = str(file_path)
        
        if first_file_path is None:
            first_file_path = file_path
            
        log_message(f"[DEBUG] Processing file for album art: {file_path}", log_type="debug")
            
        # Get album art
        audio = get_audio_file(file_path)
        if audio:
            current_art = extract_album_art_from_file(file_path, audio)
            if current_art:
                log_message(f"[DEBUG] Found album art in file: {file_path} ({len(current_art)} bytes)", log_type="debug")
                if not found_album_art:
                    # First art found
                    art_data = current_art
                    found_album_art = True
                elif not different_art:
                    # Compare with first art
                    # Simplified check: just compare if bytes are identical
                    # This will work reliably for files with identical art
                    if art_data != current_art:
                        log_message(f"[DEBUG] Different album art found in file: {file_path}", log_type="debug")
                        different_art = True
            else:
                log_message(f"[DEBUG] No album art found in file: {file_path}", log_type="debug")
    
    # Handle album art based on our checks
    if found_album_art and not different_art:
        # All files have the same album art
        current_album_art_bytes = art_data
        photo = image_update_album_art_display(
            art_data,
            label=album_cover_label,
            size=Config.ALBUM_ART["COVER_SIZE"],
            load_default_func=load_default_album_art
        )
        current_album_art = photo
    else:
        # Files have different album art or no art found
        if different_art:
            log_message("[COVER] Selected files have different album art", log_type="processing")
        else:
            log_message("[COVER] No album art found in selected files", log_type="debug")
        load_default_album_art()
        current_album_art_bytes = None
    
    # Process metadata fields
    process_metadata_fields(selected_items, values_by_field)

def process_metadata_fields(selected_items, values_by_field):
    """Process metadata fields for the selected items."""
    
    # Get the original values directly from file metadata instead of the table
    for item in selected_items:
        values = file_table.item(item)['values']
        file_path = values[8]  # File path is the last column
        
        if file_path and os.path.exists(file_path):
            # Get metadata directly from file instead of table values
            audio = get_audio_file(file_path)
            if audio:
                field_mapping = {
                    "Artist": get_tag_value(audio, "artist", ""),
                    "Title": get_tag_value(audio, "title", ""),
                    "Album": get_tag_value(audio, "album", ""),
                    "Catalog Number": get_tag_value(audio, "catalognumber", ""),
                    "Album Artist": get_tag_value(audio, "albumartist", ""),
                    "Year": get_tag_value(audio, "date", ""),
                    "Track": get_tag_value(audio, "tracknumber", ""),
                    "Genre": get_tag_value(audio, "genre", "")
                }
                
                # Add values to their respective lists
                for field, value in field_mapping.items():
                    values_by_field[field].append(value)
            else:
                # Fallback to table values if file can't be read
                table_mapping = {
                    "Artist": values[0],
                    "Title": values[1],
                    "Album": values[2],
                    "Catalog Number": values[3],
                    "Album Artist": values[4],
                    "Year": values[5],
                    "Track": values[6],
                    "Genre": values[7]
                }
                for field, value in table_mapping.items():
                    values_by_field[field].append(str(value) if value is not None else "")
        else:
            # Fallback if no file path or file doesn't exist
            table_mapping = {
                "Artist": values[0],
                "Title": values[1],
                "Album": values[2],
                "Catalog Number": values[3],
                "Album Artist": values[4],
                "Year": values[5],
                "Track": values[6],
                "Genre": values[7]
            }
            for field, value in table_mapping.items():
                values_by_field[field].append(str(value) if value is not None else "")
    
    # Set values in all fields (unchanged)
    for field, var in basic_field_vars.items():
        values = values_by_field[field]
        
        if not any(values):  # If all values are empty
            var.set("")
        elif len(set(values)) == 1:
            # All values are the same
            var.set(values[0])
        else:
            # Different values
            var.set("<different values>")

def apply_basic_fields():
    """Apply metadata from basic fields to selected files."""
    global pending_album_art, shared_album_art_files
    
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
    album_art_updated_files = []  # Track files that had album art updated
    
    for item in selected_items:
        values = file_table.item(item)['values']
        table_metadata = [values[0], values[1], values[2], values[4]]  # Artist, Title, Album, Album Artist
        
        # Find matching file using cached metadata - with improved numeric matching
        matching_file = None
        for file_path, metadata in file_metadata_cache.items():
            current_metadata = [
                metadata.get("artist", ""),
                metadata.get("title", ""),
                metadata.get("album", ""),
                metadata.get("albumartist", "")
            ]
            
            # Check if all values match, with special handling for numeric values
            is_match = True
            for a, b in zip(current_metadata, table_metadata):
                a_str = str(a).strip()
                b_str = str(b).strip()
                
                # If strings are equal, they match
                if a_str == b_str:
                    continue
                
                # Try numeric comparison if both can be converted to numbers
                try:
                    if a_str and b_str and float(a_str) == float(b_str):
                        continue
                except ValueError:
                    pass
                
                # If we reach here, values don't match
                is_match = False
                break
            
            if is_match:
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
                                        desc='Front Cover',
                                        data=pending_album_art
                                    )
                                )
                                updated = True
                                album_art_updated_files.append(matching_file)
                                log_message(f"[SUCCESS] Updated album art for {os.path.basename(matching_file)}")
                            elif isinstance(audio, mutagen.flac.FLAC):
                                # Clear existing pictures
                                audio.clear_pictures()
                                
                                # Create new picture
                                picture = mutagen.flac.Picture()
                                picture.type = 3  # Front cover
                                picture.mime = mime_type
                                picture.desc = 'Front Cover'
                                picture.data = pending_album_art
                                
                                # Add picture to FLAC file
                                audio.add_picture(picture)
                                updated = True
                                album_art_updated_files.append(matching_file)
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
                                album_art_updated_files.append(matching_file)
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
    
    # Update shared_album_art_files set with files that now have the same album art
    if album_art_updated_files:
        # Update the set of files that share the art
        for file_path in album_art_updated_files:
            # Normalize the path to ensure consistent comparison
            normalized_path = os.path.normpath(file_path)
            shared_album_art_files.add(normalized_path)
        log_message(f"[COVER] Total files with shared album art: {len(shared_album_art_files)}", log_type="debug")
    
    # Reset pending album art only if we've successfully applied all updates
    # We'll keep the shared_album_art_files list for use in other functions
    pending_album_art = None
    
    if updated_count > 0:
        log_message(f"[INFO] Successfully updated {updated_count} files")
    else:
        log_message("[WARNING] No files were updated")

# Add save button below album art
save_metadata_button = create_styled_button(left_panel, "SAVE METADATA", apply_basic_fields)
save_metadata_button.pack(fill="x", padx=10, pady=5)

# Add Clear buttons at the bottom of left panel - CREATE THESE FIRST
try:
    # Create a container for the Clear buttons - side by side
    buttons_container = ttk.Frame(left_panel)
    buttons_container.pack(side="bottom", fill="x", padx=10, pady=5)
    
    # Use lambda functions to delay the lookup of the functions until they're clicked
    create_button_pair(
        buttons_container, 
        "CLEAR FILES", 
        lambda: clear_file_list(),  # Use lambda to delay lookup
        "CLEAR LOGS", 
        lambda: clear_logs()
    )
except Exception as e:
    log_message(f"[ERROR] Failed to add clear buttons: {str(e)}")

# Add processing log to left panel - AFTER the buttons are created
processing_container = ttk.Frame(left_panel)
processing_container.pack(fill="both", expand=True, padx=10, pady=5)

processing_listbox = create_styled_text(processing_container, height=25, state="disabled", wrap="word")
processing_listbox.pack(side="left", fill="both", expand=True)

# Configure tags for colored text
configure_text_tags(processing_listbox)

# Add scrollbar to processing listbox with autohide - using default style
processing_scrollbar = ttk.Scrollbar(processing_container, orient="vertical", command=processing_listbox.yview)
processing_listbox.configure(yscrollcommand=lambda f, l: autohide_scrollbar(processing_scrollbar, f, l))
processing_scrollbar.pack(side="right", fill="y")

# Update the logger with the processing widget
logger.set_processing_widget(processing_listbox)

def load_default_album_art():
    """Load the default album art image when no art is available."""
    global current_album_art, current_album_art_bytes
    
    # Reset the bytes data
    current_album_art_bytes = None
    
    # Use the function from utils.image_handling
    photo = image_load_default_album_art(
        default_image_path=Config.ALBUM_ART["DEFAULT_IMAGE"],
        label=album_cover_label,
        size=(Config.ALBUM_ART["COVER_SIZE"], Config.ALBUM_ART["COVER_SIZE"])
    )
    
    # Keep a reference to prevent garbage collection
    current_album_art = photo
    return photo is not None

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
        
    # Show the folder format dialog with callback
    show_folder_format_dialog(app, custom_font, organize_files_with_format)

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
        
        # Handle genre with backslash or semicolon separator - take only the part before first separator
        if "\\" in genre:
            genre = genre.split("\\")[0].strip()
            log_message(f"[INFO] Using first genre component: {genre}")
        elif ";" in genre:
            genre = genre.split(";")[0].strip()
            log_message(f"[INFO] Using first genre component: {genre}")
            
        # Sanitize values for use in paths
        from utils.file_operations import sanitize_filename as file_ops_sanitize_filename
        safe_genre = file_ops_sanitize_filename(genre)
        safe_year = file_ops_sanitize_filename(year)
        safe_catalognumber = file_ops_sanitize_filename(catalognumber)
        safe_albumartist = file_ops_sanitize_filename(albumartist)
        safe_album = file_ops_sanitize_filename(album)
        safe_artist = file_ops_sanitize_filename(artist)
        safe_title = file_ops_sanitize_filename(title)
        
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

    def execute_move():
        """Execute the file move operation."""
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

    # Show the confirmation dialog
    show_move_confirmation_dialog(app, custom_font, files_to_move, skipped_files, execute_move)

# Bind table events after all functions are defined
file_table.bind('<Double-1>', start_editing)
file_table.bind('<Escape>', cancel_editing)
file_table.bind('<Return>', finish_editing)
file_table.bind('<<TreeviewSelect>>', 
    lambda e: (file_table_selection_callback(file_table, file_count_var), update_basic_fields(e)))
# Update these bindings to pass None instead of the event
file_table.bind('<Control-a>', lambda e: select_all_visible(file_table, file_count_var, filter_var.get()))

# Add Ctrl+S shortcut for saving metadata
app.bind('<Control-s>', lambda e: apply_basic_fields())

# Configure the root window background
app.configure(bg=Config.COLORS["BACKGROUND"])

# Apply style to main paned window
main_paned.configure(style='Dark.TPanedwindow')

# Initialize the album art display with the default image
load_default_album_art()

# Set the window to start maximized (windowed fullscreen)
app.state('zoomed')  # For Windows systems

def update_table():
    """Update the table with current file list and metadata."""
    table_ops_update_table(file_table, apply_filter, file_count_var, columns)  # Use renamed import
    # Force UI update
    app.update_idletasks()

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
    # Use the new logger's clear_logs method
    logger.clear_logs(app, debug_scrollbar, processing_scrollbar)

# Add the function to open the folder in explorer
def show_in_explorer():
    """Open the folder containing the selected files in Windows Explorer."""
    selected_items = file_table.selection()
    if not selected_items:
        log_message("[ERROR] No files selected to show in explorer", log_type="processing")
        return
        
    # Get all directories from selected files
    directories = set()
    for item in selected_items:
        values = file_table.item(item)['values']
        file_path = values[8]  # File path is in position 8 (9th element, 0-indexed)
        
        if not file_path:
            log_message("[ERROR] Missing file path", log_type="processing")
            continue
            
        # Get the directory of the file
        directory = os.path.dirname(file_path)
        if os.path.exists(directory):
            directories.add(directory)
    
    # If all files are in the same directory, open it
    if len(directories) == 1:
        directory = next(iter(directories))
        try:
            if platform.system() == 'Windows':
                os.startfile(directory)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', directory])
            else:  # Linux and others
                subprocess.call(['xdg-open', directory])
            
            log_message(f"[SUCCESS] Opened folder: {directory}", log_type="processing")
        except Exception as e:
            log_message(f"[ERROR] Failed to open folder: {str(e)}", log_type="processing")
    else:
        log_message("[ERROR] Selected files are in different folders", log_type="processing")

# Function to check if the "Show in Explorer" option should be enabled
def update_explorer_menu_state():
    """Enable or disable the Show in Explorer menu item based on selection."""
    selected_items = file_table.selection()
    
    # Default to disabled
    file_table_context_menu.entryconfig("Show in Explorer", state="disabled")
    
    if not selected_items:
        return
        
    # Get all directories from selected files
    directories = set()
    for item in selected_items:
        values = file_table.item(item)['values']
        file_path = values[8]  # File path is in position 8
        
        if not file_path:
            continue
            
        # Get the directory of the file
        directory = os.path.dirname(file_path)
        if os.path.exists(directory):
            directories.add(directory)
    
    # Enable if all files are in the same directory
    if len(directories) == 1:
        file_table_context_menu.entryconfig("Show in Explorer", state="normal")

app.mainloop()
