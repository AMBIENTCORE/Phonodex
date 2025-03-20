"""
File operation utilities for the application.
Provides functionality for file handling, selection, path manipulation, and audio file operations.
"""

import os
import sys
import re
import shutil
from tkinter import filedialog
from utils.logging import log_message

# Import Mutagen libraries for audio file handling
import mutagen
from mutagen.mp3 import MP3
from mutagen.flac import FLAC
from mutagen.mp4 import MP4
from mutagen.oggvorbis import OggVorbis
from mutagen.asf import ASF
from mutagen.wave import WAVE

def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller.
    
    Args:
        relative_path: Relative path to the resource
        
    Returns:
        Absolute path to the resource
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def select_files(file_type_description, supported_extensions, file_list_var=None, count_var=None, update_table_func=None):
    """
    Open file dialog to select audio files.
    
    Args:
        file_type_description: Description of the file type for the dialog
        supported_extensions: List of supported file extensions
        file_list_var: Reference to the list where selected files will be stored
        count_var: Optional StringVar to update with file count
        update_table_func: Optional function to call to update the UI table
    
    Returns:
        List of selected file paths
    """
    files = filedialog.askopenfilenames(filetypes=[(file_type_description, "*" + supported_extensions[0])])
    
    selected_files = list(files)  # Convert to list
    
    if files and file_list_var is not None:
        file_list_var.extend(selected_files)
        
        # Update counter if provided
        if count_var:
            count_var.set(f"{len(file_list_var)}/{len(file_list_var)}")
            
        # Update UI if function provided
        if update_table_func:
            update_table_func()
            
    return selected_files

def select_folder(update_table_func=None, file_list_var=None, metadata_cache=None, 
                  processed_files=None, updated_files=None, selected_folders_var=None,
                  supported_extensions=None, count_var=None):
    """
    Open a dialog to select a folder and add all audio files inside it, including subfolders.
    
    Args:
        update_table_func: Function to call to update the UI table
        file_list_var: Reference to the list where selected files will be stored
        metadata_cache: Dictionary to cache file metadata
        processed_files: Set of already processed files
        updated_files: Set of updated files
        selected_folders_var: Set to store selected folder paths
        supported_extensions: List of supported file extensions
        count_var: StringVar to update with file count
    
    Returns:
        List of found audio files
    """
    folder_selected = filedialog.askdirectory()
    if not folder_selected:
        return []
        
    # Clear existing data if variables provided
    if file_list_var is not None:
        file_list_var.clear()
    if metadata_cache is not None:
        metadata_cache.clear()
    if processed_files is not None:
        processed_files.clear()
    if updated_files is not None:
        updated_files.clear()
        
    # Store the selected folder for potential refresh operations
    if selected_folders_var is not None:
        selected_folders_var.add(folder_selected)
        
    # Find all matching files recursively
    found_files = []
    for root, _, files in os.walk(folder_selected):
        for file in files:
            if any(file.lower().endswith(ext) for ext in supported_extensions):
                found_files.append(os.path.join(root, file))
                
    # Update file list if provided
    if file_list_var is not None:
        file_list_var.extend(found_files)
        
        # Update counter if provided
        if count_var:
            count_var.set(f"{len(file_list_var)}/{len(file_list_var)}")
            
        # Update UI if function provided  
        if update_table_func:
            update_table_func()
            
    return found_files

def handle_drop(files, file_list_var=None, processed_files=None, updated_files=None, 
               selected_folders_var=None, metadata_cache=None, table=None,
               supported_extensions=None, count_var=None, update_table_func=None):
    """
    Handle dropped files and add them to the file list.
    
    Args:
        files: File paths from drag and drop event
        file_list_var: Reference to the list where files will be stored
        processed_files: Set of already processed files
        updated_files: Set of updated files
        selected_folders_var: Set to store selected folder paths
        metadata_cache: Dictionary to cache file metadata
        table: Treeview widget to clear
        supported_extensions: List of supported file extensions
        count_var: StringVar to update with file count
        update_table_func: Function to update the UI table
        
    Returns:
        List of added files
    """
    # Clear all data structures if provided
    if file_list_var is not None:
        file_list_var.clear()
    if processed_files is not None:
        processed_files.clear()
    if updated_files is not None:
        updated_files.clear()
    if selected_folders_var is not None:
        selected_folders_var.clear()
    if metadata_cache is not None:
        metadata_cache.clear()
    
    # Clear the table if provided
    if table:
        table.delete(*table.get_children())
    
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
                # It's a folder - add it to selected folders if tracking
                if selected_folders_var is not None:
                    selected_folders_var.add(path)
                
                # Find all audio files recursively
                for root, _, files in os.walk(path):
                    for file in files:
                        if any(file.lower().endswith(ext) for ext in supported_extensions):
                            full_path = os.path.join(root, file)
                            all_files.append(full_path)
            elif os.path.isfile(path) and any(path.lower().endswith(ext) for ext in supported_extensions):
                # It's a supported audio file
                all_files.append(path)
        except Exception as e:
            log_message(f"[ERROR] Failed to process path {path}: {str(e)}")
    
    # Update file list if provided
    if file_list_var is not None:
        file_list_var.extend(all_files)
        
        # Update counter if provided
        if count_var:
            count_var.set(f"{len(file_list_var)}/{len(file_list_var)}")
    
    # Update UI if function provided
    if update_table_func:
        update_table_func()
        
    return all_files

def get_audio_file(file_path):
    """
    Helper function to safely get an audio file object with appropriate tag handling.
    
    Args:
        file_path: Path to the audio file
        
    Returns:
        Audio file object with the appropriate type based on file extension
    """
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
            log_message(f"[ERROR] Unsupported file type: {ext}")
            return None
    except Exception as e:
        log_message(f"[ERROR] Failed to load audio file {file_path}: {str(e)}")
        return None

def sanitize_filename(filename):
    """
    Sanitize filename by removing or replacing invalid characters.
    
    Args:
        filename: Original filename
        
    Returns:
        Sanitized filename safe for use in file systems
    """
    # Replace characters that are invalid in filenames
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename

def move_file_to_destination(source_path, dest_path, create_dirs=True):
    """
    Move a file to a destination path, creating directories if needed.
    
    Args:
        source_path: Source file path
        dest_path: Destination file path
        create_dirs: Whether to create destination directories if they don't exist
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create destination directories if needed
        if create_dirs:
            dest_dir = os.path.dirname(dest_path)
            os.makedirs(dest_dir, exist_ok=True)
            
        # Move the file
        shutil.move(source_path, dest_path)
        return True
    except Exception as e:
        log_message(f"[ERROR] Failed to move file from {source_path} to {dest_path}: {str(e)}")
        return False

def copy_file_to_destination(source_path, dest_path, create_dirs=True):
    """
    Copy a file to a destination path, creating directories if needed.
    
    Args:
        source_path: Source file path
        dest_path: Destination file path
        create_dirs: Whether to create destination directories if they don't exist
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create destination directories if needed
        if create_dirs:
            dest_dir = os.path.dirname(dest_path)
            os.makedirs(dest_dir, exist_ok=True)
            
        # Copy the file
        shutil.copy2(source_path, dest_path)
        return True
    except Exception as e:
        log_message(f"[ERROR] Failed to copy file from {source_path} to {dest_path}: {str(e)}")
        return False
