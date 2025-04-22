"""
Dialog windows for the Phonodex application.
"""

import tkinter as tk
from tkinter import ttk, StringVar, messagebox
from config import Config, folder_format, DEFAULT_FOLDER_FORMAT, save_settings
from ui.styles import style_button, create_styled_entry, style_label
import os

def show_folder_format_dialog(parent_window, custom_font, on_continue_callback):
    """Show a dialog to edit the folder structure format.
    
    Args:
        parent_window: The parent window (main app window)
        custom_font: The font to use for the dialog
        on_continue_callback: Function to call when Continue is clicked
    """
    global folder_format
    
    format_dialog = tk.Toplevel(parent_window)
    format_dialog.title("Configure Folder Structure")
    format_dialog.geometry("1000x550")  # Increased height from 300 to 400
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
    
    # Create a frame to hold the entry field and save button
    format_entry_frame = ttk.Frame(format_dialog)
    format_entry_frame.pack(fill="x", padx=10, pady=5)
    
    format_var = StringVar(value=folder_format)
    
    # Use the create_styled_entry function for consistency
    format_entry = create_styled_entry(
        format_entry_frame, 
        textvariable=format_var,
        justify="center"
    )
    format_entry.pack(side="left", fill="x", expand=True)
    
    # Set initial background to valid (green) since we loaded it successfully
    format_entry.configure(bg=Config.COLORS["VALID_ENTRY"])
    
    # Create an error message label (initially empty)
    error_var = StringVar(value="")
    error_label = tk.Label(format_dialog, 
                          textvariable=error_var,
                          font=('Consolas', Config.FONTS["TABLE_SIZE"]),
                          bg=Config.COLORS["BACKGROUND"],
                          fg="#F44336")  # Red color for errors
    error_label.pack(fill="x", padx=10, pady=(5, 0))
    
    # Function to update the error message
    def show_error(message):
        error_var.set(message)
        format_entry.configure(bg=Config.COLORS["INVALID_ENTRY"])
    
    # Function to clear the error message
    def clear_error():
        error_var.set("")
    
    # Add reset to default button in its own container
    def reset_to_default():
        format_var.set(DEFAULT_FOLDER_FORMAT)
        # Reset background to valid (green) when setting to default
        format_entry.configure(bg=Config.COLORS["VALID_ENTRY"])
        clear_error()  # Clear any error message
    
    # Function to validate the folder format
    def validate_folder_format(format_string):
        """Validate that the folder format makes sense."""
        # Check if format ends with a directory separator
        if format_string.endswith('\\') or format_string.endswith('/'):
            return False, "Format cannot end with a directory separator (\\, /)"
            
        # Check if format includes at least one metadata placeholder
        placeholders = ['%genre%', '%year%', '%catalognumber%', '%albumartist%', '%album%', '%artist%', '%title%']
        if not any(placeholder in format_string for placeholder in placeholders):
            return False, "Format must include at least one metadata placeholder"
            
        # Check if the format has a filename component
        has_title = '%title%' in format_string
        has_extension = any(ext in format_string.lower() for ext in ['.mp3', '.flac', '.m4a', '.ogg', '.wav'])
        
        if not (has_title or has_extension):
            return False, "Format must include a filename component (%title% or a file extension)"
            
        # Check for basic Windows path validity
        if ':' in format_string:
            # If there's a drive letter, make sure it's correctly formatted
            import re
            drive_pattern = r'^[a-zA-Z]:\\.*'
            if not re.match(drive_pattern, format_string):
                return False, "Invalid drive format. Must be like 'C:\\...'"
                
        # Check for invalid filename characters
        invalid_chars = ['<', '>', ':', '"', '|', '?', '*']
        for char in invalid_chars:
            if char in format_string.replace(':', '').replace('\\', ''):  # Allow ':' in drive letter and '\' in path
                return False, f"Invalid character in format: '{char}'"
                
        return True, ""
    
    # Add save button - use 💾 icon like the API key save button
    def save_format():
        global folder_format
        
        # Clear any previous error
        clear_error()
        
        # Validate format first
        format_string = format_var.get()
        is_valid, error_message = validate_folder_format(format_string)
        
        if not is_valid:
            show_error(error_message)
            return
            
        try:
            # Update global variable
            folder_format = format_string
            
            # Save to file - with explicit file path for debugging
            settings_path = os.path.abspath(Config.FOLDER_STRUCTURE["SETTINGS_FILE"])
            
            # Debug message about path
            print(f"Attempting to save to: {settings_path}")
            
            # Try to directly write to the file instead of using save_settings
            import json
            with open(settings_path, 'w') as f:
                settings = {
                    'folder_format': folder_format
                }
                json.dump(settings, f, indent=4)
            
            # Just verify the file exists without showing a popup
            if not os.path.exists(settings_path):
                # Show error in the label
                show_error(f"Settings file not found after save: {settings_path}")
            else:
                # Success - set background to green
                format_entry.configure(bg=Config.COLORS["VALID_ENTRY"])
                
        except Exception as e:
            error_msg = f"Error saving settings: {str(e)}"
            print(error_msg)  # Print to console
            show_error(error_msg)
            
    # Create the save button with the same style as the API key save button
    save_button = tk.Button(format_entry_frame, text="💾", 
                         width=Config.DIMENSIONS["SAVE_BUTTON_WIDTH"],
                         command=save_format)
    style_button(save_button)
    save_button.pack(side="left", padx=(Config.PADDING["SMALL"], 0))
    
    # Add a button container frame
    button_container = ttk.Frame(format_dialog)
    button_container.pack(pady=5, padx=10, anchor="e")
    
    reset_button = tk.Button(button_container, 
                          text="RESET TO DEFAULT",
                          command=reset_to_default)
    style_button(reset_button)
    reset_button.pack(side="left", padx=0)
    
    # Add button frame at the bottom
    button_frame = ttk.Frame(format_dialog)
    button_frame.pack(fill="x", pady=10, padx=10, side="bottom")
    
    # Function to save and continue
    def save_and_continue():
        global folder_format
        
        # Clear any previous error
        clear_error()
        
        # Validate format first
        format_string = format_var.get()
        is_valid, error_message = validate_folder_format(format_string)
        
        if not is_valid:
            show_error(error_message)
            return
            
        try:
            # Save the new format
            folder_format = format_string
            
            # Save to file - with explicit file path for debugging
            settings_path = os.path.abspath(Config.FOLDER_STRUCTURE["SETTINGS_FILE"])
            
            # Debug message about path
            print(f"Saving format before continuing: {folder_format}")
            print(f"Saving to: {settings_path}")
            
            # Try to directly write to the file instead of using save_settings
            import json
            with open(settings_path, 'w') as f:
                settings = {
                    'folder_format': folder_format
                }
                json.dump(settings, f, indent=4)
                
            # Success - set background to green (even though we'll be closing the dialog)
            format_entry.configure(bg=Config.COLORS["VALID_ENTRY"])
                
        except Exception as e:
            error_msg = f"Error saving settings before continue: {str(e)}"
            print(error_msg)
            show_error(error_msg)
            return
            
        # Close the dialog
        format_dialog.destroy()
        # Continue with organization
        on_continue_callback()
    
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
    parent_window.wait_window(format_dialog)
    return False  # Dialog was canceled

def show_move_confirmation_dialog(parent_window, custom_font, files_to_move, skipped_files, on_confirm_callback):
    """Show confirmation dialog for file moves.
    
    Args:
        parent_window: The parent window
        custom_font: Font to use for the dialog
        files_to_move: List of (source, destination) tuples
        skipped_files: List of files that will be skipped
        on_confirm_callback: Function to call when user confirms
    """
    # Create confirmation dialog
    confirmation_dialog = tk.Toplevel(parent_window)
    confirmation_dialog.title("Confirm File Organization")
    confirmation_dialog.geometry("1200x500")  # Increased height from 400 to 500
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
    def confirm():
        confirmation_dialog.destroy()
        on_confirm_callback()
    
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
             command=confirm,
             font=custom_font,
             bg=Config.COLORS["SECONDARY_BACKGROUND"],
             fg="#006400",
             padx=Config.STYLES["WIDGET_PADDING"],
             pady=Config.STYLES["WIDGET_PADDING"]).pack(side="right", fill="x", expand=True, padx=(5, 0))
    
    # Wait for the dialog to close
    parent_window.wait_window(confirmation_dialog)
