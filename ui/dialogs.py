"""
Dialog windows for the Phonodex application.
"""

import tkinter as tk
from tkinter import ttk, StringVar, messagebox
from config import Config, folder_format, DEFAULT_FOLDER_FORMAT, save_settings
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
