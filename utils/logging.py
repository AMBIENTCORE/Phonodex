"""
Logging utilities for the application.
Provides functionality for logging messages to UI text widgets or console.
"""

import tkinter as tk

class Logger:
    """
    Logger class that handles sending messages to the debug log widget with
    fallback to the console when the widget isn't available yet.
    """
    def __init__(self):
        """Initialize the logger with an empty widget reference."""
        self.debug_widget = None
        
    def set_debug_widget(self, widget):
        """Set the widget for log messages."""
        self.debug_widget = widget
    
    def clear_logs(self, app=None, debug_scrollbar=None):
        """Clear the debug log box and reset its scrollbar."""
        # Clear debug log box if it exists
        if self.debug_widget:
            self.debug_widget.configure(state="normal")
            self.debug_widget.delete("1.0", "end")
            self.debug_widget.configure(state="disabled")
            
            # Force scrollbar update for debug log
            if debug_scrollbar:
                self.debug_widget.yview_moveto(0)
                autohide_scrollbar(debug_scrollbar, 0, 1)
            
        # Update the UI if app is provided
        if app:
            app.update_idletasks()
    
    def log(self, message, log_type="debug"):
        """
        Log a message to the debug log widget.
        
        Args:
            message: The message text to log
            log_type: Retained for backward compatibility; all messages now go
                to the single debug log widget.
        """
        # Handle the case when the UI isn't ready yet (early startup)
        if self.debug_widget is None:
            print(f"Early log: {message}")
            return

        self.debug_widget.configure(state="normal")
        self.debug_widget.insert("end", message + "\n")
        self.debug_widget.configure(state="disabled")
        self.debug_widget.see("end")  # Auto-scroll to the latest message


def autohide_scrollbar(scrollbar, first, last):
    """
    Hide scrollbar if not needed, show if needed.
    Used by text widgets to automatically hide/show scrollbars.
    """
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
        print(f"[ERROR] Scrollbar error: {str(e)}")  # Use print to avoid potential recursion


# Create a global logger instance for the application
logger = Logger()

# Function for backward compatibility
def log_message(message, log_type="debug"):
    """
    Compatibility function to maintain backward compatibility with existing code.
    Forwards to the global logger instance.
    """
    logger.log(message, log_type)
