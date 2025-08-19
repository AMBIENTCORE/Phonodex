"""
Logging utilities for the application.
Provides functionality for logging messages to UI text widgets or console.
"""

import tkinter as tk

class Logger:
    """
    Logger class that handles sending messages to UI text widgets with fallback to console.
    Supports different log types and special formatting for success/failure messages.
    """
    def __init__(self):
        """Initialize the logger with empty widget references."""
        self.debug_widget = None
        self.processing_widget = None
        
    def set_debug_widget(self, widget):
        """Set the widget for debug messages."""
        self.debug_widget = widget
        
    def set_processing_widget(self, widget):
        """Set the widget for processing messages."""
        self.processing_widget = widget
    
    def clear_logs(self, app=None, debug_scrollbar=None, processing_scrollbar=None):
        """Clear both log widgets and reset their scrollbars."""
        # Clear processing log box if it exists
        if self.processing_widget:
            self.processing_widget.configure(state="normal")
            self.processing_widget.delete("1.0", "end")
            self.processing_widget.configure(state="disabled")
            
            # Force scrollbar update for processing log
            if processing_scrollbar:
                self.processing_widget.yview_moveto(0)
                autohide_scrollbar(processing_scrollbar, 0, 1)
        
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
        Log messages in the appropriate text box based on type.
        
        Args:
            message: The message text to log
            log_type: Either "debug" (for technical messages) or "processing" (for operation results)
                - "debug" messages will appear in the debug widget (technical information)
                - "processing" messages will appear in the processing widget (success/failure results)
        """
        # Handle the case when UI elements aren't defined yet (early startup)
        if log_type == "debug" and self.debug_widget is None:
            print(f"Early log: {message}")
            return
            
        # Use debug widget as fallback if processing widget isn't defined yet
        if log_type == "processing" and self.processing_widget is None:
            if self.debug_widget is None:
                print(f"Early log: {message}")
                return
            target_widget = self.debug_widget
        else:
            # Determine which widget to use based on message type
            if (message.startswith("[OK]") or message.startswith("[NOK]") or 
                message.startswith("[INFO] API Calls:")):
                # Only OK/NOK messages and API counter go to processing widget
                target_widget = self.processing_widget
            else:
                # Everything else goes to debug widget
                target_widget = self.debug_widget
            
        target_widget.configure(state="normal")
        
        # Special handling for OK/NOK tags in processing messages
        if message.startswith("[OK]") and target_widget == self.processing_widget:
            target_widget.insert("end", "[OK] ", "ok")
            target_widget.insert("end", message[4:] + "\n")
        elif message.startswith("[NOK]") and target_widget == self.processing_widget:
            target_widget.insert("end", "[NOK] ", "nok")
            target_widget.insert("end", message[5:] + "\n")
        elif message.startswith("[INFO] API Calls:"):
            target_widget.insert("end", message + "\n", "api_call")
        else:
            target_widget.insert("end", message + "\n")
            
        target_widget.configure(state="disabled")
        target_widget.see("end")  # Auto-scroll to the latest message


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
