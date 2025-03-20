"""
Styles management for the application.
Provides functionality for configuring and updating UI styles.
"""

from config import Config
import tkinter as tk
from tkinter import ttk, font

# Add a global reference for the custom font
_app_custom_font = None

def set_custom_font(custom_font):
    """Store a reference to the application's custom font for use in styling functions."""
    global _app_custom_font
    _app_custom_font = custom_font

def style_button(button, is_danger=False, is_process_button=False):
    """Apply standard button styling with optional variants.
    
    Args:
        button: The tk.Button to style
        is_danger: If True, apply danger (red) text color
        is_process_button: If True, apply process button styling
    """
    button.configure(
        font=_app_custom_font,
        bg=Config.COLORS["SECONDARY_BACKGROUND"],
        fg="#990000" if is_danger else Config.COLORS["TEXT"],
        padx=Config.STYLES["WIDGET_PADDING"],
        pady=Config.STYLES["WIDGET_PADDING"]
    )

def style_entry(entry, font_size=None):
    """Apply standard entry field styling.
    
    Args:
        entry: The tk.Entry to style
        font_size: Optional custom font size
    """
    font_to_use = ('Consolas', font_size if font_size else Config.FONTS["TABLE_SIZE"])
    entry.configure(
        font=font_to_use,
        bg=Config.COLORS["SECONDARY_BACKGROUND"],
        fg=Config.COLORS["TEXT"],
        insertbackground=Config.COLORS["TEXT"]
    )

def style_label(label, use_smaller_font=False):
    """Apply standard label styling.
    
    Args:
        label: The tk.Label to style
        use_smaller_font: If True, use a font size 1pt smaller than default
    """
    if _app_custom_font:
        current_font = font.nametofont(_app_custom_font.name)
        current_size = current_font.cget("size")
        font_size = current_size - 1 if use_smaller_font else current_size
        font_family = current_font.cget("family")
    else:
        font_family = Config.STYLES["CUSTOM_FONT"]["FAMILY"]
        font_size = Config.FONTS["DEFAULT_SIZE"] - (1 if use_smaller_font else 0)
        
    label.configure(
        font=(font_family, font_size),
        bg=Config.COLORS["BACKGROUND"],
        fg=Config.COLORS["TEXT"],
        bd=0
    )

def style_text_widget(text_widget):
    """Apply standard text widget styling.
    
    Args:
        text_widget: The tk.Text widget to style
    """
    text_widget.configure(
        font=('Consolas', Config.FONTS["TABLE_SIZE"]),
        bg=Config.COLORS["SECONDARY_BACKGROUND"],
        fg=Config.COLORS["TEXT"],
        insertbackground=Config.COLORS["TEXT"]
    )

def style_checkbutton(checkbutton):
    """Apply standard checkbutton styling.
    
    Args:
        checkbutton: The tk.Checkbutton to style
    """
    checkbutton.configure(
        font=_app_custom_font,
        bg=Config.COLORS["BACKGROUND"],
        fg=Config.COLORS["TEXT"],
        selectcolor=Config.COLORS["SECONDARY_BACKGROUND"],
        activebackground=Config.COLORS["BACKGROUND"],
        activeforeground=Config.COLORS["TEXT"]
    )

def configure_context_menu(menu):
    """Apply styling to a context menu.
    
    Args:
        menu: The tk.Menu to style
    """
    menu.configure(
        bg=Config.COLORS["SECONDARY_BACKGROUND"],
        fg=Config.COLORS["TEXT"],
        activebackground=Config.COLORS["BACKGROUND"],
        activeforeground=Config.COLORS["TEXT"],
        tearoff=0
    )

def configure_text_tags(text_widget):
    """Configure standard text tags for a text widget.
    
    Args:
        text_widget: The tk.Text widget to configure tags for
    """
    text_widget.tag_config("ok", foreground="#006400")  # Dark green
    text_widget.tag_config("nok", foreground="#8B0000")  # Dark red
    text_widget.tag_config("api_call", foreground="#0000CD")  # Medium blue

def configure_styles(style, custom_font):
    """Configure all ttk styles for the application.
    
    Args:
        style: The ttk.Style object to configure
        custom_font: The custom font object to use
    """
    # Store the custom font for use in other functions
    set_custom_font(custom_font)
    
    # Set theme
    style.theme_use(Config.STYLES["THEME"])
    
    # Configure dark theme styles
    style.configure('Dark.TPanedwindow', background=Config.COLORS["BACKGROUND"], sashwidth=0)
    style.configure('TFrame', background=Config.COLORS["BACKGROUND"])
    style.configure('TButton', padding=Config.STYLES["WIDGET_PADDING"], font=custom_font, 
                   background=Config.COLORS["SECONDARY_BACKGROUND"], foreground=Config.COLORS["TEXT"], 
                   relief="solid", borderwidth=1)
    style.map('TButton',
        relief=[('pressed', 'sunken'), ('!pressed', 'solid')],
        borderwidth=[('pressed', 1), ('!pressed', 1)])
    style.configure('TEntry', padding=Config.STYLES["WIDGET_PADDING"], 
                   fieldbackground=Config.COLORS["SECONDARY_BACKGROUND"], foreground=Config.COLORS["TEXT"])
    style.configure('TLabel', background=Config.COLORS["BACKGROUND"], foreground=Config.COLORS["TEXT"], font=custom_font)
    style.configure('TText', padding=Config.STYLES["WIDGET_PADDING"], 
                   background=Config.COLORS["SECONDARY_BACKGROUND"], foreground=Config.COLORS["TEXT"])
    
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

    # Configure table borders and remove extra spacing
    style.layout("Treeview", [
        ('Treeview.treearea', {'sticky': 'nswe'})
    ])
    
    style.layout("Treeview.Heading", [
        ("Treeview.Heading.cell", {'sticky': 'nswe'}),
        ("Treeview.Heading.padding", {'sticky': 'nswe', 'children': [
            ("Treeview.Heading.image", {'side': 'right', 'sticky': ''}),
            ("Treeview.Heading.text", {'sticky': 'we'})
        ]})
    ])

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

def update_progress_bar_style(style, progress, bar_type="file"):
    """Update progress bar color based on progress value.
    
    Args:
        style: The ttk.Style object
        progress: For file progress, 0-100 percentage. For API, number of used calls.
        bar_type: Either "file" or "api"
    """
    # Calculate color based on progress
    if progress < 50:
        # Green to Yellow (mix more yellow as progress increases)
        green = 255
        red = int((progress / 50) * 255)
    else:
        # Yellow to Red (reduce green as progress increases)
        red = 255
        green = int(((100 - progress) / 50) * 255)
    
    color = f'#{red:02x}{green:02x}00'
    
    # Apply color to appropriate style
    if bar_type == "file":
        style.configure("Gradient.Horizontal.TProgressbar", background=color)
    else:
        style.configure("API.Horizontal.TProgressbar", background=color)

def set_api_entry_style(entry, is_valid):
    """Set the style of the API entry based on validity.
    
    Args:
        entry: The tk.Entry for API key
        is_valid: Boolean indicating if the API key is valid
    """
    if is_valid:
        entry.configure(bg=Config.COLORS["VALID_ENTRY"])
    else:
        entry.configure(bg=Config.COLORS["INVALID_ENTRY"])

def configure_table_columns(table, columns, column_widths, hide_columns=None):
    """Configure table columns with proper widths and visibility.
    
    Args:
        table: The ttk.Treeview widget
        columns: List of column names
        column_widths: Dictionary mapping column names to widths
        hide_columns: List of column names to hide
    """
    hide_columns = hide_columns or []
    for col in columns:
        table.column(col, anchor="center", width=column_widths.get(col, 100))
        
        # Hide specified columns
        if col in hide_columns:
            table.column(col, width=0, minwidth=0, stretch=False)

def configure_table_tags(table):
    """Configure tags for table rows.
    
    Args:
        table: The ttk.Treeview widget
    """
    table.tag_configure('oddrow', background=Config.COLORS["BACKGROUND"])
    table.tag_configure('evenrow', background=Config.COLORS["SECONDARY_BACKGROUND"])
    table.tag_configure('hidden', foreground='#FFFFFF', background='#FFFFFF')  # Make text invisible
    table.tag_configure("updated", background=Config.COLORS["UPDATED_ROW"])
    table.tag_configure("failed", background=Config.COLORS["FAILED_ROW"])

def create_styled_button(parent, text, command, is_danger=False):
    """Create and return a styled button.
    
    Args:
        parent: The parent widget
        text: Button text
        command: Button command
        is_danger: If True, apply danger (red) text color
    
    Returns:
        The styled tk.Button
    """
    button = tk.Button(parent, text=text, command=command)
    style_button(button, is_danger=is_danger)
    return button

def create_styled_entry(parent, textvariable=None, width=None, justify=None):
    """Create and return a styled entry field.
    
    Args:
        parent: The parent widget
        textvariable: Optional StringVar to connect
        width: Optional width
        justify: Optional text justification
    
    Returns:
        The styled tk.Entry
    """
    entry = tk.Entry(parent, textvariable=textvariable, width=width, justify=justify)
    style_entry(entry)
    return entry

def create_styled_text(parent, width=None, height=None, state="normal", wrap="word"):
    """Create and return a styled text widget.
    
    Args:
        parent: The parent widget
        width: Optional width
        height: Optional height
        state: Initial state ("normal" or "disabled")
        wrap: Text wrapping mode
    
    Returns:
        The styled tk.Text
    """
    text = tk.Text(parent, width=width, height=height, state=state, wrap=wrap)
    style_text_widget(text)
    return text

def create_button_pair(container, button1_text, button1_command, button2_text, button2_command):
    """Create two styled buttons side by side in a container.
    
    Args:
        container: The parent container
        button1_text: Text for first button
        button1_command: Command for first button
        button2_text: Text for second button
        button2_command: Command for second button
    
    Returns:
        Tuple of the two created buttons
    """
    button1 = create_styled_button(container, button1_text, button1_command)
    button1.pack(side="left", fill="x", expand=True, padx=(0, 2))
    
    button2 = create_styled_button(container, button2_text, button2_command)
    button2.pack(side="left", fill="x", expand=True, padx=(2, 0))
    
    return button1, button2
