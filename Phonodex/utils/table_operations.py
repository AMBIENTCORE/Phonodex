import os
from config import Config

def auto_adjust_column_widths(file_table, columns):
    """Calculate and set optimal column widths based on content.
    
    Args:
        file_table: The ttk.Treeview widget to adjust
        columns: List of column names in the table
    """
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
    column_widths["File Path"] = 0  # Always keep File Path column hidden
    
    # Apply the calculated widths
    for col in columns:
        file_table.column(col, width=column_widths[col])
        
        # Ensure File Path column stays completely hidden
        if col == "File Path":
            file_table.column(col, width=0, minwidth=0, stretch=False)

def treeview_sort_column(tv, col, reverse, columns):
    """Sort treeview content when a column header is clicked.
    
    Args:
        tv: The ttk.Treeview widget to sort
        col: The column to sort by
        reverse: Whether to reverse the sort order
        columns: List of column names in the table
    """
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
    
    # Update header arrows
    for column in columns:
        if column == col:
            tv.heading(column, text=f"{column} {'↓' if reverse else '↑'}")
        else:
            tv.heading(column, text=column)
    
    return not reverse  # Return new sort order 

def select_all_visible(table, count_var, filter_text=''):
    """Select all visible items in the table.
    
    Args:
        table: The Treeview widget
        count_var: StringVar to update with selection count
        filter_text: Current filter text (optional)
    """
    # Deselect all items first
    table.selection_remove(table.selection())
    
    # Get all visible items
    visible_items = []
    for item in table.get_children():
        # If there's a filter, check if the item should be visible
        if filter_text:
            values = table.item(item)['values']
            # Convert all values to strings and check if any contain the filter text
            if any(filter_text.lower() in str(value).lower() for value in values):
                visible_items.append(item)
        else:
            visible_items.append(item)
    
    # Select all visible items
    if visible_items:
        table.selection_add(visible_items)
    
    # Update the count display
    count_var.set(f"{len(table.selection())}/{len(table.get_children())}")

def file_table_selection_callback(table, count_var):
    """Update the file count when selection changes.
    
    Args:
        table: The ttk.Treeview widget
        count_var: The StringVar to update with the count
    """
    selected_count = len(table.selection())
    total_count = len(table.get_children())
    count_var.set(f"{selected_count}/{total_count}")

def update_table(file_table, apply_filter_func, file_count_var, columns):
    """Update the table with current file list and metadata.
    
    Args:
        file_table: The ttk.Treeview widget to update
        apply_filter_func: Function to apply the current filter
        file_count_var: StringVar for updating file count display
        columns: List of column names
    """
    # Clear the current table
    file_table.delete(*file_table.get_children())
    
    # Apply the current filter to show the correct items
    apply_filter_func()
    
    # Update file count label to show total files - use actual table items
    selected_count = len(file_table.selection())
    total_count = len(file_table.get_children())  # Count actual visible items
    file_count_var.set(f"{selected_count}/{total_count}")
    
    # Auto-adjust column widths after updating the table
    auto_adjust_column_widths(file_table, columns) 

def apply_filter(file_table, filter_text, file_list, file_metadata_cache, get_audio_file, get_tag_value, updated_files, processed_files, file_count_var, columns):
    """Filter table contents based on filter text.
    
    Args:
        file_table: The ttk.Treeview widget
        filter_text: Text to filter by (lowercase)
        file_list: List of files to display
        file_metadata_cache: Cache of file metadata
        get_audio_file: Function to get audio file object
        get_tag_value: Function to get tag value from audio file
        updated_files: Set of updated file paths
        processed_files: Set of processed file paths
        file_count_var: StringVar for count display
        columns: List of column names
    """
    # Clear the current table
    file_table.delete(*file_table.get_children())
    
    # Repopulate with filtered items in the same order as file_list
    for idx, file_path in enumerate(file_list):
        # Skip files that no longer exist
        if not os.path.exists(file_path):
            continue
            
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
                metadata.get("genre", ""),
                file_path  # Add file_path as the last value
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
                item = file_table.insert("", "end", values=["Error", "", "", "", "", "", "", "", ""])
                file_table.tag_configure("failed", background=Config.COLORS["FAILED_ROW"])
                file_table.item(item, tags=("failed",))
    
    # Update file count label
    selected_count = len(file_table.selection())
    total_count = len(file_table.get_children())  # Count actual visible items
    file_count_var.set(f"{selected_count}/{total_count}")
    
    # Auto-adjust column widths after filtering
    auto_adjust_column_widths(file_table, columns) 

def remove_selected_items(file_table, file_list, file_metadata_cache, processed_files, updated_files, file_count_var, log_message):
    """Remove selected items from the file list and update related data structures.
    
    Args:
        file_table: The ttk.Treeview widget
        file_list: List of files to maintain
        file_metadata_cache: Cache of file metadata
        processed_files: Set of processed file paths
        updated_files: Set of updated file paths
        file_count_var: StringVar for count display
        log_message: Function to log messages
    """
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
    
    log_message(f"[INFO] Removed {len(items_to_remove)} items from the list") 