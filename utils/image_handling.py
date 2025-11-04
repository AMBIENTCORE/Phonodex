"""
Image handling utilities for the application.
Provides functionality for image processing, conversion, and clipboard operations.
"""

import io
import os
from PIL import Image, ImageTk, ImageGrab
import win32clipboard
import win32con
from utils.logging import log_message
from utils.file_operations import resource_path

# Global variable to store the original image data for internal copy-paste
# This allows us to bypass clipboard compression/decompression entirely
_original_image_data = None

def get_image_from_clipboard():
    """
    Retrieves an image from the system clipboard.
    
    Returns:
        bytes: Image data in bytes if an image is on the clipboard, None otherwise
    """
    global _original_image_data
    
    # PRIMARY METHOD: Use cached original image data if available (lossless)
    if _original_image_data is not None:
        log_message("[INFO] Using cached original image data (lossless transfer)")
        data = _original_image_data
        _original_image_data = None  # Clear after using once
        return data
        
    # FALLBACK: Handle external clipboard data if no cached image is available
    try:
        win32clipboard.OpenClipboard()
        
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
            data = win32clipboard.GetClipboardData(win32con.CF_DIB)
            
            # Process clipboard bitmap data
            stream = io.BytesIO(data)
            img = Image.open(stream)
            
            # Convert to JPEG format for album art
            output = io.BytesIO()
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img.save(output, format='JPEG')
            image_data = output.getvalue()
            
            log_message("[INFO] Retrieved external image from clipboard (converted to JPEG)")
            return image_data
        else:
            log_message("[INFO] No image data found on clipboard")
            return None
    except Exception as e:
        log_message(f"[ERROR] Failed to get image from clipboard: {str(e)}")
        return None
    finally:
        win32clipboard.CloseClipboard()

def copy_image_to_clipboard(image_data):
    """
    Copies an image to the system clipboard.
    
    Args:
        image_data: Image data in bytes
        
    Returns:
        bool: True if successful, False otherwise
    """
    global _original_image_data
    
    # PRIMARY METHOD: Store original image bytes for direct internal paste (lossless)
    _original_image_data = image_data
    log_message("[INFO] Stored original image bytes for lossless internal transfers")
    
    # FALLBACK: Also place on system clipboard for external applications
    clipboard_opened = False
    try:
        # Create an image from the data
        img = Image.open(io.BytesIO(image_data))
        
        # Convert to bitmap format for Windows clipboard
        output = io.BytesIO()
        img.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]  # The file header for BMP is 14 bytes
        
        # Place on system clipboard
        win32clipboard.OpenClipboard()
        clipboard_opened = True
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
        
        log_message("[INFO] Also placed image on system clipboard for external apps")
        return True
    except Exception as e:
        log_message(f"[ERROR] Failed to copy image to system clipboard: {str(e)}")
        # Even if system clipboard fails, internal transfers will still work
        return _original_image_data is not None
    finally:
        if clipboard_opened:
            win32clipboard.CloseClipboard()

def resize_image(image_data, size=(240, 240)):
    """
    Resize an image to the specified dimensions.
    
    Args:
        image_data: Image data in bytes
        size: Tuple of (width, height)
        
    Returns:
        bytes: Resized image data
    """
    try:
        # Create an image from the data
        img = Image.open(io.BytesIO(image_data))
        
        # Resize the image
        resized_img = img.resize(size, Image.Resampling.LANCZOS)
        
        # Convert back to bytes
        output = io.BytesIO()
        # Determine format to use - default to JPEG for album art
        if img.format and img.format.lower() in ('png', 'gif') and has_alpha(resized_img):
            # Keep PNG only if it has transparency
            resized_img.save(output, format=img.format)
        else:
            # Convert to RGB if needed
            if resized_img.mode == 'RGBA':
                resized_img = resized_img.convert('RGB')
            # Use JPEG for all other images (better for album art)
            resized_img.save(output, format='JPEG')
        
        return output.getvalue()
    except Exception as e:
        log_message(f"[ERROR] Failed to resize image: {str(e)}")
        return image_data  # Return original on failure

def has_alpha(img):
    """Check if an image has an alpha channel that's in use."""
    if img.mode == 'RGBA':
        # Check if image actually uses the alpha channel
        return img.split()[3].getextrema()[0] < 255
    return False

def create_photo_image(image_data, size=(240, 240)):
    """
    Create a PhotoImage object from image data for display in tkinter.
    
    Args:
        image_data: Image data in bytes
        size: Tuple of (width, height) for resizing
        
    Returns:
        PhotoImage: A PhotoImage object for display in tkinter
    """
    try:
        # Open the image from bytes
        img = Image.open(io.BytesIO(image_data))
        
        # Resize the image
        img = img.resize(size, Image.Resampling.LANCZOS)
        
        # Create a PhotoImage
        photo = ImageTk.PhotoImage(img)
        return photo
    except Exception as e:
        log_message(f"[ERROR] Failed to create PhotoImage: {str(e)}")
        return None

def load_default_album_art(default_image_path, label=None, size=(240, 240)):
    """
    Load the default album art image.
    
    Args:
        default_image_path: Path to the default image
        label: Optional tkinter Label widget to update with the image
        size: Tuple of (width, height) for the image
        
    Returns:
        PhotoImage: The loaded image as a PhotoImage object
    """
    try:
        # Try to load the placeholder image from resources
        placeholder_path = resource_path(default_image_path)
        if os.path.exists(placeholder_path):
            img = Image.open(placeholder_path)
            img = img.resize(size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            # Update the label if provided
            if label:
                label.configure(image=photo)
                label.image = photo  # Keep a reference!
                
            return photo
        else:
            log_message(f"[WARNING] Default album art not found at {placeholder_path}")
            if label:
                label.configure(image='')
            return None
    except Exception as e:
        log_message(f"[ERROR] Failed to load default album art: {str(e)}")
        if label:
            label.configure(image='')
        return None

def update_album_art_display(image_data, label, size=240, load_default_func=None):
    """
    Update the album art display with the provided image data.
    
    Args:
        image_data: Image data in bytes
        label: The tkinter Label widget to update
        size: Size of the album art (square)
        load_default_func: Optional function to call if loading fails
        
    Returns:
        PhotoImage: The created PhotoImage object, or None if failed
    """
    try:
        log_message(f"[COVER] Processing image data: {len(image_data)} bytes")
        
        # Open the image data
        img_buffer = io.BytesIO(image_data)
        img = Image.open(img_buffer)
        log_message(f"[COVER] Image opened successfully: {img.format}, {img.size}, {img.mode}")
        
        # Instead of thumbnail which may leave empty space, we'll resize with padding
        # to ensure the image fills the entire space while maintaining aspect ratio
        
        # Calculate the scaling factor to fill the container
        width, height = img.size
        width_ratio = size / width
        height_ratio = size / height
        
        # Use the larger ratio to ensure the image fills the space
        ratio = max(width_ratio, height_ratio)
        
        # Calculate new dimensions
        new_width = round(width * ratio)
        new_height = round(height * ratio)
        
        # Resize the image (will be larger than container in one dimension)
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # If image is larger than container, crop it to center
        if new_width > size or new_height > size:
            left = (new_width - size) // 2
            top = (new_height - size) // 2
            right = left + size
            bottom = top + size
            img = img.crop((left, top, right, bottom))
        
        log_message(f"[COVER] Image resized and cropped to fill {size}x{size}")
        
        # Create a PhotoImage object
        try:
            photo = ImageTk.PhotoImage(img)
            log_message(f"[COVER] PhotoImage created successfully")
            
            # Update the label
            label.configure(image=photo)
            log_message(f"[COVER] Album cover label updated with new image")
            
            return photo
        except Exception as e:
            log_message(f"[COVER] Failed to create or apply PhotoImage: {e}")
            return None
        
    except Exception as e:
        log_message(f"[COVER] Failed to update album art display: {e}")
        # Load default image if we can't display the provided image
        if load_default_func:
            default_result = load_default_func()
            log_message(f"[COVER] Loaded default album art: {default_result}")
            return default_result
        return None

def paste_image_from_clipboard():
    """
    Paste image from clipboard and return it as bytes.
    
    Returns:
        bytes: Image data in bytes if successful, None otherwise
    """
    global _original_image_data
    
    # PRIMARY METHOD: Use cached original image data if available (lossless)
    if _original_image_data is not None:
        log_message("[COVER] Using cached original image data (lossless transfer)", log_type="processing")
        data = _original_image_data
        _original_image_data = None  # Clear after using once
        return data
    
    # FALLBACK: Handle external clipboard data
    try:
        img = ImageGrab.grabclipboard()
        if img is None:
            log_message("[COVER] No image found in clipboard", log_type="processing")
            return None
        
        if not isinstance(img, Image.Image):
            log_message("[COVER] Clipboard content is not an image", log_type="processing")
            return None
        
        # Process external clipboard image
        if img.mode == 'RGBA':
            img = img.convert('RGB')
            
        img_buffer = io.BytesIO()
        img.save(img_buffer, format='JPEG')
        image_data = img_buffer.getvalue()
        
        log_message("[COVER] Retrieved external image from clipboard (converted to JPEG)", log_type="processing")
        return image_data
    except Exception as e:
        log_message(f"[ERROR] Failed to paste image from clipboard: {str(e)}")
        return None

def extract_album_art_from_file(file_path, audio_file=None):
    """
    Extract album art from an audio file.
    
    Args:
        file_path: Path to the audio file
        audio_file: Optional pre-loaded audio file object
        
    Returns:
        bytes: Image data in bytes if found, None otherwise
    """
    try:
        import mutagen
        from mutagen.mp3 import MP3
        from mutagen.flac import FLAC
        from mutagen.mp4 import MP4
        from mutagen.oggvorbis import OggVorbis
        from mutagen.asf import ASF
        
        # If audio_file is not provided, load it
        if audio_file is None:
            # Get the file extension
            if not file_path:
                log_message(f"[ERROR] Invalid file path: empty or None")
                return None
                
            ext = os.path.splitext(file_path)[1].lower()
            if not ext:
                log_message(f"[ERROR] File has no extension: {file_path}")
                return None
            
            # Use appropriate handler based on file type
            if ext == '.mp3':
                audio_file = MP3(file_path)
            elif ext == '.flac':
                audio_file = FLAC(file_path)
            elif ext in ['.m4a', '.mp4']:
                audio_file = MP4(file_path)
            elif ext == '.ogg':
                audio_file = OggVorbis(file_path)
            elif ext == '.wma':
                audio_file = ASF(file_path)
            else:
                log_message(f"[ERROR] Unsupported file type: {ext} for {os.path.basename(file_path)}")
                return None
        
        # Extract album art based on file type
        if isinstance(audio_file, MP3):
            # MP3 files use ID3 tags
            if audio_file.tags:
                for tag in audio_file.tags.values():
                    if tag.FrameID == 'APIC':
                        return tag.data
        
        elif isinstance(audio_file, FLAC):
            # FLAC files store pictures directly
            if audio_file.pictures:
                return audio_file.pictures[0].data
        
        elif isinstance(audio_file, MP4):
            # MP4 files use 'covr' atom
            if 'covr' in audio_file:
                return audio_file['covr'][0]
        
        elif isinstance(audio_file, OggVorbis):
            # Ogg files might have METADATA_BLOCK_PICTURE
            if 'metadata_block_picture' in audio_file:
                import base64
                data = base64.b64decode(audio_file['metadata_block_picture'][0])
                # Skip the header to get just the image data
                # This is a simplification - proper implementation would parse the header
                return data[32:]  # Skip the FLAC picture header
        
        elif isinstance(audio_file, ASF):
            # WMA files use WM/Picture
            if 'WM/Picture' in audio_file:
                # The format is complex, this is a simplification
                picture_data = audio_file['WM/Picture'][0].value
                # Skip the ASF picture header to get just the image data
                # This is a simplification - proper implementation would parse the header
                return picture_data[50:]  # Skip the ASF picture header
        
        log_message(f"[INFO] No album art found in file: {os.path.basename(file_path)}")
        return None
    
    except Exception as e:
        log_message(f"[ERROR] Failed to extract album art: {str(e)}")
        return None
