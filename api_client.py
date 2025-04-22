"""
API client module for handling external API requests and rate limiting.
"""

import requests
import time
import os
from utils.logging import log_message
from config import Config

# API rate limiting
rate_limit_total = Config.API["RATE_LIMIT"]
rate_limit_used = 0
rate_limit_remaining = rate_limit_total
first_request_time = 0  # Track when the first request was made in the current window

def make_api_request(url, params, max_retries=3, retry_delay=2):
    """Make an API request with retries.
    
    Args:
        url: The API endpoint URL
        params: Dictionary of query parameters
        max_retries: Maximum number of retry attempts
        retry_delay: Delay between retries in seconds
        
    Returns:
        tuple: (JSON response data, response headers) or (None, None) if request failed
    """
    attempts = 0
    while attempts < max_retries:
        try:
            response = requests.get(url, params=params, timeout=10)
            
            if response.status_code == 429:  # Too Many Requests
                retry_after = int(response.headers.get('Retry-After', retry_delay))
                log_message(f"[API] Rate limit exceeded. Waiting {retry_after} seconds before retry.")
                time.sleep(retry_after)
                attempts += 1
                continue
                
            response.raise_for_status()  # Raise exception for 4xx/5xx status codes
            return response.json(), response.headers
            
        except requests.exceptions.RequestException as e:
            log_message(f"[ERROR] API request failed: {str(e)}")
            attempts += 1
            if attempts < max_retries:
                log_message(f"[API] Retrying in {retry_delay} seconds ({attempts}/{max_retries})...")
                time.sleep(retry_delay)
            else:
                log_message(f"[ERROR] API request failed after {max_retries} attempts.")
                break
    
    return None, None

def update_api_progress(state=None, verbose=False, progress_callback=None):
    """Update API progress based on rate limit usage
    
    Args:
        state: Optional state parameter ("start", "complete", or None)
        verbose: Whether to output detailed debug messages
        progress_callback: Function to call with progress updates (for UI)
    """
    global rate_limit_remaining, rate_limit_used, rate_limit_total, first_request_time
    
    # Debug print (only if verbose)
    if verbose:
        print(f"DEBUG: update_api_progress called - Total: {rate_limit_total}, Remaining: {rate_limit_remaining}, State: {state}")
    
    # Handle special states
    if state == "start":
        # When starting an API call, show a temporary visual indicator
        # without modifying the global counters
        temp_used = rate_limit_used + 1
        
        if verbose:
            print(f"DEBUG: Starting API call - Showing temporary count: {temp_used}")
            
        if progress_callback:
            progress_callback(temp_used, "api", verbose=verbose)
        return
    elif state == "complete":
        # The actual values should have been set by update_rate_limits_from_headers
        # We just need to update the progress bar with current values
        if verbose:
            print(f"DEBUG: API call complete - Setting progress bar to {rate_limit_used} out of {rate_limit_total}")
        if progress_callback:
            progress_callback(rate_limit_used, "api", verbose=verbose)
        return
    
    # Only update the window time check if this is a direct call
    if state is None:
        # If no requests in 60 seconds, reset the window
        current_time = time.time()
        if current_time - first_request_time > Config.API_RATE_LIMIT_WAIT:
            rate_limit_remaining = rate_limit_total
            rate_limit_used = 0
        
        # Update progress bar to show used requests
        if verbose:
            print(f"DEBUG: Setting API progress bar to {rate_limit_used} out of {rate_limit_total}")
        if progress_callback:
            progress_callback(rate_limit_used, "api", verbose=verbose)

def enforce_api_limit(app_update_callback=None):
    """Ensure we do not exceed API rate limit.
    
    Args:
        app_update_callback: Optional function to update the UI during waits
        
    Returns:
        bool: True if the API call should proceed, False if it should be blocked.
    """
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
        
        if app_update_callback:
            app_update_callback()  # Update UI if callback provided
            
        time.sleep(wait_time)
        # Reset counters after waiting
        first_request_time = time.time()
        rate_limit_remaining = rate_limit_total
        rate_limit_used = 0
        log_message(Config.MESSAGES["API_RESUMING"], log_type="debug")
    
    if app_update_callback:
        app_update_callback()  # Final UI update
    
    # Return True if we can proceed with the API call
    return rate_limit_remaining > 0

def update_rate_limits_from_headers(headers, update_progress=True, verbose=False, progress_callback=None):
    """Update rate limit tracking based on API response headers.
    
    Args:
        headers: API response headers dictionary
        update_progress: Whether to update the API progress bar
        verbose: Whether to output detailed debug messages
        progress_callback: Optional callback for updating progress UI
    """
    global rate_limit_total, rate_limit_used, rate_limit_remaining, first_request_time
    
    # Debug print
    if verbose:
        print("DEBUG: update_rate_limits_from_headers called with headers:", headers)
    
    # Make sure headers is a dictionary-like object
    if not headers or not hasattr(headers, 'get'):
        if verbose:
            print("DEBUG: No valid headers provided to update_rate_limits_from_headers")
        # Just increment used count as a failsafe
        rate_limit_used += 1
        rate_limit_remaining = max(0, rate_limit_remaining - 1)
        # Update the progress indicator if requested
        if update_progress and progress_callback:
            update_api_progress(verbose=verbose, progress_callback=progress_callback)
        return
    
    # Update rate limit info from headers
    if 'X-Discogs-Ratelimit' in headers:
        # Use the exact values from API headers - they are the authoritative source
        rate_limit_total = int(headers.get('X-Discogs-Ratelimit', 60))
        rate_limit_used = int(headers.get('X-Discogs-Ratelimit-Used', 0))
        rate_limit_remaining = int(headers.get('X-Discogs-Ratelimit-Remaining', rate_limit_total - rate_limit_used))
        
        # Debug print
        if verbose:
            print(f"DEBUG: Rate limits updated from API headers - Used: {rate_limit_used}, Total: {rate_limit_total}, Remaining: {rate_limit_remaining}")
        
        # If this is the first request in a new window
        if first_request_time == 0:
            first_request_time = time.time()
        
        # Update progress bar
        if update_progress and progress_callback:
            update_api_progress("complete", verbose=verbose, progress_callback=progress_callback)  # Changed from just update_api_progress() to indicate completion
            log_message(f"[INFO] API Calls: {rate_limit_used}/{rate_limit_total} (Remaining: {rate_limit_remaining})")
    else:
        # If headers don't contain rate limit info, just increment the used count
        if verbose:
            print("DEBUG: Headers did not contain X-Discogs-Ratelimit")
        rate_limit_used += 1
        rate_limit_remaining = max(0, rate_limit_remaining - 1)
        # Update the progress indicator if requested
        if update_progress and progress_callback:
            update_api_progress(verbose=verbose, progress_callback=progress_callback)

def update_api_entry_style(is_valid, api_entry=None):
    """Update the validation style of the API key entry field.
    
    Args:
        is_valid: Boolean indicating if the API key is valid
        api_entry: The tkinter entry widget to update (passed from main.py)
    """
    if api_entry:
        api_entry.configure(bg=Config.COLORS["VALID_ENTRY"] if is_valid else Config.COLORS["INVALID_ENTRY"])

def save_api_key(api_key_var=None, api_entry=None, update_global_token=None):
    """Save API Key to file and update visual state.
    
    Args:
        api_key_var: StringVar containing the API key (passed from main.py)
        api_entry: The tkinter entry widget to update (passed from main.py)
        update_global_token: Callback function to update global DISCOGS_API_TOKEN variable
    """
    if not api_key_var:
        log_message("[ERROR] API key variable not provided", log_type="debug")
        return False
        
    new_api_key = api_key_var.get().strip()

    if not new_api_key:
        if api_entry:
            update_api_entry_style(False, api_entry)
        log_message("[ERROR] API Key cannot be empty", log_type="processing")
        return False

    # Test the API key with a simple request
    try:
        test_response = requests.get(
            Config.DISCOGS_SEARCH_URL,
            params={"token": new_api_key, "q": "test", "per_page": 1},
            timeout=10
        )

        if test_response.status_code != 200:
            if api_entry:
                update_api_entry_style(False, api_entry)
            log_message("[ERROR] Invalid API Key - Authentication failed", log_type="processing")
            return False
    except requests.RequestException as e:
        if api_entry:
            update_api_entry_style(False, api_entry)
        log_message(f"[ERROR] API key validation failed: {str(e)}", log_type="processing")
        return False

    # Save the API key to file
    try:
        with open(Config.API_KEY_FILE, "w") as f:
            f.write(new_api_key)
    except Exception as e:
        log_message(f"[ERROR] Failed to save API key to file: {str(e)}", log_type="processing")
        return False
    
    # Update global token if callback provided
    if update_global_token:
        update_global_token(new_api_key)
    
    if api_entry:
        update_api_entry_style(True, api_entry)
    log_message("[SUCCESS] API Key validated and saved", log_type="processing")
    return True
