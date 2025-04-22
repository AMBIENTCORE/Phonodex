"""
Metadata handling utilities for the application.
Provides functionality for reading, writing, and fetching audio metadata across different formats.
"""

import mutagen
from mutagen.mp3 import MP3
from mutagen.flac import FLAC, Picture
from mutagen.mp4 import MP4, MP4Cover
from mutagen.oggvorbis import OggVorbis
from mutagen.asf import ASF
from mutagen.id3 import ID3, APIC, TPE1, TIT2, TALB, TPE2, TXXX, TDRC, TRCK, TCON
from utils.logging import log_message
import requests
from collections import Counter
import time
import threading
import os
from services.api_client import make_api_request

# Cache for metadata results
album_catalog_cache = {}
failed_search_cache = set()  # Cache for artist-album combinations that returned no results
cache_lock = threading.Lock()  # Lock for thread-safe cache access

def get_tag_value(audio, tag_name, default=""):
    """Helper function to get tag value across different audio formats."""
    try:
        # MP3
        if isinstance(audio, MP3):
            if not audio.tags:
                return default
                
            id3_mapping = {
                "artist": "TPE1",
                "title": "TIT2",
                "album": "TALB",
                "albumartist": "TPE2",
                "catalognumber": "TXXX:CATALOGNUMBER",
                "date": "TDRC",  # Year/Date
                "tracknumber": "TRCK",  # Track number
                "genre": "TCON"  # Genre
            }
            mapped_tag = id3_mapping.get(tag_name)
            
            if not mapped_tag:
                return default
                
            # Handle TXXX frames specially
            if mapped_tag.startswith("TXXX:"):
                desc = mapped_tag.split(":")[1]
                for tag in audio.tags.getall("TXXX"):
                    if tag.desc == desc:
                        return str(tag.text[0])
            # Handle regular ID3 frames
            elif mapped_tag in audio.tags:
                return str(audio.tags[mapped_tag].text[0])
                
            return default
            
        # FLAC
        elif isinstance(audio, FLAC):
            flac_mapping = {
                "date": "date",
                "tracknumber": "tracknumber",
                "genre": "genre"
            }
            mapped_tag = flac_mapping.get(tag_name, tag_name)
            return audio.get(mapped_tag, [default])[0]
        # MP4/M4A
        elif isinstance(audio, MP4):
            mp4_mapping = {
                "artist": "©ART",
                "title": "©nam",
                "album": "©alb",
                "albumartist": "aART",
                "catalognumber": "----:com.apple.iTunes:CATALOGNUMBER",
                "date": "©day",
                "tracknumber": "trkn",
                "genre": "©gen"
            }
            mapped_tag = mp4_mapping.get(tag_name)
            if mapped_tag and mapped_tag in audio:
                # Special handling for track number in MP4
                if mapped_tag == "trkn" and audio.get(mapped_tag):
                    return str(audio[mapped_tag][0][0])  # Track numbers are stored as tuples
                # Special handling for custom iTunes tags (bytes data)
                elif mapped_tag.startswith("----"):
                    # Custom iTunes tags are stored as bytes and need to be decoded
                    try:
                        byte_value = audio[mapped_tag][0]
                        if isinstance(byte_value, bytes):
                            return byte_value.decode('utf-8')
                        return str(byte_value)
                    except Exception as e:
                        log_message(f"[ERROR] Failed to decode MP4 custom tag {mapped_tag}: {e}")
                        return default
                return str(audio[mapped_tag][0])
        # OGG
        elif isinstance(audio, OggVorbis):
            ogg_mapping = {
                "date": "date",
                "tracknumber": "tracknumber",
                "genre": "genre"
            }
            mapped_tag = ogg_mapping.get(tag_name, tag_name)
            return audio.get(mapped_tag, [default])[0]
        # WMA
        elif isinstance(audio, ASF):
            asf_mapping = {
                "artist": "Author",
                "title": "Title",
                "album": "WM/AlbumTitle",
                "albumartist": "WM/AlbumArtist",
                "catalognumber": "WM/CatalogNo",
                "date": "WM/Year",
                "tracknumber": "WM/TrackNumber",
                "genre": "WM/Genre"
            }
            mapped_tag = asf_mapping.get(tag_name)
            if mapped_tag and mapped_tag in audio:
                return str(audio[mapped_tag][0])
            
        return default
    except Exception as e:
        log_message(f"[ERROR] Failed to get tag {tag_name}: {str(e)}")
        return default

def set_tag_value(audio, tag_name, value):
    """Helper function to set tag value across different audio formats."""
    try:
        # FLAC
        if isinstance(audio, FLAC):
            # FLAC tags need to be lists
            audio[tag_name] = [value]
        # MP4/M4A
        elif isinstance(audio, mutagen.mp4.MP4):
            mp4_mapping = {
                "artist": "©ART",
                "title": "©nam",
                "album": "©alb",
                "albumartist": "aART",
                "catalognumber": "----:com.apple.iTunes:CATALOGNUMBER",
                "date": "©day",  # Year/date
                "tracknumber": "trkn",  # Track number
                "genre": "©gen"  # Genre
            }
            mapped_tag = mp4_mapping.get(tag_name)
            if mapped_tag:
                if mapped_tag == "trkn":
                    # Special handling for track numbers in MP4
                    audio[mapped_tag] = [(int(value), 0)]  # Format as (track_number, total_tracks)
                elif mapped_tag.startswith("----"):
                    # Special handling for custom iTunes tags (like CATALOGNUMBER)
                    try:
                        # Custom iTunes tags need to be encoded as bytes with a special format
                        tag_parts = mapped_tag.split(":")
                        namespace = tag_parts[1]  # e.g., "com.apple.iTunes"
                        name = tag_parts[2]       # e.g., "CATALOGNUMBER"
                        
                        # Create a properly formatted custom tag
                        log_message(f"[DEBUG] Setting iTunes custom tag: {namespace}:{name}={value}")
                        audio[mapped_tag] = [value.encode("utf-8")]
                    except Exception as e:
                        log_message(f"[ERROR] Failed to set custom MP4 tag {mapped_tag}: {e}")
                else:
                    # Regular MP4 tags
                    audio[mapped_tag] = [value]
        # OGG
        elif isinstance(audio, mutagen.oggvorbis.OggVorbis):
            audio[tag_name] = [value]
        # WMA
        elif isinstance(audio, mutagen.asf.ASF):
            asf_mapping = {
                "artist": "Author",
                "title": "Title",
                "album": "WM/AlbumTitle",
                "albumartist": "WM/AlbumArtist",
                "catalognumber": "WM/CatalogNo",
                "date": "WM/Year",  # Year/date
                "tracknumber": "WM/TrackNumber",  # Track number
                "genre": "WM/Genre"  # Genre
            }
            mapped_tag = asf_mapping.get(tag_name)
            if mapped_tag:
                audio[mapped_tag] = [value]
        # MP3
        elif isinstance(audio, mutagen.mp3.MP3):
            id3_mapping = {
                "artist": (TPE1, lambda v: TPE1(encoding=3, text=[v])),
                "title": (TIT2, lambda v: TIT2(encoding=3, text=[v])),
                "album": (TALB, lambda v: TALB(encoding=3, text=[v])),
                "albumartist": (TPE2, lambda v: TPE2(encoding=3, text=[v])),
                "catalognumber": (TXXX, lambda v: TXXX(encoding=3, desc="CATALOGNUMBER", text=[v])),
                "date": (TDRC, lambda v: TDRC(encoding=3, text=[v])),
                "tracknumber": (TRCK, lambda v: TRCK(encoding=3, text=[v])),
                "genre": (TCON, lambda v: TCON(encoding=3, text=[v]))
            }
            if tag_name in id3_mapping:
                frame_class, frame_creator = id3_mapping[tag_name]
                if audio.tags is None:
                    audio.add_tags()
                
                # Remove existing frame before adding new one
                if tag_name == "catalognumber":
                    # Special handling for TXXX frames
                    for tag in list(audio.tags.getall("TXXX")):
                        if tag.desc == "CATALOGNUMBER":
                            audio.tags.delall("TXXX:" + tag.desc)
                else:
                    # Regular ID3 frames
                    frame_name = frame_class.__name__
                    if frame_name in audio.tags:
                        audio.tags.delall(frame_name)
                
                # Always add the frame with the new value (even if empty)
                audio.tags.add(frame_creator(value))
                
        audio.save()
        return True
    except Exception as e:
        log_message(f"[ERROR] Failed to set tag {tag_name}: {str(e)}")
        return False

def select_by_frequency(releases):
    """Helper function to select a release based on catalog number frequency.
    
    Args:
        releases: List of release dictionaries from Discogs API
        
    Returns:
        tuple: (selected_release, normalized_catalog_number)
    """
    # First pass: collect all catalog numbers
    log_message(f"[DEBUG] --- Processing all {len(releases)} releases to find catalog numbers ---")
    all_catalog_numbers = []
    
    # Debug raw catalog values before filtering
    raw_catalogs = [release.get("catno", "MISSING") for release in releases]
    log_message(f"[DEBUG] Raw catalog values: {raw_catalogs}")
    
    for release in releases:
        catno = release.get("catno", "").strip()
        if catno and catno.upper() != "NONE":  # Explicitly exclude NONE values
            all_catalog_numbers.append(catno.upper())
            log_message(f"[DEBUG] Found catalog number: {catno}")
    
    if not all_catalog_numbers:
        log_message(f"[WARNING] No valid catalog numbers found in the filtered releases.")
        
        # Pick first release with ANY catalog value, even if it's "NONE"
        for release in releases:
            catno = release.get("catno", "").strip()
            if catno:  # Any non-empty catalog, even "NONE"
                log_message(f"[DEBUG] Falling back to using catalog: {catno}")
                normalized_catalog = catno.replace(" ", "").upper()
                return release, normalized_catalog
                
        # If still nothing, just use the first release and assign a placeholder
        if releases:
            log_message(f"[DEBUG] Last resort: using first release with placeholder catalog")
            release = releases[0]
            return release, "UNKNOWN"
            
        return None, None
        
    # Count occurrences of all catalog numbers
    catalog_counts = Counter(all_catalog_numbers)
    log_message(f"[DEBUG] --- Analyzing frequency of {len(catalog_counts)} unique catalog numbers ---")
    
    # Get the top 2 most common catalog numbers
    most_common = catalog_counts.most_common(2)
    
    # Start with the most common
    most_common_catalog = most_common[0][0]
    normalized_catalog = most_common_catalog.replace(" ", "").upper()
    log_message(f"[DEBUG] Most common catalog: '{most_common_catalog}' (occurs {most_common[0][1]} times)")
    
    # If the second most common is not "NONE" and the first one is very similar to "NONE",
    # use the second one instead
    if len(most_common) > 1 and normalized_catalog == "NONE" or not normalized_catalog:
        second_common = most_common[1][0]
        second_normalized = second_common.replace(" ", "").upper()
        log_message(f"[DEBUG] First catalog is '{normalized_catalog}', trying second most common: '{second_normalized}' (occurs {most_common[1][1]} times)")
        normalized_catalog = second_normalized
    
    log_message(f"[DEBUG] Selected catalog number by frequency: {normalized_catalog}")
    
    # Find the release with this catalog number
    matching_release = next((release for release in releases 
                           if release.get("catno", "").strip().upper().replace(" ", "") == normalized_catalog), None)
    
    # If no matching release found (shouldn't happen but just in case)
    if not matching_release and releases:
        log_message(f"[DEBUG] No release found with catalog {normalized_catalog}, using first release")
        matching_release = releases[0]
    
    return matching_release, normalized_catalog

def fetch_metadata(artist, album, title=None, api_token=None, search_url=None):
    """Fetch the most common catalog number and essential metadata for an album.
    
    Args:
        artist: The artist name
        album: The album name
        title: The track title (optional) - used as fallback search
        api_token: Discogs API token
        search_url: Discogs search URL endpoint
        
    Returns:
        tuple: (metadata dictionary, response headers) or (None, None) if not found
    """
    if not album:
        log_message("[WARNING] No album metadata found, skipping.")
        return None, None
    
    if not api_token or not search_url:
        log_message("[ERROR] API token or search URL not provided.")
        return None, None
        
    # Clean up artist and album names for better search results
    artist = artist.strip()
    album = album.strip()
    if title:
        title = title.strip()
    
    # Use consistent cache key that includes both artist and album
    cache_key = f"{artist.lower()}|{album.lower()}"
    
    # Thread-safe cache access
    with cache_lock:
        if cache_key in album_catalog_cache:
            log_message(f"[INFO] Using cached catalog number for '{artist} - {album}'.")
            return album_catalog_cache[cache_key], None  # No headers with cached data
        if cache_key in failed_search_cache:
            log_message(f"[INFO] Skipping known failed search for '{artist} - {album}'.")
            return None, None
    
    log_message(f"[API CALL] Requesting Discogs for: Artist='{artist}', Album='{album}'")
    
    # Try first with exact search but use q parameter instead of separate fields
    response_data, response_headers = make_api_request(
        search_url,
        {
            "q": f'"{artist}" "{album}"',  # Quote the terms for exact matching
            "token": api_token,
            "type": "release"  # Ensure we're only getting releases
        }
    )
    
    if not response_data or not response_data.get("results"):
        # If no results, try a more lenient search
        log_message(f"[INFO] No exact matches found, trying broader search...")
        response_data, response_headers = make_api_request(
            search_url,
            {
                "q": f"{artist} {album}",  # Search all fields
                "token": api_token,
                "type": "release"
            }
        )
        
        # If still no results and we have a title that's different from the album name
        if (not response_data or not response_data.get("results")) and title and title.lower() != album.lower():
            log_message(f"[INFO] No matches found with album name, trying with title: {title}")
            response_data, response_headers = make_api_request(
                search_url,
                {
                    "q": f"{artist} {title}",  # Search using title instead of album
                    "token": api_token,
                    "type": "release"
                }
            )
    
    if not response_data or not response_data.get("results"):
        # Cache the failed search
        with cache_lock:
            failed_search_cache.add(cache_key)
            log_message(f"[INFO] Caching failed search for '{artist} - {album}'")
        return None, None
        
    releases = response_data.get("results", [])
    
    # Enhanced logging to show all matches for debugging
    total_results = response_data.get("pagination", {}).get("items", len(releases))
    per_page = response_data.get("pagination", {}).get("per_page", len(releases))
    current_page = response_data.get("pagination", {}).get("page", 1)
    log_message(f"[INFO] Discogs reports {total_results} total matches, showing page {current_page} with {per_page} per page")
    log_message(f"[INFO] Looking through {len(releases)} releases received from Discogs:")
    for idx, release in enumerate(releases[:10], 1):  # Show first 10 for debugging
        log_message(f"[INFO] Match {idx}: '{release.get('title', '')}' ({release.get('year', 'Unknown')}), Catalog: '{release.get('catno', 'Unknown')}'")
    
    # First filter for EXACT album matches, not just artist
    exact_album_matches = []
    for release in releases:
        release_title = release.get('title', '').lower()
        # Split on " - " to separate artist and album title
        if ' - ' in release_title:
            parts = release_title.split(' - ')
            release_artist = parts[0].strip()
            # The album title is everything after the first " - "
            release_album = ' - '.join(parts[1:]).strip()
            
            # Check if both artist and album match (using fuzzy matching to accommodate minor variations)
            artist_match = release_artist.lower() == artist.lower() or release_artist.lower() in artist.lower() or artist.lower() in release_artist.lower()
            album_match = release_album.lower() == album.lower() or release_album.lower() in album.lower() or album.lower() in release_album.lower()
            
            if artist_match and album_match:
                # Verify catalog number is preserved
                catno = release.get("catno", "").strip()
                if catno:
                    log_message(f"[DEBUG] Found exact album match with catalog {catno}: {release.get('title')}")
                else:
                    log_message(f"[DEBUG] Found exact album match WITHOUT catalog: {release.get('title')}")
                exact_album_matches.append(release)
        elif not ' - ' in release_title:
            # Some releases might not follow the "Artist - Title" format
            # Try fuzzy matching on the whole title
            title_match = album.lower() in release_title or release_title in album.lower()
            if title_match:
                log_message(f"[DEBUG] Found title-only match: {release.get('title')}")
                exact_album_matches.append(release)
    
    # If we have exact album matches, use ONLY those instead of all releases
    if exact_album_matches:
        log_message(f"[DEBUG] Using {len(exact_album_matches)} exact album matches instead of all {len(releases)} search results")
        # CRITICAL DEBUG: Verify catalog numbers are preserved in exact matches
        exact_catalogs = [r.get("catno", "") for r in exact_album_matches if r.get("catno", "").strip()]
        log_message(f"[DEBUG] Catalog numbers in exact matches: {exact_catalogs}")
        target_releases = exact_album_matches
    else:
        log_message(f"[WARNING] No exact album matches found. Results may be less accurate.")
        # Even with no exact matches, still try to find any artist matches at least
        exact_artist_matches = []
        for release in releases:
            release_title = release.get('title', '').lower()
            if ' - ' in release_title:
                release_artist = release_title.split(' - ')[0].strip()
                if release_artist.lower() == artist.lower() or release_artist.lower() in artist.lower() or artist.lower() in release_artist.lower():
                    exact_artist_matches.append(release)
                    log_message(f"[DEBUG] Found artist-only match: {release.get('title')}")
        
        target_releases = exact_artist_matches if exact_artist_matches else releases
        if exact_artist_matches:
            log_message(f"[DEBUG] Using {len(exact_artist_matches)} artist-only matches as fallback")
            # CRITICAL DEBUG: Verify catalog numbers are preserved in artist matches
            artist_catalogs = [r.get("catno", "") for r in exact_artist_matches if r.get("catno", "").strip()]
            log_message(f"[DEBUG] Catalog numbers in artist matches: {artist_catalogs}")
    
    # NEW STEP: Check for exact album title matches BEFORE filtering by catalog number
    log_message(f"[DEBUG] Looking for exact album title matches for: '{album}'")
    exact_album_title_matches = []
    
    for release in target_releases:
        release_title = release.get('title', '').lower()
        log_message(f"[DEBUG] Checking release: '{release_title}'")
        
        # Extract just the album part from "Artist - Album"
        if ' - ' in release_title:
            release_album = ' - '.join(release_title.split(' - ')[1:]).strip()
            log_message(f"[DEBUG] Extracted album part: '{release_album}'")
        else:
            release_album = release_title
            log_message(f"[DEBUG] No album part extraction possible, using whole title: '{release_album}'")
        
        # Check for exact album name match
        if release_album.lower() == album.lower():
            log_message(f"[INFO] Found exact album title match: '{release.get('title')}'")
            exact_album_title_matches.append(release)
    
    # If we found exact album title matches, use only those regardless of catalog number
    if exact_album_title_matches:
        log_message(f"[INFO] Using {len(exact_album_title_matches)} exact album title matches - these take priority over catalog number")
        filtered_releases = exact_album_title_matches
        
        # Check if ANY of these exact title matches have non-NONE catalogs
        non_none_exact_matches = [r for r in exact_album_title_matches if r.get("catno", "").strip().upper() != "NONE" and r.get("catno", "").strip() != ""]
        if non_none_exact_matches:
            log_message(f"[INFO] Found {len(non_none_exact_matches)} exact album title matches with valid catalog numbers")
            filtered_releases = non_none_exact_matches
        else:
            log_message(f"[INFO] No exact album title matches have valid catalog numbers, keeping all exact album title matches anyway")
    else:
        log_message(f"[WARNING] No exact album title matches found, proceeding with standard catalog number filtering")
        # Only now apply the NONE catalog filter if we didn't find exact album title matches
        non_none_releases = [r for r in target_releases if r.get("catno", "").strip().upper() != "NONE" and r.get("catno", "").strip() != ""]
        
        # Use non-none releases if available, otherwise fall back to all target releases
        if non_none_releases:
            log_message(f"[DEBUG] Found {len(non_none_releases)} releases with non-NONE catalog numbers")
            filtered_releases = non_none_releases
            # CRITICAL DEBUG: Verify catalog numbers are preserved after NONE filtering
            filtered_catalogs = [r.get("catno", "") for r in filtered_releases]
            log_message(f"[DEBUG] Catalog numbers after filtering: {filtered_catalogs}")
        else:
            log_message(f"[WARNING] All releases have NONE or empty catalog numbers, using all target releases")
            filtered_releases = target_releases
    
    # If no valid releases, return None
    if not filtered_releases:
        log_message(f"[WARNING] No valid releases found for '{album}'.")
        return None, None
    
    # Sort releases by year (oldest first) if year information is available
    releases_with_year = [r for r in filtered_releases if r.get("year") and str(r.get("year")).isdigit()]
    if releases_with_year:
        log_message(f"[DEBUG] --------------------------------------------------")
        log_message(f"[DEBUG] Sorting {len(releases_with_year)} releases by year (oldest first)")
        releases_with_year.sort(key=lambda r: int(r.get("year", 9999)))
        
        # Use the oldest release that has a valid catalog number AND matches artist
        oldest_release = None
        for release in releases_with_year:
            # Re-verify artist match to ensure it's not an unrelated old album
            release_title = release.get('title', '').lower()
            if ' - ' in release_title:
                release_artist = release_title.split(' - ')[0].strip()
                artist_match = (release_artist.lower() == artist.lower() or 
                               release_artist.lower() in artist.lower() or 
                               artist.lower() in release_artist.lower())
                
                if artist_match:
                    oldest_release = release
                    break
            else:
                # If no artist info in title, only use if we had exact album matches initially
                if exact_album_matches and release in exact_album_matches:
                    oldest_release = release
                    break
        
        # If no valid artist match found, fall back to the first release but log warning
        if not oldest_release and releases_with_year:
            oldest_release = releases_with_year[0]
            log_message(f"[WARNING] Oldest release may not match artist '{artist}'. Title: '{oldest_release.get('title')}'")
        
        # Continue with the rest of your existing logic for catalog number validation
        if oldest_release:
            catno = oldest_release.get("catno", "").strip()
            
            # CRITICAL FIX: Check if the catalog number exists and is not "NONE" before using it
            log_message(f"[DEBUG] Oldest release from year {oldest_release.get('year')} has catalog: '{catno}'")
            
            if catno and catno.upper() != "NONE":
                log_message(f"[INFO] Selected oldest release from year {oldest_release.get('year')} with catalog number: {catno}")
                selected_release = oldest_release
                normalized_catalog = catno.replace(" ", "").upper()
            else:
                # Even if the catalog is NONE, if this is an exact album title match, still use it
                if exact_album_title_matches and oldest_release in exact_album_title_matches:
                    log_message(f"[INFO] Selected oldest release with NONE catalog because it's an exact album title match")
                    selected_release = oldest_release
                    normalized_catalog = "NONE"  # Use a standardized placeholder
                else:
                    # Fall back to frequency-based selection if the oldest doesn't have a valid catalog
                    log_message(f"[INFO] Oldest release doesn't have a valid catalog number, falling back to frequency-based selection")
                    selected_release, normalized_catalog = select_by_frequency(filtered_releases)
    else:
        # If no year information, fall back to frequency-based selection
        log_message(f"[DEBUG] --------------------------------------------------")
        log_message(f"[DEBUG] No year information available, using frequency-based selection")
        selected_release, normalized_catalog = select_by_frequency(filtered_releases)
    
    if not selected_release or not normalized_catalog:
        log_message(f"[WARNING] Could not select a valid catalog number for '{album}'.")
        # Last resort: just use the first release with any catalog number
        for release in filtered_releases:
            catno = release.get("catno", "").strip()
            if catno and catno.upper() != "NONE":
                log_message(f"[INFO] Last resort: using first available catalog number: {catno}")
                selected_release = release
                normalized_catalog = catno.replace(" ", "").upper()
                break
        
        # If still no catalog, return None
        if not selected_release or not normalized_catalog:
            return None, None
    
    # Extract essential metadata
    metadata = {
        "catalog_number": normalized_catalog,  # Use normalized version
        "year": selected_release.get("year", ""),
        "album": selected_release.get("title", album),  # Use API's title if available, otherwise use original
        "cover_image": selected_release.get("cover_image", ""),
        "thumb": selected_release.get("thumb", "")
    }
    
    # Thread-safe cache update
    with cache_lock:
        album_catalog_cache[cache_key] = metadata
        
    log_message(f"[INFO] Found metadata for '{artist} - {album}': {metadata}")
    return metadata, response_headers

def update_album_metadata(file_path, metadata, audio_file=None, options=None, callbacks=None):
    """Update an audio file's metadata based on provided options.
    
    Args:
        file_path: Path to the audio file
        metadata: Dictionary containing metadata to update (catalog_number, year, cover_image, etc.)
        audio_file: Optional pre-loaded audio file object
        options: Dictionary with boolean flags for what to update (catalog, year, art)
        callbacks: Dictionary of callback functions (log_message, mark_updated, mark_processed)
        
    Returns:
        bool: True if any updates were made, False otherwise
    """
    # Default options if none provided
    if options is None:
        options = {
            'catalog': True,
            'year': True,
            'art': True
        }
    
    # Default callbacks
    if callbacks is None:
        callbacks = {
            'log_message': lambda msg: None,
            'mark_updated': lambda path: None,
            'mark_processed': lambda path: None
        }
    
    log_message = callbacks.get('log_message', lambda msg: None)
    mark_updated = callbacks.get('mark_updated', lambda path: None)
    mark_processed = callbacks.get('mark_processed', lambda path: None)
    
    try:
        # Load audio file if not provided
        if audio_file is None:
            from utils.file_operations import get_audio_file
            audio_file = get_audio_file(file_path)
            
        if not audio_file:
            log_message(f"[ERROR] Failed to load audio file: {file_path}")
            return False

        updated = False  # Track if any updates were made
        normalized_path = os.path.normpath(file_path)

        # Update catalog number if selected
        if options.get('catalog', True) and metadata.get("catalog_number"):
            try:
                set_tag_value(audio_file, "catalognumber", metadata["catalog_number"])
                updated = True
                log_message(f"[SUCCESS] Updated catalog number for {os.path.basename(file_path)}")
            except Exception as e:
                log_message(f"[ERROR] Failed to update catalog number: {e}")

        # Update year if selected
        if options.get('year', True) and metadata.get("year"):
            try:
                set_tag_value(audio_file, "date", str(metadata["year"]))
                updated = True
                log_message(f"[SUCCESS] Updated year to {metadata['year']} for {os.path.basename(file_path)}")
            except Exception as e:
                log_message(f"[ERROR] Failed to update year: {e}")

        # Save changes if any were made
        if updated:
            audio_file.save()
            mark_updated(normalized_path)

        # Update album art if selected
        if options.get('art', True) and (metadata.get("cover_image") or metadata.get("thumb")):
            try:
                cover_url = metadata.get("cover_image") or metadata.get("thumb")
                
                # Add API token if provided
                headers = {
                    'User-Agent': 'Phonodex/1.0',
                    'Referer': 'https://www.discogs.com/'
                }
                
                if metadata.get('api_token'):
                    headers['Authorization'] = f'Discogs token={metadata["api_token"]}'
                
                # For MP3 files, always remove existing art first
                if isinstance(audio_file, MP3):
                    if audio_file.tags is None:
                        audio_file.add_tags()
                        log_message(f"[COVER] Added new ID3 tags to file")
                    
                    # Always remove existing cover art first
                    existing_apic = audio_file.tags.getall("APIC")
                    if existing_apic:
                        log_message(f"[COVER] Found {len(existing_apic)} existing APIC frames, removing them")
                        audio_file.tags.delall("APIC")
                    else:
                        log_message("[COVER] No existing APIC frames found")
                
                response = requests.get(cover_url, headers=headers, timeout=10)
                if response.status_code == 200:
                    # Handle FLAC files
                    if isinstance(audio_file, FLAC):
                        # Clear existing pictures
                        audio_file.clear_pictures()
                        
                        # Create new picture
                        picture = Picture()
                        picture.type = 3  # Front cover
                        picture.mime = response.headers.get('content-type', 'image/jpeg')
                        picture.desc = 'Front Cover'
                        picture.data = response.content
                        
                        # Add picture to FLAC file
                        audio_file.add_picture(picture)
                        audio_file.save()
                        updated = True
                        log_message(f"[SUCCESS] Updated cover art for {os.path.basename(file_path)}")
                    
                    # Handle MP3 files
                    elif isinstance(audio_file, MP3):
                        log_message(f"[COVER] Updating cover art for MP3 file")
                        
                        # Add new cover art
                        try:
                            mime_type = response.headers.get('content-type', 'image/jpeg')
                            log_message(f"[COVER] Adding new cover art: {len(response.content)} bytes, mime: {mime_type}")
                            
                            # Always use type 3 (front cover) for new cover art
                            audio_file.tags.add(
                                APIC(
                                    encoding=3,
                                    mime=mime_type,
                                    type=3,  # Front cover
                                    desc='Front Cover',
                                    data=response.content
                                )
                            )
                            log_message(f"[COVER] Successfully added front cover APIC frame")
                            audio_file.save()
                            updated = True
                            log_message(f"[SUCCESS] Updated cover art for {os.path.basename(file_path)}")
                        except Exception as e:
                            log_message(f"[COVER] Error adding APIC frame: {e}")
                    # Handle MP4/M4A files
                    elif isinstance(audio_file, MP4):
                        log_message(f"[COVER] Updating cover art for MP4/M4A file")
                        
                        try:
                            # Get the image data and content type
                            image_data = response.content
                            mime_type = response.headers.get('content-type', 'image/jpeg')
                            log_message(f"[COVER] Adding cover art: {len(image_data)} bytes, mime: {mime_type}")
                            
                            # Determine correct cover format based on mime type
                            if mime_type.endswith('png'):
                                cover_format = MP4Cover.FORMAT_PNG
                            else:
                                cover_format = MP4Cover.FORMAT_JPEG
                                
                            # Create MP4Cover object and set it
                            cover = MP4Cover(image_data, cover_format)
                            audio_file['covr'] = [cover]
                            
                            # Save the file
                            audio_file.save()
                            updated = True
                            log_message(f"[SUCCESS] Updated cover art for {os.path.basename(file_path)}")
                        except Exception as e:
                            log_message(f"[COVER] Error updating MP4 cover art: {e}")
                    else:
                        log_message(f"[COVER] Album art update not supported for this file type: {type(audio_file).__name__}")
                else:
                    log_message(f"[ERROR] Failed to download cover image (Status {response.status_code})")
            except Exception as e:
                log_message(f"[ERROR] Failed to update cover art: {str(e)}")

        if updated:
            mark_updated(normalized_path)
            mark_processed(normalized_path)
        return updated

    except Exception as e:
        log_message(f"[ERROR] Failed to update metadata for {os.path.basename(file_path)}: {str(e)}")
        return False

def update_tag_by_column(file_path, column_num, new_value, audio_file=None, column_to_tag_mapping=None, callbacks=None):
    """Update a specific tag in an audio file based on column index.
    
    Args:
        file_path: Path to the audio file
        column_num: The column index that was edited
        new_value: The new value to set for the tag
        audio_file: Optional pre-loaded audio file object
        column_to_tag_mapping: Dictionary mapping column indices to tag names
        callbacks: Dictionary of callback functions (log_message, mark_updated)
        
    Returns:
        bool: True if update was successful, False otherwise
    """
    # Default column to tag mapping if none provided
    if column_to_tag_mapping is None:
        column_to_tag_mapping = {
            0: "artist",
            1: "title",
            2: "album",
            3: "catalognumber",
            4: "albumartist",
            5: "date",
            6: "tracknumber",
            7: "genre",
        }
    
    # Default callbacks
    if callbacks is None:
        callbacks = {
            'log_message': lambda msg: None,
            'mark_updated': lambda path: None
        }
    
    log_message = callbacks.get('log_message', lambda msg: None)
    mark_updated = callbacks.get('mark_updated', lambda path: None)
    
    try:
        # Load audio file if not provided
        if audio_file is None:
            from utils.file_operations import get_audio_file
            audio_file = get_audio_file(file_path)
            
        if not audio_file:
            log_message(f"[ERROR] Failed to load audio file: {file_path}")
            return False
        
        # Get the tag name from the mapping
        if column_num not in column_to_tag_mapping:
            log_message(f"[ERROR] No tag mapping for column {column_num}")
            return False
            
        tag = column_to_tag_mapping[column_num]
        
        # Set the tag value
        if set_tag_value(audio_file, tag, new_value):
            mark_updated(file_path)
            log_message(f"[SUCCESS] Updated {os.path.basename(file_path)} {tag}: {new_value}")
            return True
        else:
            log_message(f"[ERROR] Failed to update {tag} for {os.path.basename(file_path)}")
            return False
            
    except Exception as e:
        log_message(f"[ERROR] Failed to update metadata for {os.path.basename(file_path)}: {str(e)}")
        return False

def update_mp3_metadata(file_path, column_num, new_value, callbacks=None):
    """Update the audio file's metadata based on the edited column.
    
    Args:
        file_path: Path to the audio file
        column_num: The column index that was edited
        new_value: The new value to set for the tag
        callbacks: Dictionary of callback functions (log_message, mark_updated)
        
    Returns:
        bool: True if update was successful, False otherwise
    """
    # Use the existing utility function
    return update_tag_by_column(file_path, column_num, new_value, callbacks=callbacks)
