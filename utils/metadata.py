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
from mutagen.id3 import ID3, APIC, TPE1, TIT2, TALB, TPE2, TXXX, TDRC, TRCK, TCON, COMM
from utils.logging import log_message
import requests
from collections import Counter
import time
import threading
import os
import re
from services.api_client import make_api_request
from utils.matching import (
    normalize, similarity, extract_year, extract_track_number,
    split_discogs_title, strip_discogs_artist_suffix, parse_filename,
    score_release_candidate,
)

# Cache for full Discogs release details (tracklist/images/labels), keyed by
# release id. Populated by fetch_release_details and reused across files that
# resolve to the same release so we only pay for one /releases/{id} call.
release_details_cache = {}

# Confidence at/above which identify_release stops issuing further search
# strategies (a strong match is already in hand).
IDENTIFY_HIGH_CONFIDENCE = 0.9

# Suggestion field -> underlying tag name used by set_tag_value. "year" maps to
# the "date" tag. Cover art is handled separately (see apply_release_suggestion).
SUGGESTION_TEXT_FIELDS = {
    "artist": "artist",
    "title": "title",
    "album": "album",
    "albumartist": "albumartist",
    "tracknumber": "tracknumber",
    "year": "date",
    "catalognumber": "catalognumber",
}

# Cache for metadata results
album_catalog_cache = {}
failed_search_cache = set()  # Cache for artist-album combinations that returned no results
album_cover_image_cache = {}  # Cache for downloaded cover images (stores actual image bytes)
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
                "genre": "TCON",  # Genre
                "comment": "COMM"  # Comment (COMM frame -- has lang/desc)
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
            # Handle COMM (comment) frames specially: they carry lang + desc.
            # Prefer the frame with empty desc (main comment); fall back to first.
            elif mapped_tag == "COMM":
                comm_frames = audio.tags.getall("COMM")
                for tag in comm_frames:
                    if tag.desc == "":
                        return str(tag.text[0])
                if comm_frames:
                    return str(comm_frames[0].text[0])
            # Handle regular ID3 frames
            elif mapped_tag in audio.tags:
                return str(audio.tags[mapped_tag].text[0])
                
            return default
            
        # FLAC
        elif isinstance(audio, FLAC):
            flac_mapping = {
                "date": "date",
                "tracknumber": "tracknumber",
                "genre": "genre",
                "comment": "comment"
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
                "genre": "©gen",
                "comment": "©cmt"
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
                "genre": "genre",
                "comment": "comment"
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
                "genre": "WM/Genre",
                "comment": "WM/Comment"
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
                "genre": "©gen",  # Genre
                "comment": "©cmt"  # Comment
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
                "genre": "WM/Genre",  # Genre
                "comment": "WM/Comment"  # Comment
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
                "genre": (TCON, lambda v: TCON(encoding=3, text=[v])),
                "comment": (COMM, lambda v: COMM(encoding=3, lang='eng', desc='', text=[v]))
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
                elif tag_name == "comment":
                    # COMM frames key includes lang/desc; delall clears all of them
                    audio.tags.delall("COMM")
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


def _build_discogs_api_base(search_url):
    """Derive the Discogs API base URL from the configured search endpoint."""
    if "/database/search" in search_url:
        return search_url.split("/database/search")[0]
    return "https://api.discogs.com"


def _score_art_candidate(image_data):
    """Score a Discogs image candidate and return (score, width, height)."""
    score = 0.0
    image_type = str(image_data.get("type", "")).lower()
    uri = str(image_data.get("uri", "")).lower()
    uri150 = str(image_data.get("uri150", "")).lower()
    url_blob = f"{uri} {uri150}"

    width = int(image_data.get("width", 0) or 0)
    height = int(image_data.get("height", 0) or 0)

    # Prefer Discogs primary images.
    if image_type == "primary":
        score += 4.0

    # Prefer larger images.
    longest_side = max(width, height)
    if longest_side >= 1200:
        score += 4.0
    elif longest_side >= 800:
        score += 2.0
    elif longest_side >= 500:
        score += 0.5
    else:
        score -= 3.0

    # Prefer near-square images (typical album covers).
    if width > 0 and height > 0:
        ratio = width / height
        square_distance = abs(1.0 - ratio)
        if square_distance <= 0.08:
            score += 2.0
        elif square_distance <= 0.2:
            score += 1.0
        elif square_distance > 0.5:
            score -= 2.5

    # Penalize likely non-front-cover assets.
    penalized_terms = (
        "vinyl", "record", "disc", "label", "back", "rear", "spine",
        "insert", "booklet", "runout", "matrix"
    )
    if re.search(r"\b(" + "|".join(penalized_terms) + r")\b", url_blob):
        score -= 4.0

    return score, width, height


def _pick_best_art_image(images, source_label):
    """Pick the highest-scoring image candidate from a Discogs images array."""
    if not images:
        return None

    best = None
    best_score = float("-inf")
    for image_data in images:
        uri = image_data.get("uri")
        if not uri:
            continue

        score, width, height = _score_art_candidate(image_data)
        candidate = {
            "url": uri,
            "thumb": image_data.get("uri150", ""),
            "score": score,
            "width": width,
            "height": height,
            "source": source_label,
            "type": image_data.get("type", "")
        }
        if score > best_score:
            best = candidate
            best_score = score

    return best


def _is_candidate_poor(candidate):
    """Heuristic check to decide if we should try master images."""
    if not candidate:
        return True
    longest_side = max(candidate.get("width", 0), candidate.get("height", 0))
    return candidate.get("score", 0) < 2.0 or longest_side < 800


def _resolve_cover_art_from_discogs(selected_release, api_token, search_url):
    """Resolve best cover art via release details, then master, then search fallback."""
    release_id = selected_release.get("id")
    if not release_id:
        return selected_release.get("cover_image", ""), selected_release.get("thumb", "")
    api_base = _build_discogs_api_base(search_url)
    release_details_url = f"{api_base}/releases/{release_id}"
    release_data, _ = make_api_request(release_details_url, {"token": api_token})

    best_candidate = None
    master_id = None

    if release_data:
        master_id = release_data.get("master_id")
        best_candidate = _pick_best_art_image(release_data.get("images", []), "release")
        if best_candidate:
            log_message(
                f"[COVER] Best release image score={best_candidate['score']:.2f}, "
                f"size={best_candidate['width']}x{best_candidate['height']}, "
                f"type={best_candidate['type']}"
            )
    else:
        log_message(f"[COVER] Could not fetch release details for release id {release_id}")

    # If release art is poor/absent, attempt master art as fallback upgrade.
    if (not best_candidate or _is_candidate_poor(best_candidate)) and master_id:
        master_url = f"{api_base}/masters/{master_id}"
        master_data, _ = make_api_request(master_url, {"token": api_token})
        if master_data:
            master_candidate = _pick_best_art_image(master_data.get("images", []), "master")
            if master_candidate:
                log_message(
                    f"[COVER] Best master image score={master_candidate['score']:.2f}, "
                    f"size={master_candidate['width']}x{master_candidate['height']}, "
                    f"type={master_candidate['type']}"
                )
                if not best_candidate or master_candidate["score"] > best_candidate["score"]:
                    best_candidate = master_candidate
        else:
            log_message(f"[COVER] Could not fetch master details for master id {master_id}")

    if best_candidate and best_candidate.get("url"):
        return best_candidate["url"], best_candidate.get("thumb", "")

    # Final fallback chain for robustness.
    return selected_release.get("cover_image", ""), selected_release.get("thumb", "")


def _resolve_cover_art(artist, album, selected_release, api_token, search_url):
    """Resolve cover art using iTunes first (if enabled), then Discogs as fallback."""
    # Try iTunes Search API first - it almost always returns the real front cover
    # rather than vinyl photos / back covers / labels that Discogs sometimes serves.
    try:
        from config import Config
        prefer_itunes = Config.ALBUM_ART.get("PREFER_ITUNES", True)
    except Exception:
        prefer_itunes = True

    if prefer_itunes:
        try:
            from services.itunes_client import fetch_album_art_url
            itunes_cover, itunes_thumb = fetch_album_art_url(artist, album)
            if itunes_cover:
                log_message(f"[COVER] Using iTunes art for '{artist} - {album}'")
                return itunes_cover, itunes_thumb
        except Exception as e:
            log_message(f"[COVER] iTunes lookup raised, falling back to Discogs: {e}", log_type="debug")

    # Discogs fallback (also the path when PREFER_ITUNES is disabled)
    return _resolve_cover_art_from_discogs(selected_release, api_token, search_url)


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
    
    # Art-only enhancement: resolve better artwork after release selection.
    # Pass the original album name (not the Discogs release title, which is in
    # "Artist - Album" format) so the iTunes matcher scores against a clean title.
    resolved_cover, resolved_thumb = _resolve_cover_art(
        artist, album, selected_release, api_token, search_url
    )

    # Extract essential metadata
    metadata = {
        "catalog_number": normalized_catalog,  # Use normalized version
        "year": selected_release.get("year", ""),
        "album": selected_release.get("title", album),  # Use API's title if available, otherwise use original
        "cover_image": resolved_cover or selected_release.get("cover_image", ""),
        "thumb": resolved_thumb or selected_release.get("thumb", "")
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
                
                # Check if we already have cached image data for this URL
                image_data = None
                mime_type = 'image/jpeg'
                
                with cache_lock:
                    if cover_url in album_cover_image_cache:
                        cached_image = album_cover_image_cache[cover_url]
                        image_data = cached_image['data']
                        mime_type = cached_image['mime']
                        log_message(f"[COVER] Using cached image data: {len(image_data)} bytes (lossless transfer)")
                
                # If not cached, download it
                if image_data is None:
                    # Add API token if provided
                    headers = {
                        'User-Agent': 'Phonodex/1.0',
                        'Referer': 'https://www.discogs.com/'
                    }
                    
                    if metadata.get('api_token'):
                        headers['Authorization'] = f'Discogs token={metadata["api_token"]}'
                    
                    log_message(f"[COVER] Downloading cover art from: {cover_url}")
                    response = requests.get(cover_url, headers=headers, timeout=10)
                    
                    if response.status_code == 200:
                        image_data = response.content
                        mime_type = response.headers.get('content-type', 'image/jpeg')
                        
                        # Cache the downloaded image for future use
                        with cache_lock:
                            album_cover_image_cache[cover_url] = {
                                'data': image_data,
                                'mime': mime_type
                            }
                        log_message(f"[COVER] Downloaded and cached: {len(image_data)} bytes, mime: {mime_type}")
                    else:
                        log_message(f"[ERROR] Failed to download cover image (Status {response.status_code})")
                        image_data = None
                
                # If we have image data (either cached or freshly downloaded), apply it
                if image_data is not None:
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
                    # Handle FLAC files
                    if isinstance(audio_file, FLAC):
                        # Clear existing pictures
                        audio_file.clear_pictures()
                        
                        # Create new picture
                        picture = Picture()
                        picture.type = 3  # Front cover
                        picture.mime = mime_type
                        picture.desc = 'Front Cover'
                        picture.data = image_data
                        
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
                            log_message(f"[COVER] Adding new cover art: {len(image_data)} bytes, mime: {mime_type}")
                            
                            # Always use type 3 (front cover) for new cover art
                            audio_file.tags.add(
                                APIC(
                                    encoding=3,
                                    mime=mime_type,
                                    type=3,  # Front cover
                                    desc='Front Cover',
                                    data=image_data
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
            # 8 is File Path (hidden, not editable)
            9: "comment",
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


# ===========================================================================
# Release identification pipeline ("find the release no matter what tags exist")
# ===========================================================================

def identify_release(known, api_token, search_url, request_hooks=None, max_searches=4):
    """Identify candidate Discogs releases from whatever tags are available.

    Runs a cascade of structured Discogs searches, ordered best->worst by
    expected precision, using only the fields present in ``known``. Results from
    every strategy are pooled, de-duplicated by release id, scored against the
    known tags, and returned ranked by confidence. The cascade stops early once
    a strong match (>= IDENTIFY_HIGH_CONFIDENCE) is found or ``max_searches``
    API calls have been issued, so we don't burn the rate limit needlessly.

    Args:
        known: dict with any of artist, albumartist, album, title, year,
            catalognumber.
        api_token: Discogs API token.
        search_url: Discogs /database/search endpoint.
        request_hooks: optional dict with callables ``before`` -> bool (return
            False to abort, e.g. rate-limited) and ``after`` -> None (receives
            response headers) so the caller can drive the rate-limit UI.
        max_searches: hard cap on the number of search API calls.

    Returns:
        (ranked_candidates, last_headers). Each candidate is a dict:
        {id, master_id, title, artist, album, year, catno, cover_image, thumb,
         score, strategy}.
    """
    if not api_token or not search_url:
        log_message("[ERROR] API token or search URL not provided.")
        return [], None

    artist = (known.get("artist") or "").strip()
    albumartist = (known.get("albumartist") or "").strip()
    album = (known.get("album") or "").strip()
    title = (known.get("title") or "").strip()
    year = extract_year(known.get("year"))
    catno = (known.get("catalognumber") or "").strip()
    primary_artist = artist or albumartist

    # Build the ordered strategy list from what we actually know.
    strategies = []
    if catno:
        strategies.append(("catno", {"catno": catno}))
    if primary_artist and album:
        strategies.append(("artist+album", {"artist": primary_artist, "release_title": album}))
    if albumartist and album and albumartist.lower() != artist.lower():
        strategies.append(("albumartist+album", {"artist": albumartist, "release_title": album}))
    if primary_artist and title and not album:
        strategies.append(("artist+track", {"artist": primary_artist, "track": title}))
    if album and year:
        strategies.append(("album+year", {"release_title": album, "year": year}))
    free_terms = " ".join(t for t in [primary_artist, album or title] if t).strip()
    if free_terms:
        strategies.append(("q", {"q": free_terms}))

    if not strategies:
        log_message("[IDENTIFY] Nothing to search on (no usable tags).")
        return [], None

    candidates_by_id = {}
    last_headers = None
    searches = 0

    for name, params in strategies:
        if searches >= max_searches:
            break
        if request_hooks and request_hooks.get("before"):
            if not request_hooks["before"]():
                log_message("[IDENTIFY] Aborting search cascade (rate limited).")
                break
        query = dict(params)
        query["token"] = api_token
        query["type"] = "release"
        data, headers = make_api_request(search_url, query)
        searches += 1
        last_headers = headers or last_headers
        if request_hooks and request_hooks.get("after"):
            request_hooks["after"](headers)

        results = (data or {}).get("results", []) if data else []
        log_message(f"[IDENTIFY] strategy '{name}' -> {len(results)} results")
        for release in results:
            release_id = release.get("id")
            if not release_id:
                continue
            candidate_score = score_release_candidate(known, release)
            existing = candidates_by_id.get(release_id)
            if existing and existing["score"] >= candidate_score:
                continue
            r_artist, r_album = split_discogs_title(release.get("title", ""))
            candidates_by_id[release_id] = {
                "id": release_id,
                "master_id": release.get("master_id"),
                "title": release.get("title", ""),
                "artist": r_artist,
                "album": r_album,
                "year": str(release.get("year", "") or ""),
                "catno": release.get("catno", ""),
                "cover_image": release.get("cover_image", ""),
                "thumb": release.get("thumb", ""),
                "score": candidate_score,
                "strategy": name,
            }

        best = max((c["score"] for c in candidates_by_id.values()), default=0.0)
        if best >= IDENTIFY_HIGH_CONFIDENCE:
            log_message(f"[IDENTIFY] strong match {best:.2f} found, stopping cascade early.")
            break

    ranked = sorted(candidates_by_id.values(), key=lambda c: c["score"], reverse=True)
    log_message(f"[IDENTIFY] {len(ranked)} candidate(s); best score "
                f"{ranked[0]['score']:.2f}" if ranked else "[IDENTIFY] no candidates")
    return ranked, last_headers


def fetch_release_details(release_id, api_token, search_url, request_hooks=None):
    """Fetch full Discogs release details (tracklist/images/labels), cached.

    Returns (release_data, headers). ``release_data`` is None on failure. Cached
    per release id so multiple files on the same release share one API call.
    """
    if not release_id:
        return None, None
    with cache_lock:
        if release_id in release_details_cache:
            return release_details_cache[release_id], None
    if request_hooks and request_hooks.get("before"):
        if not request_hooks["before"]():
            return None, None
    api_base = _build_discogs_api_base(search_url)
    url = f"{api_base}/releases/{release_id}"
    data, headers = make_api_request(url, {"token": api_token})
    if request_hooks and request_hooks.get("after"):
        request_hooks["after"](headers)
    if data:
        with cache_lock:
            release_details_cache[release_id] = data
    return data, headers


def _flatten_tracklist(release_data):
    """Return only actual playable tracks, skipping headings/index entries."""
    tracks = []
    for entry in (release_data.get("tracklist", []) or []):
        # Discogs marks section headers with type_ == "heading"; real tracks are
        # "track" (or occasionally missing the field on older data).
        if entry.get("type_") in (None, "", "track"):
            tracks.append(entry)
    return tracks


def map_file_to_track(known, release_data):
    """Match a file to a track on the release by track number, then fuzzy title.

    Returns the matched track dict, or None if nothing lines up confidently.
    """
    tracks = _flatten_tracklist(release_data)
    if not tracks:
        return None

    known_tn = extract_track_number(known.get("tracknumber"))
    known_title = known.get("title") or ""

    # 1) Exact track-number match against the numeric part of each position.
    if known_tn:
        for track in tracks:
            if extract_track_number(track.get("position")) == known_tn:
                return track
        # Positional fallback: the Nth track, but only trust it if the title
        # also plausibly matches (or we have no title to check against).
        try:
            n = int(known_tn)
        except ValueError:
            n = 0
        if 1 <= n <= len(tracks):
            candidate = tracks[n - 1]
            if not known_title or similarity(known_title, candidate.get("title", "")) >= 0.4:
                return candidate

    # 2) Best fuzzy title match.
    if known_title:
        best, best_score = None, 0.0
        for track in tracks:
            score = similarity(known_title, track.get("title", ""))
            if score > best_score:
                best, best_score = track, score
        if best and best_score >= 0.6:
            return best

    return None


def has_embedded_art(audio_file):
    """Return True if the loaded audio file already has embedded cover art."""
    try:
        if isinstance(audio_file, MP3):
            return bool(audio_file.tags and audio_file.tags.getall("APIC"))
        if isinstance(audio_file, FLAC):
            return bool(audio_file.pictures)
        if isinstance(audio_file, MP4):
            return bool(audio_file.get("covr"))
    except Exception:
        return False
    return False


def _resolve_art_for_suggestion(artist, album, candidate, release_data):
    """Pick the best cover URL without extra Discogs calls.

    iTunes first (free, usually the true front cover), then the already-fetched
    release images, then the candidate's own thumbnail.
    """
    try:
        from config import Config
        prefer_itunes = Config.ALBUM_ART.get("PREFER_ITUNES", True)
    except Exception:
        prefer_itunes = True

    if prefer_itunes and album:
        try:
            from services.itunes_client import fetch_album_art_url
            cover, thumb = fetch_album_art_url(artist, album)
            if cover:
                return cover, thumb
        except Exception as e:
            log_message(f"[COVER] iTunes lookup failed, using Discogs art: {e}", log_type="debug")

    if release_data:
        best = _pick_best_art_image(release_data.get("images", []), "release")
        if best and best.get("url"):
            return best["url"], best.get("thumb", "")

    return candidate.get("cover_image", ""), candidate.get("thumb", "")


# The user standardizes every Various-Artists compilation to the single word
# "Various". Discogs (and other sources) may instead return "Various Artists",
# "V.A.", "VA", etc. — map those spellings to the canonical form. Matching is
# exact on the normalized name, so real artists like "Various Cruelties" are
# left untouched.
_VARIOUS_ALIASES = {"various", "various artists", "various artist", "va", "v a"}


def _canonicalize_various(name):
    """Return 'Various' for any Various-Artists spelling, else the name unchanged."""
    if name and normalize(name) in _VARIOUS_ALIASES:
        return "Various"
    return name


def build_track_suggestion(known, candidate, release_data, api_token, search_url):
    """Build the proposed tag values for a single file from a chosen release.

    Combines release-level fields (album, year, album artist, catalog) with the
    matched track's fields (title, per-track artist, position) and resolves a
    cover URL. Returns a dict with keys: artist, title, album, albumartist,
    tracknumber, year, catalognumber, cover_image, thumb (missing pieces = "").
    """
    catalognumber = candidate.get("catno", "") or ""
    if release_data:
        album = release_data.get("title", "") or candidate.get("album", "")
        year = str(release_data.get("year", "") or candidate.get("year", "") or "")
        albumartist = ""
        artists = release_data.get("artists", []) or []
        if artists:
            albumartist = strip_discogs_artist_suffix(artists[0].get("name", ""))
        if not catalognumber:
            for label in (release_data.get("labels", []) or []):
                label_catno = (label.get("catno") or "").strip()
                if label_catno and label_catno.upper() != "NONE":
                    catalognumber = label_catno
                    break
    else:
        album = candidate.get("album", "")
        year = candidate.get("year", "")
        albumartist = candidate.get("artist", "")

    # Normalize the catalog number the same way the regular PROCESS flow does
    # (select_by_frequency / fetch_metadata): strip spaces and uppercase, and fall
    # back to the standardized "NONE" placeholder when the release has no catalog
    # number, so behavior and folder naming stay consistent with regular PROCESS.
    catalognumber = catalognumber.replace(" ", "").upper()
    if not catalognumber:
        catalognumber = "NONE"

    artist = albumartist or candidate.get("artist", "")
    title = ""
    tracknumber = extract_track_number(known.get("tracknumber"))

    track = map_file_to_track(known, release_data) if release_data else None
    if track:
        title = track.get("title", "")
        track_artists = track.get("artists", []) or []
        if track_artists:
            per_track_artist = strip_discogs_artist_suffix(track_artists[0].get("name", ""))
            if per_track_artist:
                artist = per_track_artist
        position = extract_track_number(track.get("position"))
        if position:
            tracknumber = position

    cover_image, thumb = _resolve_art_for_suggestion(
        artist or albumartist, album, candidate, release_data
    )

    # Standardize Various-Artists spellings to the user's canonical "Various" for
    # the ALBUM artist only. The track artist is never "Various" — that's the
    # album artist's role — so if the only artist we could resolve is the VA
    # placeholder (no real per-track artist), leave it blank and suggest no
    # artist change rather than writing "Various" into the track artist.
    albumartist = _canonicalize_various(albumartist)
    if _canonicalize_various(artist) == "Various":
        artist = ""

    return {
        "artist": artist,
        "title": title,
        "album": album,
        "albumartist": albumartist,
        "tracknumber": tracknumber,
        "year": year,
        "catalognumber": catalognumber,
        "cover_image": cover_image,
        "thumb": thumb,
    }


def _download_cover_bytes(cover_url, api_token=None):
    """Download cover image bytes (cached). Returns (data, mime) or (None, None)."""
    with cache_lock:
        if cover_url in album_cover_image_cache:
            cached = album_cover_image_cache[cover_url]
            return cached["data"], cached["mime"]

    headers = {"User-Agent": "Phonodex/1.0", "Referer": "https://www.discogs.com/"}
    if api_token:
        headers["Authorization"] = f"Discogs token={api_token}"
    try:
        response = requests.get(cover_url, headers=headers, timeout=10)
    except Exception as e:
        log_message(f"[ERROR] Cover download failed: {e}")
        return None, None
    if response.status_code != 200:
        log_message(f"[ERROR] Failed to download cover image (Status {response.status_code})")
        return None, None

    data = response.content
    mime = response.headers.get("content-type", "image/jpeg")
    with cache_lock:
        album_cover_image_cache[cover_url] = {"data": data, "mime": mime}
    return data, mime


def apply_cover_art(audio_file, file_path, cover_url, api_token=None, log_fn=None):
    """Download (cached) and embed cover art into an already-loaded audio file.

    Returns True on success. Supports MP3, FLAC and MP4/M4A (same formats the
    original updater handled); other formats are skipped with a log note.
    """
    log_fn = log_fn or (lambda msg: None)
    image_data, mime_type = _download_cover_bytes(cover_url, api_token)
    if image_data is None:
        return False

    try:
        if isinstance(audio_file, FLAC):
            audio_file.clear_pictures()
            picture = Picture()
            picture.type = 3  # Front cover
            picture.mime = mime_type
            picture.desc = "Front Cover"
            picture.data = image_data
            audio_file.add_picture(picture)
            audio_file.save()
        elif isinstance(audio_file, MP3):
            if audio_file.tags is None:
                audio_file.add_tags()
            audio_file.tags.delall("APIC")
            audio_file.tags.add(
                APIC(encoding=3, mime=mime_type, type=3, desc="Front Cover", data=image_data)
            )
            audio_file.save()
        elif isinstance(audio_file, MP4):
            cover_format = MP4Cover.FORMAT_PNG if mime_type.endswith("png") else MP4Cover.FORMAT_JPEG
            audio_file["covr"] = [MP4Cover(image_data, cover_format)]
            audio_file.save()
        else:
            log_fn(f"[COVER] Album art not supported for this file type: {type(audio_file).__name__}")
            return False
        log_fn(f"[SUCCESS] Updated cover art for {os.path.basename(file_path)}")
        return True
    except Exception as e:
        log_fn(f"[ERROR] Failed to embed cover art: {e}")
        return False


def apply_release_suggestion(file_path, suggestion, accepted, api_token=None,
                             audio_file=None, callbacks=None):
    """Write the accepted fields of a suggestion to a file.

    Args:
        file_path: path to the audio file.
        suggestion: dict of proposed values (see build_track_suggestion);
            cover art comes from its "cover_image"/"thumb" keys.
        accepted: dict field->bool. Recognized keys are the text fields in
            SUGGESTION_TEXT_FIELDS plus "art".
        api_token: Discogs token (used only for cover download auth).
        audio_file: optional pre-loaded mutagen object.
        callbacks: optional dict with log_message, mark_updated, mark_processed.

    Returns True if at least one field was written.
    """
    callbacks = callbacks or {}
    log_fn = callbacks.get("log_message", lambda msg: None)
    mark_updated = callbacks.get("mark_updated", lambda path: None)
    mark_processed = callbacks.get("mark_processed", lambda path: None)

    if audio_file is None:
        from utils.file_operations import get_audio_file
        audio_file = get_audio_file(file_path)
    if not audio_file:
        log_fn(f"[ERROR] Failed to load audio file: {file_path}")
        return False

    normalized_path = os.path.normpath(file_path)
    updated = False

    for field, tag in SUGGESTION_TEXT_FIELDS.items():
        if not accepted.get(field):
            continue
        value = suggestion.get(field, "")
        if value is None or str(value) == "":
            continue
        try:
            # set_tag_value saves the file itself (consistent with inline edits).
            if set_tag_value(audio_file, tag, str(value)):
                updated = True
                log_fn(f"[SUCCESS] {field} -> '{value}' for {os.path.basename(file_path)}")
            else:
                log_fn(f"[ERROR] Failed to set {field} for {os.path.basename(file_path)}")
        except Exception as e:
            log_fn(f"[ERROR] Failed to set {field}: {e}")

    if accepted.get("art"):
        cover_url = suggestion.get("cover_image") or suggestion.get("thumb")
        if cover_url and apply_cover_art(audio_file, file_path, cover_url,
                                         api_token=api_token, log_fn=log_fn):
            updated = True

    if updated:
        mark_updated(normalized_path)
    mark_processed(normalized_path)
    return updated
