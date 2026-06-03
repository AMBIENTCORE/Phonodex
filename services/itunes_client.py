"""
iTunes Search API client for album art lookup.

The iTunes Search API is free, requires no authentication, and returns
high-quality (1200x1200) front cover art that is almost always the real
album cover rather than a vinyl photo. Used as the preferred album-art
source, with Discogs as the fallback.

Docs: https://performance-partners.apple.com/search-api
"""

import re
import unicodedata
import urllib.parse
import requests

from utils.logging import log_message

ITUNES_SEARCH_URL = "https://itunes.apple.com/search"
REQUEST_TIMEOUT = 10  # seconds


def _normalize(text):
    """Lowercase, strip accents, collapse whitespace, drop punctuation for matching."""
    if not text:
        return ""
    # NFKD splits accented chars into base + combining marks; we then drop the marks.
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.lower()
    # Drop common bracketed qualifiers like "(remastered)", "[deluxe edition]" for matching
    text = re.sub(r"[\(\[][^)\]]*[\)\]]", " ", text)
    # Strip "the " prefix which is wildly inconsistent across catalogs
    text = re.sub(r"^the\s+", "", text)
    # Reduce non-alphanumerics to spaces
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def _upgrade_artwork_url(url, target_size=1200):
    """
    iTunes returns 100x100 thumbnails via `artworkUrl100`. The URL ends in
    something like `.../source/100x100bb.jpg`. Replacing the size token
    yields the high-resolution image straight from Apple's CDN.
    """
    if not url:
        return url
    return re.sub(r"/\d+x\d+bb\.(jpg|png)", f"/{target_size}x{target_size}bb.\\1", url)


def _score_match(result, norm_artist, norm_album):
    """Score how well an iTunes result matches the requested artist/album."""
    score = 0.0
    r_artist = _normalize(result.get("artistName", ""))
    r_album = _normalize(result.get("collectionName", ""))

    if not r_album:
        return float("-inf")

    # Album name match (most important)
    if r_album == norm_album:
        score += 10.0
    elif norm_album and norm_album in r_album:
        score += 5.0
    elif norm_album and r_album in norm_album:
        score += 3.0
    else:
        score -= 2.0

    # Artist match
    if norm_artist:
        if r_artist == norm_artist:
            score += 6.0
        elif norm_artist in r_artist or r_artist in norm_artist:
            score += 3.0
        else:
            score -= 4.0

    # Mildly de-prefer obvious compilations / best-of when an exact match exists
    suspicious_terms = ("greatest hits", "best of", "very best", "essential")
    if any(term in r_album for term in suspicious_terms):
        score -= 1.5

    return score


def fetch_album_art_url(artist, album, limit=10, target_size=1200):
    """
    Look up album art on the iTunes Search API.

    Args:
        artist: Artist name.
        album: Album name.
        limit: Max results to fetch from iTunes (default 10).
        target_size: Desired pixel size for the high-res image (default 1200).

    Returns:
        (cover_url, thumb_url) tuple. Both are None if no good match found.
    """
    if not album:
        return None, None

    norm_artist = _normalize(artist)
    norm_album = _normalize(album)

    term = urllib.parse.quote_plus(f"{artist or ''} {album}".strip())
    url = f"{ITUNES_SEARCH_URL}?term={term}&entity=album&limit={limit}"

    try:
        response = requests.get(url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        results = response.json().get("results", [])
    except requests.exceptions.RequestException as e:
        log_message(f"[COVER][iTunes] Request failed: {e}", log_type="debug")
        return None, None
    except ValueError as e:
        log_message(f"[COVER][iTunes] Bad JSON response: {e}", log_type="debug")
        return None, None

    if not results:
        log_message(f"[COVER][iTunes] No results for '{artist} - {album}'", log_type="debug")
        return None, None

    # Score every result and pick the best
    best = None
    best_score = float("-inf")
    for result in results:
        score = _score_match(result, norm_artist, norm_album)
        if score > best_score:
            best = result
            best_score = score

    # Require a reasonable confidence threshold so we don't slap random art on unmatched albums.
    # 8.0 ~= album substring match + artist substring match. Below this, fall back to Discogs.
    if best is None or best_score < 8.0:
        log_message(
            f"[COVER][iTunes] Best match score {best_score:.1f} below threshold for "
            f"'{artist} - {album}' (got '{best.get('artistName') if best else None}' - "
            f"'{best.get('collectionName') if best else None}')",
            log_type="debug"
        )
        return None, None

    thumb_url = best.get("artworkUrl100") or best.get("artworkUrl60")
    if not thumb_url:
        return None, None

    cover_url = _upgrade_artwork_url(thumb_url, target_size=target_size)
    log_message(
        f"[COVER][iTunes] Matched '{best.get('artistName')}' - '{best.get('collectionName')}' "
        f"(score {best_score:.1f}) -> {cover_url}",
        log_type="debug"
    )
    return cover_url, thumb_url
