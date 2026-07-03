"""
Genre autocomplete dictionary.

Reads a plain-text list of genres (one per line, `#`-prefixed lines are comments)
from Config.GENRES_FILE and provides a case-insensitive prefix-match completion
helper used by the Genre entry ghost-text autocomplete in the sidepane.
"""

import os
from pathlib import Path

import mutagen

from config import Config
from utils.metadata import get_tag_value


_STARTER_CONTENT = (
    "# Phonodex genre dictionary\n"
    "# One genre per line. Lines starting with '#' are comments and are ignored.\n"
    "# Blank lines are ignored. Case is preserved from this file when a\n"
    "# completion is offered, so 'Deep House' typed as 'deep h' will complete\n"
    "# to 'Deep House'.\n"
    "\n"
    "# Example entries (delete or replace with your own):\n"
    "# Ambient\n"
    "# Techno\n"
    "# House\n"
    "# Deep House\n"
    "# Jazz\n"
)


def load_genres():
    """Load genre dictionary from disk.

    Returns a list of unique genres sorted alphabetically (case-insensitive).
    Creates the file with a starter template on first run if it does not exist.
    Never raises -- returns [] on any read/parse error.
    """
    path = Path(Config.GENRES_FILE)
    try:
        if not path.exists():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(_STARTER_CONTENT, encoding="utf-8")
            return []

        genres = set()
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if line and not line.startswith("#"):
                genres.add(line)
        return sorted(genres, key=str.lower)
    except Exception:
        return []


def find_completion(prefix, genres):
    """Return the alphabetically-first genre that starts with `prefix`
    (case-insensitive) and is strictly longer than `prefix`, or None.
    Empty prefix returns None.
    """
    if not prefix:
        return None
    p = prefix.lower()
    for g in genres:
        if len(g) > len(prefix) and g.lower().startswith(p):
            return g
    return None


def find_all_completions(prefix, genres, limit=50, exclude=None):
    """Return matching genres for the dropdown, ordered as:
      1. Prefix matches (case-insensitive `startswith(prefix)`), including
         exact matches so the entry the user just finished typing stays
         visible. These are the candidates the ghost-text picks from.
      2. Substring matches (`prefix` appears somewhere other than the start).
         These enrich the dropdown but are NOT used as ghost suggestions
         (ghost-text only makes sense for prefix completions).

    Args:
        prefix: what the user has typed in the current segment.
        genres: the loaded dictionary list (already alphabetical).
        limit: cap on total returned items.
        exclude: optional iterable of genre strings already present in
            other segments of the entry -- those are filtered out so the
            dropdown doesn't suggest duplicates. Matching is
            case-insensitive.

    Returns a flat list (prefix matches first, then substring matches).
    Empty prefix returns [].
    """
    if not prefix:
        return []
    p = prefix.lower()
    excluded_lower = {e.lower() for e in exclude} if exclude else set()

    prefix_matches = []
    substring_matches = []
    for g in genres:
        gl = g.lower()
        if gl in excluded_lower:
            continue
        if gl.startswith(p):
            if len(g) >= len(prefix):
                prefix_matches.append(g)
        elif p in gl:
            substring_matches.append(g)

    combined = prefix_matches + substring_matches
    return combined[:limit]


def find_prefix_completions(prefix, genres, exclude=None):
    """Return only genres whose lowercased form starts with `prefix`,
    excluding any in `exclude`. Used by the ghost-text logic which must
    only ever pick prefix completions.
    """
    if not prefix:
        return []
    p = prefix.lower()
    excluded_lower = {e.lower() for e in exclude} if exclude else set()
    return [g for g in genres if g.lower() not in excluded_lower and g.lower().startswith(p)]


def _split_genre_value(value):
    """Split a raw genre tag value into individual genre strings.

    Splits on ';' (Phonodex convention) as well as '/' and '\\' which some
    taggers use. Strips whitespace and drops empties.
    """
    if not value:
        return []
    parts = [value]
    for sep in (";", "/", "\\"):
        parts = [p for chunk in parts for p in chunk.split(sep)]
    return [p.strip() for p in parts if p and p.strip()]


def scan_folder_for_genres(folder_path, supported_extensions, progress_callback=None):
    """Recursively scan a folder for audio files and collect all unique genres.

    Args:
        folder_path: Root folder to scan.
        supported_extensions: Iterable of lowercase extensions including the dot
            (e.g. ('.mp3', '.flac', ...)).
        progress_callback: Optional callable(scanned_count, current_path, total_count)
            invoked periodically. `total_count` is the total number of matching
            audio files discovered during the initial count pass -- known before
            any file is opened, so the caller can size a determinate progress
            bar. May be None.

    Returns:
        A tuple (genres_set, scanned_count, error_count) where genres_set is a
        set of unique genre strings found on disk (case preserved from tags).
    """
    genres = set()
    scanned = 0
    errors = 0
    exts = tuple(e.lower() for e in supported_extensions)

    # Pre-count pass: fast (only stat, no file open) and lets the UI show
    # a real determinate progress bar instead of a bouncing indeterminate one.
    total = 0
    for root, _dirs, files in os.walk(folder_path):
        for fname in files:
            if fname.lower().endswith(exts):
                total += 1

    # Emit an initial callback with the total so the UI can configure its
    # progress bar's maximum before the first file is actually processed.
    if progress_callback is not None:
        try:
            progress_callback(0, "", total)
        except Exception:
            pass

    for root, _dirs, files in os.walk(folder_path):
        for fname in files:
            if not fname.lower().endswith(exts):
                continue
            full = os.path.join(root, fname)
            scanned += 1
            try:
                audio = mutagen.File(full)
                if audio is None:
                    continue
                raw = get_tag_value(audio, "genre", "")
                for g in _split_genre_value(raw):
                    genres.add(g)
            except Exception:
                errors += 1
            if progress_callback is not None and (scanned % 25 == 0 or scanned == total):
                try:
                    progress_callback(scanned, full, total)
                except Exception:
                    pass

    return genres, scanned, errors


def merge_genres_into_file(new_genres):
    """Append any genres not already present in Config.GENRES_FILE.

    Preserves the existing file contents (comments, ordering, blank lines).
    Comparison is case-insensitive against existing non-comment entries so we
    don't create duplicates that differ only in casing. New genres are appended
    under a dated section header and sorted alphabetically (case-insensitive)
    within that section.

    Returns:
        A tuple (added_list, skipped_count) where added_list is the sorted
        list of genres actually written and skipped_count is how many were
        already known.
    """
    from datetime import datetime

    path = Path(Config.GENRES_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)

    existing_lower = set()
    existing_text = ""
    if path.exists():
        existing_text = path.read_text(encoding="utf-8")
        for raw in existing_text.splitlines():
            line = raw.strip()
            if line and not line.startswith("#"):
                existing_lower.add(line.lower())
    else:
        existing_text = _STARTER_CONTENT

    to_add = []
    seen = set()
    for g in new_genres:
        key = g.lower()
        if key in existing_lower or key in seen:
            continue
        seen.add(key)
        to_add.append(g)

    skipped = len(new_genres) - len(to_add)

    if to_add:
        to_add.sort(key=str.lower)
        header = f"\n# Added by folder scan on {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        # Ensure the file ends with a newline before appending.
        if existing_text and not existing_text.endswith("\n"):
            existing_text += "\n"
        new_block = header + "\n".join(to_add) + "\n"
        path.write_text(existing_text + new_block, encoding="utf-8")

    return to_add, skipped
