"""
Matching utilities for release identification.

Pure-stdlib helpers (no third-party deps) used by the release-matching
pipeline: text normalization, fuzzy similarity, year/track parsing, Discogs
"Artist - Title" splitting, filename parsing, and candidate scoring.

Keeping this dependency-free means it bundles cleanly with PyInstaller and can
be imported from both the metadata engine and the review dialog without
pulling in network or GUI code.
"""

import os
import re
import unicodedata
from difflib import SequenceMatcher


def normalize(text):
    """Lowercase, strip accents, drop bracketed qualifiers and punctuation.

    Mirrors the normalizer used by the iTunes cover matcher so scoring is
    consistent across the app: NFKD accent stripping, removal of "(remastered)"
    / "[deluxe]" style qualifiers, a leading "the ", and reduction of any
    remaining non-alphanumerics to single spaces.
    """
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.lower()
    text = re.sub(r"[\(\[][^)\]]*[\)\]]", " ", text)  # (remastered) [deluxe edition]
    text = re.sub(r"^the\s+", "", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def similarity(a, b):
    """Return a 0..1 similarity score combining sequence ratio and token overlap.

    Uses the higher of a character-sequence ratio and a token Jaccard overlap so
    that reordered or partially-matching titles ("Greatest Hits" vs "The Greatest
    Hits") still score well without any external fuzzy-match dependency.
    """
    na, nb = normalize(a), normalize(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    seq = SequenceMatcher(None, na, nb).ratio()
    ta, tb = set(na.split()), set(nb.split())
    token = (len(ta & tb) / len(ta | tb)) if (ta and tb) else 0.0
    return max(seq, token)


def extract_year(value):
    """Return the first 19xx/20xx year found in ``value`` as a string, else ''."""
    if not value:
        return ""
    match = re.search(r"(?:19|20)\d{2}", str(value))
    return match.group(0) if match else ""


def extract_track_number(value):
    """Return the leading integer of a track/position value as a string.

    Handles "4", "04", "4/12", "A1", "1-1" -> "4"/"1" etc. Vinyl side letters
    (the "A" in "A1") are ignored; the numeric component is what we align on.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if "/" in s:
        s = s.split("/", 1)[0].strip()
    match = re.search(r"\d+", s)
    return str(int(match.group(0))) if match else ""


def split_discogs_title(title):
    """Split a Discogs "Artist - Album" title into ``(artist, album)``.

    Everything after the first " - " is treated as the album, matching how the
    existing metadata code interprets Discogs release titles.
    """
    if not title:
        return "", ""
    if " - " in title:
        artist, album = title.split(" - ", 1)
        return artist.strip(), album.strip()
    return "", title.strip()


def strip_discogs_artist_suffix(name):
    """Remove Discogs' numeric disambiguation suffix, e.g. "Nirvana (2)" -> "Nirvana"."""
    if not name:
        return ""
    return re.sub(r"\s*\(\d+\)\s*$", "", str(name)).strip()


_TRACK_PREFIX = re.compile(r"^\s*(\d{1,3})[\s._\-\)]+")


def parse_filename(file_path):
    """Best-effort extraction of ``{artist, title, album, tracknumber}`` from a path.

    Handles common patterns: a leading track number ("04 - ", "04. ", "04 "),
    an "Artist - Title" stem, and the parent folder as an album guess (stripping
    a leading "[catalog] " and a trailing "(year)"). Any field that can't be
    inferred is returned as an empty string. This is only a fallback for files
    whose tags are essentially empty.
    """
    result = {"artist": "", "title": "", "album": "", "tracknumber": ""}
    stem = os.path.splitext(os.path.basename(file_path))[0].strip()

    match = _TRACK_PREFIX.match(stem)
    if match:
        result["tracknumber"] = str(int(match.group(1)))
        stem = stem[match.end():].strip()

    if " - " in stem:
        left, right = stem.split(" - ", 1)
        result["artist"] = left.strip()
        result["title"] = right.strip()
    else:
        result["title"] = stem

    parent = os.path.basename(os.path.dirname(file_path)).strip()
    if parent:
        parent = re.sub(r"^\[[^\]]*\]\s*", "", parent)          # [CAT-001] prefix
        parent = re.sub(r"\s*\((?:19|20)\d{2}\)\s*$", "", parent)  # (2001) suffix
        if " - " in parent:
            _, album_part = parent.split(" - ", 1)
            result["album"] = album_part.strip()
        else:
            result["album"] = parent
    return result


def score_release_candidate(known, release):
    """Score a Discogs search result against the known file tags (0..1 confidence).

    ``known`` may contain any of: artist, albumartist, album, title, year,
    catalognumber. ``release`` is a Discogs /database/search result dict whose
    ``title`` is "Artist - Album". Only the fields we actually know contribute to
    the score (weighted average over present signals), so a file with just an
    album still gets a meaningful confidence, and a catalog-number hit scores
    very high.
    """
    r_artist, r_album = split_discogs_title(release.get("title", ""))
    score = 0.0
    weight = 0.0

    known_artist = known.get("artist") or known.get("albumartist") or ""
    if known_artist:
        weight += 3.0
        score += 3.0 * similarity(known_artist, r_artist)

    known_album = known.get("album") or ""
    if known_album:
        weight += 4.0
        score += 4.0 * similarity(known_album, r_album)
    elif known.get("title"):
        # No album tag: the track title often loosely matches the release/album.
        weight += 1.5
        score += 1.5 * similarity(known.get("title"), r_album)

    known_year = extract_year(known.get("year"))
    r_year = extract_year(release.get("year"))
    if known_year and r_year:
        weight += 1.0
        score += 1.0 if known_year == r_year else 0.0

    known_catno = normalize(known.get("catalognumber")).replace(" ", "")
    r_catno = normalize(release.get("catno")).replace(" ", "")
    # "none" is the app's placeholder for "this release has no catalog number",
    # not a real value. Treat it as absent so it neither matches nor penalizes
    # actual catalog numbers (otherwise a stale "NONE" tag scores cataloged
    # releases *lower* than uncataloged ones).
    if known_catno == "none":
        known_catno = ""
    if r_catno == "none":
        r_catno = ""
    if known_catno and r_catno:
        weight += 2.0
        score += 2.0 if known_catno == r_catno else 0.0

    if weight == 0.0:
        return 0.0
    return round(score / weight, 3)
