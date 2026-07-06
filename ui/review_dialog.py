"""
Release-review dialog for the "smart process" (identify & tag) flow.

Presents the release matches gathered from Discogs grouped per album, with a
per-file / per-field accept toggle so the user reviews every suggested change
before anything is written. This dialog is display-only: it renders the
prepared ``groups`` structure and flips the per-field ``accept`` booleans in
place. Writing is done by the caller's ``on_confirm`` callback.

groups: list of group dicts, each:
    {
        "album_key": str,
        "display_name": str,                 # header label for the album group
        "candidates": [candidate, ...],       # ranked Discogs releases
        "selected_index": int,                # index into candidates
        "release_data": dict | None,          # details for the selected candidate
        "files": [ filerec, ... ],
    }
filerec:
    {
        "file_path": str,
        "suggestion": {field: value, ...},    # raw values (incl cover_image/thumb)
        "changes": [field, ...],              # fields that differ / are fillable
        "display": {field: (current, suggested)},
        "accept": {field: bool},
    }
"""

import io
import os
import threading
import tkinter as tk
from tkinter import ttk

from config import Config

try:
    from PIL import Image, ImageTk
    _PIL_OK = True
except Exception:  # pragma: no cover - PIL is a hard dep elsewhere, but be safe
    _PIL_OK = False

# Accept-state glyphs. Plain ASCII so they render in any font (the tree uses
# Consolas, which lacks the fancy ballot-box code points).
GLYPH_ON = "[x]"
GLYPH_OFF = "[ ]"
GLYPH_PARTIAL = "[~]"
EM_DASH = "\u2014"  # kept out of f-string expressions for older-Python portability

# Display order + human labels for the fields the tagger can write.
FIELD_ORDER = ["artist", "title", "album", "albumartist",
               "tracknumber", "year", "catalognumber", "art"]
FIELD_LABELS = {
    "artist": "Artist",
    "title": "Title",
    "album": "Album",
    "albumartist": "Album Artist",
    "tracknumber": "Track",
    "year": "Year",
    "catalognumber": "Catalog Number",
    "art": "Album Art",
}


def _candidate_label(candidate):
    """One-line description of a release candidate for the match combobox."""
    score = int(round(candidate.get("score", 0) * 100))
    title = candidate.get("title") or f"{candidate.get('artist','')} - {candidate.get('album','')}"
    year = candidate.get("year") or "?"
    catno = candidate.get("catno") or "no catalog"
    return f"{score:>3}%  {title}  ({year})  [{catno}]"


def show_review_dialog(parent, custom_font, groups, on_confirm,
                       on_rematch=None, load_thumb=None):
    """Show the modal review dialog.

    Args:
        parent: parent Tk window.
        custom_font: chrome font used elsewhere in the app.
        groups: the prepared groups structure (mutated in place as the user toggles).
        on_confirm: callable(groups) invoked when the user confirms writing.
        on_rematch: optional callable(group, candidate_index) -> bool, run on a
            worker thread to rebuild a group's suggestions for a different release.
        load_thumb: optional callable(url) -> image bytes | None for cover previews.
    """
    colors = Config.COLORS
    data_font = ("Consolas", Config.FONTS["TABLE_SIZE"])

    dialog = tk.Toplevel(parent)
    dialog.title("Review Release Matches")
    dialog.geometry("1280x760")
    dialog.minsize(900, 520)
    dialog.resizable(True, True)
    dialog.configure(bg=colors["BACKGROUND"])
    dialog.grab_set()

    # ---- header / intro -------------------------------------------------
    total_files = sum(len(g["files"]) for g in groups)
    total_changes = sum(len(f["changes"]) for g in groups for f in g["files"])
    tk.Label(
        dialog,
        text=f"Reviewing {total_files} file(s) across {len(groups)} album group(s) "
             f"\u2014 {total_changes} suggested change(s).",
        font=custom_font, bg=colors["BACKGROUND"], fg=colors["TEXT"],
    ).pack(pady=(10, 4), padx=10, anchor="w")
    tk.Label(
        dialog,
        text="Ticked rows will be written. Empty fields are pre-ticked; overwrites of "
             "existing values start unticked. Click a row's checkbox to toggle it.",
        font=data_font, bg=colors["BACKGROUND"], fg=colors["SECONDARY_TEXT"],
    ).pack(pady=(0, 8), padx=10, anchor="w")

    body = ttk.Frame(dialog)
    body.pack(fill="both", expand=True, padx=10, pady=(0, 6))

    # ---- right: match preview panel (packed first so the tree fills rest) --
    preview = ttk.Frame(body, width=320)
    preview.pack(side="right", fill="y", padx=(8, 0))
    preview.pack_propagate(False)

    tk.Label(preview, text="MATCHED RELEASE", font=custom_font,
             bg=colors["BACKGROUND"], fg=colors["TEXT"]).pack(anchor="w", pady=(0, 6))

    # Fixed-pixel holder so the preview reserves 280x280 regardless of whether an
    # image is loaded (a bare tk.Label sizes in text units, not pixels).
    cover_holder = tk.Frame(preview, width=280, height=280, bg=colors["SECONDARY_BACKGROUND"])
    cover_holder.pack(pady=(0, 8))
    cover_holder.pack_propagate(False)
    cover_label = tk.Label(cover_holder, bg=colors["SECONDARY_BACKGROUND"],
                           fg=colors["SECONDARY_TEXT"])
    cover_label.pack(fill="both", expand=True)

    info_var = tk.StringVar(value="Select an album group to preview its match.")
    tk.Label(preview, textvariable=info_var, font=data_font, justify="left",
             wraplength=300, bg=colors["BACKGROUND"], fg=colors["SECONDARY_TEXT"]).pack(
        anchor="w", pady=(0, 8), fill="x")

    tk.Label(preview, text="Change match:", font=data_font,
             bg=colors["BACKGROUND"], fg=colors["TEXT"]).pack(anchor="w")
    match_var = tk.StringVar()
    match_combo = ttk.Combobox(preview, textvariable=match_var, state="readonly")
    match_combo.pack(fill="x", pady=(2, 0))

    # ---- left: the tree -------------------------------------------------
    tree_frame = ttk.Frame(body)
    tree_frame.pack(side="left", fill="both", expand=True)

    tree = ttk.Treeview(
        tree_frame,
        columns=("accept", "current", "suggested"),
        show="tree headings",
        selectmode="browse",
    )
    tree.heading("#0", text="Album / File / Field")
    tree.heading("accept", text="Write")
    tree.heading("current", text="Current")
    tree.heading("suggested", text="Suggested")
    tree.column("#0", width=360, stretch=True)
    tree.column("accept", width=60, anchor="center", stretch=False)
    tree.column("current", width=260, stretch=True)
    tree.column("suggested", width=260, stretch=True)

    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    tree_frame.rowconfigure(0, weight=1)
    tree_frame.columnconfigure(0, weight=1)

    # Field rows are colored by accept state: accepted (will be written) uses the
    # green "updated" row color; not-accepted uses the red "failed" row color so
    # pending overwrites stand out. Album headers stay neutral.
    tree.tag_configure("accepted", background=colors["UPDATED_ROW"])
    tree.tag_configure("rejected", background=colors["FAILED_ROW"])
    tree.tag_configure("album", background=colors["SECONDARY_BACKGROUND"])

    # Node bookkeeping: item id -> (kind, gidx, fidx, field)
    node_meta = {}
    album_items = {}
    file_items = {}
    field_items = {}

    # ---- state helpers --------------------------------------------------
    def _file_state(gidx, fidx):
        filerec = groups[gidx]["files"][fidx]
        changes = filerec["changes"]
        if not changes:
            return "off"
        vals = [bool(filerec["accept"].get(fld)) for fld in changes]
        if all(vals):
            return "on"
        if not any(vals):
            return "off"
        return "partial"

    def _album_state(gidx):
        files = groups[gidx]["files"]
        states = [_file_state(gidx, fidx) for fidx in range(len(files))]
        if not states or all(s == "off" for s in states):
            return "off"
        if all(s == "on" for s in states):
            return "on"
        return "partial"

    def _glyph(state):
        return {"on": GLYPH_ON, "off": GLYPH_OFF, "partial": GLYPH_PARTIAL}[state]

    def _refresh_glyphs(gidx, fidx=None):
        """Update the accept-column glyphs for a file (and its album), or a whole album."""
        if fidx is not None:
            filerec = groups[gidx]["files"][fidx]
            for field in filerec["changes"]:
                item = field_items.get((gidx, fidx, field))
                if item:
                    on = bool(filerec["accept"].get(field))
                    tree.set(item, "accept", GLYPH_ON if on else GLYPH_OFF)
                    tree.item(item, tags=("accepted" if on else "rejected",))
            fi = file_items.get((gidx, fidx))
            if fi:
                tree.set(fi, "accept", _glyph(_file_state(gidx, fidx)))
        else:
            for fj in range(len(groups[gidx]["files"])):
                _refresh_glyphs(gidx, fj)
        ai = album_items.get(gidx)
        if ai:
            tree.set(ai, "accept", _glyph(_album_state(gidx)))

    # ---- tree building --------------------------------------------------
    def _build_group_children(gidx):
        group = groups[gidx]
        album_item = album_items[gidx]
        for fidx, filerec in enumerate(group["files"]):
            fname = os.path.basename(filerec["file_path"])
            file_item = tree.insert(
                album_item, "end", text=fname,
                values=(_glyph(_file_state(gidx, fidx)), "", ""),
                open=True,
            )
            node_meta[file_item] = ("file", gidx, fidx, None)
            file_items[(gidx, fidx)] = file_item

            if not filerec["changes"]:
                note = tree.insert(file_item, "end", text="(no changes)",
                                   values=("", "", ""))
                node_meta[note] = ("note", gidx, fidx, None)
                continue

            for field in FIELD_ORDER:
                if field not in filerec["changes"]:
                    continue
                current, suggested = filerec["display"].get(field, ("", ""))
                on = bool(filerec["accept"].get(field))
                field_item = tree.insert(
                    file_item, "end", text=FIELD_LABELS.get(field, field),
                    values=(GLYPH_ON if on else GLYPH_OFF, current, suggested),
                    tags=("accepted" if on else "rejected",),
                )
                node_meta[field_item] = ("field", gidx, fidx, field)
                field_items[(gidx, fidx, field)] = field_item

    def _render_all():
        tree.delete(*tree.get_children())
        node_meta.clear()
        album_items.clear()
        file_items.clear()
        field_items.clear()
        for gidx, group in enumerate(groups):
            conf = 0
            if group["candidates"]:
                conf = int(round(group["candidates"][group["selected_index"]]["score"] * 100))
            label = f"{group['display_name']}   ({len(group['files'])} file(s), {conf}% match)"
            if not group["candidates"]:
                label = f"{group['display_name']}   (no match found)"
            album_item = tree.insert(
                "", "end", text=label,
                values=(_glyph(_album_state(gidx)), "", ""),
                open=True, tags=("album",),
            )
            node_meta[album_item] = ("album", gidx, None, None)
            album_items[gidx] = album_item
            _build_group_children(gidx)

    # ---- preview panel --------------------------------------------------
    _cover_imgref = {"img": None}

    def _selected_group_index():
        sel = tree.selection()
        if not sel:
            return None
        meta = node_meta.get(sel[0])
        return meta[1] if meta else None

    def _set_cover_from_bytes(data):
        if not (_PIL_OK and data):
            cover_label.configure(image="", text="(no preview)", fg=colors["SECONDARY_TEXT"])
            _cover_imgref["img"] = None
            return
        try:
            img = Image.open(io.BytesIO(data)).convert("RGB")
            img = img.resize((280, 280), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            _cover_imgref["img"] = photo
            cover_label.configure(image=photo, text="")
        except Exception:
            cover_label.configure(image="", text="(no preview)", fg=colors["SECONDARY_TEXT"])
            _cover_imgref["img"] = None

    def _load_cover_async(gidx):
        group = groups[gidx]
        if not group["candidates"]:
            _set_cover_from_bytes(None)
            return
        cand = group["candidates"][group["selected_index"]]
        url = ""
        if group["files"]:
            sugg = group["files"][0]["suggestion"]
            url = sugg.get("thumb") or sugg.get("cover_image") or ""
        url = url or cand.get("thumb") or cand.get("cover_image") or ""
        if not url or not load_thumb:
            _set_cover_from_bytes(None)
            return

        def worker():
            data = None
            try:
                data = load_thumb(url)
            except Exception:
                data = None
            dialog.after(0, lambda: _set_cover_from_bytes(data))

        cover_label.configure(image="", text="Loading\u2026", fg=colors["SECONDARY_TEXT"])
        threading.Thread(target=worker, daemon=True).start()

    def _refresh_preview():
        gidx = _selected_group_index()
        if gidx is None:
            info_var.set("Select an album group to preview its match.")
            match_combo.configure(values=[])
            match_var.set("")
            _set_cover_from_bytes(None)
            return
        group = groups[gidx]
        cands = group["candidates"]
        if not cands:
            info_var.set(f"{group['display_name']}\nNo Discogs match was found.")
            match_combo.configure(values=[])
            match_var.set("")
            _set_cover_from_bytes(None)
            return
        cand = cands[group["selected_index"]]
        catno_text = cand.get('catno') or EM_DASH
        info_var.set(
            f"{cand.get('title','')}\n"
            f"Year: {cand.get('year') or '?'}\n"
            f"Catalog: {catno_text}\n"
            f"Confidence: {int(round(cand.get('score',0)*100))}%   "
            f"(via {cand.get('strategy','?')})"
        )
        labels = [_candidate_label(c) for c in cands]
        match_combo.configure(values=labels)
        # Widen the dropdown to fit the longest entry (capped so it stays sane).
        if labels:
            match_combo.configure(width=min(100, max(len(lbl) for lbl in labels)))
        match_combo.current(group["selected_index"])
        _load_cover_async(gidx)

    # ---- interactions ---------------------------------------------------
    def _on_tree_click(event):
        if tree.identify_region(event.x, event.y) != "cell":
            return
        if tree.identify_column(event.x) != "#1":  # only the "accept" column toggles
            return
        item = tree.identify_row(event.y)
        meta = node_meta.get(item)
        if not meta:
            return
        kind, gidx, fidx, field = meta
        if kind == "field":
            filerec = groups[gidx]["files"][fidx]
            filerec["accept"][field] = not filerec["accept"].get(field)
            _refresh_glyphs(gidx, fidx)
        elif kind == "file":
            new_on = _file_state(gidx, fidx) != "on"
            filerec = groups[gidx]["files"][fidx]
            for fld in filerec["changes"]:
                filerec["accept"][fld] = new_on
            _refresh_glyphs(gidx, fidx)
        elif kind == "album":
            new_on = _album_state(gidx) != "on"
            for fj, filerec in enumerate(groups[gidx]["files"]):
                for fld in filerec["changes"]:
                    filerec["accept"][fld] = new_on
            _refresh_glyphs(gidx)

    def _set_all(state):
        for gidx, group in enumerate(groups):
            for filerec in group["files"]:
                for fld in filerec["changes"]:
                    filerec["accept"][fld] = state
            _refresh_glyphs(gidx)

    def _on_match_change(event=None):
        gidx = _selected_group_index()
        if gidx is None or not on_rematch:
            return
        new_index = match_combo.current()
        if new_index < 0 or new_index == groups[gidx]["selected_index"]:
            return
        match_combo.configure(state="disabled")
        info_var.set("Loading new match\u2026")

        def worker():
            ok = False
            try:
                ok = on_rematch(groups[gidx], new_index)
            except Exception:
                ok = False

            def finish():
                match_combo.configure(state="readonly")
                if ok:
                    # Rebuild just this album's subtree.
                    album_item = album_items.get(gidx)
                    if album_item:
                        for child in tree.get_children(album_item):
                            _purge_node(child)
                        conf = int(round(
                            groups[gidx]["candidates"][groups[gidx]["selected_index"]]["score"] * 100))
                        tree.item(album_item, text=f"{groups[gidx]['display_name']}   "
                                                   f"({len(groups[gidx]['files'])} file(s), {conf}% match)")
                        _build_group_children(gidx)
                        _refresh_glyphs(gidx)
                _refresh_preview()

            dialog.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _purge_node(item):
        """Recursively drop a subtree from the tree and bookkeeping maps."""
        for child in tree.get_children(item):
            _purge_node(child)
        meta = node_meta.pop(item, None)
        if meta:
            kind, gidx, fidx, field = meta
            if kind == "file":
                file_items.pop((gidx, fidx), None)
            elif kind == "field":
                field_items.pop((gidx, fidx, field), None)
        tree.delete(item)

    tree.bind("<Button-1>", _on_tree_click, add="+")
    tree.bind("<<TreeviewSelect>>", lambda e: _refresh_preview())
    match_combo.bind("<<ComboboxSelected>>", _on_match_change)

    # ---- bottom buttons -------------------------------------------------
    button_frame = ttk.Frame(dialog)
    button_frame.pack(fill="x", padx=10, pady=10, side="bottom")

    def _confirm():
        dialog.destroy()
        on_confirm(groups)

    tk.Button(button_frame, text="CANCEL", command=dialog.destroy,
              font=custom_font, bg=colors["SECONDARY_BACKGROUND"], fg="#990000",
              padx=Config.STYLES["WIDGET_PADDING"], pady=Config.STYLES["WIDGET_PADDING"]).pack(
        side="left", padx=(0, 5))
    tk.Button(button_frame, text="ALL OFF", command=lambda: _set_all(False),
              font=custom_font, bg=colors["SECONDARY_BACKGROUND"], fg=colors["TEXT"],
              padx=Config.STYLES["WIDGET_PADDING"], pady=Config.STYLES["WIDGET_PADDING"]).pack(
        side="left", padx=5)
    tk.Button(button_frame, text="ALL ON", command=lambda: _set_all(True),
              font=custom_font, bg=colors["SECONDARY_BACKGROUND"], fg=colors["TEXT"],
              padx=Config.STYLES["WIDGET_PADDING"], pady=Config.STYLES["WIDGET_PADDING"]).pack(
        side="left", padx=5)
    tk.Button(button_frame, text="WRITE ACCEPTED", command=_confirm,
              font=custom_font, bg=colors["SECONDARY_BACKGROUND"], fg="#006400",
              padx=Config.STYLES["WIDGET_PADDING"], pady=Config.STYLES["WIDGET_PADDING"]).pack(
        side="right", padx=(5, 0))

    _render_all()
    # Select the first album so the preview populates immediately.
    first = tree.get_children()
    if first:
        tree.selection_set(first[0])
        tree.focus(first[0])
    _refresh_preview()

    parent.wait_window(dialog)
