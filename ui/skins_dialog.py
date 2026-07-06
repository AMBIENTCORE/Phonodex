"""
Skins selection dialog.

Each available skin is shown as a preview card: the skin's name rendered in
its own font over its own background, plus a row of color swatches from its
palette. Applying a skin calls the provided callback (which saves the choice
and restarts the app). No screenshots required -- the card *is* the preview.
"""
import tkinter as tk

from config import Config, SKINS

# Palette keys shown as swatches, in order. SUCCESS/ERROR are intentionally
# excluded -- they are fixed semantic status colors (green = updated,
# red = failed), not part of a skin's aesthetic identity.
_SWATCH_KEYS = [
    "BACKGROUND",
    "SECONDARY_BACKGROUND",
    "BORDER",
    "TEXT",
    "SECONDARY_TEXT",
]


def show_skins_dialog(parent, current_skin, skin_fonts, on_apply):
    """Open the skins picker.

    Args:
        parent: the root window.
        current_skin: name of the currently active skin.
        skin_fonts: dict mapping skin name -> tk font family (for previews).
        on_apply: callback(skin_name) invoked when the user applies a skin.
    """
    dlg = tk.Toplevel(parent)
    dlg.title("Skins")
    dlg.configure(bg=Config.COLORS["BACKGROUND"])
    dlg.transient(parent)
    dlg.resizable(False, False)
    try:
        dlg.grab_set()
    except Exception:
        pass

    header_font = skin_fonts.get(current_skin, "TkDefaultFont")
    tk.Label(
        dlg, text="PICK A SKIN",
        bg=Config.COLORS["BACKGROUND"], fg=Config.COLORS["TEXT"],
        font=(header_font, 14),
    ).pack(padx=16, pady=(14, 2), anchor="w")

    tk.Label(
        dlg, text="Applying a skin will restart Phonodex.",
        bg=Config.COLORS["BACKGROUND"], fg=Config.COLORS["SECONDARY_TEXT"],
        font=("Consolas", 8),
    ).pack(padx=16, pady=(0, 10), anchor="w")

    # Scrollable card list -- there can be many skins, so cap the height and
    # let the user scroll rather than growing the window off-screen.
    list_wrap = tk.Frame(dlg, bg=Config.COLORS["BACKGROUND"])
    list_wrap.pack(padx=16, pady=(0, 8), fill="both", expand=True)

    canvas = tk.Canvas(
        list_wrap, bg=Config.COLORS["BACKGROUND"], highlightthickness=0,
        bd=0, width=860, height=460,
    )
    vbar = tk.Scrollbar(
        list_wrap, orient="vertical", command=canvas.yview,
        troughcolor=Config.COLORS["BACKGROUND"], bg=Config.COLORS["BORDER"],
        activebackground=Config.COLORS["SECONDARY_TEXT"], bd=0,
        highlightthickness=0,
    )
    canvas.configure(yscrollcommand=vbar.set)
    vbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    cards = tk.Frame(canvas, bg=Config.COLORS["BACKGROUND"])
    cards_window = canvas.create_window((0, 0), window=cards, anchor="nw")

    def _on_cards_configure(_event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    cards.bind("<Configure>", _on_cards_configure)

    def _on_canvas_configure(event):
        # Make the inner frame track the canvas width so cards fill it.
        canvas.itemconfigure(cards_window, width=event.width)

    canvas.bind("<Configure>", _on_canvas_configure)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _on_close():
        # Drop the global wheel binding so it doesn't linger after close.
        try:
            canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        dlg.destroy()

    dlg.protocol("WM_DELETE_WINDOW", _on_close)

    # Lay the cards out 4 per row (equal-width columns) so the picker is wide
    # and short rather than a tall single column.
    _COLS = 4
    for col in range(_COLS):
        cards.columnconfigure(col, weight=1, uniform="skincol")

    for idx, (name, skin) in enumerate(SKINS.items()):
        colors = skin.get("colors", {})
        bg = colors.get("BACKGROUND", "#1e1e1e")
        fg = colors.get("TEXT", "#ffffff")
        border = colors.get("BORDER", "#3e3e42")
        # Skin accent: an explicit per-skin color (its Pantone). NOT SUCCESS,
        # which is a fixed semantic green. Falls back to a neutral if unset.
        accent = skin.get("accent") or colors.get("ACCENT") or fg
        fam = skin_fonts.get(name, "TkDefaultFont")
        is_current = (name == current_skin)

        card = tk.Frame(
            cards, bg=bg, highlightthickness=2,
            highlightbackground=(accent if is_current else border),
        )
        card.grid(row=idx // _COLS, column=idx % _COLS, padx=6, pady=6, sticky="nsew")

        # Skin name in its own font, over its own background.
        tk.Label(
            card, text=skin.get("display_name", name),
            bg=bg, fg=fg, font=(fam, 16), wraplength=185, justify="left",
        ).pack(anchor="w", padx=12, pady=(10, 6))

        # Swatch row.
        sw = tk.Frame(card, bg=bg)
        sw.pack(anchor="w", padx=14, pady=(0, 10))
        swatch_colors = [colors.get(k) for k in _SWATCH_KEYS]
        # Append the skin's accent (its Pantone) as the final swatch, if set.
        if skin.get("accent") or colors.get("ACCENT"):
            swatch_colors.append(skin.get("accent") or colors.get("ACCENT"))
        for c in swatch_colors:
            if not c:
                continue
            chip = tk.Frame(sw, bg=c, width=22, height=16,
                            highlightthickness=1, highlightbackground=border)
            chip.pack(side="left", padx=(0, 3))
            chip.pack_propagate(False)

        # Apply / current row.
        row = tk.Frame(card, bg=bg)
        row.pack(fill="x", padx=14, pady=(0, 12))
        if is_current:
            tk.Label(row, text="Current", bg=bg, fg=accent,
                     font=("Consolas", 9, "bold")).pack(side="right")
        else:
            tk.Button(
                row, text="Apply", command=lambda n=name: on_apply(n),
                bg=colors.get("SECONDARY_BACKGROUND", "#252526"), fg=fg,
                activebackground=border, activeforeground=fg,
                relief="solid", bd=1, padx=12, pady=2, font=("Consolas", 9),
                cursor="hand2",
            ).pack(side="right")

    # Size the scroll area to the content (capped) so a short skin list gives a
    # compact window and a long one still scrolls.
    cards.update_idletasks()
    canvas.configure(height=min(cards.winfo_reqheight(), 620))

    tk.Button(
        dlg, text="Close", command=_on_close,
        bg=Config.COLORS["SECONDARY_BACKGROUND"], fg=Config.COLORS["TEXT"],
        relief="solid", bd=1, padx=12, pady=2, font=("Consolas", 9),
        cursor="hand2",
    ).pack(pady=(2, 14))

    # Center over the parent window.
    dlg.update_idletasks()
    try:
        px = parent.winfo_rootx() + (parent.winfo_width() - dlg.winfo_width()) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(px, 0)}+{max(py, 0)}")
    except Exception:
        pass

    return dlg
