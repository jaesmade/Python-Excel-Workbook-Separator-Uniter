"""Appearance-aware colors and widget styles for the desktop app."""

import tkinter as tk
from tkinter import ttk

import customtkinter as ctk


BACKGROUND = ("#EFF4F2", "#131D1A")
SURFACE = ("#FFFFFF", "#1B2924")
INSET = ("#F3F7F5", "#23352E")
BORDER = ("#DBE6E0", "#3A5046")
TEXT = ("#142A23", "#EBF4EE")
MUTED = ("#4F675D", "#ABBBB2")
ACCENT = ("#08734F", "#18815E")
ACCENT_HOVER = ("#075A40", "#13694C")
SELECTED = ("#D8EDE3", "#324B3E")

CARD = dict(fg_color=SURFACE, border_color=BORDER, border_width=1, corner_radius=14)
PRIMARY_BUTTON = dict(fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color="#FFFFFF")
SECONDARY_BUTTON = dict(fg_color=INSET, hover_color=BORDER, text_color=TEXT)
FIELD = dict(fg_color=INSET, border_color=BORDER, text_color=TEXT)
OPTION = dict(
    fg_color=INSET, button_color=BORDER, button_hover_color=SELECTED, text_color=TEXT,
    dropdown_fg_color=SURFACE, dropdown_hover_color=INSET, dropdown_text_color=TEXT,
)
CHECKBOX = dict(
    fg_color=ACCENT, hover_color=ACCENT_HOVER, border_color=BORDER, text_color=TEXT,
    checkmark_color="#FFFFFF", checkbox_width=18, checkbox_height=18, border_width=2,
)


class WrappingCheckBox(ctk.CTkCheckBox):
    """Wrap the built-in label so the whole group name stays clickable."""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        tk.Frame.bind(self, "<Configure>", self._wrap_label, add="+")

    def _wrap_label(self, event):
        # Tk's grid and native label both use physical pixels.
        indicator_width = self.grid_bbox(0, 0, 1, 0)[2]
        width = max(1, event.width - indicator_width)
        if int(self._text_label.cget("wraplength")) != width:
            self._text_label.configure(wraplength=width)


def native_color(colors):
    return colors[ctk.get_appearance_mode() == "Dark"]


def listbox_colors():
    return dict(
        bg=native_color(SURFACE), fg=native_color(TEXT),
        selectbackground=native_color(SELECTED), selectforeground=native_color(TEXT),
        relief="flat", highlightthickness=1, highlightbackground=native_color(BORDER),
    )


def table_style(widget):
    style = ttk.Style(widget)
    style.configure(
        "Arranger.Treeview", background=native_color(SURFACE), fieldbackground=native_color(SURFACE),
        foreground=native_color(TEXT), rowheight=27, font=("Segoe UI", 10),
    )
    style.configure(
        "Arranger.Treeview.Heading", background=native_color(INSET), foreground=native_color(TEXT),
        font=("Segoe UI", 10, "bold"), padding=(7, 5),
    )
    style.map(
        "Arranger.Treeview", background=[("selected", native_color(SELECTED))],
        foreground=[("selected", native_color(TEXT))],
    )
    return "Arranger.Treeview"
