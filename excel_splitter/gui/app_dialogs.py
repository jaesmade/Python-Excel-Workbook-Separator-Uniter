"""Preferences and help for the desktop workspace."""

import json
import os
import tempfile
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

from excel_splitter.gui import styles


APPEARANCE_MODES = ("System", "Light", "Dark")
PAGE_SIZES = (25, 50, 100)
PROJECT_URL = "https://github.com/jaesmade/Python-Excel-Workbook-Separator-Uniter"


def open_project_link(parent, title, url):
    try:
        if webbrowser.open_new_tab(url):
            return
    except (OSError, webbrowser.Error):
        pass
    messagebox.showinfo(title, f"Open this address in your browser:\n\n{url}", parent=parent)


@dataclass(frozen=True)
class AppPreferences:
    appearance: str = "Light"
    page_size: int = 50

    @classmethod
    def load(cls, path: Path):
        defaults = cls()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return defaults
        if not isinstance(data, dict):
            return defaults

        appearance = data.get("appearance")
        page_size = data.get("page_size")
        return cls(
            appearance=appearance if appearance in APPEARANCE_MODES else defaults.appearance,
            page_size=page_size if type(page_size) is int and page_size in PAGE_SIZES else defaults.page_size,
        )

    def save(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", dir=path.parent, delete=False
            ) as handle:
                temporary = Path(handle.name)
                json.dump({"appearance": self.appearance, "page_size": self.page_size}, handle, indent=2)
                handle.write("\n")
            os.replace(temporary, path)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, app):
        super().__init__(app, fg_color=styles.SURFACE)
        self.app = app
        self.title("Settings")
        self.transient(app)
        self.resizable(False, False)
        self.grid_columnconfigure(1, weight=1)
        self.bind("<Escape>", lambda _event: self.destroy())

        ctk.CTkLabel(self, text="Settings", text_color=styles.TEXT, font=ctk.CTkFont(size=24, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=24, pady=(20, 4), sticky="w"
        )
        ctk.CTkLabel(self, text="Preferences are saved on this device.", text_color=styles.MUTED).grid(
            row=1, column=0, columnspan=2, padx=24, pady=(0, 14), sticky="w"
        )
        self.appearance_var = ctk.StringVar(master=self, value=app.preferences.appearance)
        self.page_size_var = ctk.StringVar(master=self, value=str(app.preferences.page_size))
        for row, (label, variable, values) in enumerate((
            ("Appearance", self.appearance_var, list(APPEARANCE_MODES)),
            ("Groups per page", self.page_size_var, [str(size) for size in PAGE_SIZES]),
        ), start=2):
            ctk.CTkLabel(self, text=label, text_color=styles.TEXT).grid(
                row=row, column=0, padx=(24, 32), pady=7, sticky="w"
            )
            ctk.CTkOptionMenu(self, values=values, variable=variable, width=160, height=34, **styles.OPTION).grid(
                row=row, column=1, padx=(0, 24), pady=7, sticky="e"
            )

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=4, column=0, columnspan=2, padx=24, pady=(16, 22), sticky="e")
        ctk.CTkButton(actions, text="Cancel", width=82, height=34, command=self.destroy,
                      **styles.SECONDARY_BUTTON).pack(side="left", padx=(0, 8))
        self.save_button = ctk.CTkButton(actions, text="Save settings", width=116, height=34,
                                        command=self._save, **styles.PRIMARY_BUTTON)
        self.save_button.pack(side="left")
        self.grab_set()

    def _save(self):
        preferences = AppPreferences(self.appearance_var.get(), int(self.page_size_var.get()))
        try:
            preferences.save(self.app.preferences_path)
        except OSError as exc:
            messagebox.showwarning("Settings", f"Could not save your preferences.\n\n{exc}", parent=self)
            return
        self.destroy()
        self.app.apply_preferences(preferences)


class HelpDialog(ctk.CTkToplevel):
    def __init__(self, app):
        super().__init__(app, fg_color=styles.SURFACE)
        self.title("Excel Arranger Help")
        self.transient(app)
        self.resizable(False, False)
        self.bind("<Escape>", lambda _event: self.destroy())

        ctk.CTkLabel(self, text="Excel Arranger Help", text_color=styles.TEXT,
                     font=ctk.CTkFont(size=24, weight="bold")).pack(
            anchor="w", padx=24, pady=(20, 10)
        )
        for title, instructions in (
            ("Open your source files", "Use Browse files or Ctrl+O to open Excel or CSV files. Choose a worksheet, or use All sheets to combine your sources."),
            ("Choose how to group rows", "Select one or more grouping columns. Use the review tools to check similar spellings or possible duplicate records."),
            ("Review and export", "Select the groups you want. Preview worksheets to check their contents, then generate and save a new workbook."),
        ):
            ctk.CTkLabel(self, text=title, text_color=styles.TEXT,
                         font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=24)
            ctk.CTkLabel(self, text=instructions, text_color=styles.MUTED, wraplength=410, justify="left").pack(
                anchor="w", padx=24, pady=(0, 10)
            )

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.pack(fill="x", padx=24, pady=(8, 22))
        ctk.CTkButton(actions, text="Close", width=82, height=34, command=self.destroy,
                      **styles.SECONDARY_BUTTON).pack(side="right")
        self.grab_set()
