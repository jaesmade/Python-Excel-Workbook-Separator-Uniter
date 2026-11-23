"""Review and remove remembered spelling choices."""

import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from excel_splitter.gui import styles


class SavedSpellingsDialog(ctk.CTkToplevel):
    def __init__(self, app, store):
        choices = store.load()
        super().__init__(app)
        self.store = store
        self.choices = choices
        self.title("Saved spellings")
        self.geometry("700x430")
        self.minsize(480, 300)
        self.configure(fg_color=styles.BACKGROUND)
        self.transient(app)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        heading = ctk.CTkFrame(self, fg_color="transparent")
        heading.grid(row=0, column=0, columnspan=2, padx=24, pady=(22, 14), sticky="ew")
        ctk.CTkLabel(heading, text="Saved spellings", text_color=styles.TEXT,
                     font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(heading, text="Column: original → preferred. Removing a choice affects future workbooks.",
                     text_color=styles.MUTED, font=ctk.CTkFont(family="Segoe UI", size=12),
                     wraplength=650, justify="left").pack(anchor="w", pady=(4, 0))
        self.choices_list = tk.Listbox(self, exportselection=False, font=("Segoe UI", 11),
                                       activestyle="none", **styles.listbox_colors())
        self.choices_list.grid(row=1, column=0, padx=(24, 0), sticky="nsew")
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.choices_list.yview)
        scroll.grid(row=1, column=1, padx=(0, 24), sticky="ns")
        self.choices_list.configure(yscrollcommand=scroll.set)
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, columnspan=2, padx=24, pady=18, sticky="ew")
        ctk.CTkButton(footer, text="Remove selected", command=self._remove, height=36, corner_radius=9,
                      font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                      **styles.PRIMARY_BUTTON).pack(side="left")
        ctk.CTkButton(footer, text="Close", command=self.destroy, height=36, corner_radius=9,
                      font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                      **styles.SECONDARY_BUTTON).pack(side="right")
        self._refresh()

    def _refresh(self):
        self.keys = sorted(self.choices)
        self.choices_list.delete(0, tk.END)
        for column, source in self.keys:
            self.choices_list.insert(tk.END, f"{column}: {source} → {self.choices[(column, source)]}")

    def _remove(self):
        selection = self.choices_list.curselection()
        if not selection:
            return
        column, source = self.keys[selection[0]]
        try:
            self.store.remove(column, source)
        except (OSError, ValueError) as exc:
            messagebox.showwarning("Saved spellings", f"Choice could not be removed: {exc}", parent=self)
            return
        del self.choices[(column, source)]
        self._refresh()
