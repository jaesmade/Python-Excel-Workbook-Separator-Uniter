from __future__ import annotations

import os
import queue
import shutil
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

from excel_splitter.core import (
    ALL_SHEETS,
    SORT_ALPHABETICAL,
    SORT_COUNT_ASCENDING,
    SORT_COUNT_DESCENDING,
    WorkbookProcessingError,
    _iter_input_tables,
    _normalize_input_tables,
    build_combined_group_map,
    ensure_output_is_not_input,
    merge_group_values,
    preview_group_items,
    sort_group_items,
    split_input_tables,
    source_file_labels,
)


from excel_splitter.spelling_review import KEEP_SEPARATE, SpellingReview
from excel_splitter.manual_merge import category_values, merge_categories
from excel_splitter.saved_spellings import SavedSpellings
from excel_splitter.gui import styles
from excel_splitter.gui.app_dialogs import AppPreferences, HelpDialog, SettingsDialog
from excel_splitter.gui.saved_spellings_dialog import SavedSpellingsDialog
from excel_splitter.gui.duplicate_records_dialog import DuplicateRecordsDialog


class ManualMergeDialog(ctk.CTkToplevel):
    """Search two values in one grouping column and merge them."""

    def __init__(self, app, columns):
        super().__init__(app)
        self.app = app
        self.columns = tuple(columns)
        self.title("Manually merge categories")
        self.geometry("720x570")
        self.minsize(560, 450)
        self.transient(app)
        self.grab_set()
        self.grid_columnconfigure((0, 1), weight=1)
        self.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(self, text="Manually merge categories",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(
                         row=0, column=0, columnspan=2, padx=18, pady=(16, 4), sticky="w")
        ctk.CTkLabel(self, text="Select a column and search for two values. The source stays unchanged; export can optionally correct merged values.",
                     wraplength=670, justify="left").grid(
                         row=1, column=0, columnspan=2, padx=18, pady=(0, 8), sticky="w")
        self.column_var = ctk.StringVar(value=self.columns[0])
        ctk.CTkOptionMenu(self, values=list(self.columns), variable=self.column_var,
                          command=self._column_changed).grid(
                              row=2, column=0, columnspan=2, padx=18, pady=(0, 8), sticky="ew")
        self.search_vars = []
        self.value_lists = []
        for index, label in enumerate(("First value", "Second value")):
            frame = ctk.CTkFrame(self)
            frame.grid(row=3, column=index, padx=(18 if index == 0 else 6, 18 if index == 1 else 6),
                       pady=4, sticky="nsew")
            frame.grid_columnconfigure(0, weight=1)
            frame.grid_rowconfigure(2, weight=1)
            ctk.CTkLabel(frame, text=label, anchor="w").grid(row=0, column=0, padx=8, pady=(8, 2), sticky="ew")
            search = ctk.StringVar(value="")
            self.search_vars.append(search)
            ctk.CTkEntry(frame, textvariable=search, placeholder_text="Search values").grid(
                row=1, column=0, padx=8, pady=(0, 6), sticky="ew")
            values = tk.Listbox(frame, exportselection=False, selectmode="browse",
                                font=("Segoe UI", 10), **styles.listbox_colors())
            values.grid(row=2, column=0, padx=(8, 0), sticky="nsew")
            scroll = ttk.Scrollbar(frame, orient="vertical", command=values.yview)
            scroll.grid(row=2, column=1, padx=(0, 8), sticky="ns")
            horizontal_scroll = ttk.Scrollbar(frame, orient="horizontal", command=values.xview)
            horizontal_scroll.grid(row=3, column=0, padx=(8, 0), pady=(0, 8), sticky="ew")
            values.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal_scroll.set)
            self.value_lists.append(values)
            search.trace_add("write", lambda *_args, i=index: self._refresh_values(i))

        ctk.CTkLabel(self, text="Preferred name (optional; defaults to the first value)", anchor="w").grid(
            row=4, column=0, columnspan=2, padx=18, pady=(8, 2), sticky="ew")
        self.preferred_var = ctk.StringVar(value="")
        ctk.CTkEntry(self, textvariable=self.preferred_var).grid(
            row=5, column=0, columnspan=2, padx=18, pady=(0, 8), sticky="ew")
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=6, column=0, columnspan=2, padx=18, pady=(0, 16), sticky="ew")
        ctk.CTkButton(footer, text="Cancel", command=self.destroy).pack(side="right", padx=(8, 0))
        ctk.CTkButton(footer, text="Merge", command=self._merge).pack(side="right")
        self._column_changed()

    def _column_changed(self, _value=None):
        for index in range(2):
            self._refresh_values(index)

    def _refresh_values(self, index):
        values = category_values(self.app.groups, self.columns, self.column_var.get(), self.search_vars[index].get())
        box = self.value_lists[index]
        box.delete(0, tk.END)
        for value in values:
            box.insert(tk.END, value)

    def _merge(self):
        selected = [box.curselection() for box in self.value_lists]
        if not all(selected):
            messagebox.showinfo("Manual merge", "Select one value from each list.", parent=self)
            return
        first, second = (box.get(selection[0]) for box, selection in zip(self.value_lists, selected))
        preferred = self.preferred_var.get().strip() or first
        try:
            groups, mappings, selected_groups = merge_categories(
                self.app.groups, self.columns, self.app.value_mappings, self.app.selected_group_keys,
                self.column_var.get(), first, second, preferred)
        except ValueError as exc:
            messagebox.showinfo("Manual merge", str(exc), parent=self)
            return
        self.app.groups = groups
        self.app.value_mappings = mappings
        self.app._last_spelling_review = None
        self.app.group_page = 0
        self.app._render_groups(selected_groups)
        count = sum(len(rows) for rows in groups.values())
        self.app.progress_var.set(f"{len(groups)} groups detected from {count:,} records.")
        self.destroy()


class SpellingReviewDialog(ctk.CTkToplevel):
    """Review all suggested pairs before applying the chosen merges."""

    def __init__(self, app, review):
        super().__init__(app)
        self.app = app
        self.review = review
        self.title("Review spelling")
        self.geometry("760x620")
        self.minsize(620, 400)
        self.transient(app)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(self, text="Review spelling suggestions",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=18, pady=(16, 4), sticky="w")
        ctk.CTkLabel(self, text="Choose a spelling to merge each pair, or keep it separate. The source stays unchanged; export can optionally correct merged values.",
                     wraplength=700, justify="left").grid(row=1, column=0, padx=18, pady=(0, 8), sticky="w")
        self.list_frame = ctk.CTkScrollableFrame(self)
        self.list_frame.grid(row=2, column=0, padx=18, pady=4, sticky="nsew")
        self.list_frame.grid_columnconfigure(0, weight=1)
        self.suggestion_rows = []
        self._row_controls = []
        for index, (column, source, target, source_count, target_count) in enumerate(self.review.suggestions):
            row = ctk.CTkFrame(self.list_frame)
            row.grid(row=index, column=0, padx=4, pady=5, sticky="ew")
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(row, text=column, font=ctk.CTkFont(weight="bold"), anchor="w").grid(
                row=0, column=0, padx=12, pady=(8, 2), sticky="ew")
            ctk.CTkLabel(row, text=f'First: "{source}" ({source_count:,} records)\nSecond: "{target}" ({target_count:,} records)',
                         anchor="w", justify="left", wraplength=650).grid(row=1, column=0, padx=12, sticky="ew")
            buttons = ctk.CTkFrame(row, fg_color="transparent")
            buttons.grid(row=2, column=0, padx=8, pady=(6, 8), sticky="w")
            controls = []
            for text, choice in (("Use first", source), ("Use second", target), ("Keep separate", KEEP_SEPARATE)):
                button = ctk.CTkButton(buttons, text=text, width=120,
                                       command=lambda i=index, value=choice: self._choose(i, value))
                button.pack(side="left", padx=4)
                controls.append(button)
            status = ctk.CTkLabel(row, text="Awaiting choice", anchor="w")
            status.grid(row=3, column=0, padx=12, pady=(0, 8), sticky="ew")
            self.suggestion_rows.append(row)
            self._row_controls.append((controls, status))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=3, column=0, padx=18, pady=(8, 16), sticky="ew")
        self.undo_button = ctk.CTkButton(footer, text="Undo last choice", command=self._undo, state="disabled")
        self.undo_button.pack(side="left")
        ctk.CTkButton(footer, text="Cancel", command=self.destroy).pack(side="right", padx=(8, 0))
        ctk.CTkButton(footer, text="Apply choices", command=self._apply).pack(side="right")
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._refresh_choices()

    def _refresh_choices(self):
        for index, (controls, status) in enumerate(self._row_controls):
            answer = self.review.decisions.get(index)
            for button in controls:
                button.configure(state="disabled" if answer is not None else "normal")
            if answer is KEEP_SEPARATE:
                label = "Keeping separate"
            elif answer is not None:
                column = self.review.suggestions[index][0]
                label = f'Merged as "{self.review.current_value(column, answer)}"'
            else:
                label = "Awaiting choice"
            status.configure(text=label)
        self.undo_button.configure(state="normal" if self.review.can_undo else "disabled")

    def _choose(self, index, answer):
        try:
            self.review.choose(index, answer)
        except ValueError as exc:
            messagebox.showinfo("Review spelling", str(exc), parent=self)
            return
        self._refresh_choices()

    def _undo(self):
        self.review.undo()
        self._refresh_choices()

    def _apply(self):
        def approved(review):
            if review is None or not review._history:
                return {}
            initial = review._history[0][2]
            return {key: value for key, value in review.mappings.items()
                    if initial.get(key) != value}

        previous = approved(self.app._last_spelling_review)
        current = approved(self.review)
        additions = {key: value for key, value in current.items()
                     if previous.get(key) != value}
        try:
            stored = self.app.saved_spellings.load()
            removals = [key for key, value in previous.items()
                        if key not in current and stored.get(key) == value]
            if additions or removals:
                self.app.saved_spellings.update(additions, removals)
        except (OSError, ValueError) as exc:
            messagebox.showwarning("Saved spellings", f"Choices could not be saved for future workbooks: {exc}", parent=self)
        self.app.groups = self.review.groups
        self.app.value_mappings = self.review.mappings
        self.app._last_spelling_review = self.review
        self.app.group_page = 0
        self.app._render_groups(self.review.selected)
        count = sum(len(rows) for rows in self.app.groups.values())
        self.app.progress_var.set(f"{len(self.app.groups)} groups detected from {count:,} records.")
        self.destroy()


class ExcelSplitterApp(ctk.CTk):
    def __init__(self):
        ctk.set_default_color_theme("green")
        super().__init__(fg_color=styles.BACKGROUND)
        self.title("Excel Arranger")
        if sys.platform.startswith("win"):
            self.iconbitmap(str(Path(__file__).resolve().parents[1] / "assets" / "app.ico"))
        self.geometry("1080x730")
        self.minsize(920, 650)
        ttk.Style(self).theme_use("clam")

        self.file_path = None
        self.file_paths = []
        self.workbook = {}
        self.input_tables = []
        self.sheet_options = {}
        self.selected_sheet_name = None
        self.headers = []
        self.groups = {}
        self.group_checkboxes = {}
        self.selected_group_keys = set()
        self.group_page = 0
        self.group_column_controls = []
        self.value_mappings = {}
        self._last_spelling_review = None
        self.saved_spellings = SavedSpellings()
        self.preferences_path = self.saved_spellings.path.with_name("preferences.json")
        self.preferences = AppPreferences.load(self.preferences_path)
        ctk.set_appearance_mode(self.preferences.appearance)

        self._build_ui()
        self.bind("<Control-o>", self._open_files_shortcut)
        self.bind("<F1>", self._help_shortcut)

    def _open_files_shortcut(self, _event):
        self.browse_file()
        return "break"

    def _help_shortcut(self, _event):
        self.show_help()
        return "break"

    def show_settings(self):
        self.wait_window(SettingsDialog(self))

    def show_help(self):
        self.wait_window(HelpDialog(self))

    def apply_preferences(self, preferences):
        page_size_changed = self.preferences.page_size != preferences.page_size
        self.preferences = preferences
        ctk.set_appearance_mode(preferences.appearance)
        if self.groups and page_size_changed:
            self.group_page = 0
            self._render_groups(self.selected_group_keys)

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_header()
        self._build_setup_panel()
        self._build_groups_panel()
        self._build_export_bar()
        self._render_groups()

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="#103A30", corner_radius=0)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            header, text="▦", width=40, height=40, corner_radius=9, fg_color="#20664F",
            text_color="#F1FFF7", font=ctk.CTkFont(family="Segoe UI Symbol", size=25),
        ).grid(row=0, column=0, rowspan=2, padx=(22, 12), pady=15)
        ctk.CTkLabel(
            header, text="Excel Arranger", text_color="#F1FFF7", font=ctk.CTkFont(size=23, weight="bold"),
        ).grid(row=0, column=1, pady=(11, 0), sticky="sw")
        ctk.CTkLabel(
            header, text="Split and combine Excel or CSV files", text_color="#B6D7C8", height=18,
        ).grid(row=1, column=1, pady=(0, 13), sticky="nw")
        self.settings_button = ctk.CTkButton(
            header, text="Settings", width=90, height=34, command=self.show_settings,
            fg_color="#205447", hover_color="#2B6958", text_color="#F1FFF7",
        )
        self.settings_button.grid(row=0, column=2, rowspan=2, padx=(10, 8))
        self.help_button = ctk.CTkButton(
            header, text="Help", width=72, height=34, command=self.show_help,
            fg_color="#205447", hover_color="#2B6958", text_color="#F1FFF7",
        )
        self.help_button.grid(row=0, column=3, rowspan=2, padx=(0, 22))

    def _build_setup_panel(self):
        panel = ctk.CTkFrame(self, **styles.CARD)
        panel.grid(row=1, column=0, padx=(20, 12), pady=16, sticky="nsew")
        panel.grid_columnconfigure((0, 1), weight=1, uniform="setup")

        ctk.CTkLabel(panel, text="Source files", text_color=styles.TEXT,
                     font=ctk.CTkFont(size=17, weight="bold")).grid(
            row=0, column=0, padx=(16, 0), pady=(15, 8), sticky="w"
        )
        ctk.CTkButton(
            panel, text="Browse files", command=self.browse_file, width=105, height=32, **styles.PRIMARY_BUTTON,
        ).grid(row=0, column=1, padx=(0, 16), pady=(15, 8), sticky="e")
        self.file_var = ctk.StringVar(value="No files selected")
        self.file_entry = ctk.CTkEntry(
            panel, textvariable=self.file_var, state="readonly", height=34, **styles.FIELD,
        )
        self.file_entry.grid(row=1, column=0, columnspan=2, padx=16, sticky="ew")

        ctk.CTkLabel(panel, text="Worksheet", height=20, text_color=styles.MUTED).grid(
            row=2, column=0, padx=16, pady=(6, 2), sticky="w"
        )
        self.sheet_var = ctk.StringVar(value="")
        self.sheet_dropdown = ctk.CTkOptionMenu(
            panel, values=[], variable=self.sheet_var, command=self.on_sheet_selected,
            height=34, dynamic_resizing=False, **styles.OPTION,
        )
        self.sheet_dropdown.grid(row=3, column=0, columnspan=2, padx=16, sticky="ew")
        ctk.CTkFrame(panel, height=1, fg_color=styles.BORDER).grid(
            row=4, column=0, columnspan=2, padx=16, pady=(10, 8), sticky="ew"
        )

        ctk.CTkLabel(panel, text="Group by", text_color=styles.TEXT,
                     font=ctk.CTkFont(size=17, weight="bold")).grid(
            row=5, column=0, padx=(16, 0), pady=(0, 6), sticky="w"
        )
        self.add_column_button = ctk.CTkButton(
            panel, text="+ Add column", command=self._add_group_column, width=105, height=29,
            **styles.SECONDARY_BUTTON,
        )
        self.add_column_button.grid(row=5, column=1, padx=(0, 16), pady=(0, 6), sticky="e")
        group_column_viewport = ctk.CTkFrame(panel, height=88, fg_color="transparent")
        group_column_viewport.grid(row=6, column=0, columnspan=2, padx=16, sticky="ew")
        group_column_viewport.grid_propagate(False)
        group_column_viewport.grid_columnconfigure(0, weight=1)
        group_column_viewport.grid_rowconfigure(0, weight=1)
        self.group_columns_frame = ctk.CTkScrollableFrame(
            group_column_viewport, fg_color=styles.INSET, corner_radius=8,
            scrollbar_button_color=styles.BORDER, scrollbar_button_hover_color=styles.MUTED,
        )
        self.group_columns_frame.grid(row=0, column=0, sticky="nsew")
        self.group_columns_frame.grid_columnconfigure(0, weight=1)
        self._add_group_column()

        self.combine_duplicates_var = ctk.BooleanVar(value=True)
        self.combine_duplicates_checkbox = ctk.CTkCheckBox(
            panel, text="Combine duplicate groups", variable=self.combine_duplicates_var,
            command=self.on_column_selection_changed, onvalue=True, offvalue=False, **styles.CHECKBOX,
        )
        self.combine_duplicates_checkbox.grid(
            row=7, column=0, columnspan=2, padx=16, pady=(7, 0), sticky="w"
        )
        self.correct_merged_values_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            panel, text="Correct merged values in output", variable=self.correct_merged_values_var,
            onvalue=True, offvalue=False, **styles.CHECKBOX,
        ).grid(row=8, column=0, columnspan=2, padx=16, pady=(1, 7), sticky="w")

        ctk.CTkFrame(panel, height=1, fg_color=styles.BORDER).grid(
            row=9, column=0, columnspan=2, padx=16, pady=(1, 6), sticky="ew"
        )
        ctk.CTkLabel(panel, text="Review tools", text_color=styles.TEXT,
                     font=ctk.CTkFont(size=17, weight="bold")).grid(
            row=10, column=0, columnspan=2, padx=16, pady=(0, 6), sticky="w"
        )
        for index, (label, command) in enumerate((
            ("Review spelling", self.review_spelling_suggestions),
            ("Manual merge", self.manual_merge_categories),
            ("Saved spellings", self.manage_saved_spellings),
            ("Review duplicate records", self.review_duplicate_records),
        )):
            column = index % 2
            ctk.CTkButton(panel, text=label, command=command, width=148, height=30,
                          **styles.SECONDARY_BUTTON).grid(
                row=11 + index // 2, column=column, padx=(16, 4) if column == 0 else (4, 16),
                pady=(0, 5), sticky="ew"
            )

    def _build_groups_panel(self):
        panel = ctk.CTkFrame(self, **styles.CARD)
        panel.grid(row=1, column=1, padx=(0, 20), pady=16, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(panel, text="Detected groups", text_color=styles.TEXT,
                     font=ctk.CTkFont(size=20, weight="bold")).grid(
            row=0, column=0, padx=18, pady=(14, 0), sticky="w"
        )
        self.group_selection_var = ctk.StringVar(value="")
        ctk.CTkLabel(panel, textvariable=self.group_selection_var, text_color=styles.MUTED, height=20).grid(
            row=1, column=0, padx=18, pady=(0, 9), sticky="w"
        )
        self.group_sort_var = ctk.StringVar(value="Alphabetical")
        self.group_sort_dropdown = ctk.CTkOptionMenu(
            panel, values=["Alphabetical", "Record count: ascending", "Record count: descending"],
            variable=self.group_sort_var, command=self._on_group_sort_changed,
            width=190, height=34, dynamic_resizing=False, **styles.OPTION,
        )
        self.group_sort_dropdown.grid(row=0, column=1, rowspan=2, padx=(8, 18), pady=(12, 7), sticky="e")

        tools = ctk.CTkFrame(panel, fg_color="transparent")
        tools.grid(row=2, column=0, columnspan=2, padx=18, pady=(0, 10), sticky="ew")
        tools.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(tools, text="Search groups", text_color=styles.MUTED).grid(
            row=0, column=0, padx=(0, 8), sticky="w"
        )
        self.group_search_var = ctk.StringVar(value="")
        self.group_search_entry = ctk.CTkEntry(
            tools, textvariable=self.group_search_var, height=34, **styles.FIELD,
        )
        self.group_search_entry.grid(row=0, column=1, padx=(0, 8), sticky="ew")
        self.group_select_all = ctk.CTkButton(
            tools, text="Select all matches", width=115, height=34,
            command=lambda: self._select_matching_groups(True), **styles.SECONDARY_BUTTON,
        )
        self.group_select_all.grid(row=0, column=2, padx=(0, 6))
        self.group_deselect_all = ctk.CTkButton(
            tools, text="Deselect all matches", width=130, height=34,
            command=lambda: self._select_matching_groups(False), **styles.SECONDARY_BUTTON,
        )
        self.group_deselect_all.grid(row=0, column=3)

        columns = ctk.CTkFrame(panel, fg_color=styles.INSET, corner_radius=8)
        columns.grid(row=3, column=0, columnspan=2, padx=18, sticky="ew")
        columns.grid_columnconfigure(0, weight=1)
        for index, label in enumerate(("Group", "Records")):
            ctk.CTkLabel(columns, text=label, text_color=styles.MUTED,
                         font=ctk.CTkFont(size=12, weight="bold")).grid(
                row=0, column=index, padx=(14, 30), pady=3, sticky="w" if index == 0 else "e"
            )

        self.groups_frame = ctk.CTkScrollableFrame(
            panel, fg_color=styles.SURFACE, scrollbar_button_color=styles.BORDER,
            scrollbar_button_hover_color=styles.MUTED, height=100,
        )
        self.groups_frame.grid(row=4, column=0, columnspan=2, padx=18, pady=(3, 6), sticky="nsew")
        self.groups_frame.grid_columnconfigure(0, weight=1)

        pages = ctk.CTkFrame(panel, fg_color="transparent")
        pages.grid(row=5, column=0, columnspan=2, padx=18, pady=(0, 13), sticky="ew")
        pages.grid_columnconfigure(0, weight=1)
        self.group_page_var = ctk.StringVar(value="")
        ctk.CTkLabel(pages, textvariable=self.group_page_var, text_color=styles.MUTED).grid(
            row=0, column=0, sticky="w"
        )
        self.group_previous = ctk.CTkButton(
            pages, text="Previous", width=78, command=lambda: self._change_group_page(-1),
            **styles.SECONDARY_BUTTON,
        )
        self.group_previous.grid(row=0, column=1, padx=(8, 6))
        self.group_next = ctk.CTkButton(
            pages, text="Next", width=64, command=lambda: self._change_group_page(1),
            **styles.SECONDARY_BUTTON,
        )
        self.group_next.grid(row=0, column=2)
        self.group_search_var.trace_add("write", self._on_group_search_changed)

    def _build_export_bar(self):
        bottom = ctk.CTkFrame(self, **styles.CARD)
        bottom.grid(row=2, column=0, columnspan=2, padx=20, pady=(0, 16), sticky="ew")
        bottom.grid_columnconfigure(0, weight=1)
        self.progress_var = ctk.StringVar(value="Ready to open a workbook")
        self.progress_label = ctk.CTkLabel(
            bottom, textvariable=self.progress_var, text_color=styles.MUTED, anchor="w",
        )
        self.progress_label.grid(row=0, column=0, padx=(16, 20), pady=(9, 0), sticky="ew")
        self.progress_bar = ctk.CTkProgressBar(bottom, height=6, fg_color=styles.INSET, progress_color=styles.ACCENT)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=1, column=0, padx=(16, 20), pady=(0, 13), sticky="ew")

        ctk.CTkButton(
            bottom, text="Preview worksheets", command=self.preview_worksheets, width=156, height=38,
            font=ctk.CTkFont(size=13, weight="bold"), **styles.SECONDARY_BUTTON,
        ).grid(row=0, column=1, rowspan=2, padx=(0, 8), pady=11)
        self.generate_btn = ctk.CTkButton(
            bottom, text="Generate workbook", command=self.start_generation, width=166, height=38,
            font=ctk.CTkFont(size=13, weight="bold"), **styles.PRIMARY_BUTTON,
        )
        self.generate_btn.grid(row=0, column=2, rowspan=2, padx=(0, 13), pady=11)

    def preview_worksheets(self):
        selected_groups = set(self.selected_group_keys)
        if not selected_groups:
            messagebox.showinfo("Preview worksheets", "Select at least one detected group to preview.")
            return

        items = preview_group_items(
            self.groups, selected_groups, self._selected_sort_mode(), headers=self.headers,
            grouping_columns=self._selected_group_columns(), value_mappings=self.value_mappings,
            correct_merged_values=self.correct_merged_values_var.get(),
        )
        if not items:
            messagebox.showinfo("Preview worksheets", "There are no selected worksheets to preview.")
            return

        window = ctk.CTkToplevel(self)
        window.title("Preview Worksheets")
        window.geometry("1000x600")
        window.minsize(650, 400)
        window.transient(self)
        window.grab_set()
        window.grid_columnconfigure(1, weight=1)
        window.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(window, text="Worksheets", font=ctk.CTkFont(size=15, weight="bold")).grid(
            row=0, column=0, padx=16, pady=(16, 8), sticky="w"
        )
        title_var = ctk.StringVar(value="")
        ctk.CTkLabel(window, textvariable=title_var, font=ctk.CTkFont(size=15, weight="bold")).grid(
            row=0, column=1, padx=16, pady=(16, 8), sticky="w"
        )

        sheet_frame = ctk.CTkFrame(window, fg_color="transparent")
        sheet_frame.grid(row=1, column=0, padx=(16, 8), pady=(0, 12), sticky="ns")
        sheet_frame.grid_rowconfigure(0, weight=1)
        sheet_list = tk.Listbox(sheet_frame, width=26, exportselection=False, activestyle="none",
                                font=("Segoe UI", 10), **styles.listbox_colors())
        sheet_list.grid(row=0, column=0, sticky="nsew")
        sheet_scroll = ttk.Scrollbar(sheet_frame, orient="vertical", command=sheet_list.yview)
        sheet_scroll.grid(row=0, column=1, sticky="ns")
        sheet_horizontal = ttk.Scrollbar(sheet_frame, orient="horizontal", command=sheet_list.xview)
        sheet_horizontal.grid(row=1, column=0, sticky="ew")
        sheet_list.configure(yscrollcommand=sheet_scroll.set, xscrollcommand=sheet_horizontal.set)
        for name, _key, _rows in items:
            sheet_list.insert("end", name)

        table_frame = ctk.CTkFrame(window, fg_color="transparent")
        table_frame.grid(row=1, column=1, padx=(8, 16), pady=(0, 12), sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        columns = [f"column_{index}" for index in range(len(self.headers))]
        table = ttk.Treeview(table_frame, columns=columns, show="headings", style=styles.table_style(table_frame))
        for column, header in zip(columns, self.headers):
            table.heading(column, text=header)
            table.column(column, width=150, minwidth=80, stretch=False)
        table.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=table.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)

        controls = ctk.CTkFrame(window, fg_color="transparent")
        controls.grid(row=2, column=1, padx=16, pady=(0, 16), sticky="ew")
        page_var = ctk.StringVar(value="")
        page_size = 100
        state = {"sheet": 0, "page": 0}

        def render_page():
            name, _key, rows = items[state["sheet"]]
            page_count = max(1, (len(rows) + page_size - 1) // page_size)
            state["page"] = max(0, min(state["page"], page_count - 1))
            start = state["page"] * page_size
            table.delete(*table.get_children())
            for row in rows[start:start + page_size]:
                table.insert("", "end", values=["" if pd.isna(value) else str(value) for value in row])
            title_var.set(name)
            page_var.set(f"Page {state['page'] + 1} of {page_count}  ·  {len(rows):,} records")
            previous.configure(state="normal" if state["page"] else "disabled")
            following.configure(state="normal" if state["page"] + 1 < page_count else "disabled")
            table.yview_moveto(0)

        def change_page(delta):
            state["page"] += delta
            render_page()

        def select_sheet(_event):
            selection = sheet_list.curselection()
            if selection:
                state["sheet"] = selection[0]
                state["page"] = 0
                render_page()

        previous = ctk.CTkButton(controls, text="Previous", width=90, command=lambda: change_page(-1))
        previous.pack(side="left")
        ctk.CTkLabel(controls, textvariable=page_var).pack(side="left", padx=12)
        following = ctk.CTkButton(controls, text="Next", width=90, command=lambda: change_page(1))
        following.pack(side="left")
        sheet_list.bind("<<ListboxSelect>>", select_sheet)
        sheet_list.selection_set(0)
        render_page()

    def _add_group_column(self):
        row_frame = ctk.CTkFrame(self.group_columns_frame, fg_color="transparent")
        row_frame.grid(row=len(self.group_column_controls), column=0, padx=3, pady=2, sticky="ew")
        row_frame.grid_columnconfigure(0, weight=1)
        variable = ctk.StringVar(value="")
        dropdown = ctk.CTkOptionMenu(
            row_frame,
            values=self.headers or ["Select a column"],
            variable=variable,
            command=self.on_column_selection_changed,
            height=30,
            dynamic_resizing=False,
            **styles.OPTION,
        )
        dropdown.grid(row=0, column=0, sticky="ew")
        remove_button = None
        if self.group_column_controls:
            remove_button = ctk.CTkButton(
                row_frame,
                text="Remove",
                command=lambda: self._remove_group_column(row_frame, variable),
                width=70,
                height=30,
                **styles.SECONDARY_BUTTON,
            )
            remove_button.grid(row=0, column=1, padx=(8, 0))
        self.group_column_controls.append((row_frame, variable, dropdown, remove_button))
        if remove_button is not None:
            self.after_idle(self.group_columns_frame.master.yview_moveto, 1)

    def _remove_group_column(self, row_frame, variable):
        if len(self.group_column_controls) <= 1:
            return
        row_frame.destroy()
        self.group_column_controls = [control for control in self.group_column_controls if control[1] is not variable]
        for index, (frame, *_rest) in enumerate(self.group_column_controls):
            frame.grid_configure(row=index)
        self.on_column_selection_changed()

    def _selected_group_columns(self):
        return [variable.get() for _frame, variable, _dropdown, _remove in self.group_column_controls if variable.get()]

    def browse_file(self):
        filetypes = [("Excel and CSV files", "*.xlsx *.xlsm *.xls *.csv")]
        selected = filedialog.askopenfilenames(title="Select Excel or CSV files", filetypes=filetypes)
        if not selected:
            return

        self.file_paths = list(selected)
        self.file_path = self.file_paths[0]
        if len(self.file_paths) == 1:
            self.file_var.set(self.file_path)
        else:
            names = ", ".join(os.path.basename(path) for path in self.file_paths[:3])
            if len(self.file_paths) > 3:
                names += f", +{len(self.file_paths) - 3} more"
            self.file_var.set(f"{len(self.file_paths)} files selected: {names}")
        self.load_workbook(self.file_path)

    def choose_output_location(self, source_file=None):
        if source_file is None:
            source_file = self.file_paths[0] if self.file_paths else ""
        source_stem = os.path.splitext(os.path.basename(source_file))[0]
        initial_name = f"{source_stem} - SPLIT.xlsx" if source_stem else "Separated_Workbook.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save separated workbook as",
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if path:
            return path
        return None

    def load_workbook(self, file_path):
        self.progress_var.set("Loading workbook...")
        self.progress_bar.set(0)
        self.update_idletasks()
        self.workbook = {}
        self.input_tables = []
        self.sheet_options = {}
        self.sheet_var.set("")
        self.sheet_dropdown.configure(values=[])
        self._clear_worksheet_selection()

        try:
            if not self.file_paths:
                self.file_paths = [file_path]
            self.file_path = self.file_paths[0]
            input_tables = []
            workbook_tables = {}
            for source_path, file_label in zip(self.file_paths, source_file_labels(self.file_paths)):
                for sheet_name, table in _iter_input_tables(source_path, ALL_SHEETS):
                    source_label = f"{file_label} - {sheet_name}"
                    input_tables.append((source_label, table))
                    if len(self.file_paths) == 1:
                        workbook_tables[sheet_name] = table

            self.input_tables = input_tables
            self.workbook = workbook_tables
            if len(self.file_paths) > 1:
                self.sheet_options = {"All sheets": ALL_SHEETS}
                default_label = "All sheets"
            elif os.path.splitext(file_path)[1].lower() == ".csv":
                self.sheet_options = {"CSV file": "CSV file"}
                default_label = "CSV file"
            else:
                all_sheets_label = "All sheets"
                while all_sheets_label in workbook_tables:
                    all_sheets_label += " (combined)"
                self.sheet_options = {all_sheets_label: ALL_SHEETS}
                self.sheet_options.update({name: name for name in workbook_tables})
                default_label = all_sheets_label if len(workbook_tables) > 1 else next(iter(workbook_tables), all_sheets_label)

            self.sheet_dropdown.configure(values=list(self.sheet_options))
            self.sheet_var.set(default_label)
            self.on_sheet_selected(default_label)
        except Exception as exc:
            messagebox.showerror("Unable to load workbook", f"{exc}")
            self.progress_var.set("Failed to load workbook.")
            self.workbook = {}
            self.input_tables = []
            self.sheet_options = {}
            self.sheet_var.set("")
            self.sheet_dropdown.configure(values=[])
            self._clear_worksheet_selection()

    def _clear_worksheet_selection(self):
        self._last_spelling_review = None
        self.value_mappings = {}
        self.selected_sheet_name = None
        self.headers = []
        self.groups = {}
        self.selected_group_keys = set()
        self.group_search_var.set("")
        for _frame, variable, dropdown, _remove in self.group_column_controls:
            variable.set("")
            dropdown.configure(values=["Select a column"])

    def on_sheet_selected(self, sheet_name):
        if not sheet_name or not self.file_paths:
            return False

        self._last_spelling_review = None
        try:
            self.value_mappings = self.saved_spellings.load()
        except (OSError, ValueError) as exc:
            self.value_mappings = {}
            messagebox.showwarning("Saved spellings", f"Saved choices could not be loaded: {exc}", parent=self)
        self.selected_sheet_name = self.sheet_options.get(sheet_name, sheet_name)
        try:
            tables = self._selected_tables()
            if not tables:
                raise WorkbookProcessingError("The selected worksheet does not contain any data.")

            _tables, self.headers = _normalize_input_tables(tables)
            if not self.headers:
                raise WorkbookProcessingError("The selected worksheet does not contain any data.")
            for _frame, _variable, dropdown, _remove in self.group_column_controls:
                dropdown.configure(values=self.headers)

            department_default = next((h for h in self.headers if h.lower() == "department"), "")
            first_column = department_default or (self.headers[0] if self.headers else "")
            for index, (_frame, variable, _dropdown, _remove) in enumerate(self.group_column_controls):
                variable.set(first_column if index == 0 else "")

            return self._populate_groups()
        except Exception as exc:
            messagebox.showerror("Unable to process worksheet", str(exc))
            self._clear_worksheet_selection()
            self.sheet_var.set("")
            self.progress_var.set("Worksheet could not be processed.")
            return False

    def on_column_selection_changed(self, _event=None):
        if not self.file_paths or not self.selected_sheet_name:
            return
        self._populate_groups()

    def _selected_tables(self):
        if self.selected_sheet_name is ALL_SHEETS:
            return [(source_name, table) for source_name, table in self.input_tables if not table.empty]
        table = self.workbook[self.selected_sheet_name]
        if table.empty:
            return []
        source_label = f"{source_file_labels(self.file_paths)[0]} - {self.selected_sheet_name}"
        return [(source_label, table)]

    def _selected_sort_mode(self):
        return {
            "Alphabetical": SORT_ALPHABETICAL,
            "Record count: ascending": SORT_COUNT_ASCENDING,
            "Record count: descending": SORT_COUNT_DESCENDING,
        }[self.group_sort_var.get()]

    def _matching_groups(self):
        query = self.group_search_var.get().strip().casefold()
        return [item for item in sort_group_items(self.groups, self._selected_sort_mode())
                if query in " - ".join(item[0]).casefold()]

    def _on_group_search_changed(self, *_args):
        self.group_page = 0
        self._render_groups(self.selected_group_keys)

    def _select_matching_groups(self, selected):
        matching = {key for key, _rows in self._matching_groups()}
        if selected:
            self.selected_group_keys.update(matching)
        else:
            self.selected_group_keys.difference_update(matching)
        self._render_groups(self.selected_group_keys)

    def _on_group_sort_changed(self, _value=None):
        self.group_page = 0
        self._render_groups(self.selected_group_keys)

    def _change_group_page(self, delta):
        self.group_page += delta
        self._render_groups(self.selected_group_keys)

    def _set_group_selected(self, group_key):
        if self.group_checkboxes[group_key].get():
            self.selected_group_keys.add(group_key)
        else:
            self.selected_group_keys.discard(group_key)
        self._refresh_group_selection_count()

    def _refresh_group_selection_count(self):
        self.group_selection_var.set(f"{len(self.selected_group_keys):,} of {len(self.groups):,} selected")

    def _clear_group_list(self):
        for widget in self.groups_frame.winfo_children():
            widget.destroy()
        self.group_checkboxes = {}
        self.group_page_var.set("")
        self.group_selection_var.set("")
        self.group_previous.configure(state="disabled")
        self.group_next.configure(state="disabled")

    def _show_group_message(self, message):
        ctk.CTkLabel(self.groups_frame, text=message, anchor="w", text_color=styles.MUTED).grid(
            row=0, column=0, padx=14, pady=14, sticky="ew"
        )

    def _render_groups(self, selected_groups=None):
        self._clear_group_list()
        if not self.groups:
            message = "No groups detected." if self.headers else "Open an Excel or CSV file to see its groups here."
            self._show_group_message(message)
            return

        if selected_groups is None:
            selected_groups = set(self.groups)
        self.selected_group_keys = set(selected_groups)
        self._refresh_group_selection_count()

        ordered_groups = self._matching_groups()
        if not ordered_groups:
            self._show_group_message("No groups match your search.")
            return

        page_size = self.preferences.page_size
        page_count = (len(ordered_groups) + page_size - 1) // page_size
        self.group_page = max(0, min(self.group_page, page_count - 1))
        self.group_page_var.set(f"Page {self.group_page + 1} of {page_count}  ·  {len(ordered_groups):,} matching groups")
        self.group_previous.configure(state="normal" if self.group_page else "disabled")
        self.group_next.configure(state="normal" if self.group_page + 1 < page_count else "disabled")
        self.groups_frame.master.yview_moveto(0)

        start = self.group_page * page_size
        for index, (group_key, rows) in enumerate(ordered_groups[start:start + page_size]):
            row = ctk.CTkFrame(self.groups_frame, fg_color=styles.INSET if index % 2 else "transparent",
                               corner_radius=6)
            row.grid(row=index, column=0, sticky="ew")
            row.grid_columnconfigure(0, weight=1)
            label = " - ".join(group_key)
            cb = styles.WrappingCheckBox(
                row,
                text=label,
                onvalue=True,
                offvalue=False,
                command=lambda key=group_key: self._set_group_selected(key),
                **styles.CHECKBOX,
            )
            if group_key in self.selected_group_keys:
                cb.select()
            cb.grid(row=0, column=0, sticky="ew", padx=(10, 4), pady=5)
            ctk.CTkLabel(row, text=f"{len(rows):,}", width=68, anchor="e", text_color=styles.MUTED).grid(
                row=0, column=1, padx=(4, 13), sticky="e"
            )
            self.group_checkboxes[group_key] = cb

    def _populate_groups(self):
        self._last_spelling_review = None
        self.groups = {}
        self.selected_group_keys = set()
        self.group_search_var.set("")
        if not self.headers:
            return False

        try:
            grouping_columns = self._selected_group_columns()
            if not grouping_columns:
                self._clear_group_list()
                self._show_group_message("Select at least one grouping column.")
                return False

            _headers, self.groups = build_combined_group_map(
                self._selected_tables(),
                grouping_columns[0],
                grouping_columns=grouping_columns,
                combine_duplicates=self.combine_duplicates_var.get(),
            )
            self.groups = merge_group_values(self.groups, grouping_columns, self.value_mappings)
            self.group_page = 0
            self._render_groups()
            record_count = sum(len(rows) for rows in self.groups.values())
            self.progress_var.set(f"{len(self.groups)} groups detected from {record_count:,} records.")
            return True
        except Exception as exc:
            self.groups = {}
            self.selected_group_keys = set()
            self._clear_group_list()
            self._show_group_message(f"Unable to detect groups: {exc}")
            self.progress_var.set("Worksheet could not be processed.")
            return False

    def review_duplicate_records(self):
        if not self.selected_sheet_name or not self.headers:
            messagebox.showinfo("Review duplicate records", "Select a worksheet first.")
            return
        window = DuplicateRecordsDialog(self, self._selected_tables(), self.headers)
        self.wait_window(window)

    def review_spelling_suggestions(self):
        """Review similar grouping values together before applying choices."""
        if not self.groups:
            messagebox.showinfo("Review spelling", "Select a worksheet and grouping columns first.")
            return
        columns = self._selected_group_columns()
        review = self._last_spelling_review
        if (review is None or review.columns != tuple(columns) or review.groups != self.groups
                or review.mappings != self.value_mappings or review.selected != self.selected_group_keys):
            review = SpellingReview(self.groups, columns, self.value_mappings, self.selected_group_keys)
        else:
            review = review.copy()
        if not review.suggestions:
            messagebox.showinfo("Review spelling", "No similar grouping values were found.")
            return
        self._review_spelling_list(review)

    def _review_spelling_list(self, review):
        window = SpellingReviewDialog(self, review)
        self.wait_window(window)

    def manage_saved_spellings(self):
        try:
            window = SavedSpellingsDialog(self, self.saved_spellings)
        except (OSError, ValueError) as exc:
            messagebox.showwarning("Saved spellings", f"Saved choices could not be loaded: {exc}", parent=self)
            return
        self.wait_window(window)

    def manual_merge_categories(self):
        if not self.groups:
            messagebox.showinfo("Manual merge", "Select a worksheet and grouping columns first.")
            return
        window = ManualMergeDialog(self, self._selected_group_columns())
        self.wait_window(window)

    def start_generation(self):
        if not self.file_paths:
            messagebox.showerror("Missing file", "Please select at least one Excel file first.")
            return
        if not self.selected_sheet_name:
            messagebox.showerror("Missing worksheet", "Please select a worksheet to process.")
            return
        grouping_columns = self._selected_group_columns()
        if not grouping_columns:
            messagebox.showerror("Missing grouping column", "Please select at least one grouping column.")
            return
        selected_groups = set(self.selected_group_keys)
        if not selected_groups:
            messagebox.showerror("No groups selected", "Please choose at least one group to generate.")
            return

        self.generate_btn.configure(state="disabled")
        self.progress_var.set("Processing...")
        self.progress_bar.set(0)
        self.update_idletasks()

        input_tables = self._selected_tables()
        source_files = tuple(self.file_paths)
        source_file = source_files[0]
        combine_duplicates = self.combine_duplicates_var.get()
        value_mappings = dict(self.value_mappings)
        correct_merged_values = self.correct_merged_values_var.get()
        updates = queue.Queue()
        thread = threading.Thread(
            target=self._generate_workbook,
            args=(input_tables, tuple(grouping_columns), selected_groups,
                  combine_duplicates, self._selected_sort_mode(), source_file, updates, value_mappings,
                  correct_merged_values, source_files),
            daemon=True,
        )
        thread.start()
        self.after(50, self._poll_generation, updates)

    def _generate_workbook(self, input_tables, grouping_columns, selected_groups, combine_duplicates, sort_mode, source_file, updates, value_mappings=None, correct_merged_values=False, source_files=()):
        temporary_path = None
        try:
            temporary_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            temporary_path = temporary_file.name
            temporary_file.close()

            def progress_update(percent, message):
                updates.put(("progress", percent, message))

            result = split_input_tables(
                input_tables,
                grouping_columns[0],
                grouping_columns[1] if len(grouping_columns) > 1 else None,
                temporary_path,
                selected_groups,
                progress_callback=progress_update,
                combine_duplicates=combine_duplicates,
                grouping_columns=grouping_columns,
                sort_mode=sort_mode,
                value_mappings=value_mappings,
                correct_merged_values=correct_merged_values,
            )

            created = result.get("created_sheets", 0)
            processed = result.get("processed_rows", 0)
            updates.put(("complete", temporary_path, created, processed, source_file, source_files))
            temporary_path = None
        except Exception as exc:
            if temporary_path and os.path.exists(temporary_path):
                os.remove(temporary_path)
            updates.put(("error", exc))

    def _poll_generation(self, updates):
        try:
            while True:
                message = updates.get_nowait()
                if message[0] == "progress":
                    self.progress_var.set(message[2])
                    self.progress_bar.set(message[1] / 100)
                elif message[0] == "complete":
                    self._save_completed_workbook(*message[1:])
                    return
                else:
                    self._generation_failed(message[1])
                    return
        except queue.Empty:
            self.after(50, self._poll_generation, updates)

    def _save_completed_workbook(self, temporary_path, created, processed, source_file, source_files=()):
        staging_path = None
        discard_generated = False
        try:
            output_path = self.choose_output_location(source_file)
            if not output_path:
                discard_generated = True
                self.progress_var.set("Workbook generated but not saved.")
                self.progress_bar.set(0)
                return

            ensure_output_is_not_input(output_path, (*source_files, *self.file_paths))
            staging_file = tempfile.NamedTemporaryFile(
                suffix=".xlsx", dir=os.path.dirname(os.path.abspath(output_path)), delete=False
            )
            staging_path = staging_file.name
            staging_file.close()
            shutil.copyfile(temporary_path, staging_path)
            os.replace(staging_path, output_path)
            staging_path = None
            discard_generated = True
            message = (
                "Workbook successfully created!\n\n"
                f"{created} worksheets created\n"
                f"{processed} records processed\n\n"
                f"Saved to:\n{output_path}"
            )
            messagebox.showinfo("Success", message)
            self.progress_var.set("Completed successfully.")
            self.progress_bar.set(1)
        except Exception as exc:
            if discard_generated:
                self._generation_failed(exc)
            else:
                self._generation_failed(RuntimeError(
                    f"{exc}\n\nThe generated workbook was retained at:\n{temporary_path}"
                ))
        finally:
            if staging_path and os.path.exists(staging_path):
                os.remove(staging_path)
            if discard_generated and os.path.exists(temporary_path):
                os.remove(temporary_path)
            self.generate_btn.configure(state="normal")

    def _generation_failed(self, exc):
        messagebox.showerror("Unable to process the workbook", f"Unable to process the workbook.\n\nReason:\n{exc}")
        self.progress_var.set("Processing failed.")
        self.progress_bar.set(0)
        self.generate_btn.configure(state="normal")


def main():
    app = ExcelSplitterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
