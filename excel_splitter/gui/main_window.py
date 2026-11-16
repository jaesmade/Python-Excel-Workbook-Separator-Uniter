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
    preview_group_items,
    sort_group_items,
    split_input_tables,
    source_file_labels,
)


class ExcelSplitterApp(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color="#F9F6EF")
        self.title("Excel Splitter & Merger")
        if sys.platform.startswith("win"):
            self.iconbitmap(str(Path(__file__).resolve().parents[1] / "assets" / "app.ico"))
        self.geometry("930x780")
        self.minsize(860, 680)

        self.configure(fg_color="#F9F6EF")

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

        self._build_ui()

    def _build_ui(self):
        teal = "#12454B"
        teal_dark = "#0C3035"
        teal_mid = "#1D646A"
        ivory = "#F9F6EF"
        ivory_alt = "#F3EEE3"
        line = "#D7D0C3"
        text = "#14353A"
        muted = "#5B6E70"

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(self, corner_radius=18, fg_color="#FFFFFF", border_color=line, border_width=1)
        top.grid(row=0, column=0, padx=16, pady=16, sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(top, text="Excel File", text_color=text, font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, padx=(18, 12), pady=(18, 10), sticky="w")
        self.file_var = ctk.StringVar(value="No file selected")
        self.file_entry = ctk.CTkEntry(
            top,
            textvariable=self.file_var,
            state="readonly",
            fg_color=ivory_alt,
            border_color=line,
            text_color=text,
            height=38,
        )
        self.file_entry.grid(row=0, column=1, padx=(0, 10), pady=(18, 10), sticky="ew")
        ctk.CTkButton(
            top,
            text="Browse",
            command=self.browse_file,
            width=110,
            height=38,
            fg_color=teal,
            hover_color=teal_dark,
            text_color=ivory,
            border_color=teal,
        ).grid(row=0, column=2, padx=(0, 18), pady=(18, 10))

        ctk.CTkLabel(top, text="Worksheet", text_color=text, font=ctk.CTkFont(size=13, weight="bold")).grid(row=1, column=0, padx=(18, 12), pady=(8, 10), sticky="w")
        self.sheet_var = ctk.StringVar(value="")
        self.sheet_dropdown = ctk.CTkOptionMenu(
            top,
            values=[],
            variable=self.sheet_var,
            command=self.on_sheet_selected,
            fg_color=ivory_alt,
            button_color=teal,
            button_hover_color=teal_dark,
            text_color=text,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=text,
            dropdown_hover_color=ivory_alt,
            height=38,
        )
        self.sheet_dropdown.grid(row=1, column=1, columnspan=2, padx=(0, 18), pady=(8, 10), sticky="ew")

        ctk.CTkLabel(top, text="Grouping Columns", text_color=text, font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, padx=(18, 12), pady=(8, 10), sticky="nw")
        self.group_columns_frame = ctk.CTkFrame(top, fg_color="transparent")
        self.group_columns_frame.grid(row=2, column=1, padx=(0, 10), pady=(4, 6), sticky="ew")
        self.group_columns_frame.grid_columnconfigure(0, weight=1)
        self.add_column_button = ctk.CTkButton(
            top,
            text="+ Add column",
            command=self._add_group_column,
            width=110,
            height=34,
            fg_color=teal,
            hover_color=teal_dark,
            text_color=ivory,
        )
        self.add_column_button.grid(row=2, column=2, padx=(0, 18), pady=(8, 10), sticky="n")
        self._add_group_column()

        self.combine_duplicates_var = ctk.BooleanVar(value=True)
        self.combine_duplicates_checkbox = ctk.CTkCheckBox(
            top,
            text="Combine duplicate groups",
            variable=self.combine_duplicates_var,
            command=self.on_column_selection_changed,
            onvalue=True,
            offvalue=False,
            fg_color=teal,
            hover_color=teal_mid,
            text_color=text,
            border_color=teal,
        )
        self.combine_duplicates_checkbox.grid(row=3, column=0, columnspan=3, padx=(18, 18), pady=(8, 16), sticky="w")

        body = ctk.CTkFrame(self, corner_radius=18, fg_color="#FFFFFF", border_color=line, border_width=1)
        body.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=0)
        body.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(body, text="Detected Groups", text_color=text, font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=(18, 8), pady=(18, 10), sticky="w")
        self.group_sort_var = ctk.StringVar(value="Alphabetical")
        self.group_sort_dropdown = ctk.CTkOptionMenu(
            body,
            values=["Alphabetical", "Record count: ascending", "Record count: descending"],
            variable=self.group_sort_var,
            command=self._on_group_sort_changed,
            fg_color=ivory_alt,
            button_color=teal,
            button_hover_color=teal_dark,
            text_color=text,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=text,
            dropdown_hover_color=ivory_alt,
            width=210,
            height=34,
        )
        self.group_sort_dropdown.grid(row=0, column=1, padx=(8, 18), pady=(14, 10), sticky="e")
        self.groups_frame = ctk.CTkScrollableFrame(body, corner_radius=12, fg_color=ivory)
        self.groups_frame.grid(row=1, column=0, columnspan=2, padx=18, pady=(0, 18), sticky="nsew")
        self.groups_frame.grid_columnconfigure(0, weight=1)

        group_pages = ctk.CTkFrame(body, fg_color="transparent")
        group_pages.grid(row=2, column=0, columnspan=2, padx=18, pady=(0, 16), sticky="w")
        self.group_previous = ctk.CTkButton(group_pages, text="Previous", width=90, command=lambda: self._change_group_page(-1))
        self.group_previous.pack(side="left")
        self.group_previous.configure(state="disabled")
        self.group_page_var = ctk.StringVar(value="")
        ctk.CTkLabel(group_pages, textvariable=self.group_page_var).pack(side="left", padx=12)
        self.group_next = ctk.CTkButton(group_pages, text="Next", width=90, command=lambda: self._change_group_page(1))
        self.group_next.pack(side="left")
        self.group_next.configure(state="disabled")

        bottom = ctk.CTkFrame(self, corner_radius=18, fg_color="#FFFFFF", border_color=line, border_width=1)
        bottom.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
        bottom.grid_columnconfigure(0, weight=1)

        self.progress_var = ctk.StringVar(value="Ready")
        self.progress_bar = ctk.CTkProgressBar(bottom, width=600, height=16, fg_color=ivory_alt, progress_color=teal)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, padx=18, pady=(18, 8), sticky="ew")
        self.progress_label = ctk.CTkLabel(bottom, textvariable=self.progress_var, text_color=muted, anchor="w")
        self.progress_label.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="ew")

        self.generate_btn = ctk.CTkButton(
            bottom,
            text="GENERATE WORKBOOK",
            command=self.start_generation,
            height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=teal,
            hover_color=teal_dark,
            text_color=ivory,
            border_color=teal,
        )
        ctk.CTkButton(
            bottom,
            text="Preview Worksheets",
            command=self.preview_worksheets,
            height=38,
            fg_color=ivory_alt,
            hover_color=line,
            text_color=text,
        ).grid(row=2, column=0, padx=18, pady=(0, 8), sticky="ew")
        self.generate_btn.grid(row=3, column=0, padx=18, pady=(0, 18), sticky="ew")

    def preview_worksheets(self):
        selected_groups = set(self.selected_group_keys)
        if not selected_groups:
            messagebox.showinfo("Preview worksheets", "Select at least one detected group to preview.")
            return

        items = preview_group_items(self.groups, selected_groups, self._selected_sort_mode())
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
        sheet_list = tk.Listbox(sheet_frame, width=26, exportselection=False, activestyle="none")
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
        table = ttk.Treeview(table_frame, columns=columns, show="headings")
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
        teal = "#12454B"
        teal_dark = "#0C3035"
        ivory_alt = "#F3EEE3"
        text = "#14353A"
        row_frame = ctk.CTkFrame(self.group_columns_frame, fg_color="transparent")
        row_frame.grid(row=len(self.group_column_controls), column=0, pady=2, sticky="ew")
        row_frame.grid_columnconfigure(0, weight=1)
        variable = ctk.StringVar(value="")
        dropdown = ctk.CTkOptionMenu(
            row_frame,
            values=self.headers or ["Select a column"],
            variable=variable,
            command=self.on_column_selection_changed,
            fg_color=ivory_alt,
            button_color=teal,
            button_hover_color=teal_dark,
            text_color=text,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=text,
            dropdown_hover_color=ivory_alt,
            height=34,
        )
        dropdown.grid(row=0, column=0, sticky="ew")
        remove_button = None
        if self.group_column_controls:
            remove_button = ctk.CTkButton(
                row_frame,
                text="Remove",
                command=lambda: self._remove_group_column(row_frame, variable),
                width=78,
                height=34,
                fg_color="#D7D0C3",
                hover_color="#B9AEA0",
                text_color="#14353A",
            )
            remove_button.grid(row=0, column=1, padx=(8, 0))
        self.group_column_controls.append((row_frame, variable, dropdown, remove_button))

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
        self.selected_sheet_name = None
        self.headers = []
        self.groups = {}
        self.selected_group_keys = set()
        for _frame, variable, dropdown, _remove in self.group_column_controls:
            variable.set("")
            dropdown.configure(values=["Select a column"])
        self._clear_group_list()

    def on_sheet_selected(self, sheet_name):
        if not sheet_name or not self.file_paths:
            return False

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
        if self.selected_sheet_name == ALL_SHEETS:
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

    def _clear_group_list(self):
        for widget in self.groups_frame.winfo_children():
            widget.destroy()
        self.group_checkboxes = {}
        self.group_page_var.set("")
        self.group_previous.configure(state="disabled")
        self.group_next.configure(state="disabled")

    def _render_groups(self, selected_groups=None):
        self._clear_group_list()
        if not self.groups:
            ctk.CTkLabel(self.groups_frame, text="No groups detected.", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)
            return

        if selected_groups is None:
            selected_groups = set(self.groups)
        self.selected_group_keys = set(selected_groups)

        ordered_groups = sort_group_items(self.groups, self._selected_sort_mode())
        page_size = 50
        page_count = max(1, (len(ordered_groups) + page_size - 1) // page_size)
        self.group_page = max(0, min(self.group_page, page_count - 1))
        self.group_page_var.set(f"Page {self.group_page + 1} of {page_count}  ·  {len(self.groups):,} groups")
        self.group_previous.configure(state="normal" if self.group_page else "disabled")
        self.group_next.configure(state="normal" if self.group_page + 1 < page_count else "disabled")

        teal = "#12454B"
        teal_dark = "#0C3035"
        for group_key, rows in ordered_groups[self.group_page * page_size:(self.group_page + 1) * page_size]:
            label = " - ".join(group_key)
            cb = ctk.CTkCheckBox(
                self.groups_frame,
                text=f"{label}     {len(rows)} records",
                onvalue=True,
                offvalue=False,
                text_color="black",
                fg_color=teal,
                hover_color=teal_dark,
                border_color=teal,
                command=lambda key=group_key: self._set_group_selected(key),
            )
            if group_key in selected_groups:
                cb.select()
            cb.grid(sticky="w", padx=16, pady=6)
            self.group_checkboxes[group_key] = cb

    def _populate_groups(self):
        self._clear_group_list()
        self.groups = {}
        self.selected_group_keys = set()
        if not self.headers:
            return False

        try:
            grouping_columns = self._selected_group_columns()
            if not grouping_columns:
                ctk.CTkLabel(self.groups_frame, text="Please select at least one grouping column.", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)
                return False

            _headers, self.groups = build_combined_group_map(
                self._selected_tables(),
                grouping_columns[0],
                grouping_columns=grouping_columns,
                combine_duplicates=self.combine_duplicates_var.get(),
            )
            self.group_page = 0
            self._render_groups()
            record_count = sum(len(rows) for rows in self.groups.values())
            self.progress_var.set(f"{len(self.groups)} groups detected from {record_count:,} records.")
            return True
        except Exception as exc:
            self.groups = {}
            self.group_checkboxes = {}
            ctk.CTkLabel(self.groups_frame, text=f"Unable to detect groups: {exc}", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)
            self.progress_var.set("Worksheet could not be processed.")
            return False

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
        source_file = self.file_paths[0]
        combine_duplicates = self.combine_duplicates_var.get()
        updates = queue.Queue()
        thread = threading.Thread(
            target=self._generate_workbook,
            args=(input_tables, tuple(grouping_columns), selected_groups,
                  combine_duplicates, self._selected_sort_mode(), source_file, updates),
            daemon=True,
        )
        thread.start()
        self.after(50, self._poll_generation, updates)

    def _generate_workbook(self, input_tables, grouping_columns, selected_groups, combine_duplicates, sort_mode, source_file, updates):
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
            )

            created = result.get("created_sheets", 0)
            processed = result.get("processed_rows", 0)
            updates.put(("complete", temporary_path, created, processed, source_file))
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

    def _save_completed_workbook(self, temporary_path, created, processed, source_file):
        try:
            output_path = self.choose_output_location(source_file)
            if not output_path:
                self.progress_var.set("Workbook generated but not saved.")
                self.progress_bar.set(0)
                return

            shutil.copyfile(temporary_path, output_path)
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
            self._generation_failed(exc)
        finally:
            if os.path.exists(temporary_path):
                os.remove(temporary_path)
            self.generate_btn.configure(state="normal")

    def _generation_failed(self, exc):
        messagebox.showerror("Unable to process the workbook", f"Unable to process the workbook.\n\nReason:\n{exc}")
        self.progress_var.set("Processing failed.")
        self.progress_bar.set(0)
        self.generate_btn.configure(state="normal")


def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    app = ExcelSplitterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
