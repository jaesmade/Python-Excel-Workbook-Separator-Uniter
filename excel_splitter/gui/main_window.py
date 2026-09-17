from __future__ import annotations

import os
import shutil
import tempfile
import threading
from collections import OrderedDict
from tkinter import filedialog, messagebox

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
    _read_input_file,
    _select_excel_engine,
    build_group_map,
    sort_group_items,
    split_workbooks,
)


class ExcelSplitterApp(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color="#F9F6EF")
        self.title("Excel Splitter & Merger")
        self.geometry("930x780")
        self.minsize(860, 680)

        self.configure(fg_color="#F9F6EF")

        self.file_path = None
        self.file_paths = []
        self.workbook = None
        self.selected_sheet_name = None
        self.headers = []
        self.groups = {}
        self.group_checkboxes = {}
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
        body.grid_rowconfigure(0, weight=1)

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
        self.generate_btn.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="ew")

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

    def choose_output_location(self):
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

        try:
            workbook_sheet_names = []
            extensions = {os.path.splitext(path)[1].lower() for path in self.file_paths}
            has_excel = bool(extensions & {".xlsx", ".xlsm", ".xls"})
            has_csv = ".csv" in extensions
            all_excel = has_excel and not has_csv and len(self.file_paths) > 1
            if has_excel and has_csv or all_excel:
                sheet_names = [ALL_SHEETS]
                self.workbook = {}
            elif os.path.splitext(file_path)[1].lower() == ".csv":
                sheet_names = ["CSV file"]
                self.workbook = {sheet_names[0]: _read_input_file(file_path)}
            else:
                engine = _select_excel_engine(file_path)
                with pd.ExcelFile(file_path, engine=engine) as workbook:
                    workbook_sheet_names = workbook.sheet_names
                    sheet_names = [ALL_SHEETS, *workbook_sheet_names]
                    self.workbook = {
                        name: pd.read_excel(file_path, sheet_name=name, engine=engine)
                        for name in workbook_sheet_names
                    }
            self.selected_sheet_name = None
            self.headers = []
            self.groups = {}
            self.sheet_dropdown.configure(values=sheet_names)
            default_sheet = ALL_SHEETS if has_excel and has_csv or all_excel else ("CSV file" if os.path.splitext(file_path)[1].lower() == ".csv" else (workbook_sheet_names[0] if workbook_sheet_names else ""))
            self.sheet_var.set(default_sheet)
            for _frame, variable, _dropdown, _remove in self.group_column_controls:
                variable.set("")
            self._clear_group_list()
            if default_sheet:
                self.on_sheet_selected(default_sheet)
            self.progress_var.set("Workbook loaded successfully.")
        except Exception as exc:
            messagebox.showerror("Unable to load workbook", f"{exc}")
            self.progress_var.set("Failed to load workbook.")
            self.sheet_dropdown.configure(values=[])
            for _frame, _variable, dropdown, _remove in self.group_column_controls:
                dropdown.configure(values=[])

    def on_sheet_selected(self, sheet_name):
        if not sheet_name or not self.file_paths:
            return

        self.selected_sheet_name = sheet_name
        try:
            if sheet_name == ALL_SHEETS:
                tables = self._selected_tables()
                if not tables:
                    raise WorkbookProcessingError("The selected workbook does not contain any data.")
                sheet = pd.concat(tables, ignore_index=True)
            else:
                sheet = self.workbook[sheet_name]
            if sheet.empty:
                raise WorkbookProcessingError("The selected worksheet does not contain any data.")

            headers = list(sheet.columns)
            self.headers = ["" if value is None else str(value).strip() for value in headers]
            for _frame, _variable, dropdown, _remove in self.group_column_controls:
                dropdown.configure(values=self.headers)

            department_default = next((h for h in self.headers if h.lower() == "department"), "")
            first_column = department_default or (self.headers[0] if self.headers else "")
            for index, (_frame, variable, _dropdown, _remove) in enumerate(self.group_column_controls):
                variable.set(first_column if index == 0 else "")

            self._populate_groups(sheet)
        except Exception as exc:
            messagebox.showerror("Unable to process worksheet", str(exc))
            self._clear_group_list()
            self.progress_var.set("Worksheet could not be processed.")

    def on_column_selection_changed(self, _event=None):
        if not self.file_paths or not self.selected_sheet_name:
            return
        if self.selected_sheet_name == ALL_SHEETS:
            tables = self._selected_tables()
            if not tables:
                return
            sheet = pd.concat(tables, ignore_index=True)
        else:
            sheet = self.workbook[self.selected_sheet_name]
        self._populate_groups(sheet)

    def _selected_tables(self):
        input_tables = []
        for file_path in self.file_paths:
            for source_sheet_name, table in _iter_input_tables(file_path, self.selected_sheet_name):
                if table.empty:
                    continue
                input_tables.append((f"{os.path.basename(file_path)} - {source_sheet_name}", table))
        normalized_tables, _headers = _normalize_input_tables(input_tables)
        return [table for _source_name, table in normalized_tables]

    def _selected_sort_mode(self):
        return {
            "Alphabetical": SORT_ALPHABETICAL,
            "Record count: ascending": SORT_COUNT_ASCENDING,
            "Record count: descending": SORT_COUNT_DESCENDING,
        }[self.group_sort_var.get()]

    def _on_group_sort_changed(self, _value=None):
        selected_groups = {key for key, checkbox in self.group_checkboxes.items() if checkbox.get()}
        self._render_groups(selected_groups)

    def _clear_group_list(self):
        for widget in self.groups_frame.winfo_children():
            widget.destroy()
        self.group_checkboxes = {}

    def _render_groups(self, selected_groups=None):
        self._clear_group_list()
        if not self.groups:
            ctk.CTkLabel(self.groups_frame, text="No groups detected.", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)
            return

        if selected_groups is None:
            selected_groups = set(self.groups)

        teal = "#12454B"
        teal_dark = "#0C3035"
        for group_key, rows in sort_group_items(self.groups, self._selected_sort_mode()):
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
            )
            if group_key in selected_groups:
                cb.select()
            cb.grid(sticky="w", padx=16, pady=6)
            self.group_checkboxes[group_key] = cb

    def _populate_groups(self, sheet):
        self._clear_group_list()
        if not self.headers:
            return

        try:
            grouping_columns = self._selected_group_columns()
            if not grouping_columns:
                self.groups = {}
                ctk.CTkLabel(self.groups_frame, text="Please select at least one grouping column.", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)
                return

            if self.selected_sheet_name == ALL_SHEETS:
                tables = self._selected_tables()
                if not tables:
                    raise WorkbookProcessingError("No data was found in the selected worksheets.")
                groups = OrderedDict()
                for source_sheet in tables:
                    _, source_groups = build_group_map(source_sheet, grouping_columns[0], grouping_columns=grouping_columns)
                    for group_key, rows in source_groups.items():
                        groups.setdefault(group_key, []).extend(rows)
            elif len(self.file_paths) == 1:
                headers, groups = build_group_map(sheet, grouping_columns[0], grouping_columns=grouping_columns)
            else:
                groups = OrderedDict()
                for file_path in self.file_paths:
                    engine = _select_excel_engine(file_path)
                    source_sheet = _read_input_file(file_path, self.selected_sheet_name)
                    _, source_groups = build_group_map(source_sheet, grouping_columns[0], grouping_columns=grouping_columns)
                    for group_key, rows in source_groups.items():
                        groups.setdefault(group_key, []).extend(rows)
            self.groups = groups
            if not groups:
                self._render_groups()
                return
            self._render_groups()
        except Exception as exc:
            self.groups = {}
            self.group_checkboxes = {}
            ctk.CTkLabel(self.groups_frame, text=f"Unable to detect groups: {exc}", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)

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
        selected_groups = {key for key, cb in self.group_checkboxes.items() if cb.get()}
        if not selected_groups:
            messagebox.showerror("No groups selected", "Please choose at least one group to generate.")
            return

        self.generate_btn.configure(state="disabled")
        self.progress_var.set("Processing...")
        self.progress_bar.set(0)
        self.update_idletasks()

        thread = threading.Thread(
            target=self._generate_workbook,
            args=(selected_groups, self._selected_sort_mode()),
            daemon=True,
        )
        thread.start()

    def _generate_workbook(self, selected_groups, sort_mode):
        temporary_path = None
        try:
            grouping_columns = self._selected_group_columns()
            temporary_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            temporary_path = temporary_file.name
            temporary_file.close()

            def progress_update(percent, message):
                self.progress_var.set(message)
                self.progress_bar.set(percent / 100)
                self.update_idletasks()

            result = split_workbooks(
                self.file_paths,
                self.selected_sheet_name,
                grouping_columns[0],
                grouping_columns[1] if len(grouping_columns) > 1 else None,
                temporary_path,
                selected_groups,
                progress_callback=progress_update,
                combine_duplicates=self.combine_duplicates_var.get(),
                grouping_columns=grouping_columns,
                sort_mode=sort_mode,
            )

            created = result.get("created_sheets", 0)
            processed = result.get("processed_rows", 0)
            self.after(0, self._save_completed_workbook, temporary_path, created, processed)
            temporary_path = None
        except Exception as exc:
            if temporary_path and os.path.exists(temporary_path):
                os.remove(temporary_path)
            self.after(0, self._generation_failed, exc)

    def _save_completed_workbook(self, temporary_path, created, processed):
        try:
            output_path = self.choose_output_location()
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
