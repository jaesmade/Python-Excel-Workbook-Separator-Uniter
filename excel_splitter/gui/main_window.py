from __future__ import annotations

import os
import shutil
import tempfile
import threading
from collections import OrderedDict
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

from excel_splitter.core import WorkbookProcessingError, _select_excel_engine, build_group_map, split_workbooks


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

        ctk.CTkLabel(top, text="Department Column", text_color=text, font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, padx=(18, 12), pady=(8, 10), sticky="w")
        self.department_var = ctk.StringVar(value="")
        self.department_dropdown = ctk.CTkOptionMenu(
            top,
            values=[],
            variable=self.department_var,
            command=self.on_column_selection_changed,
            fg_color=ivory_alt,
            button_color=teal,
            button_hover_color=teal_dark,
            text_color=text,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=text,
            dropdown_hover_color=ivory_alt,
            height=38,
        )
        self.department_dropdown.grid(row=2, column=1, columnspan=2, padx=(0, 18), pady=(8, 10), sticky="ew")

        ctk.CTkLabel(top, text="ID Number / Status Column", text_color=text, font=ctk.CTkFont(size=13, weight="bold")).grid(row=3, column=0, padx=(18, 12), pady=(8, 10), sticky="w")
        self.id_var = ctk.StringVar(value="")
        self.id_dropdown = ctk.CTkOptionMenu(
            top,
            values=[],
            variable=self.id_var,
            command=self.on_column_selection_changed,
            fg_color=ivory_alt,
            button_color=teal,
            button_hover_color=teal_dark,
            text_color=text,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=text,
            dropdown_hover_color=ivory_alt,
            height=38,
        )
        self.id_dropdown.grid(row=3, column=1, columnspan=2, padx=(0, 18), pady=(8, 10), sticky="ew")

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
        self.combine_duplicates_checkbox.grid(row=4, column=0, columnspan=3, padx=(18, 18), pady=(8, 16), sticky="w")

        body = ctk.CTkFrame(self, corner_radius=18, fg_color="#FFFFFF", border_color=line, border_width=1)
        body.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        ctk.CTkLabel(body, text="Detected Groups", text_color=text, font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=18, pady=(18, 10), sticky="w")
        self.groups_frame = ctk.CTkScrollableFrame(body, corner_radius=12, fg_color=ivory)
        self.groups_frame.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
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

    def browse_file(self):
        filetypes = [("Excel files", "*.xlsx *.xlsm *.xls")]
        selected = filedialog.askopenfilenames(title="Select Excel workbooks", filetypes=filetypes)
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
            engine = _select_excel_engine(file_path)
            workbook = pd.ExcelFile(file_path, engine=engine)
            sheet_names = workbook.sheet_names
            self.workbook = {name: pd.read_excel(file_path, sheet_name=name, engine=engine) for name in sheet_names}
            self.selected_sheet_name = None
            self.headers = []
            self.groups = {}
            self.sheet_dropdown.configure(values=sheet_names)
            self.sheet_var.set(sheet_names[0] if sheet_names else "")
            self.department_var.set("")
            self.id_var.set("")
            self._clear_group_list()
            if sheet_names:
                self.on_sheet_selected(sheet_names[0])
            self.progress_var.set("Workbook loaded successfully.")
        except Exception as exc:
            messagebox.showerror("Unable to load workbook", f"{exc}")
            self.progress_var.set("Failed to load workbook.")
            self.sheet_dropdown.configure(values=[])
            self.department_dropdown.configure(values=[])
            self.id_dropdown.configure(values=[])

    def on_sheet_selected(self, sheet_name):
        if not sheet_name or not self.workbook:
            return

        self.selected_sheet_name = sheet_name
        try:
            sheet = self.workbook[sheet_name]
            if sheet.empty:
                raise WorkbookProcessingError("The selected worksheet does not contain any data.")

            headers = list(sheet.columns)
            self.headers = ["" if value is None else str(value).strip() for value in headers]
            self.department_dropdown.configure(values=self.headers)
            self.id_dropdown.configure(values=self.headers)

            department_default = next((h for h in self.headers if h.lower() == "department"), "")
            id_default = next((h for h in self.headers if h.lower() in {"id number", "id_number", "employment status", "status"}), "")

            if department_default:
                self.department_var.set(department_default)
            elif self.headers:
                self.department_var.set(self.headers[0])

            if id_default:
                self.id_var.set(id_default)
            elif self.headers:
                self.id_var.set(self.headers[0])

            self._populate_groups(sheet)
        except Exception as exc:
            messagebox.showerror("Unable to process worksheet", str(exc))
            self._clear_group_list()
            self.progress_var.set("Worksheet could not be processed.")

    def on_column_selection_changed(self, _event=None):
        if not self.workbook or not self.selected_sheet_name:
            return
        sheet = self.workbook[self.selected_sheet_name]
        self._populate_groups(sheet)

    def _clear_group_list(self):
        for widget in self.groups_frame.winfo_children():
            widget.destroy()

    def _populate_groups(self, sheet):
        self._clear_group_list()
        if not self.headers:
            return

        teal = "#12454B"
        teal_dark = "#0C3035"

        try:
            dep_name = self.department_var.get()
            id_name = self.id_var.get()
            if not dep_name or not id_name:
                self.groups = {}
                ctk.CTkLabel(self.groups_frame, text="Please select both required columns.", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)
                return

            if len(self.file_paths) == 1:
                headers, groups = build_group_map(sheet, dep_name, id_name)
            else:
                groups = OrderedDict()
                for file_path in self.file_paths:
                    engine = _select_excel_engine(file_path)
                    source_sheet = pd.read_excel(
                        file_path,
                        sheet_name=self.selected_sheet_name,
                        engine=engine,
                    )
                    _, source_groups = build_group_map(source_sheet, dep_name, id_name)
                    for group_key, rows in source_groups.items():
                        groups.setdefault(group_key, []).extend(rows)
            self.groups = groups
            if not groups:
                ctk.CTkLabel(self.groups_frame, text="No groups detected.", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)
                return

            self.group_checkboxes = {}
            for (department, status), rows in groups.items():
                label = f"{department} - {status}"
                count = len(rows)
                cb = ctk.CTkCheckBox(
                    self.groups_frame,
                    text=f"{label}     {count} records",
                    onvalue=True,
                    offvalue=False,
                    text_color="black",
                    fg_color=teal,
                    hover_color=teal_dark,
                    border_color=teal,
                )
                cb.select()
                cb.grid(sticky="w", padx=16, pady=6)
                self.group_checkboxes[(department, status)] = cb
        except Exception as exc:
            ctk.CTkLabel(self.groups_frame, text=f"Unable to detect groups: {exc}", anchor="w", text_color="black").grid(sticky="ew", padx=16, pady=8)

    def start_generation(self):
        if not self.file_paths:
            messagebox.showerror("Missing file", "Please select at least one Excel file first.")
            return
        if not self.selected_sheet_name:
            messagebox.showerror("Missing worksheet", "Please select a worksheet to process.")
            return
        if not self.department_var.get():
            messagebox.showerror("Missing Department column", "Please select the Department column.")
            return
        if not self.id_var.get():
            messagebox.showerror("Missing ID Number column", "Please select the ID Number / Employment Status column.")
            return

        selected_groups = {key for key, cb in self.group_checkboxes.items() if cb.get()}
        if not selected_groups:
            messagebox.showerror("No groups selected", "Please choose at least one group to generate.")
            return

        self.generate_btn.configure(state="disabled")
        self.progress_var.set("Processing...")
        self.progress_bar.set(0)
        self.update_idletasks()

        thread = threading.Thread(target=self._generate_workbook, args=(selected_groups,), daemon=True)
        thread.start()

    def _generate_workbook(self, selected_groups):
        temporary_path = None
        try:
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
                self.department_var.get(),
                self.id_var.get(),
                temporary_path,
                selected_groups,
                progress_callback=progress_update,
                combine_duplicates=self.combine_duplicates_var.get(),
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
