from __future__ import annotations

import os
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

from excel_splitter.core import WorkbookProcessingError, build_group_map, split_workbooks


class ExcelSplitterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Department & Employment Status Splitter")
        self.geometry("930x780")
        self.minsize(860, 680)

        self.file_path = None
        self.file_paths = []
        self.workbook = None
        self.selected_sheet_name = None
        self.headers = []
        self.groups = {}
        self.group_checkboxes = {}

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(self, corner_radius=12)
        top.grid(row=0, column=0, padx=16, pady=16, sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(top, text="Excel File", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, padx=(16, 12), pady=(16, 8), sticky="w")
        self.file_var = ctk.StringVar(value="No file selected")
        self.file_entry = ctk.CTkEntry(top, textvariable=self.file_var, state="readonly", fg_color="#1f1f1f")
        self.file_entry.grid(row=0, column=1, padx=(0, 10), pady=(16, 8), sticky="ew")
        ctk.CTkButton(top, text="Browse", command=self.browse_file, width=110).grid(row=0, column=2, padx=(0, 16), pady=(16, 8))

        ctk.CTkLabel(top, text="Worksheet", font=ctk.CTkFont(size=13, weight="bold")).grid(row=1, column=0, padx=(16, 12), pady=(8, 8), sticky="w")
        self.sheet_var = ctk.StringVar(value="")
        self.sheet_dropdown = ctk.CTkOptionMenu(top, values=[], variable=self.sheet_var, command=self.on_sheet_selected)
        self.sheet_dropdown.grid(row=1, column=1, columnspan=2, padx=(0, 16), pady=(8, 8), sticky="ew")

        ctk.CTkLabel(top, text="Department Column", font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, padx=(16, 12), pady=(8, 8), sticky="w")
        self.department_var = ctk.StringVar(value="")
        self.department_dropdown = ctk.CTkOptionMenu(top, values=[], variable=self.department_var, command=self.on_column_selection_changed)
        self.department_dropdown.grid(row=2, column=1, columnspan=2, padx=(0, 16), pady=(8, 8), sticky="ew")

        ctk.CTkLabel(top, text="ID Number / Status Column", font=ctk.CTkFont(size=13, weight="bold")).grid(row=3, column=0, padx=(16, 12), pady=(8, 8), sticky="w")
        self.id_var = ctk.StringVar(value="")
        self.id_dropdown = ctk.CTkOptionMenu(top, values=[], variable=self.id_var, command=self.on_column_selection_changed)
        self.id_dropdown.grid(row=3, column=1, columnspan=2, padx=(0, 16), pady=(8, 8), sticky="ew")

        ctk.CTkLabel(top, text="Output File", font=ctk.CTkFont(size=13, weight="bold")).grid(row=4, column=0, padx=(16, 12), pady=(8, 16), sticky="w")
        self.output_var = ctk.StringVar(value="Separated_Workbook.xlsx")
        self.output_entry = ctk.CTkEntry(top, textvariable=self.output_var, fg_color="#1f1f1f")
        self.output_entry.grid(row=4, column=1, padx=(0, 10), pady=(8, 16), sticky="ew")
        ctk.CTkButton(top, text="Browse", command=self.choose_output_location, width=110).grid(row=4, column=2, padx=(0, 16), pady=(8, 16))

        body = ctk.CTkFrame(self, corner_radius=12)
        body.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        ctk.CTkLabel(body, text="Detected Groups", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=16, pady=(16, 8), sticky="w")
        self.groups_frame = ctk.CTkScrollableFrame(body, corner_radius=10)
        self.groups_frame.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")
        self.groups_frame.grid_columnconfigure(0, weight=1)

        bottom = ctk.CTkFrame(self, corner_radius=12)
        bottom.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
        bottom.grid_columnconfigure(0, weight=1)

        self.progress_var = ctk.StringVar(value="Ready")
        self.progress_bar = ctk.CTkProgressBar(bottom, width=600, height=18)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, padx=16, pady=(16, 6), sticky="ew")
        self.progress_label = ctk.CTkLabel(bottom, textvariable=self.progress_var, anchor="w")
        self.progress_label.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="ew")

        self.generate_btn = ctk.CTkButton(bottom, text="GENERATE WORKBOOK", command=self.start_generation, height=44, font=ctk.CTkFont(size=15, weight="bold"))
        self.generate_btn.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")

    def browse_file(self):
        filetypes = [("Excel files", "*.xlsx *.xlsm *.xls")]
        selected = filedialog.askopenfilenames(title="Select Excel workbooks", filetypes=filetypes)
        if not selected:
            return

        self.file_paths = list(selected)
        self.file_path = self.file_paths[0]
        if len(self.file_paths) == 1:
            self.file_var.set(self.file_path)
            self._set_default_output_name(self.file_path)
        else:
            names = ", ".join(os.path.basename(path) for path in self.file_paths[:3])
            if len(self.file_paths) > 3:
                names += f", +{len(self.file_paths) - 3} more"
            self.file_var.set(f"{len(self.file_paths)} files selected: {names}")
            self._set_default_output_name(self.file_paths[0])
        self.load_workbook(self.file_path)

    def choose_output_location(self):
        path = filedialog.asksaveasfilename(
            title="Save separated workbook as",
            defaultextension=".xlsx",
            initialfile=self.output_var.get() or "Separated_Workbook.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _set_default_output_name(self, source_file: str):
        if not source_file:
            self.output_var.set("Separated_Workbook.xlsx")
            return
        stem = os.path.splitext(os.path.basename(source_file))[0]
        if not stem:
            self.output_var.set("Separated_Workbook.xlsx")
            return
        self.output_var.set(f"{stem} - SPLIT.xlsx")

    def load_workbook(self, file_path):
        self.progress_var.set("Loading workbook...")
        self.progress_bar.set(0)
        self.update_idletasks()

        try:
            workbook = pd.ExcelFile(file_path)
            sheet_names = workbook.sheet_names
            self.workbook = {name: pd.read_excel(file_path, sheet_name=name) for name in sheet_names}
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

        try:
            dep_name = self.department_var.get()
            id_name = self.id_var.get()
            if not dep_name or not id_name:
                self.groups = {}
                ctk.CTkLabel(self.groups_frame, text="Please select both required columns.", anchor="w").grid(sticky="ew", padx=16, pady=8)
                return

            headers, groups = build_group_map(sheet, dep_name, id_name)
            self.groups = groups
            if not groups:
                ctk.CTkLabel(self.groups_frame, text="No groups detected.", anchor="w").grid(sticky="ew", padx=16, pady=8)
                return

            self.group_checkboxes = {}
            for (department, status), rows in groups.items():
                label = f"{department} - {status}"
                count = len(rows)
                cb = ctk.CTkCheckBox(self.groups_frame, text=f"{label}     {count} records", onvalue=True, offvalue=False, justify="left")
                cb.select()
                cb.grid(sticky="w", padx=16, pady=6)
                self.group_checkboxes[(department, status)] = cb
        except Exception as exc:
            ctk.CTkLabel(self.groups_frame, text=f"Unable to detect groups: {exc}", anchor="w").grid(sticky="ew", padx=16, pady=8)

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

        output_path = self.output_var.get().strip()
        if not output_path:
            messagebox.showerror("Missing output file", "Please enter or choose an output file name.")
            return

        self.generate_btn.configure(state="disabled")
        self.progress_var.set("Processing...")
        self.progress_bar.set(0)
        self.update_idletasks()

        thread = threading.Thread(target=self._generate_workbook, args=(output_path, selected_groups), daemon=True)
        thread.start()

    def _generate_workbook(self, output_path, selected_groups):
        try:
            directory = os.path.dirname(os.path.abspath(output_path))
            if not os.path.isdir(directory):
                raise WorkbookProcessingError("The selected output directory does not exist.")

            def progress_update(percent, message):
                self.progress_var.set(message)
                self.progress_bar.set(percent / 100)
                self.update_idletasks()

            result = split_workbooks(
                self.file_paths,
                self.selected_sheet_name,
                self.department_var.get(),
                self.id_var.get(),
                output_path,
                selected_groups,
                progress_callback=progress_update,
            )

            created = result.get("created_sheets", 0)
            processed = result.get("processed_rows", 0)
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
            messagebox.showerror("Unable to process the workbook", f"Unable to process the workbook.\n\nReason:\n{exc}")
            self.progress_var.set("Processing failed.")
            self.progress_bar.set(0)
        finally:
            self.generate_btn.configure(state="normal")


def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    app = ExcelSplitterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
