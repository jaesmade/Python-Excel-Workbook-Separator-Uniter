"""Review possible duplicate records without changing the loaded data."""

import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk
import pandas as pd

from excel_splitter.duplicate_records import find_duplicate_records
from excel_splitter.gui import styles


class DuplicateRecordsDialog(ctk.CTkToplevel):
    def __init__(self, app, tables, columns):
        super().__init__(app)
        self.tables = tables
        self.columns = list(columns)
        self.title("Review duplicate records")
        self.geometry("950x650")
        self.minsize(650, 450)
        self.configure(fg_color=styles.BACKGROUND)
        self.transient(app)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(self, text="Review possible duplicate records", text_color=styles.TEXT,
                     font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold")).grid(
                         row=0, column=0, padx=24, pady=(22, 4), sticky="w")
        ctk.CTkLabel(self, text="Select columns to compare. Text ignores case and extra spaces. Records with blank or missing matching fields are skipped. No records are removed or changed.",
                     text_color=styles.MUTED, font=ctk.CTkFont(family="Segoe UI", size=12),
                     wraplength=880, justify="left").grid(
                         row=1, column=0, padx=24, pady=(0, 14), sticky="w")
        selection = ctk.CTkFrame(self, fg_color="transparent")
        selection.grid(row=2, column=0, padx=24, sticky="ew")
        selection.grid_columnconfigure(0, weight=1)
        self.columns_list = tk.Listbox(selection, selectmode="multiple", exportselection=False, height=5,
                                       font=("Segoe UI", 10), activestyle="none", **styles.listbox_colors())
        self.columns_list.grid(row=0, column=0, sticky="ew")
        scroll = ttk.Scrollbar(selection, orient="vertical", command=self.columns_list.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.columns_list.configure(yscrollcommand=scroll.set)
        for column in self.columns:
            self.columns_list.insert(tk.END, column)
        for index, column in enumerate(self.columns):
            if all(column in table.columns for _, table in self.tables):
                self.columns_list.selection_set(index)
        ctk.CTkButton(selection, text="Find matches", command=self.find_matches, height=36,
                      corner_radius=9, font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                      **styles.PRIMARY_BUTTON).grid(
            row=0, column=2, padx=(12, 0), sticky="s")
        self.status = ctk.StringVar(value="Select matching columns, then find matches.")
        ctk.CTkLabel(self, textvariable=self.status, anchor="w", text_color=styles.MUTED,
                     font=ctk.CTkFont(family="Segoe UI", size=12)).grid(
            row=3, column=0, padx=24, pady=10, sticky="ew")

        result_frame = ctk.CTkFrame(self, fg_color="transparent")
        result_frame.grid(row=4, column=0, padx=24, sticky="nsew")
        result_frame.grid_columnconfigure(0, weight=1)
        result_frame.grid_rowconfigure(0, weight=1)
        self.results = ttk.Treeview(result_frame, columns=("match", "source", "row", "values"),
                                    show="headings", style=styles.table_style(result_frame))
        for column, title, width in (("match", "Match", 65), ("source", "Source", 230),
                                     ("row", "Loaded record", 100), ("values", "Record values", 480)):
            self.results.heading(column, text=title)
            self.results.column(column, width=width, minwidth=60, stretch=column == "values")
        self.results.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(result_frame, orient="vertical", command=self.results.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(result_frame, orient="horizontal", command=self.results.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.results.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.details = tk.Text(self, height=6, wrap="word", state="disabled",
                               font=("Segoe UI", 10), padx=10, pady=8, **styles.listbox_colors())
        self.details.grid(row=5, column=0, padx=24, pady=(10, 0), sticky="ew")
        self.results.bind("<<TreeviewSelect>>", self._show_details)
        ctk.CTkButton(self, text="Close", command=self.destroy, height=36, corner_radius=9,
                      font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                      **styles.SECONDARY_BUTTON).grid(
            row=6, column=0, padx=24, pady=18, sticky="e")

    def find_matches(self):
        columns = [self.columns[index] for index in self.columns_list.curselection()]
        if not columns:
            messagebox.showinfo("Review duplicate records", "Select at least one matching column.", parent=self)
            return
        groups = find_duplicate_records(self.tables, columns)
        self.results.delete(*self.results.get_children())
        self._show_details()
        for index, group in enumerate(groups, start=1):
            for record in group:
                values = " | ".join(f"{column}: {'' if pd.isna(value) else value}"
                                    for column, value in record.values.items())
                self.results.insert("", "end", values=(index, record.source, record.row_position, values))
        count = sum(len(group) for group in groups)
        skipped = sum(any(column not in table.columns for column in columns) for _, table in self.tables)
        self.status.set(f"{len(groups):,} possible matches, {count:,} records. {skipped:,} source(s) lack a selected column. Positions are in the loaded data.")

    def _show_details(self, _event=None):
        selection = self.results.selection()
        self.details.configure(state="normal")
        self.details.delete("1.0", tk.END)
        if selection:
            match, source, row, values = self.results.item(selection[0], "values")
            self.details.insert("1.0", f"Match {match}\nSource: {source}\nLoaded record: {row}\n{values}")
        self.details.configure(state="disabled")
