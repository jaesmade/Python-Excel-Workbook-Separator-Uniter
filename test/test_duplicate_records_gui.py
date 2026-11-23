import tkinter as tk
from tkinter import TclError
from unittest.mock import patch

import pandas as pd
import pytest

from excel_splitter.core import ALL_SHEETS
from excel_splitter.gui.main_window import ExcelSplitterApp


def test_review_shows_matching_records_and_sources_without_changing_groups():
    try:
        app = ExcelSplitterApp()
    except TclError as exc:
        pytest.skip(f"Tk display unavailable: {exc}")
    try:
        app.selected_sheet_name = ALL_SHEETS
        app.headers = ["Name", "City"]
        app.input_tables = [
            ("first.xlsx - People", pd.DataFrame({"Name": ["Alice", "Bob"], "City": ["Manila", "Cebu"]})),
            ("second.xlsx - Staff", pd.DataFrame({"Name": [" alice "], "City": ["Davao"]})),
        ]
        app.groups = {("existing",): [("unchanged",)]}
        app.selected_group_keys = {("existing",)}
        original = [table.copy(deep=True) for _, table in app.input_tables]

        def inspect(dialog):
            dialog.columns_list.selection_clear(0, tk.END)
            dialog.columns_list.selection_set(0)  # Match on Name only.
            dialog.find_matches()
            rows = [dialog.results.item(item, "values") for item in dialog.results.get_children()]
            assert len(rows) == 2
            assert "first.xlsx - People" in str(rows[0])
            assert "second.xlsx - Staff" in str(rows[1])
            assert "Manila" in str(rows[0]) and "Davao" in str(rows[1])
            dialog.destroy()

        with patch.object(app, "wait_window", side_effect=inspect):
            app.review_duplicate_records()
        assert app.groups == {("existing",): [("unchanged",)]}
        assert app.selected_group_keys == {("existing",)}
        for (_, table), expected in zip(app.input_tables, original):
            pd.testing.assert_frame_equal(table, expected)
    finally:
        app.destroy()



def test_review_defaults_to_columns_shared_by_all_sources():
    try:
        app = ExcelSplitterApp()
    except TclError as exc:
        pytest.skip(f"Tk display unavailable: {exc}")
    try:
        app.selected_sheet_name = ALL_SHEETS
        app.headers = ["Name", "Note"]
        app.input_tables = [
            ("first", pd.DataFrame({"Name": ["Alice"], "Note": ["x"]})),
            ("second", pd.DataFrame({"Name": ["Alice"]})),
        ]
        def inspect(dialog):
            assert dialog.columns_list.curselection() == (0,)
            dialog.find_matches()
            assert len(dialog.results.get_children()) == 2
            dialog.destroy()
        with patch.object(app, "wait_window", side_effect=inspect):
            app.review_duplicate_records()
    finally:
        app.destroy()
