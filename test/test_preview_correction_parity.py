import tkinter as tk
from tkinter import ttk

import pytest
from openpyxl import load_workbook

from excel_splitter.core import split_input_tables
from excel_splitter.gui.main_window import ExcelSplitterApp


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


@pytest.mark.parametrize("correct", [False, True])
def test_preview_rows_match_export_with_merged_values(tmp_path, correct):
    try:
        app = ExcelSplitterApp()
    except tk.TclError as exc:
        pytest.skip(f"Tk display unavailable: {exc}")
    try:
        source = tmp_path / "source.csv"
        source.write_text("Type,Name\nAlpha,Alice\nZulu,Bob\n", encoding="utf-8")
        app.file_paths = [str(source)]
        app.load_workbook(str(source))
        app.value_mappings = {("Type", "Zulu"): "Alpha"}
        app.on_column_selection_changed()
        app.correct_merged_values_var.set(correct)
        app.preview_worksheets()
        app.update()
        tree = next(widget for widget in descendants(app) if isinstance(widget, ttk.Treeview))
        preview_rows = [tuple(tree.item(item, "values")) for item in tree.get_children()]
        expected = [("Alpha", "Alice"), ("Alpha" if correct else "Zulu", "Bob")]
        assert preview_rows == expected

        output = tmp_path / "output.xlsx"
        split_input_tables(app._selected_tables(), "Type", None, str(output),
                           grouping_columns=["Type"], value_mappings=app.value_mappings,
                           correct_merged_values=correct)
        workbook = load_workbook(output, read_only=True, data_only=True)
        try:
            assert list(workbook.active.values)[1:] == expected
        finally:
            workbook.close()
        assert source.read_text(encoding="utf-8") == "Type,Name\nAlpha,Alice\nZulu,Bob\n"
    finally:
        app.destroy()
