"""Overlapping spelling decisions remain effective in future workbooks."""

from tkinter import TclError
from unittest.mock import patch

import pytest

from excel_splitter.gui.main_window import ExcelSplitterApp
from excel_splitter.saved_spellings import SavedSpellings


def test_overlapping_choices_apply_to_a_later_workbook(tmp_path):
    try:
        app = ExcelSplitterApp()
    except TclError as exc:
        pytest.skip(f"Tk display unavailable: {exc}")
    try:
        app.saved_spellings = SavedSpellings(tmp_path / "spellings.json")
        first = tmp_path / "first.csv"
        first.write_text(
            "Type,Name\nJOB ORDER,Alice\nJOB ORDER,Amy\nJOB ORDER,Ann\n"
            "Job Order,Bob\nJob Order,Ben\nJO ORDER,Carol\n", encoding="utf-8")
        app.file_paths = [str(first)]
        app.load_workbook(str(first))

        def choose(dialog):
            first_index = next(i for i, item in enumerate(dialog.review.suggestions)
                               if set(item[1:3]) == {"Job Order", "JOB ORDER"})
            second_index = next(i for i, item in enumerate(dialog.review.suggestions)
                                if set(item[1:3]) == {"JO ORDER", "JOB ORDER"})
            dialog._choose(first_index, "Job Order")
            dialog._choose(second_index, "JO ORDER")
            dialog._apply()

        with patch.object(app, "wait_window", side_effect=choose):
            app.review_spelling_suggestions()
        assert set(app.groups) == {("JO ORDER",)}

        later = tmp_path / "later.csv"
        later.write_text("Type,Name\nJob Order,David\n", encoding="utf-8")
        app.file_paths = [str(later)]
        app.load_workbook(str(later))
        assert set(app.groups) == {("JO ORDER",)}
    finally:
        app.destroy()
