"""Integration tests for group paging and generation snapshots."""

import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
from tkinter import TclError

from excel_splitter.gui.main_window import ExcelSplitterApp


class GuiRegressionTests(unittest.TestCase):
    def setUp(self):
        try:
            self.app = ExcelSplitterApp()
        except TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.addCleanup(self.app.destroy)
        self.app.update()

    def test_group_selection_survives_paging_and_sorting(self):
        app = self.app
        app.groups = {(f"Group {i:03d}",): [(i,)] for i in range(120)}
        app._render_groups()
        self.assertLess(len(app.group_checkboxes), 120)
        app.group_checkboxes[("Group 000",)].toggle()
        app._change_group_page(1)
        second_page_key = next(iter(app.group_checkboxes))
        app.group_checkboxes[second_page_key].toggle()
        app.group_sort_var.set("Record count: descending")
        app._on_group_sort_changed()
        self.assertNotIn(("Group 000",), app.selected_group_keys)
        self.assertNotIn(second_page_key, app.selected_group_keys)
        self.assertEqual(len(app.selected_group_keys), 118)

    def test_clearing_worksheet_clears_selection_and_page_controls(self):
        app = self.app
        app.groups = {(f"Group {i:03d}",): [(i,)] for i in range(120)}
        app._render_groups()
        app._change_group_page(1)
        app._clear_worksheet_selection()
        self.assertEqual(app.selected_group_keys, set())
        self.assertEqual(app.group_previous.cget("state"), "disabled")
        self.assertEqual(app.group_next.cget("state"), "disabled")

    def test_spelling_list_choices_undo_and_export(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "source.csv"
            output = Path(folder) / "out.xlsx"
            source.write_text(
                "ID Number,Name\nJOB ORDER,Alice\nJOB ORDER,Amy\nJO ORDER,Bob\n"
                "LABOR PAYROLL,Carol\nLABOR PAYROLL,Chris\nLABOR PAYYROLL,Dana\n",
                encoding="utf-8",
            )
            app = self.app
            app.file_paths = [str(source)]
            app.load_workbook(str(source))
            def widget_text(widget):
                try:
                    value = str(widget.cget("text"))
                except (TclError, ValueError, AttributeError):
                    value = ""
                return " ".join([value] + [widget_text(child) for child in widget.winfo_children()])

            def review_list(dialog):
                from excel_splitter.gui.main_window import SpellingReviewDialog
                from excel_splitter.spelling_review import KEEP_SEPARATE

                self.assertIsInstance(dialog, SpellingReviewDialog)
                suggestions = dialog.review.suggestions
                self.assertEqual(len(suggestions), 2)
                self.assertEqual(len(dialog.suggestion_rows), 2)
                for suggestion, row in zip(suggestions, dialog.suggestion_rows):
                    column, first, second, first_count, second_count = suggestion
                    displayed = widget_text(row)
                    for expected in (column, first, second, str(first_count), str(second_count)):
                        self.assertIn(expected, displayed)

                merge_index = next(i for i, item in enumerate(suggestions) if item[2] == "JOB ORDER")
                separate_index = 1 - merge_index
                dialog._choose(merge_index, "JOB ORDER")
                self.assertNotIn(("JO ORDER",), dialog.review.groups)
                self.assertIn(("JO ORDER",), app.groups)  # Changes are staged until Apply.
                dialog._undo()
                self.assertIn(("JO ORDER",), dialog.review.groups)
                dialog._choose(merge_index, "JOB ORDER")
                dialog._choose(separate_index, KEEP_SEPARATE)
                dialog._apply()

            with patch.object(app, "wait_window", side_effect=review_list):
                app.review_spelling_suggestions()
            self.assertEqual(set(app.groups), {("JOB ORDER",), ("LABOR PAYROLL",), ("LABOR PAYYROLL",)})

            def undo_then_cancel(dialog):
                self.assertTrue(dialog.review.can_undo)
                dialog._undo()
                dialog._undo()
                self.assertIn(("JO ORDER",), dialog.review.groups)
                self.assertNotIn(("JO ORDER",), app.groups)
                dialog.destroy()

            with patch.object(app, "wait_window", side_effect=undo_then_cancel) as wait, \
                 patch("excel_splitter.gui.main_window.messagebox.showinfo"):
                app.review_spelling_suggestions()
                wait.assert_called_once()
            self.assertEqual(set(app.groups), {("JOB ORDER",), ("LABOR PAYROLL",), ("LABOR PAYYROLL",)})

            def undo_applied_choices(dialog):
                self.assertTrue(dialog.review.can_undo)
                dialog._undo()  # Undo the keep-separate choice.
                dialog._undo()  # Undo the merge.
                self.assertIn(("JO ORDER",), dialog.review.groups)
                self.assertNotIn(("JO ORDER",), app.groups)  # Undo is staged until Apply.
                dialog._apply()

            with patch.object(app, "wait_window", side_effect=undo_applied_choices) as wait, \
                 patch("excel_splitter.gui.main_window.messagebox.showinfo"):
                app.review_spelling_suggestions()
                wait.assert_called_once()
            self.assertEqual(set(app.groups), {
                ("JOB ORDER",), ("JO ORDER",), ("LABOR PAYROLL",), ("LABOR PAYYROLL",),
            })

            def decide_again(dialog):
                from excel_splitter.spelling_review import KEEP_SEPARATE

                suggestions = dialog.review.suggestions
                merge_index = next(i for i, item in enumerate(suggestions) if item[2] == "JOB ORDER")
                dialog._choose(merge_index, "JOB ORDER")
                dialog._choose(1 - merge_index, KEEP_SEPARATE)
                dialog._apply()

            with patch.object(app, "wait_window", side_effect=decide_again) as wait, \
                 patch("excel_splitter.gui.main_window.messagebox.showinfo"):
                app.review_spelling_suggestions()
                wait.assert_called_once()
            self.assertEqual(set(app.groups), {("JOB ORDER",), ("LABOR PAYROLL",), ("LABOR PAYYROLL",)})
            with patch.object(app, "choose_output_location", return_value=str(output)), \
                 patch("excel_splitter.gui.main_window.messagebox.showinfo"), \
                 patch("excel_splitter.gui.main_window.messagebox.showerror") as showerror:
                app.start_generation()
                deadline = time.monotonic() + 20
                while not output.exists() and time.monotonic() < deadline:
                    app.update()
                    time.sleep(0.01)
                self.assertTrue(output.exists(), "Generation did not complete")
                showerror.assert_not_called()
            workbook = load_workbook(output, read_only=True, data_only=True)
            try:
                self.assertEqual(set(workbook.sheetnames), {"JOB ORDER", "LABOR PAYROLL", "LABOR PAYYROLL"})
                self.assertEqual(list(workbook["JOB ORDER"].values)[1:],
                                 [("JOB ORDER", "Alice"), ("JOB ORDER", "Amy"), ("JO ORDER", "Bob")])
            finally:
                workbook.close()

    def test_generation_uses_loaded_data_after_source_and_controls_change(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "source.csv"
            output = Path(folder) / "out.xlsx"
            source.write_text("Department,Name\nSales,Alice\n", encoding="utf-8")
            app = self.app
            app.file_paths = [str(source)]
            app.load_workbook(str(source))
            source.write_text("Department,Name\nSales,Bob\n", encoding="utf-8")
            with patch.object(app, "choose_output_location", return_value=str(output)) as choose_output, \
                 patch("excel_splitter.gui.main_window.messagebox.showinfo"), \
                 patch("excel_splitter.gui.main_window.messagebox.showerror") as showerror:
                app.start_generation()
                app.file_paths = [str(Path(folder) / "other.csv")]
                app.combine_duplicates_var.set(False)
                deadline = time.monotonic() + 20
                while not output.exists() and time.monotonic() < deadline:
                    app.update()
                    time.sleep(0.01)
                self.assertTrue(output.exists(), "Generation did not complete")
                showerror.assert_not_called()
                choose_output.assert_called_once_with(str(source))
            workbook = load_workbook(output, read_only=True, data_only=True)
            try:
                self.assertEqual(list(workbook.active.values)[1], ("Sales", "Alice"))
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
