"""Manual category merge through the application controls."""

import tempfile
import unittest
from pathlib import Path
from tkinter import TclError
from unittest.mock import patch

from excel_splitter.gui.main_window import ExcelSplitterApp


class ManualMergeGuiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.app = ExcelSplitterApp()
        except TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.addCleanup(self.app.destroy)
        self.app.update()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        source = Path(folder.name) / "source.csv"
        source.write_text("Category,Name\nAlpha,Alice\nZulu,Bob\n", encoding="utf-8")
        self.app.file_paths = [str(source)]
        self.app.load_workbook(str(source))

    def test_search_and_merge_unrelated_values_with_preferred_name(self):
        app = self.app

        def interact(dialog):
            dialog.search_vars[0].set("alp")
            dialog.search_vars[1].set("ZUL")
            self.assertEqual(dialog.value_lists[0].get(0, "end"), ("Alpha",))
            self.assertEqual(dialog.value_lists[1].get(0, "end"), ("Zulu",))
            dialog.value_lists[0].selection_set(0)
            dialog.value_lists[1].selection_set(0)
            dialog.preferred_var.set("Combined")
            self.assertEqual(set(app.groups), {("Alpha",), ("Zulu",)})
            dialog._merge()

        with patch.object(app, "wait_window", side_effect=interact):
            app.manual_merge_categories()
        self.assertEqual(set(app.groups), {("Combined",)})
        self.assertEqual(app.selected_group_keys, {("Combined",)})
        self.assertEqual(app.groups[("Combined",)], [("Alpha", "Alice"), ("Zulu", "Bob")])

    def test_cancel_leaves_groups_and_mappings_unchanged(self):
        app = self.app

        def cancel(dialog):
            dialog.search_vars[0].set("Alpha")
            dialog.value_lists[0].selection_set(0)
            dialog.destroy()

        with patch.object(app, "wait_window", side_effect=cancel):
            app.manual_merge_categories()
        self.assertEqual(set(app.groups), {("Alpha",), ("Zulu",)})
        self.assertEqual(app.value_mappings, {})

    def test_long_values_can_be_scrolled_horizontally_in_both_lists(self):
        from tkinter import ttk

        app = self.app
        long_value = "A" * 150
        app.groups = {(long_value,): [(long_value, "Alice")], ("Zulu",): [("Zulu", "Bob")]}
        app.selected_group_keys = set(app.groups)

        def inspect(dialog):
            dialog.update()
            for box in dialog.value_lists:
                scrollbars = [widget for widget in box.master.winfo_children()
                              if isinstance(widget, ttk.Scrollbar) and str(widget.cget("orient")) == "horizontal"]
                self.assertEqual(len(scrollbars), 1)
                box.xview_moveto(1.0)
                dialog.update()
                self.assertGreater(box.xview()[0], 0.0)
                self.assertGreater(scrollbars[0].get()[0], 0.0)
            dialog.destroy()

        with patch.object(app, "wait_window", side_effect=inspect):
            app.manual_merge_categories()


if __name__ == "__main__":
    unittest.main()
