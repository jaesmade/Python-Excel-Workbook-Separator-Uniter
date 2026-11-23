"""Saved spelling decisions through the review and management interfaces."""

import tempfile
import unittest
from pathlib import Path
from tkinter import TclError
from unittest.mock import patch

from excel_splitter.gui.main_window import ExcelSplitterApp


class SavedSpellingsGuiTests(unittest.TestCase):
    def setUp(self):
        from excel_splitter.saved_spellings import SavedSpellings

        self.folder = tempfile.TemporaryDirectory()
        self.addCleanup(self.folder.cleanup)
        self.store_path = Path(self.folder.name) / "saved-spellings.json"
        try:
            self.app = ExcelSplitterApp()
        except TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.addCleanup(self.app.destroy)
        self.app.saved_spellings = SavedSpellings(self.store_path)
        self.app.update()

    def load_csv(self, name, rows):
        source = Path(self.folder.name) / name
        source.write_text("Type,Name\n" + rows, encoding="utf-8")
        self.app.file_paths = [str(source)]
        self.app.load_workbook(str(source))

    def review_choice(self, apply=True):
        def choose(dialog):
            index = next(i for i, item in enumerate(dialog.review.suggestions)
                         if set(item[1:3]) == {"JO ORDER", "JOB ORDER"})
            dialog._choose(index, "JOB ORDER")
            if apply:
                dialog._apply()
            else:
                dialog.destroy()

        with patch.object(self.app, "wait_window", side_effect=choose):
            self.app.review_spelling_suggestions()

    def test_applied_choice_is_saved_and_used_for_later_source_only_workbook(self):
        from excel_splitter.saved_spellings import SavedSpellings

        self.load_csv("first.csv", "JOB ORDER,Alice\nJOB ORDER,Amy\nJO ORDER,Bob\n")
        self.review_choice()
        self.assertEqual(set(self.app.groups), {("JOB ORDER",)})
        self.assertEqual(SavedSpellings(self.store_path).load(), {("Type", "JO ORDER"): "JOB ORDER"})

        # A fresh store instance checks that the choice survived a restart.
        self.app.saved_spellings = SavedSpellings(self.store_path)
        self.load_csv("later.csv", "JO ORDER,Carol\n")
        self.assertEqual(set(self.app.groups), {("JOB ORDER",)})

    def test_cancelled_choice_is_not_saved(self):
        from excel_splitter.saved_spellings import SavedSpellings

        self.load_csv("first.csv", "JOB ORDER,Alice\nJOB ORDER,Amy\nJO ORDER,Bob\n")
        self.review_choice(apply=False)
        self.assertEqual(set(self.app.groups), {("JO ORDER",), ("JOB ORDER",)})
        self.assertEqual(SavedSpellings(self.store_path).load(), {})

        self.app.saved_spellings = SavedSpellings(self.store_path)
        self.load_csv("later.csv", "JO ORDER,Carol\n")
        self.assertEqual(set(self.app.groups), {("JO ORDER",)})

    def test_removing_a_saved_choice_stops_it_applying_to_later_workbooks(self):
        from excel_splitter.saved_spellings import SavedSpellings

        self.load_csv("first.csv", "JOB ORDER,Alice\nJOB ORDER,Amy\nJO ORDER,Bob\n")
        self.review_choice()

        def remove(dialog):
            dialog.update_idletasks()
            self.assertEqual(dialog.choices_list.size(), 1)
            text = dialog.choices_list.get(0)
            self.assertIn("JO ORDER", text)
            self.assertIn("JOB ORDER", text)
            dialog.choices_list.selection_set(0)
            dialog._remove()
            self.assertEqual(dialog.choices_list.size(), 0)
            dialog.destroy()

        with patch.object(self.app, "wait_window", side_effect=remove):
            self.app.manage_saved_spellings()
        self.app.saved_spellings = SavedSpellings(self.store_path)
        self.load_csv("later.csv", "JO ORDER,Carol\n")
        self.assertEqual(set(self.app.groups), {("JO ORDER",)})

    def test_applying_an_undo_removes_the_saved_choice(self):
        from excel_splitter.saved_spellings import SavedSpellings

        self.load_csv("first.csv", "JOB ORDER,Alice\nJOB ORDER,Amy\nJO ORDER,Bob\n")
        self.review_choice()
        self.assertEqual(SavedSpellings(self.store_path).load(), {("Type", "JO ORDER"): "JOB ORDER"})

        def undo(dialog):
            dialog._undo()
            dialog._apply()

        with patch.object(self.app, "wait_window", side_effect=undo):
            self.app.review_spelling_suggestions()
        self.assertEqual(set(self.app.groups), {("JO ORDER",), ("JOB ORDER",)})
        self.app.saved_spellings = SavedSpellings(self.store_path)
        self.load_csv("later.csv", "JO ORDER,Carol\n")
        self.assertEqual(set(self.app.groups), {("JO ORDER",)})

    def test_reapplying_unchanged_review_does_not_restore_removed_choice(self):
        from excel_splitter.saved_spellings import SavedSpellings

        self.load_csv("first.csv", "JOB ORDER,Alice\nJOB ORDER,Amy\nJO ORDER,Bob\n")
        self.review_choice()
        self.assertEqual(SavedSpellings(self.store_path).load(), {("Type", "JO ORDER"): "JOB ORDER"})
        self.app.saved_spellings.remove("Type", "JO ORDER")

        with patch.object(self.app, "wait_window", side_effect=lambda dialog: dialog._apply()):
            self.app.review_spelling_suggestions()

        self.assertEqual(SavedSpellings(self.store_path).load(), {})
        self.load_csv("later.csv", "JO ORDER,Carol\n")
        self.assertEqual(set(self.app.groups), {("JO ORDER",)})


if __name__ == "__main__":
    unittest.main()
