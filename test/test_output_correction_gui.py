"""Export correction and input-file protection through the GUI."""

import tempfile
import time
import unittest
from pathlib import Path
from tkinter import TclError
from unittest.mock import patch

from openpyxl import load_workbook

from excel_splitter.gui.main_window import ExcelSplitterApp


class OutputCorrectionGuiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.app = ExcelSplitterApp()
        except TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.addCleanup(self.app.destroy)
        self.app.update()

    def test_correction_selection_is_snapshotted_for_export(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "source.csv"
            output = Path(folder) / "out.xlsx"
            source.write_text("Type,Name\nAlpha,Alice\nZulu,Bob\n", encoding="utf-8")
            app = self.app
            app.file_paths = [str(source)]
            app.load_workbook(str(source))
            app.value_mappings = {("Type", "Zulu"): "Alpha"}
            app.on_column_selection_changed()
            app.correct_merged_values_var.set(True)
            with patch.object(app, "choose_output_location", return_value=str(output)), \
                 patch("excel_splitter.gui.main_window.messagebox.showinfo"), \
                 patch("excel_splitter.gui.main_window.messagebox.showerror") as error:
                app.start_generation()
                app.correct_merged_values_var.set(False)
                deadline = time.monotonic() + 20
                while not output.exists() and time.monotonic() < deadline:
                    app.update()
                    time.sleep(0.01)
                self.assertTrue(output.exists(), "Generation did not complete")
                error.assert_not_called()
            workbook = load_workbook(output, read_only=True, data_only=True)
            try:
                self.assertEqual(list(workbook.active.values)[1:], [("Alpha", "Alice"), ("Alpha", "Bob")])
            finally:
                workbook.close()
            self.assertEqual(source.read_text(encoding="utf-8"), "Type,Name\nAlpha,Alice\nZulu,Bob\n")

    def test_save_cannot_overwrite_any_original_input(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "source.xlsx"
            source.write_bytes(b"original input")
            temporary = Path(folder) / "generated.xlsx"
            temporary.write_bytes(b"generated output")
            self.app.file_paths = [str(Path(folder) / "different.xlsx")]
            with patch.object(self.app, "choose_output_location", return_value=str(source)), \
                 patch("excel_splitter.gui.main_window.messagebox.showerror") as error:
                self.app._save_completed_workbook(str(temporary), 1, 1, "first.xlsx", ("first.xlsx", str(source)))
            self.assertEqual(source.read_bytes(), b"original input")
            error.assert_called_once()


if __name__ == "__main__":
    unittest.main()
