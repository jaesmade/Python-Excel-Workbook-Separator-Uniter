import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from excel_splitter import core


class WorksheetPreviewTests(unittest.TestCase):
    def test_preview_and_export_use_excel_compatible_titles(self):
        prefix = "x" * 30
        departments = ["History - 2", "History", "hIsToRy", "'Ops", "Ops'", "O'Neil", prefix + "'s", "' '"]
        expected_titles = ["History - 2", "History - 3", "hIsToRy - 4", "Ops", "Ops - 2", "O'Neil", prefix, "GROUP"]
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.csv"
            output = Path(directory) / "output.xlsx"
            pd.DataFrame({"Department": departments}).to_csv(source, index=False)
            tables = list(core._iter_input_tables(str(source), core.ALL_SHEETS))
            _, groups = core.build_combined_group_map(tables, "Department")
            preview = core.preview_group_items(groups, set(groups), core.SORT_COUNT_DESCENDING)
            core.split_workbooks(
                [str(source)], core.ALL_SHEETS, "Department", None, str(output),
                sort_mode=core.SORT_COUNT_DESCENDING,
            )

            workbook = load_workbook(output, read_only=True)
            try:
                self.assertEqual([item[0] for item in preview], expected_titles)
                self.assertEqual(workbook.sheetnames, expected_titles)
                self.assertEqual([sheet["A2"].value for sheet in workbook], departments)
            finally:
                workbook.close()

    def test_preview_names_match_excel_for_case_insensitive_collisions(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.csv"
            output = Path(directory) / "output.xlsx"
            pd.DataFrame({"Department": ["Sales", "sales"]}).to_csv(source, index=False)
            tables = list(core._iter_input_tables(str(source), core.ALL_SHEETS))
            _, groups = core.build_combined_group_map(tables, "Department")
            preview = core.preview_group_items(groups, set(groups))
            core.split_workbooks([str(source)], core.ALL_SHEETS, "Department", None, str(output))
            workbook = load_workbook(output, read_only=True)
            try:
                self.assertEqual([item[0] for item in preview], workbook.sheetnames)
            finally:
                workbook.close()

    def test_preview_matches_selected_export_names_order_and_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.csv"
            output = Path(directory) / "output.xlsx"
            pd.DataFrame({
                "Department": ["A/B", "A:B", "A:B", "Excluded"],
                "Name": ["Alice", "Bob", "Cara", "Dan"],
            }).to_csv(source, index=False)
            tables = list(core._iter_input_tables(str(source), core.ALL_SHEETS))
            headers, groups = core.build_combined_group_map(tables, "Department")
            selected = {("A/B",), ("A:B",)}

            preview = core.preview_group_items(groups, selected, core.SORT_COUNT_DESCENDING)
            self.assertEqual([item[0] for item in preview], ["A-B", "A-B - 2"])
            self.assertEqual([item[1] for item in preview], [("A:B",), ("A/B",)])
            self.assertEqual(preview[0][2], [("A:B", "Bob"), ("A:B", "Cara")])

            core.split_workbooks([str(source)], core.ALL_SHEETS, "Department", None,
                                 str(output), selected_groups=selected,
                                 sort_mode=core.SORT_COUNT_DESCENDING)
            workbook = load_workbook(output, read_only=True, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, [item[0] for item in preview])
                for title, _key, rows in preview:
                    actual = list(workbook[title].values)
                    self.assertEqual(actual[0], tuple(headers))
                    self.assertEqual(actual[1:], rows)
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
