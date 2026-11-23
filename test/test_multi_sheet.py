"""Regression coverage for importing and exporting every selected worksheet.

Run with ``python -m unittest discover -s test -p test_multi_sheet.py``.
"""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from types import MethodType, SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd
import xlrd
import xlwt
from openpyxl import Workbook, load_workbook

from excel_splitter.saved_spellings import SavedSpellings
from excel_splitter.core import (
    ALL_SHEETS,
    WorkbookProcessingError,
    _iter_input_tables,
    build_combined_group_map,
    split_input_tables,
    split_workbooks,
)
from excel_splitter.gui.main_window import ExcelSplitterApp


class Variable:
    """The get/set behavior needed by the UI without starting a Tk window."""

    def __init__(self, value=None):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class MultiSheetWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary_directory.cleanup)
        self.folder = Path(self.temporary_directory.name)

    def write_workbook(self, filename, sheets):
        path = self.folder / filename
        workbook = Workbook()
        workbook.remove(workbook.active)
        for name, rows in sheets.items():
            sheet = workbook.create_sheet(name)
            for row in rows:
                sheet.append(row)
        workbook.save(path)
        workbook.close()
        return str(path)

    def output_rows(self, path):
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            result = []
            for sheet in workbook:
                rows = iter(sheet.iter_rows(values_only=True))
                headers = next(rows)
                result.extend(dict(zip(headers, row)) for row in rows)
            return result
        finally:
            workbook.close()

    def split(self, sources, selection=ALL_SHEETS, **options):
        output = str(self.folder / "output.xlsx")
        result = split_workbooks(
            sources, selection, "Department", None, output,
            grouping_columns=["Department"], **options,
        )
        return output, result

    def make_app(self, sources):
        app = SimpleNamespace(
            file_paths=sources, file_path=None, workbook={}, input_tables=[],
            sheet_options={}, selected_sheet_name=None, headers=[], groups={},
            saved_spellings=SavedSpellings(self.folder / "saved-spellings.json"),
            progress_var=Variable(), progress_bar=Mock(), sheet_var=Variable(),
            sheet_dropdown=Mock(), update_idletasks=Mock(),
            group_column_controls=[(None, Variable(""), Mock(), None)],
            combine_duplicates_var=Variable(True), group_search_var=Variable(),
            _clear_group_list=Mock(), _render_groups=Mock(),
        )
        for name in (
            "load_workbook", "_clear_worksheet_selection", "on_sheet_selected",
            "_selected_tables", "_populate_groups", "_selected_group_columns",
            "on_column_selection_changed",
        ):
            setattr(app, name, MethodType(getattr(ExcelSplitterApp, name), app))
        error_patcher = patch("excel_splitter.gui.main_window.messagebox.showerror")
        app.errors = error_patcher.start()
        self.addCleanup(error_patcher.stop)
        return app

    def test_all_sheets_and_named_sheet_export_the_selected_rows(self):
        source = self.write_workbook("staff.xlsx", {
            "January": [
                ["Department", "Name"], ["Sales", "Alice"],
            ],
            "February": [
                ["Department", "Name"],
                ["Support", "Bob"], ["Sales", "Cara"],
            ],
        })

        output, result = self.split([source])
        self.assertEqual(result["created_sheets"], 2)
        self.assertEqual(result["processed_rows"], 3)
        self.assertCountEqual(
            [row["Name"] for row in self.output_rows(output)],
            ["Alice", "Bob", "Cara"],
        )

        output, result = self.split([source], "February")
        self.assertEqual(result["processed_rows"], 2)
        self.assertCountEqual(
            [row["Name"] for row in self.output_rows(output)], ["Bob", "Cara"],
        )

    def test_default_table_selection_reads_first_sheet(self):
        source = self.write_workbook("staff.xlsx", {
            "First": [["Department", "Name"], ["Sales", "Alice"]],
            "Second": [["Department", "Name"], ["Support", "Bob"]],
        })

        tables = list(_iter_input_tables(source, None))

        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0][0], "First")
        self.assertEqual(tables[0][1]["Name"].tolist(), ["Alice"])

    def test_named_sheet_export_accepts_the_former_all_sheets_marker(self):
        source = self.write_workbook("staff.xlsx", {
            "Other": [["Department", "Name"], ["Sales", "Bob"]],
            "__ALL_SHEETS__": [["Department", "Name"], ["Sales", "Alice"]],
        })

        output, _ = self.split([source], "__ALL_SHEETS__")

        self.assertEqual(self.output_rows(output), [{"Department": "Sales", "Name": "Alice"}])

    def test_multiple_workbooks_align_reordered_and_optional_columns(self):
        first = self.write_workbook("first.xlsx", {
            "Casual": [
                ["Department", "Name", "ID Number", "Status"],
                ["Health", "Alice", "CASUAL", "C/In"],
            ],
            "Permanent": [
                ["Name", "ID Number", "Department", "Status"],
                ["Bob", "PERMANENT", "Health", "C/Out"],
            ],
        })
        second = self.write_workbook("second.xlsx", {
            "Other": [
                ["Workcode", "Name", "Department"], [0, "Cara", "Health"],
            ],
        })
        output = str(self.folder / "merged.xlsx")

        result = split_workbooks(
            [first, second], ALL_SHEETS, "Department", "ID Number", output,
            grouping_columns=["Department", "ID Number"],
        )

        self.assertEqual(result["created_sheets"], 3)
        self.assertEqual(result["processed_rows"], 3)
        by_name = {row["Name"]: row for row in self.output_rows(output)}
        self.assertEqual(list(by_name["Alice"]), [
            "Department", "Name", "ID Number", "Status", "Workcode",
        ])
        self.assertEqual(by_name["Bob"]["Department"], "Health")
        self.assertEqual(by_name["Bob"]["ID Number"], "PERMANENT")
        self.assertEqual(by_name["Bob"]["Status"], "C/Out")
        self.assertIsNone(by_name["Alice"]["Workcode"])
        self.assertEqual(by_name["Cara"]["Workcode"], 0)
        self.assertIsNone(by_name["Cara"]["Status"])
        self.assertIsNone(by_name["Cara"]["ID Number"])
        workbook = load_workbook(output, read_only=True)
        self.addCleanup(workbook.close)
        self.assertIn("Health - UNKNOWN GROUP", workbook.sheetnames)

    def test_empty_and_whitespace_only_sheets_and_rows_are_skipped(self):
        source = self.write_workbook("with_blanks.xlsx", {
            "Empty": [],
            "Headers only": [["Unused"]],
            "Whitespace only": [["Unused"], ["  "], ["\t"]],
            "Records": [
                ["Department", "Name"],
                ["Sales", "Alice"], [None, None], [" ", "\t"],
                ["Support", "Bob"],
            ],
        })

        output, result = self.split([source])

        self.assertEqual(result["processed_rows"], 2)
        self.assertEqual(result["created_sheets"], 2)
        rows = self.output_rows(output)
        self.assertCountEqual([row["Name"] for row in rows], ["Alice", "Bob"])
        self.assertEqual(list(rows[0]), ["Department", "Name"])

    def test_combined_group_map_retains_source_identity_when_requested(self):
        tables = [
            ("first.xlsx - First", pd.DataFrame({
                "Department": ["Sales"], "Name": ["Alice"],
            })),
            ("second.xlsx - Second", pd.DataFrame({
                "Name": ["Bob"], "Department": ["Sales"],
            })),
        ]

        headers, combined = build_combined_group_map(tables, "Department")
        _, separate = build_combined_group_map(
            tables, "Department", combine_duplicates=False,
        )

        self.assertEqual(headers, ["Department", "Name"])
        self.assertEqual(list(combined), [("Sales",)])
        self.assertEqual(combined[("Sales",)], [
            ("Sales", "Alice"), ("Sales", "Bob"),
        ])
        self.assertEqual(list(separate), [
            ("Sales", "first.xlsx - First"),
            ("Sales", "second.xlsx - Second"),
        ])

    def test_preview_group_subset_exports_matching_rows_with_both_combine_modes(self):
        source = self.write_workbook("staff.xlsx", {
            "First": [
                ["Department", "Name"],
                ["Sales", "Alice"], ["Support", "Cara"],
            ],
            "Second": [["Department", "Name"], ["Sales", "Bob"]],
        })
        input_tables = [
            (f"{os.path.basename(source)} - {name}", table)
            for name, table in _iter_input_tables(source, ALL_SHEETS)
        ]

        for combine in (True, False):
            with self.subTest(combine_duplicates=combine):
                _, preview_groups = build_combined_group_map(
                    input_tables, "Department", combine_duplicates=combine,
                )
                selected = next(key for key in preview_groups if key[0] == "Sales")
                output, result = self.split(
                    [source], selected_groups={selected},
                    combine_duplicates=combine,
                )

                self.assertEqual(result["created_sheets"], 1)
                expected_names = [row[1] for row in preview_groups[selected]]
                self.assertEqual(result["processed_rows"], len(expected_names))
                self.assertCountEqual(
                    [row["Name"] for row in self.output_rows(output)],
                    expected_names,
                )

    def test_unknown_selected_groups_raise_before_creating_output(self):
        source = self.write_workbook("staff.xlsx", {
            "First": [["Department", "Name"], ["Sales", "Alice"]],
        })

        for selected in ({("Missing",)}, {("Sales",), ("Missing",)}):
            with self.subTest(selected_groups=selected):
                with self.assertRaises(WorkbookProcessingError):
                    self.split([source], selected_groups=selected)
                self.assertFalse((self.folder / "output.xlsx").exists())

    def test_stale_selection_does_not_overwrite_existing_output(self):
        source = self.write_workbook("staff.xlsx", {
            "First": [["Department", "Name"], ["Sales", "Alice"]],
        })
        output = self.folder / "output.xlsx"
        original = b"previous output must remain unchanged"
        output.write_bytes(original)

        with self.assertRaises(WorkbookProcessingError):
            self.split([source], selected_groups={("Sales", "old source")})

        self.assertEqual(output.read_bytes(), original)

    def test_all_blank_inputs_raise_a_processing_error(self):
        source = self.write_workbook("blank.xlsx", {
            "Empty": [], "Whitespace": [["Department"], [" "]],
        })

        with self.assertRaises(WorkbookProcessingError):
            self.split([source])

        self.assertFalse((self.folder / "output.xlsx").exists())

    def test_legacy_xls_reads_all_sheets_and_writes_all_groups(self):
        source = str(self.folder / "legacy.xls")
        output = str(self.folder / "legacy_output.xls")
        workbook = xlwt.Workbook()
        for name, rows in {
            "First": [["Department", "Name"], ["Sales", "Alice"]],
            "Second": [
                ["Department", "Name"], ["Sales", "Bob"], ["Support", "Cara"],
            ],
        }.items():
            sheet = workbook.add_sheet(name)
            for row_index, row in enumerate(rows):
                for column_index, value in enumerate(row):
                    sheet.write(row_index, column_index, value)
        workbook.save(source)

        result = split_workbooks(
            [source], ALL_SHEETS, "Department", None, output,
            grouping_columns=["Department"],
        )

        self.assertEqual(result["created_sheets"], 2)
        self.assertEqual(result["processed_rows"], 3)
        with xlrd.open_workbook(output) as exported:
            self.assertEqual(exported.sheet_names(), ["Sales", "Support"])
            self.assertEqual(exported.sheet_by_name("Sales").col_values(1)[1:], [
                "Alice", "Bob",
            ])
            self.assertEqual(exported.sheet_by_name("Support").row_values(1), [
                "Support", "Cara",
            ])

    def test_macro_enabled_extension_reads_all_sheets(self):
        source = self.write_workbook("staff.xlsm", {
            "First": [["Department", "Name"], ["Sales", "Alice"]],
            "Second": [["Department", "Name"], ["Support", "Bob"]],
        })

        output, result = self.split([source])

        self.assertEqual(result["processed_rows"], 2)
        self.assertCountEqual(
            [row["Name"] for row in self.output_rows(output)], ["Alice", "Bob"],
        )

    def test_ui_defaults_multi_sheet_workbook_to_friendly_all_sheets_choice(self):
        source = self.write_workbook("staff.xlsx", {
            "First": [["Department", "Name"], ["Sales", "Alice"]],
            "Second": [["Department", "Name"], ["Support", "Bob"]],
        })
        app = self.make_app([source])

        app.load_workbook(source)

        app.errors.assert_not_called()
        self.assertEqual(app.sheet_var.get(), "All sheets")
        self.assertEqual(app.selected_sheet_name, ALL_SHEETS)
        self.assertEqual(list(app.sheet_options), ["All sheets", "First", "Second"])
        self.assertEqual(set(app.groups), {("Sales",), ("Support",)})

    def test_ui_sheet_selection_aligns_tables_only_for_grouping(self):
        app = self.make_app([str(self.folder / "staff.xlsx")])
        app.sheet_options = {"All sheets": ALL_SHEETS}
        app.input_tables = [
            ("Blank", pd.DataFrame([[" "]], columns=["Unused"])),
            ("First", pd.DataFrame([["Alice", "Sales"]], columns=[" Name ", "Department"])),
            ("Second", pd.DataFrame([["Support", "Engineer"]], columns=["Department", "Title"])),
        ]

        with patch.object(
            pd.DataFrame, "reindex", autospec=True, side_effect=pd.DataFrame.reindex,
        ) as reindex:
            self.assertTrue(app.on_sheet_selected("All sheets"))

        self.assertEqual(app.headers, ["Name", "Department", "Title"])
        self.assertEqual(list(app.groups), [("Sales",), ("Support",)])
        self.assertLessEqual(reindex.call_count, 2)

    def test_ui_real_sheet_named_all_sheets_remains_selectable(self):
        source = self.write_workbook("staff.xlsx", {
            "All sheets": [["Department", "Name"], ["Sales", "Alice"]],
            "All sheets (combined)": [
                ["Department", "Name"], ["Support", "Bob"],
            ],
        })
        app = self.make_app([source])

        app.load_workbook(source)

        app.errors.assert_not_called()
        self.assertEqual(app.sheet_var.get(), "All sheets (combined) (combined)")
        self.assertEqual(app.selected_sheet_name, ALL_SHEETS)
        app.on_sheet_selected("All sheets")
        self.assertEqual(app.selected_sheet_name, "All sheets")
        self.assertEqual(set(app.groups), {("Sales",)})

    def test_ui_single_sheet_workbook_keeps_named_sheet_default(self):
        source = self.write_workbook("staff.xlsx", {
            "Staff": [["Department", "Name"], ["Sales", "Alice"]],
        })
        app = self.make_app([source])

        app.load_workbook(source)

        app.errors.assert_not_called()
        self.assertEqual(app.sheet_var.get(), "Staff")
        self.assertEqual(app.selected_sheet_name, "Staff")
        self.assertEqual(set(app.groups), {("Sales",)})

    def test_ui_exports_only_the_worksheet_named_like_the_former_marker(self):
        source = self.write_workbook("staff.xlsx", {
            "Other": [["Department", "Name"], ["Sales", "Bob"]],
            "__ALL_SHEETS__": [["Department", "Name"], ["Sales", "Alice"]],
        })
        app = self.make_app([source])
        app.load_workbook(source)
        self.assertEqual(len(app.groups[("Sales",)]), 2)

        self.assertIn("__ALL_SHEETS__", app.sheet_options)
        self.assertTrue(app.on_sheet_selected("__ALL_SHEETS__"))
        output = str(self.folder / "output.xlsx")
        split_input_tables(app._selected_tables(), "Department", None, output, {("Sales",)})

        self.assertEqual(self.output_rows(output), [{"Department": "Sales", "Name": "Alice"}])
        self.assertEqual(app.groups[("Sales",)], [("Sales", "Alice")])

    def test_ui_multi_file_preview_honors_duplicate_checkbox_without_rereading(self):
        first = self.write_workbook("first.xlsx", {
            "First": [["Department", "Name"], ["Sales", "Alice"]],
        })
        second = self.write_workbook("second.xlsx", {
            "Second": [["Name", "Department"], ["Bob", "Sales"]],
        })
        app = self.make_app([first, second])
        app.load_workbook(first)
        self.assertEqual(app.sheet_var.get(), "All sheets")
        self.assertEqual(list(app.groups), [("Sales",)])
        self.assertEqual(len(app.groups[("Sales",)]), 2)

        with patch(
            "excel_splitter.gui.main_window._iter_input_tables",
            side_effect=AssertionError("The loaded worksheets should be cached"),
        ) as read_tables:
            app.combine_duplicates_var.set(False)
            app.on_column_selection_changed()

        read_tables.assert_not_called()
        app.errors.assert_not_called()
        self.assertEqual(list(app.groups), [
            ("Sales", "first.xlsx - First"), ("Sales", "second.xlsx - Second"),
        ])
        self.assertEqual(sum(len(rows) for rows in app.groups.values()), 2)


if __name__ == "__main__":
    unittest.main()
