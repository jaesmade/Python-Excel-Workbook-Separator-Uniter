import os

import pandas as pd
import pytest
import xlrd
from openpyxl import Workbook, load_workbook

from excel_splitter.core import (
    WorkbookProcessingError,
    read_workbook_sheet,
    split_input_tables,
    split_workbooks,
)


def test_file_export_does_not_overwrite_input(tmp_path):
    source = tmp_path / "source.csv"
    original = b"Type,Name\nOld,Alice\n"
    source.write_bytes(original)
    with pytest.raises(WorkbookProcessingError, match="different output"):
        split_workbooks([str(source)], None, "Type", None, str(source))
    assert source.read_bytes() == original


def test_file_export_rejects_hard_link_to_second_input(tmp_path):
    first = tmp_path / "first.csv"
    second = tmp_path / "second.csv"
    alias = tmp_path / "alias.xlsx"
    first.write_text("Type,Name\nA,Alice\n", encoding="utf-8")
    second.write_text("Type,Name\nB,Bob\n", encoding="utf-8")
    try:
        os.link(second, alias)
    except OSError as exc:
        pytest.skip(f"Hard links unavailable: {exc}")
    with pytest.raises(WorkbookProcessingError, match="different output"):
        split_workbooks([str(first), str(second)], None, "Type", None, str(alias))
    assert second.read_text(encoding="utf-8") == "Type,Name\nB,Bob\n"


@pytest.mark.parametrize("extension", ["xlsx", "xls"])
def test_opt_in_corrects_only_merged_grouping_cells_without_changing_input(tmp_path, extension):
    source = pd.DataFrame({
        "Name": ["Old", "Middle", "Unchanged"],
        "Type": [" Old ", "Middle", "Final"],
        "Area": ["Nroth", "North", "North"],
        "Number": [1, 2, 3],
    })
    original = source.copy(deep=True)
    output = tmp_path / f"corrected.{extension}"
    mappings = {("Type", "Old"): "Middle", ("Type", "Middle"): "Final",
                ("Area", "Nroth"): "North"}

    split_input_tables([("source", source)], "Type", "Area", str(output),
                       grouping_columns=["Type", "Area"], value_mappings=mappings,
                       correct_merged_values=True)

    if extension == "xlsx":
        workbook = load_workbook(output, data_only=True)
        try:
            rows = list(workbook.active.values)[1:]
        finally:
            workbook.close()
    else:
        workbook = xlrd.open_workbook(str(output))
        sheet = workbook.sheet_by_index(0)
        rows = [tuple(sheet.row_values(i)) for i in range(1, sheet.nrows)]
    assert rows == [("Old", "Final", "North", 1),
                    ("Middle", "Final", "North", 2),
                    ("Unchanged", "Final", "North", 3)]
    pd.testing.assert_frame_equal(source, original)


def test_correction_preserves_unmapped_values_and_literal_formula_text(tmp_path):
    table = pd.DataFrame({"Type": ["Old", " =Preferred ", " Other "], "Value": [1, 2, 3]})
    output = tmp_path / "out.xlsx"
    split_input_tables([("source", table)], "Type", None, str(output),
                       value_mappings={("Type", "Old"): "=Preferred"},
                       correct_merged_values=True)
    workbook = load_workbook(output, data_only=False)
    try:
        sheet = next(sheet for sheet in workbook if sheet.title.startswith("=Preferred"))
        assert sheet.cell(2, 1).value == "=Preferred"
        assert sheet.cell(2, 1).data_type == "s"
        assert sheet.cell(3, 1).value == "=Preferred"
        assert workbook["Other"].cell(2, 1).value == " Other "
    finally:
        workbook.close()


@pytest.mark.parametrize("extension", ["xlsx", "xls"])
@pytest.mark.parametrize("regions", [("North", "South"), ("South",)], ids=["all", "merged-only"])
def test_numeric_merge_correction_preserves_cell_types_and_formats(tmp_path, extension, regions):
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(("ID", "Region"))
    sheet.append((12, "North"))
    sheet.append((13, "South"))
    sheet["A2"].number_format = "0000"
    sheet["A3"].number_format = "00000"
    workbook.save(source)
    workbook.close()
    table, _, _ = read_workbook_sheet(str(source), sheet.title)
    output = tmp_path / f"output.{extension}"

    split_input_tables([("source", table)], "ID", "Region", str(output),
                       selected_groups={("12", region) for region in regions},
                       value_mappings={("ID", "13"): "12"}, correct_merged_values=True)

    formats = {"North": "0000", "South": "00000"}
    if extension == "xlsx":
        workbook = load_workbook(output, read_only=True)
        try:
            for region in regions:
                cell = workbook[f"12 - {region}"]["A2"]
                assert (cell.value, cell.data_type, cell.number_format) == (12, "n", formats[region])
        finally:
            workbook.close()
    else:
        with xlrd.open_workbook(output, formatting_info=True) as book:
            for region in regions:
                cell = book.sheet_by_name(f"12 - {region}").cell(1, 0)
                number_format = book.format_map[book.xf_list[cell.xf_index].format_key].format_str
                assert (cell.value, cell.ctype, number_format) == (12, xlrd.XL_CELL_NUMBER, formats[region])
