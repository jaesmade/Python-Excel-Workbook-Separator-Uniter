from datetime import datetime

import pytest
import xlwt
from openpyxl import Workbook, load_workbook

from excel_splitter.core import ALL_SHEETS, WorkbookProcessingError, read_workbook_sheet, split_workbooks


def test_csv_preserves_literal_values_and_formula_like_text(tmp_path):
    source = tmp_path / "staff.csv"
    output = tmp_path / "out.xlsx"
    source.write_text("Department,ID,State,Note\nSales,00123,NA,=1+1\n", encoding="utf-8")
    split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))
    workbook = load_workbook(output, data_only=False)
    try:
        row = list(workbook.active[2])
        assert [cell.value for cell in row] == ["Sales", "00123", "NA", "=1+1"]
        assert row[-1].data_type == "s"
    finally:
        workbook.close()


def test_excel_preserves_text_numbers_and_dates(tmp_path):
    source = tmp_path / "staff.xlsx"
    output = tmp_path / "out.xlsx"
    workbook = Workbook()
    workbook.active.append(["Department", "ID", "State", "Amount", "Date"])
    workbook.active.append(["Sales", "00123", "NA", 42, datetime(2024, 5, 1)])
    workbook.save(source)
    workbook.close()
    split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))
    workbook = load_workbook(output, data_only=False)
    try:
        assert list(workbook.active.values)[1] == ("Sales", "00123", "NA", 42, datetime(2024, 5, 1))
    finally:
        workbook.close()


@pytest.mark.parametrize("leading_rows", [[[]], [[], [" ", "\t"]]])
def test_xlsx_uses_first_populated_row_as_headers(tmp_path, leading_rows):
    source = tmp_path / "staff.xlsx"
    output = tmp_path / "out.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Staff"
    for row in leading_rows:
        sheet.append(row)
    sheet.append(["Department", "ID"])
    records = [("Sales", 123, "00000"), ("Support", 456, "000000")]
    for department, identifier, number_format in records:
        sheet.append([department, identifier])
        sheet.cell(sheet.max_row, 2).number_format = number_format
    workbook.save(source)
    workbook.close()

    _, _, headers = read_workbook_sheet(str(source), "Staff")
    result = split_workbooks([str(source)], ALL_SHEETS, headers[0], None, str(output))

    assert result["processed_rows"] == 2
    assert headers == ["Department", "ID"]
    workbook = load_workbook(output)
    try:
        assert workbook.sheetnames == ["Sales", "Support"]
        for department, identifier, number_format in records:
            sheet = workbook[department]
            assert list(sheet.values) == [("Department", "ID"), (department, identifier)]
            assert sheet["B2"].number_format == number_format
    finally:
        workbook.close()


@pytest.mark.parametrize("kind", ["csv", "xlsx", "xls"])
def test_duplicate_original_headers_are_rejected(tmp_path, kind):
    source = tmp_path / f"staff.{kind}"
    if kind == "csv":
        source.write_text("Department,Name,Name\nSales,Alice,Other\n", encoding="utf-8")
    elif kind == "xlsx":
        workbook = Workbook()
        workbook.active.append(["Department", "Name", "Name"])
        workbook.active.append(["Sales", "Alice", "Other"])
        workbook.save(source)
        workbook.close()
    else:
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Staff")
        for row_index, row in enumerate((["Department", "Name", "Name"], ["Sales", "Alice", "Other"])):
            for column_index, value in enumerate(row):
                sheet.write(row_index, column_index, value)
        workbook.save(str(source))
    with pytest.raises(WorkbookProcessingError, match="Duplicate column headers"):
        split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(tmp_path / "out.xlsx"))


def test_same_named_sources_remain_separate_when_combining_is_disabled(tmp_path):
    sources = []
    for folder, name in (("first", "Alice"), ("second", "Bob")):
        directory = tmp_path / folder
        directory.mkdir()
        source = directory / "staff.csv"
        source.write_text(f"Department,Name\nSales,{name}\n", encoding="utf-8")
        sources.append(str(source))
    output = tmp_path / "out.xlsx"
    result = split_workbooks(sources, ALL_SHEETS, "Department", None, str(output), combine_duplicates=False)
    workbook = load_workbook(output, data_only=True)
    try:
        assert result["created_sheets"] == 2
        assert sorted(list(sheet.values)[1][1] for sheet in workbook) == ["Alice", "Bob"]
    finally:
        workbook.close()


def test_ambiguous_file_and_sheet_labels_remain_separate(tmp_path):
    sources = []
    for filename, sheet_name, name in (
        ("a.xlsx", "b.xlsx - Sheet", "Alice"),
        ("a.xlsx - b.xlsx", "Sheet", "Bob"),
    ):
        source = tmp_path / filename
        workbook = Workbook()
        workbook.active.title = sheet_name
        workbook.active.append(["Department", "Name"])
        workbook.active.append(["Sales", name])
        workbook.save(source)
        workbook.close()
        sources.append(str(source))
    output = tmp_path / "out.xlsx"
    result = split_workbooks(sources, ALL_SHEETS, "Department", None, str(output), combine_duplicates=False)
    workbook = load_workbook(output, data_only=True)
    try:
        assert result["created_sheets"] == 2
        assert sorted(list(sheet.values)[1][1] for sheet in workbook) == ["Alice", "Bob"]
    finally:
        workbook.close()
