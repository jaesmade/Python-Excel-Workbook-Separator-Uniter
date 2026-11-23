"""Regression tests for Excel-compatible export tables."""

import pandas as pd
import pytest
import xlrd
from openpyxl import load_workbook

from excel_splitter.core import WorkbookProcessingError, split_input_tables, split_workbooks


@pytest.mark.filterwarnings("error:File may not be readable")
def test_error_like_header_is_exported_as_table_header_text(tmp_path):
    source = tmp_path / "staff.csv"
    source.write_text("Department,#N/A,=1+1\nSales,value,other\n", encoding="utf-8")
    output = tmp_path / "out.xlsx"

    split_workbooks([str(source)], None, "Department", None, str(output))

    workbook = load_workbook(output)
    try:
        sheet = workbook["Sales"]
        assert [(cell.value, cell.data_type) for cell in sheet[1]] == [
            ("Department", "s"), ("#N/A", "s"), ("=1+1", "s"),
        ]
        assert sheet.tables["Attendance_1"].column_names == ["Department", "#N/A", "=1+1"]
    finally:
        workbook.close()


def test_oversized_table_header_is_rejected_before_output_is_overwritten(tmp_path):
    source = pd.DataFrame([["Sales", "value"]], columns=["Department", "H" * 256])
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"existing output")

    with pytest.raises(WorkbookProcessingError, match=r"column 2.*255"):
        split_input_tables([("Staff", source)], "Department", None, str(output))

    assert output.read_bytes() == b"existing output"


@pytest.mark.parametrize(("extension", "length"), [("xlsx", 255), ("xls", 256)])
def test_header_at_table_limit_and_longer_legacy_header_are_preserved(tmp_path, extension, length):
    header = "H" * length
    source = pd.DataFrame([["Sales", "value"]], columns=["Department", header])
    output = tmp_path / f"out.{extension}"

    split_input_tables([("Staff", source)], "Department", None, str(output))

    if extension == "xls":
        assert xlrd.open_workbook(str(output)).sheet_by_index(0).cell_value(0, 1) == header
    else:
        workbook = load_workbook(output)
        try:
            sheet = workbook["Sales"]
            assert sheet["B1"].value == header
            assert sheet.tables["Attendance_1"].column_names[1] == header
        finally:
            workbook.close()


def test_oversized_combined_table_width_is_rejected_before_output_is_overwritten(tmp_path):
    first = pd.DataFrame([["Sales"] + [1] * 8192], columns=["Department"] + [f"A{i}" for i in range(8192)])
    second = first.rename(columns={f"A{i}": f"B{i}" for i in range(8192)})
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"existing output")

    with pytest.raises(WorkbookProcessingError, match=r"16,385.*16,384"):
        split_input_tables([("First", first), ("Second", second)], "Department", None, str(output))

    assert output.read_bytes() == b"existing output"


def test_table_at_excel_column_limit_is_exported(tmp_path):
    headers = ["Department"] + [f"Field {index}" for index in range(2, 16385)]
    source = pd.DataFrame([["Sales"] + [1] * 16383], columns=headers)
    output = tmp_path / "out.xlsx"

    split_input_tables([("Staff", source)], "Department", None, str(output))

    workbook = load_workbook(output)
    try:
        sheet = workbook["Sales"]
        table = sheet.tables["Attendance_1"]
        assert table.ref == "A1:XFD2"
        assert table.column_names == headers
        assert sheet["XFD2"].value == 1
    finally:
        workbook.close()
