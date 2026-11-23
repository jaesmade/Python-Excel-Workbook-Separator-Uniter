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


@pytest.mark.filterwarnings("error:In write-only mode you must add table columns manually")
def test_export_preserves_table_filters_and_worksheet_layout(tmp_path):
    source = pd.DataFrame(
        [("Sales", "Open", "0123456789\r\n0123456789"), ("Support", "Closed", None)],
        columns=["Department", "Status", "Note"],
    )
    output = tmp_path / "out.xlsx"

    split_input_tables([("Staff", source)], "Department", None, str(output))

    workbook = load_workbook(output)
    try:
        for index, title in enumerate(("Sales", "Support"), start=1):
            sheet = workbook[title]
            table = sheet.tables[f"Attendance_{index}"]
            assert table.ref == table.autoFilter.ref == "A1:C2"
            assert table.column_names == ["Department", "Status", "Note"]
            assert table.tableStyleInfo.name == "TableStyleMedium2"
            assert table.tableStyleInfo.showRowStripes
            assert sheet.freeze_panes == "C2"
            assert sheet.sheet_view.showGridLines is False
            assert sheet.row_dimensions[1].height == 31
            assert sheet.row_dimensions[2].height == 23
            assert sheet["A1"].font.bold
            assert sheet["B2"].alignment.horizontal == "center"
            assert sheet["C2"].alignment.horizontal == "left"
            assert sheet["C2"].font.name == "Arial"
        assert workbook["Sales"]["C2"].value == "0123456789\n0123456789"
        assert workbook["Sales"].column_dimensions["C"].width == 23
        assert workbook["Support"]["C2"].value is None
        assert workbook["Support"].column_dimensions["C"].width == 14
    finally:
        workbook.close()

    workbook = load_workbook(output, read_only=True)
    try:
        assert workbook["Sales"].calculate_dimension() == "A1:C2"
    finally:
        workbook.close()


def test_cancelled_stream_removes_temporary_worksheet_xml(tmp_path, monkeypatch):
    spool = tmp_path / "spool"
    spool.mkdir()
    monkeypatch.setattr("tempfile.tempdir", str(spool))
    source = pd.DataFrame([("Sales", "Open")], columns=["Department", "Status"])
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"previous workbook")

    def cancel(_progress, _message):
        worksheet_xml, = spool.iterdir()
        assert "Sales" in worksheet_xml.read_text(encoding="utf-8")
        raise RuntimeError("export cancelled")

    with pytest.raises(RuntimeError, match="export cancelled"):
        split_input_tables([("Staff", source)], "Department", None, str(output), progress_callback=cancel)

    assert not list(spool.iterdir())
    assert output.read_bytes() == b"previous workbook"
