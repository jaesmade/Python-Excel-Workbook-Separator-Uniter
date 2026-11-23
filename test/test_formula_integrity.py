"""Formula inputs must never be silently converted to cached values or blanks."""

from zipfile import ZIP_DEFLATED, ZipFile

import pytest
import xlwt
from openpyxl import Workbook, load_workbook

from excel_splitter.core import (
    ALL_SHEETS,
    WorkbookProcessingError,
    read_workbook_sheet,
    split_workbooks,
)


def write_xlsx(path, formula_column, cached):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Staff"
    sheet.append(["Department", "Amount"])
    sheet.append(["Sales", 5])
    sheet.cell(2, formula_column, "=1+1")
    workbook.save(path)
    workbook.close()
    if cached:
        # openpyxl writes empty cached values; simulate a workbook saved by Excel.
        with ZipFile(path) as archive:
            files = {name: archive.read(name) for name in archive.namelist()}
        xml = files["xl/worksheets/sheet1.xml"]
        assert b"<f>1+1</f><v />" in xml
        files["xl/worksheets/sheet1.xml"] = xml.replace(
            b"<f>1+1</f><v />", b"<f>1+1</f><v>2</v>"
        )
        with ZipFile(path, "w", ZIP_DEFLATED) as archive:
            for name, contents in files.items():
                archive.writestr(name, contents)


@pytest.mark.parametrize("formula_column", [1, 2], ids=["grouping", "other"])
@pytest.mark.parametrize("cached", [False, True], ids=["uncached", "cached"])
def test_xlsx_formula_is_rejected_before_output_is_overwritten(tmp_path, formula_column, cached):
    source = tmp_path / "staff.xlsx"
    output = tmp_path / "out.xlsx"
    write_xlsx(source, formula_column, cached)
    output.write_bytes(b"existing output")

    with pytest.raises(WorkbookProcessingError, match="(?i)formula"):
        split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))

    assert output.read_bytes() == b"existing output"


def test_single_sheet_reader_rejects_formulas_before_preview(tmp_path):
    source = tmp_path / "staff.xlsx"
    write_xlsx(source, 2, False)

    with pytest.raises(WorkbookProcessingError, match="(?i)formula"):
        read_workbook_sheet(str(source), "Staff")


def test_literal_formula_like_excel_text_remains_accepted(tmp_path):
    source = tmp_path / "staff.xlsx"
    output = tmp_path / "out.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Department", "Note"])
    sheet.append(["Sales", "=1+1"])
    sheet["B2"].data_type = "s"
    workbook.save(source)
    workbook.close()

    split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))

    workbook = load_workbook(output, data_only=False)
    try:
        assert workbook.active["B2"].value == "=1+1"
        assert workbook.active["B2"].data_type == "s"
    finally:
        workbook.close()


def test_xls_formula_is_rejected_in_selected_sheet(tmp_path):
    source = tmp_path / "staff.xls"
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Staff")
    sheet.write(0, 0, "Department")
    sheet.write(0, 1, "Amount")
    sheet.write(1, 0, "Sales")
    sheet.write(1, 1, xlwt.Formula("1+1"))
    workbook.save(str(source))

    with pytest.raises(WorkbookProcessingError, match="(?i)formula"):
        split_workbooks([str(source)], "Staff", "Department", None, str(tmp_path / "out.xlsx"))


def test_only_selected_worksheet_is_checked_for_formulas(tmp_path):
    source = tmp_path / "staff.xlsx"
    output = tmp_path / "out.xlsx"
    workbook = Workbook()
    clean = workbook.active
    clean.title = "Clean"
    clean.append(["Department", "Amount"])
    clean.append(["Sales", 5])
    other = workbook.create_sheet("Other")
    other.append(["Department", "Amount"])
    other.append(["Support", "=1+1"])
    workbook.save(source)
    workbook.close()

    split_workbooks([str(source)], "Clean", "Department", None, str(output))

    result = load_workbook(output, data_only=False)
    try:
        assert list(result.active.values)[1] == ("Sales", 5)
    finally:
        result.close()


def test_formula_hidden_by_stale_worksheet_dimensions_is_rejected(tmp_path):
    source = tmp_path / "staff.xlsx"
    write_xlsx(source, 2, False)
    with ZipFile(source) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    xml = files["xl/worksheets/sheet1.xml"]
    assert b'<dimension ref="A1:B2"' in xml
    files["xl/worksheets/sheet1.xml"] = xml.replace(
        b'<dimension ref="A1:B2"', b'<dimension ref="A1:A1"'
    )
    with ZipFile(source, "w", ZIP_DEFLATED) as archive:
        for name, contents in files.items():
            archive.writestr(name, contents)

    with pytest.raises(WorkbookProcessingError, match="(?i)formula"):
        split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(tmp_path / "out.xlsx"))


def test_only_selected_legacy_worksheet_is_checked_for_formulas(tmp_path):
    source = tmp_path / "staff.xls"
    output = tmp_path / "out.xlsx"
    workbook = xlwt.Workbook()
    clean = workbook.add_sheet("Clean")
    clean.write(0, 0, "Department")
    clean.write(0, 1, "Amount")
    clean.write(1, 0, "Sales")
    clean.write(1, 1, 5)
    other = workbook.add_sheet("Other")
    other.write(0, 0, "Department")
    other.write(0, 1, "Amount")
    other.write(1, 0, "Support")
    other.write(1, 1, xlwt.Formula("1+1"))
    workbook.save(str(source))

    split_workbooks([str(source)], "Clean", "Department", None, str(output))

    result = load_workbook(output, data_only=False)
    try:
        assert list(result.active.values)[1] == ("Sales", 5)
    finally:
        result.close()
