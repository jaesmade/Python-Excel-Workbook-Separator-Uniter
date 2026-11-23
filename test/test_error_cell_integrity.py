import pytest
import xlrd
import xlwt
from openpyxl import Workbook, load_workbook

from excel_splitter.core import ALL_SHEETS, read_workbook_sheet, split_workbooks


@pytest.mark.parametrize("source_format", ["xlsx", "xls"])
@pytest.mark.parametrize("output_format", ["xlsx", "xls"])
@pytest.mark.parametrize("error_code", ["#DIV/0!", "#VALUE!", "#N/A"])
def test_excel_errors_and_error_like_text_survive_preview_and_export(tmp_path, source_format, output_format, error_code):
    source = tmp_path / f"source.{source_format}"
    output = tmp_path / f"output.{output_format}"
    if source_format == "xlsx":
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Staff"
        sheet.append(["Department", "Error", "Text"])
        sheet.append(["Sales", error_code, error_code])
        sheet["C2"].data_type = "s"
        workbook.save(source)
        workbook.close()
    else:
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Staff")
        for column, value in enumerate(["Department", "Error", "Text"]):
            sheet.write(0, column, value)
        sheet.write(1, 0, "Sales")
        numeric_code = next(code for code, text in xlrd.error_text_from_code.items() if text == error_code)
        sheet.row(1).set_cell_error(1, numeric_code)
        sheet.write(1, 2, error_code)
        workbook.save(str(source))

    table, _, _ = read_workbook_sheet(str(source), "Staff")
    assert table.iloc[0, 1] == error_code
    split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))

    if output_format == "xlsx":
        result = load_workbook(output)
        try:
            error, text = result.active["B2"], result.active["C2"]
            assert (error.value, error.data_type) == (error_code, "e")
            assert (text.value, text.data_type) == (error_code, "s")
        finally:
            result.close()
    else:
        result = xlrd.open_workbook(str(output))
        error, text = result.sheet_by_index(0).cell(1, 1), result.sheet_by_index(0).cell(1, 2)
        assert (error.ctype, xlrd.error_text_from_code[error.value]) == (xlrd.XL_CELL_ERROR, error_code)
        assert (text.ctype, text.value) == (xlrd.XL_CELL_TEXT, error_code)


@pytest.mark.parametrize("output_format", ["xlsx", "xls"])
def test_new_excel_errors_remain_visible_when_exported(tmp_path, output_format):
    source = tmp_path / "source.xlsx"
    output = tmp_path / f"output.{output_format}"
    workbook = Workbook()
    workbook.active.append(["Department", "Error"])
    workbook.active.append(["Sales", "#SPILL!"])
    workbook.active["B2"].data_type = "e"
    workbook.save(source)
    workbook.close()

    split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))

    if output_format == "xlsx":
        result = load_workbook(output)
        try:
            assert (result.active["B2"].value, result.active["B2"].data_type) == ("#SPILL!", "e")
        finally:
            result.close()
    else:
        cell = xlrd.open_workbook(str(output)).sheet_by_index(0).cell(1, 1)
        assert (cell.ctype, cell.value) == (xlrd.XL_CELL_TEXT, "#SPILL!")
