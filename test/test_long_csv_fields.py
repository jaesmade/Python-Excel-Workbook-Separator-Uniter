import pytest
import xlrd
from openpyxl import load_workbook

from excel_splitter.core import ALL_SHEETS, WorkbookProcessingError, split_workbooks


@pytest.mark.parametrize("extension", ["xlsx", "xls"])
@pytest.mark.parametrize("field", ["header", "value"])
@pytest.mark.parametrize("length", [32768, 33000])
def test_oversized_csv_field_is_rejected_before_output_is_overwritten(tmp_path, extension, field, length):
    source = tmp_path / "staff.csv"
    output = tmp_path / f"out.{extension}"
    long_text = "x" * length
    header = long_text if field == "header" else "Note"
    value = long_text if field == "value" else "ordinary"
    source.write_text(f"Department,{header}\nSales,{value}\n", encoding="utf-8")
    output.write_bytes(b"existing output")

    with pytest.raises(WorkbookProcessingError, match=r"32,767") as error:
        split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))

    assert "Sales" in str(error.value)
    assert ("row 1" if field == "header" else "row 2") in str(error.value)
    assert "column 2" in str(error.value)
    assert output.read_bytes() == b"existing output"


@pytest.mark.parametrize("extension", ["xlsx", "xls"])
def test_csv_field_at_excel_limit_is_preserved(tmp_path, extension):
    source = tmp_path / "staff.csv"
    output = tmp_path / f"out.{extension}"
    value = "x" * 32767
    source.write_text(f"Department,Note\nSales,{value}\n", encoding="utf-8")

    split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))

    if extension == "xls":
        assert xlrd.open_workbook(str(output)).sheet_by_index(0).cell_value(1, 1) == value
    else:
        workbook = load_workbook(output)
        try:
            assert workbook.active["B2"].value == value
        finally:
            workbook.close()


def test_oversized_field_in_unselected_group_does_not_block_export(tmp_path):
    source = tmp_path / "staff.csv"
    output = tmp_path / "out.xlsx"
    source.write_text(f"Department,Note\nSales,valid\nSupport,{'x' * 33000}\n", encoding="utf-8")

    split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output), {("Sales",)})

    workbook = load_workbook(output)
    try:
        assert workbook.active["B2"].value == "valid"
    finally:
        workbook.close()
