"""Number formats must travel with source cells through grouping and export."""

from datetime import datetime

import pandas as pd
import pytest
import xlrd
import xlwt
from openpyxl import Workbook, load_workbook

from excel_splitter.core import (
    ALL_SHEETS,
    _iter_input_tables,
    _normalize_input_tables,
    build_group_map,
    split_input_tables,
    split_workbooks,
)


def _write_source(path, extension):
    # The second sheet reorders the columns and uses different formats for the
    # same output columns. A blank row tests source-to-output row alignment.
    sheets = [
        ("First", [
            [("Department", None), ("ID", None), ("Rate", None), ("Date", None)],
            [("Sales", None), (123, "00000"), (0.125, "0.0%"),
             (datetime(2024, 5, 1), "dd/mm/yyyy")],
            [(None, None)] * 4,
            [("Sales", None), (789, "0000"), (0.5, "0%"),
             (datetime(2024, 5, 3), "yyyy/mm/dd")],
        ]),
        ("Second", [
            [("Rate", None), ("Date", None), ("Department", None), ("ID", None)],
            [(0.25, "0.00%"), (datetime(2024, 6, 2), "mmm d, yyyy"),
             ("Sales", None), (456, "000000")],
        ]),
    ]
    if extension == "xlsx":
        book = Workbook()
        book.remove(book.active)
        for name, rows in sheets:
            sheet = book.create_sheet(name)
            for row_number, row in enumerate(rows, 1):
                for column_number, (value, number_format) in enumerate(row, 1):
                    cell = sheet.cell(row_number, column_number, value)
                    if number_format:
                        cell.number_format = number_format
        book.save(path)
        book.close()
    else:
        book = xlwt.Workbook()
        for name, rows in sheets:
            sheet = book.add_sheet(name)
            for row_number, row in enumerate(rows):
                for column_number, (value, number_format) in enumerate(row):
                    if value is not None:
                        style = xlwt.easyxf(num_format_str=number_format or "General")
                        sheet.write(row_number, column_number, value, style)
        book.save(str(path))


def _output_cells(path, extension):
    if extension == "xlsx":
        book = load_workbook(path)
        try:
            sheet = book["Sales"]
            return [[(cell.value, cell.number_format) for cell in row]
                    for row in sheet.iter_rows(min_row=2, max_row=4, min_col=2, max_col=4)]
        finally:
            book.close()
    with xlrd.open_workbook(path, formatting_info=True) as book:
        sheet = book.sheet_by_name("Sales")
        result = []
        for row_number in (1, 2, 3):
            cells = []
            for column_number in (1, 2, 3):
                cell = sheet.cell(row_number, column_number)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(value, book.datemode)
                fmt = book.format_map[book.xf_list[cell.xf_index].format_key].format_str
                cells.append((value, fmt))
            result.append(cells)
        return result


@pytest.mark.parametrize("extension", ["xlsx", "xls"])
def test_column_access_shares_immutable_number_formats(tmp_path, extension):
    source = tmp_path / f"source.{extension}"
    _write_source(source, extension)
    tables = list(_iter_input_tables(str(source), ALL_SHEETS))
    aligned_tables, _ = _normalize_input_tables(tables)

    for _, table in [*tables, *aligned_tables]:
        formats = table.attrs["number_formats"]
        column = table.iloc[:, 0]
        assert column.attrs["number_formats"] is formats
        with pytest.raises(TypeError):
            formats[0][0] = "0%"
        with pytest.raises(TypeError):
            formats[0] = ()


def test_grouping_skips_deepcopy_of_legacy_list_formats():
    class TrackedFormat(str):
        def __deepcopy__(self, memo):
            pytest.fail("Row iteration copied a legacy number format")

    table = pd.DataFrame([("Sales", 123)], columns=["Department", "ID"])
    formats = [["General", TrackedFormat("00000")]]
    table.attrs["number_formats"] = formats

    _, groups = build_group_map(table, "Department")

    row, = groups[("Sales",)]
    assert row == ("Sales", 123)
    assert row.number_formats == ["General", "00000"]
    assert table.attrs["number_formats"] is formats


@pytest.mark.parametrize("input_extension", ["xlsx", "xls"])
@pytest.mark.parametrize("output_extension", ["xlsx", "xls"])
@pytest.mark.parametrize("snapshot", [False, True], ids=["direct", "cached-tables"])
def test_source_cell_number_formats_survive_grouping_and_export(
    tmp_path, input_extension, output_extension, snapshot
):
    source = tmp_path / f"source.{input_extension}"
    output = tmp_path / f"output.{output_extension}"
    _write_source(source, input_extension)

    if snapshot:
        tables = [(name, table) for name, table in _iter_input_tables(str(source), ALL_SHEETS)]
        source.unlink()  # A cached preview must not depend on reopening the source.
        split_input_tables(tables, "Department", None, str(output))
    else:
        split_workbooks([str(source)], ALL_SHEETS, "Department", None, str(output))

    assert _output_cells(output, output_extension) == [
        [(123, "00000"), (0.125, "0.0%"), (datetime(2024, 5, 1), "dd/mm/yyyy")],
        [(789, "0000"), (0.5, "0%"), (datetime(2024, 5, 3), "yyyy/mm/dd")],
        [(456, "000000"), (0.25, "0.00%"), (datetime(2024, 6, 2), "mmm d, yyyy")],
    ]
