"""Export styling based on the FULL DATAS attendance workbooks."""

import xlwt
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.styles.borders import DEFAULT_BORDER
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo


NAVY = "203E62"
COLUMN_WIDTHS = {
    "department": 33, "name": 34, "no.": 14, "date/time": 30,
    "location id": 16, "id number": 18, "verifycode": 17,
    "cardno": 14, "workcode": 15, "status": 13,
}
CENTERED_COLUMNS = {"no.", "location id", "status"}


def column_width(header, values):
    key = str(header).strip().casefold()
    if key in COLUMN_WIDTHS:
        return COLUMN_WIDTHS[key]
    longest = max((len(str(value)) for value in values if value is not None), default=0)
    return _fitted_column_width(header, longest)


def _fitted_column_width(header, longest):
    return min(48, max(14, len(str(header)) + 3, longest + 2))


def append_xlsx_header(sheet, headers, value_lengths, row_count, table_index):
    """Set worksheet metadata before streaming the first row."""
    last_column = len(headers)
    table_ref = f"A1:{get_column_letter(last_column)}{row_count + 1}"
    # Give the write-only writer the dimensions already known from the table.
    sheet.calculate_dimension = lambda: table_ref
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = f"{get_column_letter(min(3, last_column))}2"
    sheet.row_dimensions[1].height = 31
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor=NAVY)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    white_line = Side(style="thin", color="FFFFFF")
    header_cells = []
    for index, (header, longest) in enumerate(zip(headers, value_lengths), start=1):
        key = header.strip().casefold()
        width = COLUMN_WIDTHS[key] if key in COLUMN_WIDTHS else _fitted_column_width(header, longest)
        sheet.column_dimensions[get_column_letter(index)].width = width
        cell = WriteOnlyCell(sheet, value=header)
        cell.data_type = "s"
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = Border(
            left=white_line if index > 1 else Side(),
            right=white_line if index < last_column else Side(),
            bottom=Side(style="medium", color=NAVY),
        )
        header_cells.append(cell)
    table = Table(
        displayName=f"Attendance_{table_index}", ref=table_ref, autoFilter=AutoFilter(ref=table_ref),
        tableColumns=[TableColumn(id=index, name=header) for index, header in enumerate(headers, start=1)],
        tableStyleInfo=TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        ),
    )
    # The headers are explicit and Attendance_N names are unique in the workbook.
    sheet.tables.add(table)
    sheet.append(header_cells)


def xlsx_body_styles(workbook, headers):
    """Build reusable style arrays that Cell copies when a row is written."""
    font = Font(name="Arial", size=11, color="000000")
    styles = {}
    for centered in (False, True):
        style = NamedStyle(font=font, border=DEFAULT_BORDER, alignment=Alignment(
            horizontal="center" if centered else "left", vertical="center", wrap_text=False,
        ))
        style.bind(workbook)
        styles[centered] = style.as_tuple()
    return [styles[header.strip().casefold() in CENTERED_COLUMNS] for header in headers]


def legacy_cell_style(styles, striped, centered, number_format):
    """Reuse an output style for each combination of appearance and format."""
    key = (striped, centered, number_format)
    if key not in styles:
        styles[key] = xlwt.easyxf(
            "font: name Arial, height 220;"
            f"alignment: horiz {'center' if centered else 'left'}, vert centre;"
            + ("pattern: pattern solid, fore_colour 34;" if striped else ""),
            num_format_str=number_format,
        )
    return styles[key]


def legacy_styles(workbook):
    """Use matching cell fills for BIFF8, which has no Excel table objects."""
    workbook.set_colour_RGB(0x21, 32, 62, 98)
    workbook.set_colour_RGB(0x22, 220, 230, 241)
    styles = {
        "header": xlwt.easyxf(
            "font: name Arial, height 220, bold on, colour white;"
            "pattern: pattern solid, fore_colour 33;"
            "alignment: horiz center, vert centre;"
            "borders: bottom medium, bottom_colour 33;"
        )
    }
    return styles


def format_legacy_dimensions(sheet, headers, rows):
    sheet.show_grid = False
    sheet.panes_frozen = True
    sheet.horz_split_pos = 1
    sheet.vert_split_pos = min(2, len(headers) - 1)
    sheet.row(0).height = 31 * 20
    sheet.row(0).height_mismatch = True
    for index, header in enumerate(headers):
        sheet.col(index).width = int(column_width(header, (row[index] for row in rows)) * 256)
