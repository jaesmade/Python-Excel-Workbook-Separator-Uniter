"""Export styling based on the FULL DATAS attendance workbooks."""

from datetime import date, datetime

import xlwt
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


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
    return min(48, max(14, len(str(header)) + 3, longest + 2))


def format_worksheet(sheet, table_index: int):
    """Style a populated worksheet without changing its headers or values."""
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = f"{get_column_letter(min(3, sheet.max_column))}2"
    sheet.row_dimensions[1].height = 31
    body_font = Font(name="Arial", size=11, color="000000")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    white_line = Side(style="thin", color="FFFFFF")
    for column in sheet.iter_cols():
        header = column[0]
        key = str(header.value).strip().casefold()
        sheet.column_dimensions[header.column_letter].width = column_width(
            header.value, (cell.value for cell in column[1:])
        )
        header.font = header_font
        header.fill = PatternFill("solid", fgColor=NAVY)
        header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        header.border = Border(
            left=white_line if header.column > 1 else Side(),
            right=white_line if header.column < sheet.max_column else Side(),
            bottom=Side(style="medium", color=NAVY),
        )
        alignment = Alignment(
            horizontal="center" if key in CENTERED_COLUMNS else "left",
            vertical="center", wrap_text=False,
        )
        for cell in column[1:]:
            cell.font = body_font
            cell.alignment = alignment
            if isinstance(cell.value, datetime):
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            elif isinstance(cell.value, date):
                cell.number_format = "yyyy-mm-dd"
    for row_index in range(2, sheet.max_row + 1):
        sheet.row_dimensions[row_index].height = 23
    table = Table(displayName=f"Attendance_{table_index}", ref=sheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    sheet.add_table(table)


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
    for striped in (False, True):
        for centered in (False, True):
            for is_date in (False, True):
                styles[striped, centered, is_date] = xlwt.easyxf(
                    "font: name Arial, height 220;"
                    f"alignment: horiz {'center' if centered else 'left'}, vert centre;"
                    + ("pattern: pattern solid, fore_colour 34;" if striped else ""),
                    num_format_str="YYYY-MM-DD HH:MM:SS" if is_date else "General",
                )
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
