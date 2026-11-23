"""Detect formulas before Excel readers discard their expressions."""

import os
import struct

import xlrd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from xlrd.biffh import XL_EOF, XL_FORMULA_OPCODES
from xlrd.book import bofcodes


def formula_cell(file_path, sheet_name):
    """Return the first formula's coordinate in a worksheet, or None.

    xlrd exposes only cached values, so legacy files need a BIFF record scan.
    Fail closed if the workbook stream cannot be inspected reliably.
    """
    if os.path.splitext(file_path)[1].lower() != ".xls":
        workbook = load_workbook(file_path, read_only=True, data_only=False)
        try:
            sheet = workbook[sheet_name]
            # Worksheet dimensions can be stale; pandas also ignores them.
            sheet.reset_dimensions()
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        return cell.coordinate
        finally:
            workbook.close()
        return None

    workbook = xlrd.open_workbook(file_path, on_demand=True)
    try:
        if workbook.biff_version < 50 or workbook.mem is None:
            raise ValueError("This older .xls format cannot be checked for formulas. Save it as .xlsx first.")
        position = workbook._sh_abs_posn[workbook.sheet_names().index(sheet_name)]
        depth = 0
        while position + 4 <= len(workbook.mem):
            opcode, length = struct.unpack_from("<HH", workbook.mem, position)
            position += 4
            if position + length > len(workbook.mem):
                break
            if opcode in bofcodes:
                depth += 1
            elif opcode == XL_EOF:
                depth -= 1
                if depth == 0:
                    return None
            elif depth == 1 and opcode in XL_FORMULA_OPCODES:
                if length < 4:
                    break
                row, column = struct.unpack_from("<HH", workbook.mem, position)
                return f"{get_column_letter(column + 1)}{row + 1}"
            position += length
        raise ValueError("This .xls worksheet could not be checked for formulas. Save it as .xlsx first.")
    finally:
        workbook.release_resources()
