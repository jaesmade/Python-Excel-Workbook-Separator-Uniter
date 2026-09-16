from __future__ import annotations

from collections import OrderedDict
from datetime import datetime
from typing import Callable, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from .utils import detect_header_names, normalize_group_value, sanitize_sheet_name


class WorkbookProcessingError(ValueError):
    pass


def _iter_sheet_rows(sheet):
    for row in sheet.iter_rows(values_only=True):
        yield row


def _first_row_headers(sheet) -> list[str]:
    rows = list(_iter_sheet_rows(sheet))
    if not rows:
        return []
    return detect_header_names(rows[0])


def read_workbook_sheet(file_path: str, sheet_name: str):
    workbook = load_workbook(file_path, read_only=False, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise WorkbookProcessingError(f"The worksheet '{sheet_name}' was not found in the workbook.")
        sheet = workbook[sheet_name]
        headers = _first_row_headers(sheet)
        if not headers or all((header == "" for header in headers)):
            raise WorkbookProcessingError("The selected worksheet does not contain any data.")
        return workbook, sheet, headers
    except Exception:
        workbook.close()
        raise


def build_group_map(sheet, department_column: str, id_number_column: str):
    headers = _first_row_headers(sheet)
    if not headers:
        raise WorkbookProcessingError("The selected worksheet does not contain any data.")

    if department_column not in headers:
        raise WorkbookProcessingError(f"The Department column '{department_column}' was not found in the worksheet headers.")
    if id_number_column not in headers:
        raise WorkbookProcessingError(f"The ID Number column '{id_number_column}' was not found in the worksheet headers.")

    department_index = headers.index(department_column)
    id_index = headers.index(id_number_column)
    groups = OrderedDict()

    for row in list(_iter_sheet_rows(sheet))[1:]:
        if not row or all((cell is None or str(cell).strip() == "") for cell in row):
            continue

        department_value = row[department_index] if department_index < len(row) else None
        status_value = row[id_index] if id_index < len(row) else None

        department_name = normalize_group_value(department_value, "UNKNOWN DEPARTMENT")
        status_name = normalize_group_value(status_value, "UNKNOWN STATUS")
        group_key = (department_name, status_name)

        if group_key not in groups:
            groups[group_key] = []
        groups[group_key].append(row)

    if not groups:
        raise WorkbookProcessingError("No data rows were found in the selected worksheet.")

    return headers, groups


def split_workbook(file_path: str, sheet_name: str, department_column: str, id_number_column: str,
                   output_file: str, selected_groups: set[tuple[str, str]] | None = None,
                   progress_callback: Callable[[int, str], None] | None = None):
    workbook = load_workbook(file_path, read_only=False, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise WorkbookProcessingError(f"The worksheet '{sheet_name}' was not found in the workbook.")

        source_sheet = workbook[sheet_name]
        headers, groups = build_group_map(source_sheet, department_column, id_number_column)
        if not selected_groups:
            selected_groups = set(groups.keys())

        if not selected_groups:
            raise WorkbookProcessingError("No groups were selected to generate worksheets.")

        output_wb = Workbook()
        default_sheet = output_wb.active
        output_wb.remove(default_sheet)

        used_names: set[str] = set()
        processed_rows = 0
        total_groups = len(groups)

        for index, (group_key, rows) in enumerate(groups.items(), start=1):
            if group_key not in selected_groups:
                continue

            display_name = f"{group_key[0]} - {group_key[1]}"
            sheet_title = sanitize_sheet_name(display_name, used_names)
            target_sheet = output_wb.create_sheet(title=sheet_title)

            for col_num, header in enumerate(headers, start=1):
                target_sheet.cell(row=1, column=col_num, value=header)

            for row_num, row_values in enumerate(rows, start=2):
                for col_num, value in enumerate(row_values, start=1):
                    target_sheet.cell(row=row_num, column=col_num, value=value)
                processed_rows += 1

            if progress_callback:
                progress_callback(int((index / total_groups) * 100), f"Creating: {display_name}")

        output_wb.save(output_file)
        return {
            "created_sheets": len(output_wb.worksheets),
            "processed_rows": processed_rows,
            "output_file": output_file,
        }
    finally:
        workbook.close()
