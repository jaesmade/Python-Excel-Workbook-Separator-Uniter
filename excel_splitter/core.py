from __future__ import annotations

import os
from collections import OrderedDict
from typing import Callable

import pandas as pd
import xlwt
from openpyxl import Workbook

from .utils import detect_header_names, normalize_group_value, sanitize_sheet_name


class WorkbookProcessingError(ValueError):
    pass


def _ensure_python_value(value):
    if pd.isna(value):
        return None
    return value


def _iter_sheet_rows(sheet):
    if isinstance(sheet, pd.DataFrame):
        for row in sheet.itertuples(index=False, name=None):
            yield row
        return

    for row in sheet.iter_rows(values_only=True):
        yield row


def _first_row_headers(sheet) -> list[str]:
    if isinstance(sheet, pd.DataFrame):
        return detect_header_names(list(sheet.columns))

    rows = list(_iter_sheet_rows(sheet))
    if not rows:
        return []
    return detect_header_names(rows[0])


def read_workbook_sheet(file_path: str, sheet_name: str):
    data = pd.read_excel(file_path, sheet_name=sheet_name)
    headers = _first_row_headers(data)
    if not headers or all((header == "" for header in headers)):
        raise WorkbookProcessingError("The selected worksheet does not contain any data.")
    return data, data, headers


def build_group_map(sheet, department_column: str, id_number_column: str):
    headers = _first_row_headers(sheet)
    if not headers:
        raise WorkbookProcessingError("The selected worksheet does not contain any data.")

    if isinstance(sheet, pd.DataFrame):
        columns = [str(col).strip() for col in list(sheet.columns)]
        if department_column not in columns:
            raise WorkbookProcessingError(f"The Department column '{department_column}' was not found in the worksheet headers.")
        if id_number_column not in columns:
            raise WorkbookProcessingError(f"The ID Number column '{id_number_column}' was not found in the worksheet headers.")
        department_index = columns.index(department_column)
        id_index = columns.index(id_number_column)
    else:
        if department_column not in headers:
            raise WorkbookProcessingError(f"The Department column '{department_column}' was not found in the worksheet headers.")
        if id_number_column not in headers:
            raise WorkbookProcessingError(f"The ID Number column '{id_number_column}' was not found in the worksheet headers.")
        department_index = headers.index(department_column)
        id_index = headers.index(id_number_column)

    groups = OrderedDict()

    for row in list(_iter_sheet_rows(sheet))[1:] if not isinstance(sheet, pd.DataFrame) else list(_iter_sheet_rows(sheet)):
        if not row or all((cell is None or pd.isna(cell) or str(cell).strip() == "") for cell in row):
            continue

        department_value = row[department_index] if department_index < len(row) else None
        status_value = row[id_index] if id_index < len(row) else None

        department_name = normalize_group_value(_ensure_python_value(department_value), "UNKNOWN DEPARTMENT")
        status_name = normalize_group_value(_ensure_python_value(status_value), "UNKNOWN STATUS")
        group_key = (department_name, status_name)

        if group_key not in groups:
            groups[group_key] = []
        groups[group_key].append(row)

    if not groups:
        raise WorkbookProcessingError("No data rows were found in the selected worksheet.")

    return headers, groups


def split_workbook(file_path: str, sheet_name: str, department_column: str, id_number_column: str,
                   output_file: str, selected_groups: set[tuple[str, str]] | None = None,
                   progress_callback: Callable[[int, str], None] | None = None,
                   combine_duplicates: bool = True):
    return split_workbooks([file_path], sheet_name, department_column, id_number_column, output_file,
                           selected_groups, progress_callback, combine_duplicates)


def split_workbooks(file_paths: list[str], sheet_name: str, department_column: str, id_number_column: str,
                    output_file: str, selected_groups: set[tuple[str, str]] | None = None,
                    progress_callback: Callable[[int, str], None] | None = None,
                    combine_duplicates: bool = True):
    if not file_paths:
        raise WorkbookProcessingError("No files were selected for processing.")

    combined_groups: OrderedDict[tuple[str, str], list[tuple]] = OrderedDict()
    combined_headers: list[str] | None = None
    total_files = len(file_paths)

    for file_index, file_path in enumerate(file_paths, start=1):
        source_df = pd.read_excel(file_path, sheet_name=sheet_name)
        headers, groups = build_group_map(source_df, department_column, id_number_column)
        if combined_headers is None:
            combined_headers = headers

        for group_key, rows in groups.items():
            if combine_duplicates:
                if group_key not in combined_groups:
                    combined_groups[group_key] = []
                combined_groups[group_key].extend(rows)
            else:
                unique_key = (group_key[0], group_key[1], os.path.basename(file_path))
                if unique_key not in combined_groups:
                    combined_groups[unique_key] = []
                combined_groups[unique_key].extend(rows)

        if progress_callback:
            progress_callback(int((file_index / total_files) * 40), f"Reading: {os.path.basename(file_path)}")

    if not combined_groups:
        raise WorkbookProcessingError("No valid rows were found across the selected files.")

    if selected_groups is None:
        selected_groups = set(combined_groups.keys())

    if not selected_groups:
        raise WorkbookProcessingError("No groups were selected to generate worksheets.")

    output_ext = os.path.splitext(output_file)[1].lower()
    used_names: set[str] = set()
    processed_rows = 0
    total_groups = len(combined_groups)

    if output_ext == ".xls":
        workbook = xlwt.Workbook()
        for index, (group_key, rows) in enumerate(combined_groups.items(), start=1):
            if group_key not in selected_groups:
                continue

            display_name = f"{group_key[0]} - {group_key[1]}"
            sheet_title = sanitize_sheet_name(display_name, used_names)
            sheet = workbook.add_sheet(sheet_title[:31])
            for col_num, header in enumerate(combined_headers or []):
                sheet.write(0, col_num, header)
            for row_num, row_values in enumerate(rows, start=1):
                for col_num, value in enumerate(row_values):
                    if pd.isna(value):
                        sheet.write(row_num, col_num, "")
                    else:
                        sheet.write(row_num, col_num, value)
                processed_rows += 1

            if progress_callback:
                progress_callback(40 + int((index / total_groups) * 60), f"Creating: {display_name}")

        workbook.save(output_file)
        return {
            "created_sheets": len(selected_groups),
            "processed_rows": processed_rows,
            "output_file": output_file,
        }

    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    for index, (group_key, rows) in enumerate(combined_groups.items(), start=1):
        if group_key not in selected_groups:
            continue

        display_name = f"{group_key[0]} - {group_key[1]}"
        sheet_title = sanitize_sheet_name(display_name, used_names)
        target_sheet = output_wb.create_sheet(title=sheet_title)

        for col_num, header in enumerate(combined_headers or [], start=1):
            target_sheet.cell(row=1, column=col_num, value=header)

        for row_num, row_values in enumerate(rows, start=2):
            for col_num, value in enumerate(row_values, start=1):
                if pd.isna(value):
                    target_sheet.cell(row=row_num, column=col_num, value=None)
                else:
                    target_sheet.cell(row=row_num, column=col_num, value=value)
            processed_rows += 1

        if progress_callback:
            progress_callback(40 + int((index / total_groups) * 60), f"Creating: {display_name}")

    output_wb.save(output_file)
    return {
        "created_sheets": len(output_wb.worksheets),
        "processed_rows": processed_rows,
        "output_file": output_file,
    }
