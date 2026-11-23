from __future__ import annotations

import os
from collections import Counter
from datetime import date
from collections import OrderedDict
from difflib import SequenceMatcher
from enum import Enum, auto
from numbers import Real
from typing import Callable
from itertools import chain

import pandas as pd
import xlrd
import xlwt
from openpyxl import Workbook

from .formula_check import formula_cell
from .utils import SheetNameRegistry, detect_header_names, normalize_group_value, sanitize_sheet_name
from .formatting import CENTERED_COLUMNS, format_legacy_dimensions, format_worksheet, legacy_cell_style, legacy_styles


class WorkbookProcessingError(ValueError):
    pass


class _ExcelError(str):
    """Keep Excel errors distinct from ordinary text with the same spelling."""


_LEGACY_ERROR_CODES = {text: code for code, text in xlrd.error_text_from_code.items()}


class _FormattedRow(tuple):
    """A row of ordinary values with their source number formats attached."""

    def __new__(cls, values, number_formats):
        row = super().__new__(cls, values)
        row.number_formats = number_formats
        return row


def _parse_excel_sheet(workbook, sheet_name):
    raw = workbook.parse(sheet_name=sheet_name, header=None, dtype=object, keep_default_na=False)
    # pandas converts error cells to NaN before its NA settings are applied.
    # Recover their values and types from the native reader's worksheet.
    if workbook.engine == "xlrd":
        sheet = workbook.book.sheet_by_name(sheet_name)
        errors = ((row, col, xlrd.error_text_from_code[sheet.cell_value(row, col)])
                  for row in range(sheet.nrows) for col in range(sheet.ncols)
                  if sheet.cell_type(row, col) == xlrd.XL_CELL_ERROR)
    else:
        sheet = workbook.book[sheet_name]
        if workbook.book.read_only:
            sheet.reset_dimensions()
        errors = ((cell.row - 1, cell.column - 1, cell.value)
                  for row in sheet.iter_rows() for cell in row
                  if cell.data_type == "e")
    for row, col, value in errors:
        raw.iat[row, col] = _ExcelError(value)
    if workbook.engine == "xlrd":
        book = workbook.book
        formats = [
            [book.format_map[book.xf_list[sheet.cell_xf_index(row, col)].format_key].format_str
             if row < sheet.nrows and col < sheet.ncols else None
             for col in range(raw.shape[1])]
            for row in range(raw.shape[0])
        ]
    else:
        formats = [[getattr(cell, "number_format", None) for cell in row]
                   for row in sheet.iter_rows(min_row=1, max_row=raw.shape[0],
                                              min_col=1, max_col=raw.shape[1])]
    raw.attrs["number_formats"] = formats
    return raw


def ensure_output_is_not_input(output_path, input_paths):
    """Reject output paths that refer to a loaded source, including aliases."""
    output_real = os.path.normcase(os.path.realpath(output_path))
    for input_path in input_paths:
        same_path = output_real == os.path.normcase(os.path.realpath(input_path))
        if not same_path and os.path.exists(output_path) and os.path.exists(input_path):
            same_path = os.path.samefile(output_path, input_path)
        if same_path:
            raise WorkbookProcessingError("Choose a different output file to avoid overwriting an input file.")


class _SheetSelection(Enum):
    ALL = auto()


ALL_SHEETS = _SheetSelection.ALL
SORT_ALPHABETICAL = "alphabetical"
SORT_COUNT_ASCENDING = "count_ascending"
SORT_COUNT_DESCENDING = "count_descending"


def _similar_group_values(first: str, second: str) -> bool:
    """Find small spelling differences without treating IDs as words."""
    first = " ".join(first.casefold().split())
    second = " ".join(second.casefold().split())
    if first == second:
        return True
    if min(len(first), len(second)) < 4 or abs(len(first) - len(second)) > 2:
        return False
    if SequenceMatcher(None, first, second).ratio() < 0.84:
        return False

    # Bounded Levenshtein distance: reject longer differences even when a
    # long shared prefix would otherwise make the ratio look convincing.
    previous = list(range(len(second) + 1))
    for index, char in enumerate(first, 1):
        current = [index]
        for other_index, other in enumerate(second, 1):
            current.append(min(current[-1] + 1, previous[other_index] + 1,
                               previous[other_index - 1] + (char != other)))
        previous = current
    return previous[-1] <= 2


def suggest_group_value_merges(groups, grouping_columns):
    """Suggest existing similar values; the caller must approve every merge.

    Counts represent records across all groups for each selected column. The
    most common spelling is proposed as the target, with stable tie-breaking.
    """
    suggestions = []
    for index, column in enumerate(grouping_columns):
        counts = Counter()
        for key, rows in groups.items():
            value = key[index]
            if (value.casefold() != "unknown group" and any(c.isalpha() for c in value)
                    and not any(c.isdigit() for c in value)):
                counts[value] += len(rows)
        ranked = sorted(counts, key=lambda value: (-counts[value], value.casefold(), value))
        for position, source in enumerate(ranked):
            candidates = [target for target in ranked[:position] if _similar_group_values(source, target)]
            if candidates:
                target = candidates[0]
                suggestions.append((column, source, target, counts[source], counts[target]))
    return suggestions


def merge_group_values(groups, grouping_columns, mappings):
    """Apply approved value replacements to keys, leaving source rows intact."""
    merged = OrderedDict()
    for key, rows in groups.items():
        parts = list(key)
        for index, column in enumerate(grouping_columns):
            value = parts[index]
            seen = set()
            while (column, value) in mappings:
                if value in seen:
                    raise WorkbookProcessingError("Value merge choices contain a cycle.")
                seen.add(value)
                value = mappings[(column, value)]
            parts[index] = value
        merged.setdefault(tuple(parts), []).extend(rows)
    return merged


def sort_group_items(groups, sort_mode: str = SORT_ALPHABETICAL):
    items = list(groups.items())
    if sort_mode == SORT_COUNT_ASCENDING:
        return sorted(items, key=lambda item: len(item[1]))
    if sort_mode == SORT_COUNT_DESCENDING:
        return sorted(items, key=lambda item: len(item[1]), reverse=True)
    return sorted(items, key=lambda item: " - ".join(item[0]).casefold())


def _is_numeric_group_value(value):
    return isinstance(value, Real) and not isinstance(value, bool)


def _numeric_merge_targets(groups, column_indexes, merge_targets):
    """Find numeric values behind preferred labels, even in unselected groups."""
    numeric_targets = {}
    for key, rows in groups.items():
        for target, (column, cell_index) in zip(key, column_indexes):
            target_key = (column, target)
            if target_key not in merge_targets or target_key in numeric_targets:
                continue
            for row in rows:
                value = _ensure_python_value(row[cell_index])
                if _is_numeric_group_value(value) and normalize_group_value(value, "UNKNOWN GROUP") == target:
                    numeric_targets[target_key] = value
                    break
    return numeric_targets


def preview_group_items(groups, selected_groups, sort_mode: str = SORT_ALPHABETICAL,
                        *, headers=None, grouping_columns=None, value_mappings=None,
                        correct_merged_values: bool = False):
    """Return selected worksheets and the rows that will be exported."""
    used_names = SheetNameRegistry()
    items = [
        (sanitize_sheet_name(" - ".join(key), used_names), key, rows)
        for key, rows in sort_group_items(groups, sort_mode)
        if key in selected_groups
    ]
    if not correct_merged_values or not value_mappings:
        return items

    columns = [str(column).strip() for column in grouping_columns if str(column).strip()]
    column_indexes = [(column, headers.index(column)) for column in columns]
    merge_targets = {(column, target) for (column, _), target in value_mappings.items()}
    numeric_targets = _numeric_merge_targets(groups, column_indexes, merge_targets)
    corrected_groups = []
    for title, key, rows in items:
        corrected_rows = []
        for row in rows:
            values = list(row)
            for target, (column, cell_index) in zip(key, column_indexes):
                original = _ensure_python_value(values[cell_index])
                value = normalize_group_value(original, "UNKNOWN GROUP")
                if (column, value) not in value_mappings and (column, value) not in merge_targets:
                    continue
                if _is_numeric_group_value(original):
                    if value == target:
                        continue
                    target = numeric_targets.get((column, target), target)
                values[cell_index] = target
            formats = getattr(row, "number_formats", None)
            corrected_rows.append(_FormattedRow(values, formats) if formats is not None else tuple(values))
        corrected_groups.append((title, key, corrected_rows))
    return corrected_groups


def _select_excel_engine(file_path: str) -> str | None:
    extension = os.path.splitext(file_path)[1].lower()
    if extension == ".xls":
        return "xlrd"
    if extension in {".xlsx", ".xlsm"}:
        return "openpyxl"
    return None


def _ensure_no_formulas(file_path: str, sheet_name: str):
    try:
        coordinate = formula_cell(file_path, sheet_name)
    except ValueError as exc:
        raise WorkbookProcessingError(str(exc)) from exc
    if coordinate:
        raise WorkbookProcessingError(
            f"Formula found in '{os.path.basename(file_path)}' - '{sheet_name}' at {coordinate}. "
            "Formulas cannot be safely split because references may change. "
            "Replace formulas with values in a copy of the workbook and try again."
        )


def _read_input_file(file_path: str, sheet_name: str | None = None) -> pd.DataFrame:
    extension = os.path.splitext(file_path)[1].lower()
    if extension == ".csv":
        raw = pd.read_csv(file_path, header=None, dtype=str, keep_default_na=False)
        return _table_from_raw(raw, os.path.basename(file_path))
    engine = _select_excel_engine(file_path)
    with pd.ExcelFile(file_path, engine=engine,
                      engine_kwargs={"formatting_info": True} if engine == "xlrd" else {}) as workbook:
        name = workbook.sheet_names[0] if sheet_name is None else sheet_name
        _ensure_no_formulas(file_path, name)
        raw = _parse_excel_sheet(workbook, name)
        return _table_from_raw(raw, f"{os.path.basename(file_path)} - {name}")


def _table_from_raw(raw: pd.DataFrame, source_label: str) -> pd.DataFrame:
    """Find the first populated row and validate its headers before naming columns."""
    for header_index, row in enumerate(raw.itertuples(index=False, name=None)):
        if _row_has_values(row):
            break
    else:
        return pd.DataFrame()

    original = detect_header_names(row)
    present = [name for name in original if name]
    if len(present) != len(set(present)):
        raise WorkbookProcessingError(f"Duplicate column headers found in '{source_label}'.")
    headers = [name if name else f"Unnamed: {index}" for index, name in enumerate(original)]
    if len(headers) != len(set(headers)):
        raise WorkbookProcessingError(f"Duplicate column headers found in '{source_label}'.")
    data_start = header_index + 1
    table = raw.iloc[data_start:].copy()
    table.columns = headers
    table = table.reset_index(drop=True)
    if "number_formats" in raw.attrs:
        table.attrs["number_formats"] = raw.attrs["number_formats"][data_start:]
    return table


def _iter_input_tables(file_path: str, sheet_name: str | _SheetSelection | None):
    if os.path.splitext(file_path)[1].lower() == ".csv":
        yield "CSV file", _read_input_file(file_path)
        return

    engine = _select_excel_engine(file_path)
    with pd.ExcelFile(file_path, engine=engine,
                      engine_kwargs={"formatting_info": True} if engine == "xlrd" else {}) as workbook:
        selected_sheets = workbook.sheet_names if sheet_name is ALL_SHEETS else [
            workbook.sheet_names[0] if sheet_name is None else sheet_name
        ]
        for source_sheet_name in selected_sheets:
            if source_sheet_name not in workbook.sheet_names:
                raise WorkbookProcessingError(
                    f"Worksheet '{source_sheet_name}' was not found in '{os.path.basename(file_path)}'."
                )
            _ensure_no_formulas(file_path, source_sheet_name)
            raw = _parse_excel_sheet(workbook, source_sheet_name)
            yield source_sheet_name, _table_from_raw(raw, f"{os.path.basename(file_path)} - {source_sheet_name}")


def source_file_labels(file_paths: list[str]) -> list[str]:
    """Use concise file labels unless a basename needs disambiguation."""
    basenames = [os.path.basename(path) for path in file_paths]
    counts = Counter(name.casefold() for name in basenames)
    labels = []
    used = set()
    for path, basename in zip(file_paths, basenames):
        base = os.path.abspath(path) if counts[basename.casefold()] > 1 else basename
        label = base
        suffix = 2
        while label.casefold() in used:
            label = f"{base} ({suffix})"
            suffix += 1
        used.add(label.casefold())
        labels.append(label)
    return labels


def _row_has_values(row) -> bool:
    return any(
        cell is not None and not pd.isna(cell) and str(cell).strip() != ""
        for cell in row
    )


def _normalize_input_tables(input_tables):
    tables = []
    headers: list[str] = []
    seen_headers: set[str] = set()

    for source_name, table in input_tables:
        if not any(_row_has_values(row) for row in table.itertuples(index=False, name=None)):
            continue
        normalized_table = table.copy()
        normalized_headers = detect_header_names(normalized_table.columns)
        if len(normalized_headers) != len(set(normalized_headers)):
            raise WorkbookProcessingError(f"Duplicate column headers found in '{source_name}'.")

        normalized_table.columns = normalized_headers
        for header in normalized_headers:
            if header not in seen_headers:
                headers.append(header)
                seen_headers.add(header)
        tables.append((source_name, normalized_table))

    if not tables:
        return [], headers

    aligned_tables = []
    for source_name, table in tables:
        aligned = table.reindex(columns=headers)
        if "number_formats" in table.attrs:
            positions = {header: index for index, header in enumerate(table.columns)}
            aligned.attrs["number_formats"] = [
                [formats[positions[header]] if header in positions else None for header in headers]
                for formats in table.attrs["number_formats"]
            ]
        aligned_tables.append((source_name, aligned))
    return aligned_tables, headers


def _ensure_python_value(value):
    if pd.isna(value):
        return None
    return value


def _iter_sheet_rows(sheet):
    if isinstance(sheet, pd.DataFrame):
        formats = sheet.attrs.get("number_formats")
        for index, row in enumerate(sheet.itertuples(index=False, name=None)):
            yield _FormattedRow(row, formats[index]) if formats is not None else row
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
    data = _read_input_file(file_path, sheet_name)
    headers = _first_row_headers(data)
    if not headers or all((header == "" for header in headers)):
        raise WorkbookProcessingError("The selected worksheet does not contain any data.")
    return data, data, headers


def build_group_map(
    sheet,
    department_column: str,
    id_number_column: str | None = None,
    grouping_columns: list[str] | None = None,
):
    headers = _first_row_headers(sheet)
    if not headers:
        raise WorkbookProcessingError("The selected worksheet does not contain any data.")

    if grouping_columns is None:
        grouping_columns = [department_column]
        if id_number_column:
            grouping_columns.append(id_number_column)
    grouping_columns = [str(column).strip() for column in grouping_columns if str(column).strip()]
    if not grouping_columns:
        raise WorkbookProcessingError("At least one grouping column must be selected.")

    if isinstance(sheet, pd.DataFrame):
        columns = [str(col).strip() for col in list(sheet.columns)]
        missing = [column for column in grouping_columns if column not in columns]
        if missing:
            raise WorkbookProcessingError(f"Grouping column(s) not found in the worksheet headers: {', '.join(missing)}")
        grouping_indexes = [columns.index(column) for column in grouping_columns]
    else:
        missing = [column for column in grouping_columns if column not in headers]
        if missing:
            raise WorkbookProcessingError(f"Grouping column(s) not found in the worksheet headers: {', '.join(missing)}")
        grouping_indexes = [headers.index(column) for column in grouping_columns]

    groups = OrderedDict()

    rows = iter(_iter_sheet_rows(sheet))
    if not isinstance(sheet, pd.DataFrame):
        next(rows, None)
    for row in rows:
        if not _row_has_values(row):
            continue

        group_key = tuple(
            normalize_group_value(
                _ensure_python_value(row[index]) if index < len(row) else None,
                "UNKNOWN GROUP",
            )
            for index in grouping_indexes
        )

        if group_key not in groups:
            groups[group_key] = []
        groups[group_key].append(row)

    if not groups:
        raise WorkbookProcessingError("No data rows were found in the selected worksheet.")

    return headers, groups


def build_combined_group_map(
    input_tables,
    department_column: str,
    id_number_column: str | None = None,
    grouping_columns: list[str] | None = None,
    combine_duplicates: bool = True,
):
    """Build the same source-aware groups for the preview and workbook export."""
    normalized_tables, headers = _normalize_input_tables(input_tables)
    combined_groups = OrderedDict()
    reserved_labels = {source_label for source_label, _table in normalized_tables}
    used_labels = set()
    for source_label, table in normalized_tables:
        # Different file/sheet pairs can form identical human-readable labels.
        # Make those labels unique in input order for both preview and export.
        label = source_label
        suffix = 2
        while label in used_labels:
            label = f"{source_label} ({suffix})"
            suffix += 1
            while label in reserved_labels:
                label = f"{source_label} ({suffix})"
                suffix += 1
        used_labels.add(label)
        _, groups = build_group_map(table, department_column, id_number_column, grouping_columns)
        for group_key, rows in groups.items():
            output_key = group_key if combine_duplicates else (*group_key, label)
            combined_groups.setdefault(output_key, []).extend(rows)

    if not combined_groups:
        raise WorkbookProcessingError("No valid rows were found across the selected worksheets and files.")
    return headers, combined_groups


def split_workbook(file_path: str, sheet_name: str | _SheetSelection | None,
                   department_column: str, id_number_column: str | None,
                   output_file: str, selected_groups: set[tuple[str, ...]] | None = None,
                   progress_callback: Callable[[int, str], None] | None = None,
           combine_duplicates: bool = True,
           grouping_columns: list[str] | None = None,
           sort_mode: str = SORT_ALPHABETICAL):
    return split_workbooks([file_path], sheet_name, department_column, id_number_column, output_file,
               selected_groups, progress_callback, combine_duplicates, grouping_columns, sort_mode)


def split_workbooks(file_paths: list[str], sheet_name: str | _SheetSelection | None,
                    department_column: str, id_number_column: str | None,
                    output_file: str, selected_groups: set[tuple[str, ...]] | None = None,
                    progress_callback: Callable[[int, str], None] | None = None,
                    combine_duplicates: bool = True,
                    grouping_columns: list[str] | None = None,
                    sort_mode: str = SORT_ALPHABETICAL):
    if not file_paths:
        raise WorkbookProcessingError("No files were selected for processing.")
    ensure_output_is_not_input(output_file, file_paths)

    total_files = len(file_paths)
    input_tables = []

    for file_index, (file_path, file_label) in enumerate(zip(file_paths, source_file_labels(file_paths)), start=1):
        for source_sheet_name, source_df in _iter_input_tables(file_path, sheet_name):
            if not source_df.empty:
                source_label = f"{file_label} - {source_sheet_name}"
                input_tables.append((source_label, source_df))

        if progress_callback:
            progress_callback(int((file_index / total_files) * 40), f"Reading: {os.path.basename(file_path)}")

    return split_input_tables(input_tables, department_column, id_number_column, output_file,
                              selected_groups, progress_callback, combine_duplicates,
                              grouping_columns, sort_mode)


def split_input_tables(input_tables, department_column: str, id_number_column: str | None,
                       output_file: str, selected_groups: set[tuple[str, ...]] | None = None,
                       progress_callback: Callable[[int, str], None] | None = None,
                       combine_duplicates: bool = True, grouping_columns: list[str] | None = None,
                       sort_mode: str = SORT_ALPHABETICAL, value_mappings=None,
                       correct_merged_values: bool = False):
    """Export an already-loaded set of tables, preserving the preview snapshot."""
    combined_headers, combined_groups = build_combined_group_map(
        input_tables, department_column, id_number_column, grouping_columns, combine_duplicates
    )
    if value_mappings:
        effective_columns = grouping_columns if grouping_columns is not None else [department_column] + (
            [id_number_column] if id_number_column else []
        )
        combined_groups = merge_group_values(combined_groups, effective_columns, value_mappings)

    if selected_groups is None:
        selected_groups = set(combined_groups.keys())

    if not selected_groups:
        raise WorkbookProcessingError("No groups were selected to generate worksheets.")
    if not selected_groups.issubset(combined_groups):
        raise WorkbookProcessingError("The selected groups have changed. Refresh the group preview and try again.")

    output_ext = os.path.splitext(output_file)[1].lower()
    column_count = len(combined_headers)
    if output_ext != ".xls" and column_count > 16_384:
        raise WorkbookProcessingError(
            f"The output has {column_count:,} columns, exceeding Excel's 16,384-column worksheet limit. "
            "Reduce the number of input columns and try again."
        )

    processed_rows = 0
    ordered_groups = preview_group_items(
        combined_groups, selected_groups, sort_mode, headers=combined_headers,
        grouping_columns=effective_columns if value_mappings else None,
        value_mappings=value_mappings, correct_merged_values=correct_merged_values,
    )
    total_groups = len(ordered_groups)

    # Both Excel formats limit cell text to 32,767 characters. openpyxl
    # silently truncates longer strings, so check the actual export first.
    for sheet_title, _group_key, rows in ordered_groups:
        for row_num, values in enumerate(chain([combined_headers], rows), start=1):
            for col_num, value in enumerate(values, start=1):
                if not isinstance(value, str):
                    continue
                text_length = len(value)
                if text_length > 32767:
                    raise WorkbookProcessingError(
                        f"Text in worksheet '{sheet_title}', row {row_num}, column {col_num} "
                        "exceeds Excel's 32,767-character cell limit. Shorten the value and try again."
                    )
                if row_num == 1 and output_ext != ".xls" and text_length > 255:
                    raise WorkbookProcessingError(
                        f"The table header in column {col_num} exceeds Excel's 255-character table header limit. "
                        "Shorten the header and try again."
                    )

    if output_ext == ".xls":
        workbook = xlwt.Workbook()
        styles = legacy_styles(workbook)
        for index, (sheet_title, group_key, rows) in enumerate(ordered_groups, start=1):
            display_name = " - ".join(group_key)
            sheet = workbook.add_sheet(sheet_title[:31])
            format_legacy_dimensions(sheet, combined_headers, rows)
            for col_num, header in enumerate(combined_headers or []):
                sheet.write(0, col_num, header, styles["header"])
            for row_num, row_values in enumerate(rows, start=1):
                sheet.row(row_num).height = 23 * 20
                sheet.row(row_num).height_mismatch = True
                for col_num, value in enumerate(row_values):
                    formats = getattr(row_values, "number_formats", None)
                    number_format = formats[col_num] if formats is not None else None
                    style = legacy_cell_style(
                        styles, row_num % 2 == 0,
                        combined_headers[col_num].casefold() in CENTERED_COLUMNS,
                        number_format or ("YYYY-MM-DD HH:MM:SS" if isinstance(value, date) else "General"),
                    )
                    if isinstance(value, _ExcelError) and value in _LEGACY_ERROR_CODES:
                        sheet.row(row_num).set_cell_error(col_num, _LEGACY_ERROR_CODES[value], style)
                    elif isinstance(value, _ExcelError):
                        # Newer error codes cannot be represented in legacy XLS.
                        sheet.write(row_num, col_num, str(value), style)
                    elif pd.isna(value):
                        sheet.write(row_num, col_num, "", style)
                    else:
                        sheet.write(row_num, col_num, value, style)
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

    for index, (sheet_title, group_key, rows) in enumerate(ordered_groups, start=1):
        display_name = " - ".join(group_key)
        target_sheet = output_wb.create_sheet(title=sheet_title)

        for col_num, header in enumerate(combined_headers, start=1):
            cell = target_sheet.cell(row=1, column=col_num, value=header)
            cell.data_type = "s"

        for row_num, row_values in enumerate(rows, start=2):
            for col_num, value in enumerate(row_values, start=1):
                if pd.isna(value):
                    target_sheet.cell(row=row_num, column=col_num, value=None)
                else:
                    cell = target_sheet.cell(row=row_num, column=col_num, value=value)
                    formats = getattr(row_values, "number_formats", None)
                    if formats is not None and formats[col_num - 1]:
                        cell.number_format = formats[col_num - 1]
                    if isinstance(value, _ExcelError):
                        cell.data_type = "e"
                    elif isinstance(value, str):
                        cell.data_type = "s"
            processed_rows += 1

        format_worksheet(target_sheet, index)
        if progress_callback:
            progress_callback(40 + int((index / total_groups) * 60), f"Creating: {display_name}")

    output_wb.save(output_file)
    return {
        "created_sheets": len(output_wb.worksheets),
        "processed_rows": processed_rows,
        "output_file": output_file,
    }
