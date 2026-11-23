"""Find possible matching records without changing the loaded data."""

from collections import OrderedDict
from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class DuplicateRecord:
    source: str
    row_position: int
    values: dict


def _match_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, bool):
        return ("boolean", value)
    if isinstance(value, str):
        value = " ".join(value.split()).casefold()
        return ("text", value) if value else None
    return ("value", value)


def find_duplicate_records(input_tables, columns):
    """Return matching record groups with sources and 1-based data-row positions.

    Matching is limited to the chosen columns. Blank keys are excluded, while
    original row values and their respective source columns remain intact.
    """
    columns = tuple(columns)
    if not columns:
        raise ValueError("Select at least one matching column.")

    input_tables = list(input_tables)
    reserved = {source for source, _ in input_tables}
    used = set()
    matches = OrderedDict()
    for source, table in input_tables:
        label = source
        suffix = 2
        while label in used:
            label = f"{source} ({suffix})"
            suffix += 1
            while label in reserved:
                label = f"{source} ({suffix})"
                suffix += 1
        used.add(label)
        if any(column not in table.columns for column in columns):
            continue
        for position, row in enumerate(table.itertuples(index=False, name=None), start=1):
            values = dict(zip(table.columns, row))
            key = tuple(_match_value(values[column]) for column in columns)
            if any(value is None for value in key):
                continue
            matches.setdefault(key, []).append(DuplicateRecord(label, position, values))
    return [records for records in matches.values() if len(records) > 1]
