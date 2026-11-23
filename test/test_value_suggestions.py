from collections import OrderedDict

import pandas as pd
from openpyxl import load_workbook

from excel_splitter.core import (
    build_combined_group_map,
    merge_group_values,
    split_input_tables,
    suggest_group_value_merges,
)


def test_suggestions_find_typos_and_capitalization_but_ignore_numeric_ids():
    groups = OrderedDict([
        (("JO ORDER",), [(1,)]),
        (("Job Order",), [(2,), (3,), (4,)]),
        (("LABOR PAYYROLL",), [(5,)]),
        (("LABOR PAYROLL",), [(6,), (7,)]),
        (("labor payroll",), [(8,)]),
        (("123456",), [(9,)]),
        (("123457",), [(10,)]),
        (("UNKNOWN GROUP",), [(11,)]),
        (("PERMANENT",), [(12,)]),
    ])
    suggestions = suggest_group_value_merges(groups, ["ID Number"])
    pairs = {(column, source, target) for column, source, target, _, _ in suggestions}
    assert ("ID Number", "JO ORDER", "Job Order") in pairs
    assert ("ID Number", "LABOR PAYYROLL", "LABOR PAYROLL") in pairs
    assert ("ID Number", "labor payroll", "LABOR PAYROLL") in pairs
    assert all(not source.isdigit() and not target.isdigit() for _, source, target in pairs)
    assert all(source != "UNKNOWN GROUP" and target != "UNKNOWN GROUP" for _, source, target in pairs)


def test_approved_mapping_merges_group_keys_without_modifying_rows_or_source_labels():
    groups = OrderedDict([
        (("Sales", "JO ORDER", "File A"), [("Sales", "JO ORDER")]),
        (("Sales", "JOB ORDER", "File A"), [("Sales", "JOB ORDER")]),
        (("Sales", "JO ORDER", "File B"), [("Sales", "JO ORDER")]),
    ])
    merged = merge_group_values(groups, ["Department", "ID Number"], {("ID Number", "JO ORDER"): "JOB ORDER"})
    assert list(merged) == [("Sales", "JOB ORDER", "File A"), ("Sales", "JOB ORDER", "File B")]
    assert merged[("Sales", "JOB ORDER", "File A")] == [("Sales", "JO ORDER"), ("Sales", "JOB ORDER")]
    assert len(groups) == 3


def test_export_uses_approved_mappings_and_preserves_original_cell_values(tmp_path):
    tables = [("file", pd.DataFrame({"ID Number": ["JO ORDER", "JOB ORDER"], "Name": ["A", "B"]}))]
    output = tmp_path / "output.xlsx"
    result = split_input_tables(tables, "ID Number", None, str(output),
                                selected_groups={("JOB ORDER",)},
                                grouping_columns=["ID Number"],
                                value_mappings={("ID Number", "JO ORDER"): "JOB ORDER"})
    assert result["created_sheets"] == 1
    workbook = load_workbook(output, data_only=True)
    try:
        assert workbook.sheetnames == ["JOB ORDER"]
        assert list(workbook.active.values)[1:] == [("JO ORDER", "A"), ("JOB ORDER", "B")]
    finally:
        workbook.close()
