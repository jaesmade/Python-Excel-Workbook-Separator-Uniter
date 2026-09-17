import pandas as pd

from excel_splitter.core import (
    ALL_SHEETS,
    SORT_ALPHABETICAL,
    SORT_COUNT_ASCENDING,
    SORT_COUNT_DESCENDING,
    build_group_map,
    sort_group_items,
    split_workbooks,
)


def test_sort_group_items_supports_all_modes():
    groups = {
        ("beta",): [1],
        ("Alpha",): [1, 2],
        ("gamma",): [1, 2, 3],
    }

    assert [key for key, _rows in sort_group_items(groups, SORT_ALPHABETICAL)] == [
        ("Alpha",), ("beta",), ("gamma",),
    ]
    assert [key for key, _rows in sort_group_items(groups, SORT_COUNT_ASCENDING)] == [
        ("beta",), ("Alpha",), ("gamma",),
    ]
    assert [key for key, _rows in sort_group_items(groups, SORT_COUNT_DESCENDING)] == [
        ("gamma",), ("Alpha",), ("beta",),
    ]


def test_sort_group_items_preserves_order_for_equal_counts():
    groups = {("First",): [1], ("Second",): [2]}

    assert list(sort_group_items(groups, SORT_COUNT_ASCENDING)) == list(groups.items())


def test_group_map_allows_missing_second_column():
    sheet = pd.DataFrame({
        "Department": ["Sales", "Sales", "Support"],
        "Name": ["Alice", "Bob", "Carla"],
    })

    _, groups = build_group_map(sheet, "Department", None)

    assert list(groups.keys()) == [("Sales",), ("Support",)]
    assert len(groups[("Sales",)]) == 2
    assert len(groups[("Support",)]) == 1


def test_split_workbooks_accepts_csv(tmp_path):
    source = tmp_path / "employees.csv"
    output = tmp_path / "split.xlsx"
    pd.DataFrame({
        "Department": ["Sales", "Support"],
        "Name": ["Alice", "Bob"],
    }).to_csv(source, index=False)

    result = split_workbooks(
        [str(source)],
        "CSV file",
        "Department",
        None,
        str(output),
        grouping_columns=["Department"],
    )

    assert result["created_sheets"] == 2
    assert result["processed_rows"] == 2


def test_split_workbooks_accepts_all_sheets(tmp_path):
    source = tmp_path / "employees.xlsx"
    output = tmp_path / "split.xlsx"
    with pd.ExcelWriter(source, engine="openpyxl") as writer:
        pd.DataFrame({"Department": ["Sales"], "Name": ["Alice"]}).to_excel(writer, sheet_name="January", index=False)
        pd.DataFrame({"Department": ["Support", "Sales"], "Name": ["Bob", "Cara"]}).to_excel(writer, sheet_name="February", index=False)

    result = split_workbooks(
        [str(source)],
        ALL_SHEETS,
        "Department",
        None,
        str(output),
        grouping_columns=["Department"],
    )

    assert result["created_sheets"] == 2
    assert result["processed_rows"] == 3


def test_all_sheets_aligns_different_headers(tmp_path):
    source = tmp_path / "employees.xlsx"
    output = tmp_path / "split.xlsx"
    with pd.ExcelWriter(source, engine="openpyxl") as writer:
        pd.DataFrame({"Department": ["Sales"], "Name": ["Alice"]}).to_excel(writer, sheet_name="January", index=False)
        pd.DataFrame({"Department": ["Support"], "Email": ["bob@example.com"]}).to_excel(writer, sheet_name="February", index=False)

    result = split_workbooks(
        [str(source)],
        ALL_SHEETS,
        "Department",
        None,
        str(output),
        grouping_columns=["Department"],
    )

    assert result["processed_rows"] == 2
    workbook = pd.ExcelFile(output)
    assert set(workbook.sheet_names) == {"Sales", "Support"}
    sales = pd.read_excel(output, sheet_name="Sales")
    support = pd.read_excel(output, sheet_name="Support")
    assert list(sales.columns) == ["Department", "Name", "Email"]
    assert pd.isna(sales.loc[0, "Email"])
    assert pd.isna(support.loc[0, "Name"])


def test_mixed_excel_and_csv_inputs_are_combined(tmp_path):
    excel_source = tmp_path / "employees.xlsx"
    csv_source = tmp_path / "contractors.csv"
    output = tmp_path / "split.xlsx"
    with pd.ExcelWriter(excel_source, engine="openpyxl") as writer:
        pd.DataFrame({"Department": ["Sales"], "Name": ["Alice"]}).to_excel(writer, sheet_name="January", index=False)
    pd.DataFrame({"Department": ["Support"], "Email": ["bob@example.com"]}).to_csv(csv_source, index=False)

    result = split_workbooks(
        [str(excel_source), str(csv_source)],
        ALL_SHEETS,
        "Department",
        None,
        str(output),
        grouping_columns=["Department"],
    )

    assert result["processed_rows"] == 2
    assert result["created_sheets"] == 2


def test_multiple_workbooks_with_different_sheet_names_are_combined(tmp_path):
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    output = tmp_path / "split.xlsx"
    with pd.ExcelWriter(first, engine="openpyxl") as writer:
        pd.DataFrame({"Department": ["Sales"], "Name": ["Alice"]}).to_excel(writer, sheet_name="January", index=False)
    with pd.ExcelWriter(second, engine="openpyxl") as writer:
        pd.DataFrame({"Department": ["Support"], "Name": ["Bob"]}).to_excel(writer, sheet_name="February", index=False)

    result = split_workbooks(
        [str(first), str(second)],
        ALL_SHEETS,
        "Department",
        None,
        str(output),
        grouping_columns=["Department"],
    )

    assert result["processed_rows"] == 2
    assert result["created_sheets"] == 2


def test_generated_worksheets_follow_sort_order(tmp_path):
    source = tmp_path / "employees.csv"
    output = tmp_path / "split.xlsx"
    pd.DataFrame({
        "Department": ["beta", "Alpha", "Alpha", "gamma", "gamma", "gamma"],
    }).to_csv(source, index=False)

    split_workbooks(
        [str(source)],
        "CSV file",
        "Department",
        None,
        str(output),
        grouping_columns=["Department"],
        sort_mode=SORT_COUNT_DESCENDING,
    )

    workbook = pd.ExcelFile(output)
    assert workbook.sheet_names == ["gamma", "Alpha", "beta"]
