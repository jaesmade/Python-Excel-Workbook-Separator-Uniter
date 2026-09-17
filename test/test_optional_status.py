import pandas as pd
import pytest

from excel_splitter.core import ALL_SHEETS, WorkbookProcessingError, build_group_map, split_workbooks


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


def test_all_sheets_reject_incompatible_headers(tmp_path):
    source = tmp_path / "employees.xlsx"
    output = tmp_path / "split.xlsx"
    with pd.ExcelWriter(source, engine="openpyxl") as writer:
        pd.DataFrame({"Department": ["Sales"], "Name": ["Alice"]}).to_excel(writer, sheet_name="January", index=False)
        pd.DataFrame({"Department": ["Support"], "Email": ["bob@example.com"]}).to_excel(writer, sheet_name="February", index=False)

    with pytest.raises(WorkbookProcessingError, match="do not match"):
        split_workbooks(
            [str(source)],
            ALL_SHEETS,
            "Department",
            None,
            str(output),
            grouping_columns=["Department"],
        )
