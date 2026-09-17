import pandas as pd

from excel_splitter.core import build_group_map, split_workbooks


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
