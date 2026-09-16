import pandas as pd

from excel_splitter.core import build_group_map


def test_group_map_allows_missing_second_column():
    sheet = pd.DataFrame({
        "Department": ["Sales", "Sales", "Support"],
        "Name": ["Alice", "Bob", "Carla"],
    })

    _, groups = build_group_map(sheet, "Department", None)

    assert list(groups.keys()) == [("Sales",), ("Support",)]
    assert len(groups[("Sales",)]) == 2
    assert len(groups[("Support",)]) == 1
