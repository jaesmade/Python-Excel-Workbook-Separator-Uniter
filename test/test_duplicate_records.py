import pandas as pd
import pytest
from pandas.testing import assert_frame_equal

from excel_splitter.duplicate_records import find_duplicate_records


def test_finds_matches_with_sources_rows_and_original_values():
    first = pd.DataFrame({"Name": ["  Alice  Smith ", "Bob", "alice smith"], "ID": ["001", "002", "001"]})
    second = pd.DataFrame({"ID": ["001", "003"], "Name": ["ALICE\tSMITH", "Cara"], "Note": ["new", "other"]})
    groups = find_duplicate_records([("one - Sheet", first), ("two - Sheet", second)], ["Name", "ID"])
    assert len(groups) == 1
    assert [(record.source, record.row_position) for record in groups[0]] == [
        ("one - Sheet", 1), ("one - Sheet", 3), ("two - Sheet", 1),
    ]
    assert groups[0][0].values == {"Name": "  Alice  Smith ", "ID": "001"}
    assert groups[0][2].values == {"ID": "001", "Name": "ALICE\tSMITH", "Note": "new"}


def test_ignores_blank_or_missing_keys_and_distinguishes_text_from_numbers():
    first = pd.DataFrame({"Key": [None, " ", 1, "1", 1, "1", "valid"], "Other": ["a", "a", "a", "a", "a", "a", ""]})
    second = pd.DataFrame({"Other": ["a"], "Unrelated": ["valid"]})
    groups = find_duplicate_records([("first", first), ("second", second)], ["Key", "Other"])
    assert [[r.row_position for r in group] for group in groups] == [[3, 5], [4, 6]]


def test_source_labels_are_disambiguated_without_mutating_inputs():
    table = pd.DataFrame({"Key": ["value", "value"]})
    original = table.copy(deep=True)
    groups = find_duplicate_records([("source", table), ("source", table), ("source (2)", table)], ["Key"])
    assert [r.source for r in groups[0]] == ["source", "source", "source (3)", "source (3)", "source (2)", "source (2)"]
    assert_frame_equal(table, original)


def test_requires_at_least_one_matching_column():
    with pytest.raises(ValueError, match="column"):
        find_duplicate_records([], [])


def test_boolean_values_do_not_match_equal_numbers():
    table = pd.DataFrame({"Key": [True, 1, False, 0]}, dtype=object)
    assert find_duplicate_records([("source", table)], ["Key"]) == []
