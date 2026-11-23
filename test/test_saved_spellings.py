import json

import pytest

from excel_splitter.saved_spellings import SavedSpellings


def test_choices_survive_reopening_and_are_scoped_to_the_column(tmp_path):
    path = tmp_path / "settings" / "spellings.json"
    choices = SavedSpellings(path)
    choices.save({("Type", "JO ORDER"): "JOB ORDER"})
    choices.save({("Department", "JO ORDER"): "JO ORDER OFFICE"})

    assert SavedSpellings(path).load() == {
        ("Type", "JO ORDER"): "JOB ORDER",
        ("Department", "JO ORDER"): "JO ORDER OFFICE",
    }


def test_remove_persists_without_removing_other_choices(tmp_path):
    path = tmp_path / "spellings.json"
    choices = SavedSpellings(path)
    choices.save({("Type", "JO ORDER"): "JOB ORDER", ("Type", "TEMP"): "TEMPORARY"})
    choices.remove("Type", "JO ORDER")

    assert SavedSpellings(path).load() == {("Type", "TEMP"): "TEMPORARY"}


def test_update_removes_and_adds_choices_together_without_changing_unrelated_entries(tmp_path):
    choices = SavedSpellings(tmp_path / "spellings.json")
    choices.save({("Type", "A"): "B", ("Other", "X"): "Y"})

    choices.update({("Type", "B"): "A"}, [("Type", "A")])

    assert choices.load() == {("Type", "B"): "A", ("Other", "X"): "Y"}


def test_cycle_is_rejected_without_changing_saved_choices(tmp_path):
    choices = SavedSpellings(tmp_path / "spellings.json")
    choices.save({("Type", "JO ORDER"): "JOB ORDER"})

    with pytest.raises(ValueError, match="cycle"):
        choices.save({("Type", "JOB ORDER"): "JO ORDER"})

    assert choices.load() == {("Type", "JO ORDER"): "JOB ORDER"}


def test_malformed_file_is_not_silently_overwritten(tmp_path):
    path = tmp_path / "spellings.json"
    path.write_text('{broken', encoding="utf-8")
    choices = SavedSpellings(path)

    with pytest.raises(ValueError):
        choices.save({("Type", "JO ORDER"): "JOB ORDER"})

    assert path.read_text(encoding="utf-8") == '{broken'


def test_valid_json_with_invalid_records_is_rejected(tmp_path):
    path = tmp_path / "spellings.json"
    path.write_text(json.dumps({"choices": [{"column": "Type", "source": "JO ORDER"}]}), encoding="utf-8")

    with pytest.raises(ValueError):
        SavedSpellings(path).load()
