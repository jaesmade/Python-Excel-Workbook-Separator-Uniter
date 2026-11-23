from collections import OrderedDict

import pytest

from excel_splitter.manual_merge import category_values, merge_categories
from excel_splitter.core import merge_group_values


def test_search_finds_all_current_values_in_selected_column():
    groups = {('123', 'North'): [1], ('Apples', 'South'): [2], ('apricot', 'North'): [3]}
    assert category_values(groups, ['Code', 'Area'], 'Code', 'ap') == ['Apples', 'apricot']
    assert category_values(groups, ['Code', 'Area'], 'Code', '123') == ['123']
    assert category_values(groups, ['Code', 'Area'], 'Area', 'north') == ['North']


def test_manual_merge_custom_name_preserves_rows_selection_and_mapping_chains():
    groups = OrderedDict([
        (('Old', 'North', 'file1'), [('original A',)]),
        (('Other', 'North', 'file1'), [('original B',)]),
        (('Old', 'South', 'file2'), [('original C',)]),
    ])
    mappings = {('Type', 'Earlier'): 'Old'}
    merged, updated, selected = merge_categories(groups, ['Type', 'Area'], mappings,
        {('Old', 'North', 'file1')}, 'Type', 'Old', 'Other', '  Preferred  ')
    assert merged == {
        ('Preferred', 'North', 'file1'): [('original A',), ('original B',)],
        ('Preferred', 'South', 'file2'): [('original C',)],
    }
    assert selected == {('Preferred', 'North', 'file1')}
    assert merge_group_values({('Earlier', 'North', 'file1'): [1]}, ['Type', 'Area'], updated) == {('Preferred', 'North', 'file1'): [1]}
    assert mappings == {('Type', 'Earlier'): 'Old'}
    assert ('Old', 'North', 'file1') in groups


def test_manual_merge_defaults_to_first_and_validates_choices():
    groups = {('A',): [1], ('B',): [2], ('C',): [3]}
    merged, mappings, selected = merge_categories(groups, ['Type'], {}, {('B',)}, 'Type', 'A', 'B', '')
    assert merged == {('A',): [1, 2], ('C',): [3]}
    assert mappings == {('Type', 'B'): 'A'}
    assert selected == {('A',)}
    for first, second, preferred in [('A', 'A', ''), ('A', 'missing', ''), ('A', 'B', 'C')]:
        with pytest.raises(ValueError):
            merge_categories(groups, ['Type'], {}, set(), 'Type', first, second, preferred)
    with pytest.raises(ValueError):
        merge_categories(groups, ['Type'], {('Type', 'Former'): 'A'}, set(), 'Type', 'A', 'B', 'Former')


def test_manual_merge_export_uses_preferred_name_and_keeps_original_cells(tmp_path):
    import pandas as pd
    from openpyxl import load_workbook
    from excel_splitter.core import build_combined_group_map, split_input_tables

    tables = [('file', pd.DataFrame({'Type': ['Alpha', 'Zulu'], 'Name': ['A', 'B']}))]
    _, groups = build_combined_group_map(tables, 'Type', grouping_columns=['Type'])
    merged, mappings, selected = merge_categories(groups, ['Type'], {}, set(groups),
                                                   'Type', 'Alpha', 'Zulu', 'Combined')
    assert set(merged) == {('Combined',)}
    output = tmp_path / 'manual.xlsx'
    split_input_tables(tables, 'Type', None, str(output), selected_groups=selected,
                       grouping_columns=['Type'], value_mappings=mappings)
    workbook = load_workbook(output, data_only=True)
    try:
        assert workbook.sheetnames == ['Combined']
        assert list(workbook.active.values)[1:] == [('Alpha', 'A'), ('Zulu', 'B')]
    finally:
        workbook.close()
