"""Search and merge grouping values without changing source records."""

from .core import merge_group_values


def category_values(groups, columns, column, search=""):
    """Return matching current values for one grouping column."""
    if column not in columns:
        raise ValueError("Choose a grouping column.")
    index = columns.index(column)
    query = search.casefold()
    values = {key[index] for key in groups if query in key[index].casefold()}
    return sorted(values, key=lambda value: (value.casefold(), value))


def merge_categories(groups, columns, mappings, selected, column, first, second, preferred=""):
    """Merge two current values and return updated groups, mappings and selection."""
    values = set(category_values(groups, columns, column))
    if first == second or first not in values or second not in values:
        raise ValueError("Choose two different existing values.")
    target = preferred.strip() or first
    if target in values and target not in (first, second):
        raise ValueError("The preferred name is already used by another value. Choose it as one of the two values instead.")
    if (column, target) in mappings:
        raise ValueError("The preferred name was already replaced. Choose a current value or a new name.")

    replacements = {(column, value): target for value in (first, second) if value != target}
    updated_mappings = dict(mappings)
    updated_mappings.update(replacements)
    updated_selected = set(merge_group_values(
        {key: groups[key] for key in selected if key in groups}, columns, replacements
    ))
    return merge_group_values(groups, columns, replacements), updated_mappings, updated_selected
