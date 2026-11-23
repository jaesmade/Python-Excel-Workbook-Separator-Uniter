"""Reversible decisions for a single spelling-review session."""

from .core import merge_group_values, suggest_group_value_merges


KEEP_SEPARATE = object()


class SpellingReview:
    """Keep the list stable while applying and undoing individual choices."""

    def __init__(self, groups, columns, mappings, selected):
        self.columns = tuple(columns)
        self.groups = groups.copy()
        self.mappings = dict(mappings)
        self.selected = set(selected)
        self.suggestions = suggest_group_value_merges(groups, columns)
        self.decisions = {}
        self._history = []

    @property
    def can_undo(self):
        return bool(self._history)

    def copy(self):
        """Clone session state and undo history while sharing immutable row data."""
        clone = self.__class__.__new__(self.__class__)
        clone.columns = self.columns
        clone.groups = self.groups.copy()
        clone.mappings = self.mappings.copy()
        clone.selected = self.selected.copy()
        clone.suggestions = self.suggestions.copy()
        clone.decisions = self.decisions.copy()
        clone._history = [
            (index, groups.copy(), mappings.copy(), selected.copy())
            for index, groups, mappings, selected in self._history
        ]
        return clone

    def _root(self, column, value):
        seen = set()
        while (column, value) in self.mappings:
            if value in seen:
                raise ValueError("Value merge choices contain a cycle.")
            seen.add(value)
            value = self.mappings[(column, value)]
        return value

    def current_value(self, column, value):
        """Return the spelling currently used for an original suggestion."""
        return self._root(column, value)

    def choose(self, index, answer):
        if index in self.decisions:
            raise ValueError("Undo this decision before choosing again.")
        column, source, target, _, _ = self.suggestions[index]
        if answer is not KEEP_SEPARATE and answer not in (source, target):
            raise ValueError("Choose one of the suggested spellings or keep them separate.")

        source_root = self._root(column, source)
        target_root = self._root(column, target)
        if answer is not KEEP_SEPARATE and self._root(column, answer) != answer:
            raise ValueError("That spelling was already replaced. Undo the earlier merge first or choose the other spelling.")
        if answer is KEEP_SEPARATE and source_root == target_root:
            raise ValueError("These spellings were already merged. Undo that merge first.")

        if answer is not KEEP_SEPARATE and source_root != target_root:
            for previous_index, decision in self.decisions.items():
                other_column, first, second, _, _ = self.suggestions[previous_index]
                if decision is KEEP_SEPARATE and other_column == column:
                    roots = {self._root(column, first), self._root(column, second)}
                    if roots == {source_root, target_root}:
                        raise ValueError("This would combine spellings marked to keep separate. Undo that choice first.")

        self._history.append((index, self.groups, self.mappings.copy(), self.selected.copy()))
        self.decisions[index] = answer
        if answer is KEEP_SEPARATE or source_root == target_root:
            return
        if answer == source:
            source_root, target_root = target_root, source_root
        mapping = {(column, source_root): target_root}
        self.selected = set(merge_group_values(
            {key: self.groups[key] for key in self.selected if key in self.groups},
            self.columns, mapping,
        ))
        self.mappings.update(mapping)
        self.groups = merge_group_values(self.groups, self.columns, mapping)

    def undo(self):
        if not self._history:
            return False
        index, self.groups, self.mappings, self.selected = self._history.pop()
        del self.decisions[index]
        return True
