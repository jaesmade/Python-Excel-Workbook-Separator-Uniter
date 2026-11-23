from __future__ import annotations

import re
from typing import Iterable


INVALID_SHEET_CHARS = set(':/?*[]\\')
RESERVED_SHEET_NAMES = frozenset({"history"})
SHEET_NAME_EDGES = re.compile(r"^[\s']+|[\s']+$")


class SheetNameRegistry:
    """Track case-insensitive names and the next suffix for each base name."""

    def __init__(self):
        self._names: set[str] = set(RESERVED_SHEET_NAMES)
        self._next_suffix: dict[tuple[str, ...], int] = {}

    def allocate(self, candidate: str) -> str:
        name = candidate
        if name.casefold() in self._names:
            # Unicode casefolding can expand characters, so candidates that
            # casefold equally may truncate differently when a suffix is added.
            # Share counters only when every possible suffix width produces the
            # same case-insensitive prefix.
            key = (candidate.casefold(),) + tuple(
                candidate[:28 - digits].rstrip(" -_").casefold()
                for digits in range(1, 29)
            )
            counter = self._next_suffix.get(key, 2)
            while True:
                suffix = f" - {counter}"
                name = candidate[:31 - len(suffix)].rstrip(" -_") + suffix
                counter += 1
                if name.casefold() not in self._names:
                    break
            self._next_suffix[key] = counter
        self._names.add(name.casefold())
        return name


def normalize_group_value(value, default_label: str) -> str:
    if value is None:
        return default_label
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned if cleaned else default_label
    return str(value).strip() or default_label


def sanitize_sheet_name(name: str, used_names: set[str] | SheetNameRegistry) -> str:
    cleaned = "".join("-" if char in INVALID_SHEET_CHARS else char for char in str(name or "GROUP"))
    cleaned = SHEET_NAME_EDGES.sub("", cleaned.strip().strip("."))
    # Truncation can leave an internal apostrophe at the end of the title.
    candidate = SHEET_NAME_EDGES.sub("", cleaned[:31]) or "GROUP"
    if isinstance(used_names, SheetNameRegistry):
        return used_names.allocate(candidate)
    unavailable_names = RESERVED_SHEET_NAMES | {name.casefold() for name in used_names}
    counter = 2
    final_name = candidate
    while final_name.casefold() in unavailable_names:
        suffix = f" - {counter}"
        max_length = 31 - len(suffix)
        final_name = candidate[:max_length].rstrip(" -_") + suffix
        counter += 1

    used_names.add(final_name)
    return final_name


def detect_header_names(headers: Iterable[object]) -> list[str]:
    result: list[str] = []
    for value in headers:
        if value is None:
            result.append("")
        else:
            text = str(value).strip()
            result.append(text)
    return result
