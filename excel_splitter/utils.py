from __future__ import annotations

from typing import Iterable


INVALID_SHEET_CHARS = set(':/?*[]\\')


def normalize_group_value(value, default_label: str) -> str:
    if value is None:
        return default_label
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned if cleaned else default_label
    return str(value).strip() or default_label


def sanitize_sheet_name(name: str, used_names: set[str]) -> str:
    cleaned = "".join("-" if char in INVALID_SHEET_CHARS else char for char in str(name or "GROUP"))
    cleaned = cleaned.strip().strip(".")
    if not cleaned:
        cleaned = "GROUP"

    base_name = cleaned[:31]
    if not base_name:
        base_name = "GROUP"

    candidate = base_name.strip()
    counter = 2
    final_name = candidate
    while final_name in used_names:
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
