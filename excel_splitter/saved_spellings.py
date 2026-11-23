"""Persistent, column-specific spelling choices approved by the user."""

import json
import os
from pathlib import Path
import tempfile


def _default_path():
    config = os.environ.get("APPDATA") if os.name == "nt" else os.environ.get("XDG_CONFIG_HOME")
    base = Path(config) if config else Path.home() / ".config"
    return base / "Excel-arranger" / "saved-spellings.json"


def _validate(mappings):
    for key, target in mappings.items():
        if (not isinstance(key, tuple) or len(key) != 2
                or not all(isinstance(value, str) and value for value in key)
                or not isinstance(target, str) or not target):
            raise ValueError("Saved spelling choices must contain nonempty text values.")
    for column, source in mappings:
        seen = set()
        value = source
        while (column, value) in mappings:
            if value in seen:
                raise ValueError("Saved spelling choices contain a cycle.")
            seen.add(value)
            value = mappings[(column, value)]


class SavedSpellings:
    """Read and update a local file without modifying input workbooks."""

    def __init__(self, path=None):
        self.path = Path(path) if path is not None else _default_path()

    def load(self):
        try:
            with self.path.open(encoding="utf-8") as handle:
                data = json.load(handle)
        except FileNotFoundError:
            return {}
        except (json.JSONDecodeError, UnicodeError) as exc:
            raise ValueError("Saved spelling choices could not be read.") from exc
        if not isinstance(data, dict) or not isinstance(data.get("choices"), list):
            raise ValueError("Saved spelling choices have an invalid format.")
        mappings = {}
        for entry in data["choices"]:
            if not isinstance(entry, dict) or not all(name in entry for name in ("column", "source", "target")):
                raise ValueError("Saved spelling choices have an invalid format.")
            column, source, target = entry["column"], entry["source"], entry["target"]
            if not all(isinstance(value, str) and value for value in (column, source, target)):
                raise ValueError("Saved spelling choices have an invalid format.")
            key = (column, source)
            if key in mappings:
                raise ValueError("Saved spelling choices contain duplicate entries.")
            mappings[key] = target
        _validate(mappings)
        return mappings

    def save(self, mappings):
        self.update(mappings)

    def update(self, mappings, removals=()):
        """Apply additions and removals together, preserving other choices."""
        combined = self.load()
        for key in removals:
            combined.pop(key, None)
        combined.update(mappings)
        _validate(combined)
        self._write(combined)

    def remove(self, column, source):
        self.update({}, [(column, source)])

    def _write(self, mappings):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        data = {"choices": [
            {"column": column, "source": source, "target": target}
            for (column, source), target in sorted(mappings.items())
        ]}
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.path.parent,
                                             delete=False) as handle:
                temporary = Path(handle.name)
                json.dump(data, handle, ensure_ascii=False, indent=2)
                handle.write("\n")
            os.replace(temporary, self.path)
        finally:
            if temporary is not None and temporary.exists():
                temporary.unlink()
