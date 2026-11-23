"""Keep tests from reading or changing the user's saved preferences."""

import pytest


@pytest.fixture(autouse=True)
def isolate_saved_preferences(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
