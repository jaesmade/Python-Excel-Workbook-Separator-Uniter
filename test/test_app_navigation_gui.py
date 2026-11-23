"""Exercise the desktop settings and project link through their dialogs."""

import tkinter as tk
from contextlib import contextmanager
from unittest.mock import patch

import customtkinter as ctk
import pytest

from excel_splitter.gui.main_window import ExcelSplitterApp


@contextmanager
def launch_app():
    try:
        window = ExcelSplitterApp()
    except tk.TclError as exc:
        pytest.skip(f"Tk display unavailable: {exc}")
    try:
        window.update_idletasks()
        yield window
    finally:
        window.destroy()


def test_settings_persist_and_keep_group_selection():
    def change_settings(dialog):
        dialog.appearance_var.set("Dark")
        dialog.page_size_var.set("25")
        dialog.save_button.invoke()

    groups = {(f"Group {index:02d}",): [(index,)] for index in range(60)}
    with launch_app() as app:
        app.groups = groups
        app._render_groups()
        app.group_checkboxes[("Group 00",)].toggle()
        with patch.object(app, "wait_window", side_effect=change_settings):
            app.settings_button.invoke()
        assert ctk.get_appearance_mode() == "Dark"
        assert len(app.group_checkboxes) == 25
        assert ("Group 00",) not in app.selected_group_keys
        app._change_group_page(1)
        assert next(iter(app.group_checkboxes)) == ("Group 25",)

    with launch_app() as relaunched:
        relaunched.groups = groups
        relaunched._render_groups()
        assert len(relaunched.group_checkboxes) == 25
        assert ctk.get_appearance_mode() == "Dark"


def test_help_button_opens_the_repository():
    def open_project(dialog):
        dialog.github_button.invoke()
        dialog.destroy()

    with launch_app() as app, patch("webbrowser.open_new_tab", return_value=True) as browser:
        with patch.object(app, "wait_window", side_effect=open_project):
            app.help_button.invoke()

    browser.assert_called_once_with("https://github.com/jaesmade/Python-Excel-Workbook-Separator-Uniter")
