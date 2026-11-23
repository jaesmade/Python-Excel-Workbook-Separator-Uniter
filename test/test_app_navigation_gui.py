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
    appearance_at_close = []

    def change_settings(dialog):
        def remember_appearance(event):
            if event.widget is dialog:
                appearance_at_close.append(ctk.get_appearance_mode())

        dialog.bind("<Destroy>", remember_appearance)
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
        assert appearance_at_close == ["Light"]
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


def test_review_tools_fit_at_minimum_window_size():
    with launch_app() as app:
        width, height = app.wm_minsize()
        app.wm_geometry(f"{width}x{height}")
        app.update()
        review_button = next(
            widget for panel in app.winfo_children() for widget in panel.winfo_children()
            if isinstance(widget, ctk.CTkButton) and widget.cget("text") == "Review duplicate records"
        )
        panel_bottom = review_button.master.winfo_rooty() + review_button.master.winfo_height()
        button_bottom = review_button.winfo_rooty() + review_button.winfo_height()
        assert button_bottom <= panel_bottom


def test_long_group_names_wrap_at_minimum_window_size():
    with launch_app() as app:
        width, height = app.wm_minsize()
        app.wm_geometry(f"{width}x{height}")
        short = ("Operations",)
        long = ("Operations", "National Customer Experience and Change Management", "Senior Administrator", "North")
        app.groups = {short: [(1,)], long: [(2,)]}
        app._render_groups()
        app.update()
        assert app.group_checkboxes[long].winfo_height() > app.group_checkboxes[short].winfo_height()
