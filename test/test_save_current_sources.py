"""Protect source files selected while a generated workbook is pending."""

from types import SimpleNamespace
from unittest.mock import Mock, patch

from excel_splitter.gui.main_window import ExcelSplitterApp


def test_save_does_not_overwrite_source_selected_after_generation_started(tmp_path):
    original_source = tmp_path / "original.xlsx"
    new_source = tmp_path / "new-source.xlsx"
    generated = tmp_path / "generated.xlsx"
    original_source.write_bytes(b"original source")
    new_source.write_bytes(b"newly loaded source")
    generated.write_bytes(b"generated workbook")
    app = SimpleNamespace(
        file_paths=[str(original_source)],
        choose_output_location=Mock(),
        _generation_failed=Mock(),
        generate_btn=Mock(),
        progress_var=Mock(),
        progress_bar=Mock(),
    )

    def choose_output(_source_file):
        # A modal dialog may process other UI events before it returns.
        app.file_paths = [str(new_source)]
        return str(new_source)

    app.choose_output_location.side_effect = choose_output

    with patch("excel_splitter.gui.main_window.messagebox.showinfo"):
        ExcelSplitterApp._save_completed_workbook(
            app, str(generated), 1, 1, str(original_source), (str(original_source),)
        )

    assert new_source.read_bytes() == b"newly loaded source"
    app._generation_failed.assert_called_once()
