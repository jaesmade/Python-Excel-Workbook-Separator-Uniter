"""Regression tests for saving generated workbooks safely."""

from types import SimpleNamespace
from unittest.mock import Mock, patch

from excel_splitter.gui.main_window import ExcelSplitterApp


def make_app(destination):
    return SimpleNamespace(
        choose_output_location=lambda _source: str(destination),
        file_paths=[],
        progress_var=Mock(),
        progress_bar=Mock(),
        generate_btn=Mock(),
        _generation_failed=Mock(),
    )


def test_interrupted_copy_preserves_existing_output_and_generated_workbook(tmp_path):
    generated = tmp_path / "generated.xlsx"
    generated.write_bytes(b"complete generated workbook")
    output = tmp_path / "output.xlsx"
    output.write_bytes(b"previous workbook")
    app = make_app(output)

    def interrupted_copy(_source, destination):
        with open(destination, "wb") as destination_file:
            destination_file.write(b"partial")
        raise OSError("copy interrupted")

    with patch("excel_splitter.gui.main_window.shutil.copyfile", side_effect=interrupted_copy):
        ExcelSplitterApp._save_completed_workbook(app, str(generated), 1, 1, "source.xlsx")

    assert output.read_bytes() == b"previous workbook"
    assert generated.read_bytes() == b"complete generated workbook"
    assert str(generated) in str(app._generation_failed.call_args.args[0])
    assert set(tmp_path.iterdir()) == {generated, output}


def test_successful_save_replaces_output_and_removes_generated_workbook(tmp_path):
    generated = tmp_path / "generated.xlsx"
    generated.write_bytes(b"complete generated workbook")
    output = tmp_path / "output.xlsx"
    output.write_bytes(b"previous workbook")
    app = make_app(output)

    with patch("excel_splitter.gui.main_window.messagebox.showinfo"):
        ExcelSplitterApp._save_completed_workbook(app, str(generated), 1, 1, "source.xlsx")

    assert output.read_bytes() == b"complete generated workbook"
    assert not generated.exists()
    assert set(tmp_path.iterdir()) == {output}


def test_failed_replace_preserves_existing_output_and_generated_workbook(tmp_path):
    generated = tmp_path / "generated.xlsx"
    generated.write_bytes(b"complete generated workbook")
    output = tmp_path / "output.xlsx"
    output.write_bytes(b"previous workbook")
    app = make_app(output)

    with patch("excel_splitter.gui.main_window.os.replace", side_effect=OSError("replace failed")), \
         patch("excel_splitter.gui.main_window.messagebox.showinfo"):
        ExcelSplitterApp._save_completed_workbook(app, str(generated), 1, 1, "source.xlsx")

    assert output.read_bytes() == b"previous workbook"
    assert generated.read_bytes() == b"complete generated workbook"
    assert str(generated) in str(app._generation_failed.call_args.args[0])
    assert set(tmp_path.iterdir()) == {generated, output}
