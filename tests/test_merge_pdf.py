"""Tests for the Merge PDFs workflow."""

from pathlib import Path
import sys
from types import SimpleNamespace

import fitz


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import main
from main import PDFEditorApp


class RecordingWidget:
    """Small stand-in for a CustomTkinter widget used by the merge action."""

    def __init__(self):
        self.calls = []

    def configure(self, **kwargs):
        self.calls.append(kwargs)


def create_pdf(path, page_texts):
    with fitz.open() as document:
        for text in page_texts:
            page = document.new_page()
            page.insert_text((72, 72), text)
        document.save(path)


def merge_app():
    """Return the minimum app state needed by ``merge_pdf_files``."""
    app = SimpleNamespace(
        merge_button=RecordingWidget(),
        convert_button=RecordingWidget(),
        status_label=RecordingWidget(),
        pdf_path=None,
        update_idletasks=lambda: None,
    )
    app.load_pdf_calls = []
    app.load_pdf = app.load_pdf_calls.append
    return app


def test_merge_pdf_files_combines_pages_in_selected_order(tmp_path, monkeypatch):
    first_pdf = tmp_path / "first.pdf"
    second_pdf = tmp_path / "second.pdf"
    output_pdf = tmp_path / "merged.pdf"
    create_pdf(first_pdf, ["First page", "Second page"])
    create_pdf(second_pdf, ["Third page"])

    app = merge_app()
    info_messages = []
    monkeypatch.setattr(
        main.filedialog, "askopenfilenames", lambda **_kwargs: (str(first_pdf), str(second_pdf))
    )
    monkeypatch.setattr(
        main.filedialog, "asksaveasfilename", lambda **_kwargs: str(output_pdf)
    )
    monkeypatch.setattr(main.messagebox, "showinfo", lambda *args: info_messages.append(args))
    monkeypatch.setattr(main.messagebox, "showwarning", lambda *_args: None)
    monkeypatch.setattr(main.messagebox, "showerror", lambda *_args: None)

    PDFEditorApp.merge_pdf_files(app)

    assert output_pdf.is_file()
    with fitz.open(output_pdf) as merged_pdf:
        assert merged_pdf.page_count == 3
        assert [page.get_text().strip() for page in merged_pdf] == [
            "First page",
            "Second page",
            "Third page",
        ]
    assert app.load_pdf_calls == [output_pdf]
    assert info_messages == [("Merge complete", "The merged PDF was saved successfully.")]
