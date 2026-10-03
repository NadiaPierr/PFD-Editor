"""Integration tests for the Split PDF workflow, without a running GUI."""

from pathlib import Path
import sys
from types import SimpleNamespace

import fitz


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import main
from main import PDFEditorApp


class RecordingWidget:
    def __init__(self):
        self.calls = []

    def configure(self, **kwargs):
        self.calls.append(kwargs)


class RecordingDialog:
    def __init__(self):
        self.destroyed = False

    def destroy(self):
        self.destroyed = True


def create_pdf(path, page_texts):
    with fitz.open() as document:
        for text in page_texts:
            page = document.new_page()
            page.insert_text((72, 72), text)
        document.save(path)


def split_app(source_pdf, source_path):
    app = SimpleNamespace(
        pdf_document=source_pdf,
        pdf_path=source_path,
        split_button=RecordingWidget(),
        status_label=RecordingWidget(),
        update_idletasks=lambda: None,
    )
    app.load_pdf_calls = []
    app.load_pdf = app.load_pdf_calls.append
    app.parse_page_selection = lambda selection, total_pages: PDFEditorApp.parse_page_selection(
        None, selection, total_pages
    )
    return app


def test_split_pdf_pages_writes_requested_pages_in_order(tmp_path, monkeypatch):
    source_path = tmp_path / "source.pdf"
    output_path = tmp_path / "split.pdf"
    create_pdf(source_path, ["Page one", "Page two", "Page three", "Page four"])

    dialog = RecordingDialog()
    info_messages = []
    monkeypatch.setattr(
        main.filedialog, "asksaveasfilename", lambda **_kwargs: str(output_path)
    )
    monkeypatch.setattr(main.messagebox, "showinfo", lambda *args: info_messages.append(args))
    monkeypatch.setattr(main.messagebox, "showwarning", lambda *_args: None)
    monkeypatch.setattr(main.messagebox, "showerror", lambda *_args: None)

    with fitz.open(source_path) as source_pdf:
        app = split_app(source_pdf, source_path)
        PDFEditorApp.split_pdf_pages(app, "2, 4-3", dialog)

        # A reversed range is invalid and must leave no output behind.
        assert not output_path.exists()
        assert not dialog.destroyed

        PDFEditorApp.split_pdf_pages(app, "2, 4, 2", dialog)

    assert output_path.is_file()
    with fitz.open(output_path) as split_pdf:
        assert split_pdf.page_count == 2
        assert [page.get_text().strip() for page in split_pdf] == ["Page two", "Page four"]
    assert dialog.destroyed
    assert app.load_pdf_calls == [output_path]
    assert info_messages == [("Split complete", "The split PDF was saved successfully.")]


def test_split_pdf_rejects_replacing_source_file(tmp_path, monkeypatch):
    source_path = tmp_path / "source.pdf"
    create_pdf(source_path, ["Page one", "Page two"])
    warnings = []
    monkeypatch.setattr(
        main.filedialog, "asksaveasfilename", lambda **_kwargs: str(source_path)
    )
    monkeypatch.setattr(main.messagebox, "showwarning", lambda *args: warnings.append(args))

    with fitz.open(source_path) as source_pdf:
        app = split_app(source_pdf, source_path)
        dialog = RecordingDialog()
        PDFEditorApp.split_pdf_pages(app, "1", dialog)

    assert warnings == [("Invalid output", "Save the split PDF with a different name.")]
    assert not dialog.destroyed
    assert app.load_pdf_calls == []
