"""End-to-end test for the core PDF split-and-merge workflow.

The test deliberately calls the production action methods while replacing only
the GUI widgets and file dialogs.  It therefore exercises the same PyMuPDF
read/write path as a user workflow, without requiring a display server.
"""

from pathlib import Path
import sys
from types import SimpleNamespace

import fitz


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import main
from main import PDFEditorApp


class RecordingWidget:
    def configure(self, **_kwargs):
        pass


class RecordingDialog:
    def destroy(self):
        pass


def create_source_pdf(path):
    """Create a small document representative of the application's input."""
    # The default PDF font used by PyMuPDF cannot preserve Greek characters in
    # its text extraction map.  The Greek text still makes this a realistic
    # visual input, while the ASCII IDs let the test verify page order after
    # reopening the generated PDFs on every supported machine.
    page_texts = [
        "Σελίδα ένα — ID: alpha",
        "Σελίδα δύο — ID: beta",
        "Σελίδα τρία — ID: gamma",
    ]
    with fitz.open() as document:
        for text in page_texts:
            page = document.new_page()
            page.insert_text((72, 72), text, fontsize=12)
        document.save(path)


def split_app(source_pdf, source_path):
    app = SimpleNamespace(
        pdf_document=source_pdf,
        pdf_path=source_path,
        split_button=RecordingWidget(),
        status_label=RecordingWidget(),
        update_idletasks=lambda: None,
    )
    app.load_pdf = lambda _path: None
    app.parse_page_selection = lambda selection, total_pages: PDFEditorApp.parse_page_selection(
        None, selection, total_pages
    )
    return app


def merge_app():
    app = SimpleNamespace(
        merge_button=RecordingWidget(),
        convert_button=RecordingWidget(),
        status_label=RecordingWidget(),
        pdf_path=None,
        update_idletasks=lambda: None,
    )
    app.load_pdf = lambda _path: None
    return app


def page_ids(pdf_path):
    """Reopen a PDF exactly as the application does and return stable content."""
    with fitz.open(pdf_path) as document:
        return document.page_count, [page.get_text().strip().split("ID: ")[-1] for page in document]


def test_user_can_split_then_merge_and_reopen_the_result(tmp_path, monkeypatch):
    source_path = tmp_path / "greek_source.pdf"
    split_path = tmp_path / "selected_pages.pdf"
    merged_path = tmp_path / "final.pdf"
    create_source_pdf(source_path)

    monkeypatch.setattr(main.messagebox, "showinfo", lambda *_args: None)
    monkeypatch.setattr(main.messagebox, "showwarning", lambda *_args: None)
    monkeypatch.setattr(main.messagebox, "showerror", lambda *_args: None)

    # A user selects the first and third source pages in the Split dialog.
    monkeypatch.setattr(
        main.filedialog, "asksaveasfilename", lambda **_kwargs: str(split_path)
    )
    with fitz.open(source_path) as source_pdf:
        PDFEditorApp.split_pdf_pages(
            split_app(source_pdf, source_path), "1, 3", RecordingDialog()
        )

    assert page_ids(split_path) == (2, ["alpha", "gamma"])

    # The user then merges that result with the original document.
    monkeypatch.setattr(
        main.filedialog,
        "askopenfilenames",
        lambda **_kwargs: (str(split_path), str(source_path)),
    )
    monkeypatch.setattr(
        main.filedialog, "asksaveasfilename", lambda **_kwargs: str(merged_path)
    )
    PDFEditorApp.merge_pdf_files(merge_app())

    assert merged_path.is_file()
    assert page_ids(merged_path) == (5, ["alpha", "gamma", "alpha", "beta", "gamma"])
