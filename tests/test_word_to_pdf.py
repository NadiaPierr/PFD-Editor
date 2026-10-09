"""Tests for the Word to PDF backend selection and validation."""

from pathlib import Path
import sys

import pytest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import word_to_pdf
from main import PDFEditorApp


class DropTargetStub:
    class TkStub:
        @staticmethod
        def splitlist(data):
            return [data]

    def __init__(self):
        self.tk = self.TkStub()
        self.opened_pdf = None
        self.loaded_word = None
        self.converted_word = None

    def load_pdf(self, path):
        self.opened_pdf = path

    def convert_word_file_to_pdf(self, path):
        self.converted_word = path

    def load_word_document(self, path):
        self.loaded_word = path


def test_convert_word_to_pdf_uses_fallback_backend(tmp_path, monkeypatch):
    source = tmp_path / "sample.docx"
    output = tmp_path / "sample.pdf"
    source.write_bytes(b"word")

    def unavailable(_source, _output):
        raise word_to_pdf.BackendUnavailable("Word unavailable")

    def fallback(_source, target):
        target.write_bytes(b"%PDF-1.7")

    monkeypatch.setattr(word_to_pdf, "_convert_with_microsoft_word", unavailable)
    monkeypatch.setattr(word_to_pdf, "_convert_with_libreoffice", fallback)

    assert word_to_pdf.convert_word_to_pdf(source, output) == output.resolve()
    assert output.read_bytes().startswith(b"%PDF")


def test_convert_word_to_pdf_rejects_unsupported_input(tmp_path):
    source = tmp_path / "notes.txt"
    source.write_text("hello", encoding="utf-8")

    with pytest.raises(ValueError, match=r"\.doc"):
        word_to_pdf.convert_word_to_pdf(source, tmp_path / "notes.pdf")


def test_convert_word_to_pdf_reports_all_unavailable_backends(tmp_path, monkeypatch):
    source = tmp_path / "sample.doc"
    source.write_bytes(b"word")

    def unavailable(_source, _output):
        raise word_to_pdf.BackendUnavailable("not available")

    monkeypatch.setattr(word_to_pdf, "_convert_with_microsoft_word", unavailable)
    monkeypatch.setattr(word_to_pdf, "_convert_with_libreoffice", unavailable)

    with pytest.raises(word_to_pdf.WordToPdfError, match="Microsoft Word or LibreOffice"):
        word_to_pdf.convert_word_to_pdf(source, tmp_path / "sample.pdf")


@pytest.mark.parametrize("extension", [".doc", ".docx", ".DOCX"])
def test_word_file_drop_starts_word_conversion(extension):
    target = DropTargetStub()
    event = type("DropEvent", (), {"data": f"document{extension}"})()

    PDFEditorApp.on_file_drop(target, event)

    assert target.converted_word == Path(f"document{extension}")
    assert target.loaded_word == Path(f"document{extension}")
    assert target.opened_pdf is None


def test_pdf_file_drop_still_opens_pdf():
    target = DropTargetStub()
    event = type("DropEvent", (), {"data": "document.pdf"})()

    PDFEditorApp.on_file_drop(target, event)

    assert target.opened_pdf == Path("document.pdf")
    assert target.loaded_word is None
    assert target.converted_word is None


@pytest.mark.parametrize(
    (
        "document_type",
        "expected_pdf_state",
        "expected_word_state",
        "expected_merge_state",
        "expected_images_state",
    ),
    [
        ("pdf", "normal", "disabled", "normal", "disabled"),
        ("word", "disabled", "normal", "disabled", "disabled"),
        (None, "disabled", "normal", "normal", "normal"),
    ],
)
def test_document_type_controls_relevant_actions(
    document_type,
    expected_pdf_state,
    expected_word_state,
    expected_merge_state,
    expected_images_state,
):
    class ButtonStub:
        def __init__(self):
            self.state = None

        def configure(self, state):
            self.state = state

    target = type(
        "ActionTarget",
        (),
        {
            "convert_button": ButtonStub(),
            "extract_button": ButtonStub(),
            "scan_to_word_button": ButtonStub(),
            "split_button": ButtonStub(),
            "word_to_pdf_button": ButtonStub(),
            "merge_button": ButtonStub(),
            "images_to_pdf_button": ButtonStub(),
        },
    )()

    PDFEditorApp.set_document_action_states(target, document_type)

    assert target.convert_button.state == expected_pdf_state
    assert target.extract_button.state == expected_pdf_state
    assert target.scan_to_word_button.state == expected_pdf_state
    assert target.split_button.state == expected_pdf_state
    assert target.word_to_pdf_button.state == expected_word_state
    assert target.merge_button.state == expected_merge_state
    assert target.images_to_pdf_button.state == expected_images_state
