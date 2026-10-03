"""Tests for extracting text from a selected PDF area."""

from pathlib import Path
import sys

import fitz


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from main import PDFEditorApp


def make_text_page():
    """Create a tiny PDF whose words have stable, known positions."""
    document = fitz.open()
    page = document.new_page(width=300, height=150)
    page.insert_text((50, 50), "Alpha Beta")
    page.insert_text((50, 80), "Gamma Delta")
    return document, page


def test_extract_words_from_rect_returns_intersecting_words_in_reading_order():
    document, page = make_text_page()
    try:
        beta = page.search_for("Beta")[0]
        gamma = page.search_for("Gamma")[0]

        # The rectangle begins inside "Beta", exercising the
        # intersection behaviour used while a user drags a selection.
        selection = fitz.Rect(
            beta.x0 + 1,
            beta.y0,
            gamma.x1,
            gamma.y1,
        )

        extracted = PDFEditorApp.extract_words_from_rect(None, page, selection)

        assert extracted == "Beta\nGamma"
    finally:
        document.close()


def test_extract_words_from_rect_returns_empty_text_for_an_empty_area():
    document, page = make_text_page()
    try:
        empty_area = fitz.Rect(200, 110, 280, 140)

        extracted = PDFEditorApp.extract_words_from_rect(None, page, empty_area)

        assert extracted == ""
    finally:
        document.close()
