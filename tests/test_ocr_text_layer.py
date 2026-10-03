"""Tests for deciding whether OCR is necessary for a PDF page."""

from pathlib import Path
import sys

import fitz
import pytest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ocr_to_word_worker import (
    MINIMUM_DIGITAL_ALPHANUMERIC_CHARACTERS,
    MINIMUM_DIGITAL_TEXT_CHARACTERS,
    page_has_usable_text_layer,
)


def page_with_text(text):
    """Create an in-memory PDF page with an embedded text layer."""
    document = fitz.open()
    page = document.new_page()
    page.insert_textbox(page.rect + (36, 36, -36, -36), text, fontsize=12)
    return document, page


def test_page_with_sufficient_embedded_text_skips_ocr():
    document, page = page_with_text(
        "This PDF already contains enough selectable embedded text to avoid OCR."
    )
    try:
        assert page_has_usable_text_layer(page) is True
    finally:
        document.close()


@pytest.mark.parametrize(
    "text",
    [
        "7",  # A page number alone must not bypass OCR.
        "A" * (MINIMUM_DIGITAL_TEXT_CHARACTERS - 1),
        "!" * MINIMUM_DIGITAL_TEXT_CHARACTERS,
        "A" * (MINIMUM_DIGITAL_ALPHANUMERIC_CHARACTERS - 1)
        + "!" * (MINIMUM_DIGITAL_TEXT_CHARACTERS - MINIMUM_DIGITAL_ALPHANUMERIC_CHARACTERS + 1),
    ],
)
def test_page_without_a_usable_text_layer_requires_ocr(text):
    document, page = page_with_text(text)
    try:
        assert page_has_usable_text_layer(page) is False
    finally:
        document.close()
