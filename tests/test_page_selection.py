"""Tests for parsing page selections in the Split PDF workflow."""

from pathlib import Path
import sys

import pytest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from main import PDFEditorApp


@pytest.mark.parametrize(
    ("selection", "total_pages", "expected"),
    [
        ("1", 5, [0]),
        ("1, 3, 5-7", 8, [0, 2, 4, 5, 6]),
        ("1, 1, 2-3, 2", 5, [0, 1, 2]),
        (" 2 - 4 , 6 ", 6, [1, 2, 3, 5]),
    ],
)
def test_parse_page_selection_returns_zero_based_unique_pages(
    selection, total_pages, expected
):
    assert PDFEditorApp.parse_page_selection(None, selection, total_pages) == expected


@pytest.mark.parametrize(
    ("selection", "message"),
    [
        ("", "Enter at least one page or page range."),
        ("   ", "Enter at least one page or page range."),
        ("a", "Use page numbers like 1, 3, 5-8."),
        ("1-", "Use page numbers like 1, 3, 5-8."),
        ("4-2", "Page ranges must go from smaller to larger page numbers."),
        ("0", "Pages must be between 1 and 5."),
        ("6", "Pages must be between 1 and 5."),
    ],
)
def test_parse_page_selection_rejects_invalid_input(selection, message):
    with pytest.raises(ValueError, match=message):
        PDFEditorApp.parse_page_selection(None, selection, total_pages=5)
