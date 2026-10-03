"""Tests for the OCR render scale chosen from PDF page dimensions."""

from pathlib import Path
import sys
from types import SimpleNamespace

import fitz
import pytest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ocr_to_word_worker import render_scale_for_page


def page_with_size(width, height):
    """Return the minimum page-like object needed by ``render_scale_for_page``."""
    return SimpleNamespace(rect=fitz.Rect(0, 0, width, height))


@pytest.mark.parametrize(
    ("width", "height", "expected_scale"),
    [
        # A small receipt should favour legibility but remain capped.
        (200, 300, 3.0),
        # A normal A4 page keeps its long side near the 2,400-pixel target.
        (595, 842, 2400 / 842),
        # A large page must not request an unnecessarily memory-heavy render.
        (2000, 3000, 1.75),
    ],
)
def test_render_scale_uses_safe_range_for_page_sizes(width, height, expected_scale):
    scale = render_scale_for_page(page_with_size(width, height))

    assert scale == pytest.approx(expected_scale)
    assert 1.75 <= scale <= 3.0


def test_render_scale_is_independent_of_page_orientation():
    portrait_scale = render_scale_for_page(page_with_size(595, 842))
    landscape_scale = render_scale_for_page(page_with_size(842, 595))

    assert landscape_scale == pytest.approx(portrait_scale)


def test_render_scale_handles_zero_dimensions_defensively():
    scale = render_scale_for_page(page_with_size(0, 0))

    assert scale == 3.0
