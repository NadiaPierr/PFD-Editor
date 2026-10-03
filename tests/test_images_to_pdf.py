"""Tests for image preparation used by the Images to PDF workflow."""

from io import BytesIO
from pathlib import Path
import sys

import fitz
from PIL import Image
import pytest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from main import PDFEditorApp


def image_to_single_page_pdf(image_path, output_path):
    """Exercise the same PDF-writing operations as the UI workflow."""
    image_bytes, width, height = PDFEditorApp.prepare_image_for_pdf(None, image_path)
    with fitz.open() as document:
        page = document.new_page(width=width, height=height)
        page.insert_image(page.rect, stream=image_bytes)
        document.save(output_path)


@pytest.mark.parametrize(
    ("name", "image", "expected_size"),
    [
        ("photo.jpg", Image.new("RGB", (120, 80), "navy"), (120, 80)),
        ("transparent.png", Image.new("RGBA", (75, 50), (255, 0, 0, 128)), (75, 50)),
    ],
)
def test_images_to_pdf_creates_a_page_with_the_source_dimensions(
    tmp_path, name, image, expected_size
):
    image_path = tmp_path / name
    output_path = tmp_path / "images.pdf"
    image.save(image_path)

    image_to_single_page_pdf(image_path, output_path)

    with fitz.open(output_path) as pdf:
        assert pdf.page_count == 1
        page = pdf[0]
        assert (page.rect.width, page.rect.height) == pytest.approx(expected_size)


def test_prepare_image_for_pdf_flattens_transparency_onto_white(tmp_path):
    image_path = tmp_path / "transparent.png"
    Image.new("RGBA", (20, 20), (255, 0, 0, 0)).save(image_path)

    image_bytes, width, height = PDFEditorApp.prepare_image_for_pdf(None, image_path)

    assert (width, height) == (20, 20)
    with Image.open(BytesIO(image_bytes)) as prepared_image:
        assert prepared_image.mode == "RGB"
        assert prepared_image.getpixel((10, 10)) == (255, 255, 255)


def test_prepare_image_for_pdf_applies_exif_rotation(tmp_path):
    image_path = tmp_path / "rotated.jpg"
    image = Image.new("RGB", (40, 90), "green")
    exif = Image.Exif()
    exif[274] = 6  # Rotate 90 degrees clockwise when displayed.
    image.save(image_path, exif=exif)

    _image_bytes, width, height = PDFEditorApp.prepare_image_for_pdf(None, image_path)

    assert (width, height) == (90, 40)
