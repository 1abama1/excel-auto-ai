import threading

import numpy as np
import pytest
from PIL import Image

from document_excel.application.workflows import CaseService
from document_excel.domain.errors import AppError, Cancelled
from document_excel.domain.models import ClientCase
from document_excel.infrastructure.ocr import TesseractEngine
from document_excel.infrastructure.preprocessing import ImagePreprocessingService, load_image
from document_excel.infrastructure.storage import Settings, copy_originals


@pytest.mark.parametrize("mode", ImagePreprocessingService.MODES)
def test_preprocessing_modes_preserve_source(image_factory, mode):
    path = image_factory()
    before = path.read_bytes()
    result = ImagePreprocessingService().prepare(path, mode)
    assert result.variants["original"].shape[:2] == (1100, 1600)
    assert path.read_bytes() == before


def test_exif_orientation(tmp_path):
    path = tmp_path / "orientation.jpg"
    image = Image.new("RGB", (400, 200), "white")
    exif = image.getexif()
    exif[274] = 6
    image.save(path, exif=exif)
    assert load_image(path).shape[:2] == (400, 200)


def test_bad_and_tiny_images(tmp_path):
    bad = tmp_path / "bad.jpg"
    bad.write_bytes(b"not a photo")
    with pytest.raises(AppError, match="повреждён"):
        load_image(bad)
    Image.new("RGB", (50, 50)).save(bad)
    with pytest.raises(AppError, match="маленькое"):
        load_image(bad)


def test_low_contrast_warning(image_factory):
    prepared = ImagePreprocessingService().prepare(image_factory(low_contrast=True))
    assert any("контраст" in text for text in prepared.quality.warnings)


def test_cancel_before_ocr_does_not_start_engine():
    event = threading.Event()
    event.set()
    with pytest.raises(Cancelled):
        TesseractEngine(Settings()).recognize(np.zeros((200, 200), dtype=np.uint8), "x", "original", event)


def test_bad_executable_diagnostic(tmp_path):
    diagnostic = TesseractEngine(Settings(tesseract_path=str(tmp_path / "missing.exe"))).diagnose()
    assert diagnostic.error


def test_case_grouping_move_and_merge(image_factory):
    path = image_factory()
    first, second = ClientCase(), ClientCase()
    service = CaseService()
    service.add_images(first, [path, path])
    assert len(first.images) == 1
    service.move_image(first, second, first.images[0])
    assert not first.images and len(second.images) == 1
    service.merge(first, second)
    assert len(first.images) == 1


def test_copied_images_have_no_sensitive_names(image_factory, tmp_path):
    original = image_factory(name="secret-pin-name.png")
    copied = copy_originals([original], tmp_path / "copies", "case-id")
    assert copied[0].read_bytes() == original.read_bytes()
    assert copied[0].name.startswith("image-")
