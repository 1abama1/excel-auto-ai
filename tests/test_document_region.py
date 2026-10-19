"""Synthetic geometry only: no real identity documents or personal data."""

from __future__ import annotations

import cv2
import numpy as np
from PIL import Image

from document_excel.infrastructure.preprocessing import (
    ImagePreprocessingService,
    PreparedImage,
    Quality,
    crop_document_region,
    detect_document_region,
)


def draw_card(image: np.ndarray, bbox: tuple[int, int, int, int], *, covered: bool = False) -> None:
    x, y, w, h = bbox
    cv2.rectangle(image, (x, y), (x + w - 1, y + h - 1), (196, 223, 232), -1)
    cv2.rectangle(image, (x, y), (x + w - 1, y + h - 1), (130, 154, 165), 2)
    cv2.rectangle(image, (x + 30, y + 90), (x + w // 4, y + h - 60), (150, 160, 170), -1)
    for row in range(7):
        yy = y + 65 + row * (h - 100) // 7
        cv2.putText(
            image, f"SYNTHETIC FIELD {row}", (x + w // 3, yy), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (40, 55, 65), 1
        )
    if covered:
        cv2.rectangle(image, (x + 30, y + h // 2), (x + w - 30, y + h // 2 + 65), (0, 0, 0), -1)


def draw_qr(image: np.ndarray, x: int, y: int, size: int = 300) -> None:
    cells = np.random.default_rng(42).integers(0, 2, (29, 29), dtype=np.uint8) * 255
    qr = cv2.resize(cells, (size, size), interpolation=cv2.INTER_NEAREST)
    image[y : y + size, x : x + size] = np.repeat(qr[:, :, None], 3, axis=2)


def screenshot(*, card: bool = True, covered: bool = False) -> np.ndarray:
    image = np.full((2100, 1080, 3), (242, 244, 247), np.uint8)
    cv2.rectangle(image, (0, 0), (1079, 170), (255, 255, 255), -1)
    cv2.putText(image, "PHONE INTERFACE", (80, 110), cv2.FONT_HERSHEY_SIMPLEX, 1, (60, 60, 60), 2)
    cv2.line(image, (20, 1750), (1060, 1750), (100, 100, 100), 2)
    if card:
        draw_card(image, (70, 360, 940, 593), covered=covered)
    draw_qr(image, 390, 1200)
    return image


def test_mobile_document_region_excludes_phone_chrome_and_external_qr():
    image = screenshot()
    before = image.copy()
    result = detect_document_region(image)
    assert result.detected and result.confidence >= 0.8
    x, y, w, h = result.bbox
    assert abs(x - 70) < 10 and abs(y - 360) < 10
    assert abs(w - 940) < 15 and abs(h - 593) < 15
    assert y + h < 1200
    assert np.array_equal(image, before)
    preview = crop_document_region(image, result)
    assert preview.shape[:2] == (h, w)
    preview[:] = 0
    assert np.array_equal(image, before)


def test_square_qr_is_not_a_document():
    image = screenshot(card=False)
    result = detect_document_region(image)
    assert not result.detected
    assert result.bbox == (0, 0, 1080, 2100)
    assert result.warnings


def test_document_with_large_margins():
    image = np.full((1500, 1900, 3), (65, 70, 80), np.uint8)
    draw_card(image, (440, 410, 1020, 643))
    result = detect_document_region(image)
    assert result.detected
    assert abs(result.bbox[0] - 440) < 10
    assert abs(result.bbox[1] - 410) < 10
    assert result.bbox[2] < 1040 and result.bbox[3] < 660


def test_two_separate_cards_require_manual_region_choice():
    image = np.full((2200, 1200, 3), (240, 240, 240), np.uint8)
    draw_card(image, (130, 250, 940, 593))
    draw_card(image, (130, 1150, 940, 593))
    result = detect_document_region(image)
    assert not result.detected
    assert result.bbox == (0, 0, 1200, 2200)
    assert any("несколько" in warning for warning in result.warnings)


def test_plain_image_falls_back_without_inventing_boundaries():
    image = np.full((900, 1500, 3), 220, np.uint8)
    result = detect_document_region(image)
    assert not result.detected
    assert result.confidence == 0
    assert np.array_equal(crop_document_region(image, result), image)


def test_solid_band_gets_honest_possible_obstruction_warning():
    result = detect_document_region(screenshot(covered=True))
    assert result.detected
    assert any("не перекрыт ли текст" in warning for warning in result.warnings)


def test_regular_document_photo_does_not_claim_occlusion():
    result = detect_document_region(screenshot())
    assert result.detected
    assert not any("перекрыт" in warning for warning in result.warnings)


def test_auto_ocr_uses_crop_and_can_map_boxes_to_original(tmp_path):
    path = tmp_path / "synthetic-screen.png"
    Image.fromarray(screenshot()).save(path)
    source = path.read_bytes()
    result = ImagePreprocessingService().prepare(path)
    assert result.original.shape[:2] == (2100, 1080)
    assert result.document_region is not None and result.document_region.detected
    assert "original" not in result.ocr_variants
    assert "document" in result.ocr_variants
    x, y, width, height = result.document_region.bbox
    for name, image in result.ocr_variants.items():
        assert image.shape[:2] == (height, width)
        assert result.to_original_bbox(name, (10, 20, 100, 30)) == (x + 10, y + 20, 100, 30)
    assert result.to_original_bbox("original", (10, 20, 100, 30)) == (10, 20, 100, 30)
    assert path.read_bytes() == source


def test_explicit_original_mode_retains_full_image_for_roi(tmp_path):
    path = tmp_path / "synthetic-screen.png"
    Image.fromarray(screenshot()).save(path)
    result = ImagePreprocessingService().prepare(path, "original")
    assert tuple(result.ocr_variants) == ("original",)
    assert result.original.shape[:2] == (2100, 1080)
    assert result.to_original_bbox("original", (35, 80, 50, 30)) == (35, 80, 50, 30)


def test_detection_metadata_uses_resized_original_coordinate_space(tmp_path):
    path = tmp_path / "large-synthetic-screen.png"
    large = cv2.resize(screenshot(), (2160, 4200))
    Image.fromarray(large).save(path)
    result = ImagePreprocessingService().prepare(path)
    assert max(result.original.shape[:2]) == 2600
    assert result.document_region is not None and result.document_region.detected
    x, y, width, height = result.document_region.bbox
    assert abs(x - 70 * 2600 / 2100) < 12
    assert abs(y - 360 * 2600 / 2100) < 12
    assert result.variants["document"].shape[:2] == (height, width)


def test_legacy_constructor_and_unknown_geometries_are_safe():
    image = np.ones((300, 400, 3), np.uint8)
    result = PreparedImage(image, {"original": image, "deskewed": image}, Quality(0, 0, 0, []))
    assert result.ocr_variants == result.variants
    assert result.to_original_bbox("original", (1, 2, 3, 4)) == (1, 2, 3, 4)
    assert result.to_original_bbox("deskewed", (1, 2, 3, 4)) is None
    assert result.to_original_bbox("perspective", (1, 2, 3, 4)) is None
    assert result.to_original_bbox("missing", (1, 2, 3, 4)) is None


def test_legacy_text_page_is_not_mistaken_for_one_small_text_field(image_factory):
    result = ImagePreprocessingService().prepare(image_factory())
    assert result.document_region is not None and not result.document_region.detected
    assert "original" in result.ocr_variants


def test_nearly_square_rounded_mobile_identity_layout_excludes_qr():
    # Geometry follows a common mobile layout. All content remains synthetic.
    image = np.full((1280, 576, 3), (247, 248, 250), np.uint8)
    x, y, w, h = (46, 279, 484, 509)
    cv2.rectangle(image, (x + 15, y), (x + w - 16, y + h - 1), (225, 235, 243), -1)
    cv2.rectangle(image, (x, y + 15), (x + w - 1, y + h - 16), (225, 235, 243), -1)
    for cx, cy in ((x + 15, y + 15), (x + w - 16, y + 15), (x + 15, y + h - 16), (x + w - 16, y + h - 16)):
        cv2.circle(image, (cx, cy), 15, (225, 235, 243), -1)
    for row in range(8):
        cv2.putText(
            image,
            f"SYNTHETIC {row}",
            (x + 40, y + 50 + row * 48),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.6,
            (60, 70, 80),
            1,
        )
    draw_qr(image, 220, 1110, 136)
    detection = detect_document_region(image)
    assert detection.detected
    assert abs(detection.bbox[0] - x) < 10 and abs(detection.bbox[1] - y) < 10
    assert abs(detection.bbox[2] - w) < 15 and abs(detection.bbox[3] - h) < 15
    assert detection.bbox[1] + detection.bbox[3] < 900


def test_large_square_qr_is_rejected_even_when_screen_wide():
    image = np.full((1280, 576, 3), 245, np.uint8)
    draw_qr(image, 46, 300, 484)
    detection = detect_document_region(image)
    assert not detection.detected
