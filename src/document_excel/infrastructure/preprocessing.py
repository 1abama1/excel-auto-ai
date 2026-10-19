from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageOps, UnidentifiedImageError

from document_excel.domain.errors import AppError

MAX_IMAGE_BYTES = 40 * 1024 * 1024
MAX_PIXELS = 32_000_000

BoundingBox = tuple[int, int, int, int]


@dataclass
class DocumentRegionDetection:
    """A conservative geometric proposal in the supplied image's coordinates.

    ``bbox`` uses x, y, width, height. A failed/ambiguous proposal contains the
    entire image and ``detected=False``; confidence is a geometric score, never
    a claim that all text is visible or correct. No perspective warp is applied.
    """

    bbox: BoundingBox
    confidence: float
    warnings: list[str] = field(default_factory=list)
    detected: bool = False


def _intersection_fraction(first: BoundingBox, second: BoundingBox) -> float:
    x, y, w, h = first
    a, b, c, d = second
    intersection = max(0, min(x + w, a + c) - max(x, a)) * max(0, min(y + h, b + d) - max(y, b))
    return intersection / max(1, min(w * h, c * d))


def _obstruction_warning(image: np.ndarray) -> list[str]:
    """Flag a broad solid dark band, without pretending to detect covered text.

    A phone notification/redaction can look like this; photographs, normal text
    and pale blank regions are deliberately not treated as detected occlusion.
    Other obstruction types require the employee's visual review.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY) if image.ndim == 3 else image
    h, w = gray.shape
    if min(h, w) < 100:
        return []
    mask = (gray < 28).astype(np.uint8)
    count, _, stats, _ = cv2.connectedComponentsWithStats(mask)
    for index in range(1, count):
        x, y, width, height, area = stats[index]
        if (
            width > w * 0.65
            and 0.04 * h < height < 0.3 * h
            and area / max(1, width * height) > 0.94
            and x > w * 0.02
            and x + width < w * 0.98
            and y > h * 0.03
            and y + height < h * 0.97
        ):
            return ["На документе есть сплошная тёмная область. Проверьте, не перекрыт ли текст."]
    return []


def detect_document_region(image: np.ndarray) -> DocumentRegionDetection:
    """Find an ID-shaped rectangular region and avoid isolated square QR codes.

    Analysis is limited to 1300 pixels on the long side, then mapped back into
    the *input* coordinate system. The input is never changed. Multiple equally
    plausible separate cards and weak/no boundaries fall back to the full image.
    This is an image-geometry heuristic; it does not verify document identity.
    """
    h, w = image.shape[:2]
    full = (0, 0, w, h)
    fallback = DocumentRegionDetection(
        full,
        0.0,
        [
            "Границы документа не определены уверенно. Показан оригинал; при необходимости выделите область вручную."
        ],
    )
    if min(h, w) < 100:
        return fallback
    scale = min(1.0, 1300 / max(h, w))
    sample = cv2.resize(image, (round(w * scale), round(h * scale))) if scale < 1 else image
    gray = cv2.cvtColor(sample, cv2.COLOR_RGB2GRAY) if sample.ndim == 3 else sample
    sh, sw = gray.shape
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edge_mask = cv2.Canny(blurred, 25, 90)
    edge_mask = cv2.morphologyEx(edge_mask, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))
    masks = [edge_mask]
    # Colour boundaries can be strong even when a pastel card and the phone's
    # white background have almost the same grayscale luminance.
    if sample.ndim == 3:
        border = np.concatenate((sample[0], sample[-1], sample[:, 0], sample[:, -1]))
        background = np.median(border, axis=0)
        separation = np.max(np.abs(sample.astype(np.float32) - background), axis=2)
        masks.append(
            cv2.morphologyEx(
                (separation > 14).astype(np.uint8) * 255, cv2.MORPH_CLOSE, np.ones((7, 7), np.uint8)
            )
        )
    candidates: list[tuple[float, BoundingBox]] = []
    for mask in masks:
        contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        for contour in sorted(contours, key=cv2.contourArea, reverse=True)[:80]:
            area = float(cv2.contourArea(contour))
            if area < gray.size * 0.055:
                continue
            x, y, width, height = cv2.boundingRect(contour)
            if min(width, height) < min(sw, sh) * 0.16:
                continue
            ratio = max(width, height) / min(width, height)
            rectangularity = area / (width * height)
            card_aspect = 1.28 <= ratio <= 1.95
            # Digital ID apps can arrange the same data in a large, nearly
            # square rounded card instead of the physical ID-1 proportions.
            # A QR is also square: require a substantial screen-wide region,
            # then reject binary black/white, dense code-like content below.
            mobile_card = (
                sh > sw * 1.3
                and 0.85 <= width / height <= 1.28
                and width >= sw * 0.62
                and height <= sh * 0.65
                and area >= gray.size * 0.15
            )
            if (not card_aspect and not mobile_card) or rectangularity < 0.78:
                continue
            polygon = cv2.approxPolyDP(contour, 0.025 * cv2.arcLength(contour, True), True)
            if not 4 <= len(polygon) <= 8:
                continue
            # A phone's outer panel is also rectangular. Do not turn that panel
            # (including the QR below the card) into a portrait "document".
            if sh > sw * 1.3 and height > width and height > sh * 0.6 and width > sw * 0.75:
                continue
            if mobile_card:
                content = gray[y : y + height, x : x + width]
                black = float(np.mean(content < 65))
                white = float(np.mean(content > 220))
                if black > 0.25 and white > 0.25:
                    continue
            aspect_score = 0.8 if mobile_card else max(0.0, 1 - abs(ratio - 1.586) / 0.4)
            center_score = max(0.0, 1 - abs(x + width / 2 - sw / 2) / (sw / 2))
            area_score = min(1.0, float(np.sqrt(area / (gray.size * 0.3))))
            score = 0.48 + 0.2 * min(1.0, rectangularity) + 0.13 * aspect_score
            score += 0.12 * area_score + 0.07 * center_score
            # One pixel of breathing room retains the anti-aliased border and
            # never cuts through a word sitting close to the detected edge.
            left, top = max(0, round(x / scale) - 2), max(0, round(y / scale) - 2)
            right, bottom = min(w, round((x + width) / scale) + 2), min(h, round((y + height) / scale) + 2)
            bbox = (left, top, right - left, bottom - top)
            if any(_intersection_fraction(bbox, previous) > 0.8 for _, previous in candidates):
                continue
            candidates.append((min(1.0, score), bbox))
    candidates.sort(reverse=True)
    if not candidates or candidates[0][0] < 0.8:
        return fallback
    score, bbox = candidates[0]
    for other_score, other in candidates[1:]:
        if other_score >= score - 0.08 and _intersection_fraction(bbox, other) < 0.2:
            return DocumentRegionDetection(
                full,
                0.0,
                ["Найдено несколько похожих областей. Показан оригинал; выберите нужный документ вручную."],
            )
    crop = crop_document_region(image, DocumentRegionDetection(bbox, score, detected=True))
    notes = _obstruction_warning(crop)
    x, y, width, height = bbox
    if (x <= 2 or y <= 2 or x + width >= w - 2 or y + height >= h - 2) and width * height < w * h * 0.95:
        notes.append("Документ близко к краю снимка. Проверьте, что все поля видны целиком.")
    return DocumentRegionDetection(bbox, score, notes, True)


def crop_document_region(image: np.ndarray, detection: DocumentRegionDetection | None = None) -> np.ndarray:
    """Return a preview/OCR copy, retaining the original image unchanged."""
    region = detection or detect_document_region(image)
    x, y, width, height = region.bbox
    return image[y : y + height, x : x + width].copy()


@dataclass
class Quality:
    blur: float
    brightness: float
    contrast: float
    warnings: list[str]


@dataclass
class PreparedImage:
    original: np.ndarray
    variants: dict[str, np.ndarray]
    quality: Quality
    document_region: DocumentRegionDetection | None = None
    variant_offsets: dict[str, tuple[int, int] | None] = field(default_factory=dict)
    ocr_variant_names: tuple[str, ...] = ()
    source_size: tuple[int, int] | None = None

    @property
    def ocr_variants(self) -> dict[str, np.ndarray]:
        """Only the intended OCR inputs; ``variants`` also retains the original."""
        if not self.ocr_variant_names:
            return self.variants
        return {name: self.variants[name] for name in self.ocr_variant_names}

    def to_original_bbox(self, variant: str, bbox: BoundingBox) -> BoundingBox | None:
        """Map an axis crop's OCR box to the max-2600-pixel original.

        Deskew and perspective variants have no safe translation-only mapping.
        Returning None prevents misleading evidence overlays in those modes.
        """
        if variant not in self.variants:
            return None
        offset = self.variant_offsets.get(variant, None if variant in {"perspective", "deskewed"} else (0, 0))
        if offset is None:
            return None
        x, y, width, height = bbox
        return x + offset[0], y + offset[1], width, height


def load_image(path: Path, rotation: int = 0) -> np.ndarray:
    try:
        if path.stat().st_size > MAX_IMAGE_BYTES:
            raise AppError("Изображение больше 40 МБ. Уменьшите размер файла.")
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(path) as image:
                if image.width * image.height > MAX_PIXELS:
                    raise AppError("Изображение больше 32 мегапикселей. Уменьшите разрешение.")
                if min(image.size) < 100:
                    raise AppError("Изображение слишком маленькое. Приблизьте документ и переснимите.")
                oriented = ImageOps.exif_transpose(image).convert("RGB")
                if rotation:
                    oriented = oriented.rotate(-rotation, expand=True)
                return np.array(oriented)
    except AppError:
        raise
    except (
        OSError,
        ValueError,
        UnidentifiedImageError,
        Image.DecompressionBombError,
        Image.DecompressionBombWarning,
    ) as exc:
        raise AppError("Не удалось открыть изображение. Файл повреждён или недоступен.") from exc


def assess_quality(image: np.ndarray) -> Quality:
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    sample = cv2.resize(
        gray, (min(1000, gray.shape[1]), max(1, int(gray.shape[0] * min(1, 1000 / gray.shape[1]))))
    )
    blur = float(cv2.Laplacian(sample, cv2.CV_64F).var())
    brightness, contrast = float(sample.mean()), float(sample.std())
    notes: list[str] = []
    if min(image.shape[:2]) < 600:
        notes.append("Приблизьте документ: мало пикселей для мелкого текста.")
    if blur < 45:
        notes.append("Возможно размытие: удерживайте камеру неподвижно.")
    if brightness < 65:
        notes.append("Повысите равномерное освещение.")
    if contrast < 28:
        notes.append("Низкий контраст: устраните тени и блики.")
    if float(np.mean(sample > 250)) > 0.3 and contrast < 35:
        notes.append("Возможен блик: переснимите без отражения лампы.")
    return Quality(blur, brightness, contrast, notes)


def order_corners(points: np.ndarray) -> np.ndarray:
    points = points.reshape(4, 2).astype(np.float32)
    sums = points.sum(axis=1)
    differences = np.diff(points, axis=1).ravel()
    return np.array(
        [
            points[np.argmin(sums)],
            points[np.argmin(differences)],
            points[np.argmax(sums)],
            points[np.argmax(differences)],
        ],
        dtype=np.float32,
    )


def perspective_correct(image: np.ndarray) -> np.ndarray | None:
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    edges = cv2.Canny(cv2.GaussianBlur(gray, (5, 5), 0), 50, 150)
    edges = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8))
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for contour in sorted(contours, key=cv2.contourArea, reverse=True)[:5]:
        area = cv2.contourArea(contour)
        if not 0.35 * gray.size < area < 0.97 * gray.size:
            continue
        polygon = cv2.approxPolyDP(contour, 0.02 * cv2.arcLength(contour, True), True)
        if len(polygon) != 4 or not cv2.isContourConvex(polygon):
            continue
        corners = order_corners(polygon)
        width = int(max(np.linalg.norm(corners[1] - corners[0]), np.linalg.norm(corners[2] - corners[3])))
        height = int(max(np.linalg.norm(corners[3] - corners[0]), np.linalg.norm(corners[2] - corners[1])))
        if min(width, height) < 200:
            continue
        dest = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], np.float32)
        return cv2.warpPerspective(image, cv2.getPerspectiveTransform(corners, dest), (width, height))
    return None


def deskew(image: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY) if image.ndim == 3 else image
    mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    lines = cv2.HoughLinesP(
        mask, 1, np.pi / 180, 80, minLineLength=max(50, gray.shape[1] // 5), maxLineGap=20
    )
    angles = []
    if lines is not None:
        for x1, y1, x2, y2 in lines[:, 0]:
            angle = float(np.degrees(np.arctan2(y2 - y1, x2 - x1)))
            if abs(angle) < 12:
                angles.append(angle)
    angle = float(np.median(angles)) if angles else 0
    if abs(angle) < 0.4:
        return image.copy()
    h, w = image.shape[:2]
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1)
    return cv2.warpAffine(image, matrix, (w, h), borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))


class ImagePreprocessingService:
    MODES = (
        "auto",
        "original",
        "grayscale",
        "adaptive",
        "denoised",
        "sharpened",
        "contrast",
        "deskewed",
        "perspective",
    )

    def prepare(self, path: Path, mode: str = "auto", rotation: int = 0) -> PreparedImage:
        original = load_image(path, rotation)
        # The viewer and evidence use this resized, EXIF/rotation-aware original.
        # Cropping affects OCR variants only and never touches the source file.
        h, w = original.shape[:2]
        source_size = (w, h)
        if max(h, w) > 2600:
            original = cv2.resize(original, (round(w * 2600 / max(h, w)), round(h * 2600 / max(h, w))))
        detection = detect_document_region(original)
        working = (
            crop_document_region(original, detection)
            if detection.detected and mode != "original"
            else original
        )
        quality = assess_quality(working)
        quality.warnings.extend(detection.warnings)
        gray = cv2.cvtColor(working, cv2.COLOR_RGB2GRAY)
        variants = {"original": original}
        offsets: dict[str, tuple[int, int] | None] = {"original": (0, 0)}
        offset = detection.bbox[:2] if detection.detected and mode != "original" else (0, 0)
        if working is not original:
            variants["document"] = working
            offsets["document"] = offset
        modes = [mode] if mode != "auto" else ["contrast" if quality.contrast < 45 else "grayscale"]
        for selected in modes:
            if selected == "grayscale":
                variants[selected] = gray
            elif selected == "adaptive":
                variants[selected] = cv2.adaptiveThreshold(
                    gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 12
                )
            elif selected == "denoised":
                variants[selected] = cv2.bilateralFilter(gray, 7, 50, 50)
            elif selected == "sharpened":
                variants[selected] = cv2.addWeighted(gray, 1.6, cv2.GaussianBlur(gray, (0, 0), 2), -0.6, 0)
            elif selected == "contrast":
                variants[selected] = cv2.createCLAHE(clipLimit=2, tileGridSize=(8, 8)).apply(gray)
            elif selected == "deskewed":
                variants[selected] = deskew(working)
            elif selected == "perspective":
                corrected = perspective_correct(working)
                if corrected is not None:
                    variants[selected] = corrected
                else:
                    quality.warnings.append(
                        "Границы документа не определены. Разместите его целиком на контрастном фоне."
                    )
            if selected in variants:
                offsets[selected] = None if selected in {"deskewed", "perspective"} else offset
        ocr_names = tuple(name for name in variants if name != "original" or working is original)
        return PreparedImage(original, variants, quality, detection, offsets, ocr_names, source_size)
