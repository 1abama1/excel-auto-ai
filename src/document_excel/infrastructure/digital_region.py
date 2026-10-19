"""Axis-aligned mobile ID crops; colour is used for bounds, never text removal."""

from dataclasses import dataclass
from math import ceil, floor

import cv2
import numpy as np

from document_excel.infrastructure.preprocessing import detect_document_region


@dataclass
class DigitalIDRegion:
    image: np.ndarray
    original_size: tuple[int, int]
    bbox: tuple[int, int, int, int]
    detected: bool
    confidence: float
    mobile_context: bool

    def original_bbox(self, bbox):
        if bbox is None:
            return None
        x, y, width, height = bbox
        left, top, rw, rh = self.bbox
        sx, sy = rw / self.image.shape[1], rh / self.image.shape[0]
        x0, y0 = max(0, floor(left + x * sx + 1e-9)), max(0, floor(top + y * sy + 1e-9))
        x1 = min(self.original_size[0], ceil(left + (x + width) * sx - 1e-9))
        y1 = min(self.original_size[1], ceil(top + (y + height) * sy - 1e-9))
        return x0, y0, max(0, x1 - x0), max(0, y1 - y0)

    def metrics(self):
        return {
            "working_original_size": self.original_size,
            "roi_bbox": self.bbox,
            "inference_size": self.image.shape[:2][::-1],
            "roi_detected": self.detected,
            "roi_confidence": self.confidence,
            "digital_screen": self.mobile_context,
            "scale_x": self.image.shape[1] / self.bbox[2],
            "scale_y": self.image.shape[0] / self.bbox[3],
        }


class DigitalIDRegionDetector:
    def prepare(self, original, *, forced=False, detection=None):
        h, w = original.shape[:2]
        found = detection if detection is not None else detect_document_region(original)
        x, y, rw, rh = found.bbox
        # A QR or a wide physical ID is not enough to select the digital path.
        mobile = bool(
            found.detected
            and found.confidence >= 0.8
            and h > w * 1.45
            and 0.8 <= rw / rh <= 1.3
            and rw >= w * 0.6
            and rh < h * 0.75
            and y > h * 0.05
        )
        if not forced and not mobile:
            return None
        if not found.detected or found.confidence < 0.8:
            x, y, rw, rh = 0, 0, w, h
        pixels = original[y : y + rh, x : x + rw]
        # This helps detection of existing small glyphs, without claiming new
        # source detail. Preserve aspect and never exceed the detector budget.
        scale = min(2.0, 1024 / max(1, rw), 1600 / max(rw, rh))
        scale = max(1.0, scale)
        if scale > 1.05:
            pixels = cv2.resize(pixels, (round(rw * scale), round(rh * scale)), interpolation=cv2.INTER_CUBIC)
        return DigitalIDRegion(
            pixels,
            (w, h),
            (x, y, rw, rh),
            bool(found.detected and found.confidence >= 0.8),
            found.confidence,
            mobile,
        )
