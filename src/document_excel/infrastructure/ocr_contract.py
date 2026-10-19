"""Engine independent coordinates (x, y, width, height), RGB inputs and 0..1 scores."""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from typing import Protocol

import numpy as np


@dataclass
class OcrDiagnostic:
    executable: str
    version: str
    languages: list[str]
    error: str = ""


@dataclass
class TextBlock:
    text: str
    confidence: float
    bbox: tuple[int, int, int, int]
    line_key: tuple[int, int, int]
    script: str | None = None

    @property
    def bounding_box(self) -> tuple[int, int, int, int]:
        return self.bbox


@dataclass
class OcrResult:
    blocks: list[TextBlock]
    image_id: str
    variant: str
    width: int
    height: int
    engine_name: str = "unknown"
    model_version: str = ""
    timings: dict[str, float | bool] = field(default_factory=dict)

    @property
    def source_image(self) -> str:
        return self.image_id

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines())

    def lines(self) -> list[TextBlock]:
        grouped: dict[tuple[int, int, int], list[TextBlock]] = {}
        for block in self.blocks:
            grouped.setdefault(block.line_key, []).append(block)
        output = []
        for key, words in grouped.items():
            words.sort(key=lambda b: b.bbox[0])
            left = min(b.bbox[0] for b in words)
            top = min(b.bbox[1] for b in words)
            right = max(b.bbox[0] + b.bbox[2] for b in words)
            bottom = max(b.bbox[1] + b.bbox[3] for b in words)
            scripts = {b.script for b in words}
            output.append(
                TextBlock(
                    " ".join(b.text for b in words),
                    sum(b.confidence for b in words) / len(words),
                    (left, top, right - left, bottom - top),
                    key,
                    scripts.pop() if len(scripts) == 1 else "mixed",
                )
            )
        return sorted(output, key=lambda b: (b.bbox[1], b.bbox[0]))


class OcrEngine(Protocol):
    def diagnose(self) -> OcrDiagnostic: ...

    def recognize(
        self, image: np.ndarray, image_id: str, variant: str, cancel: threading.Event
    ) -> OcrResult: ...

    def recognize_identifier_line(
        self, image: np.ndarray, image_id: str, variant: str, cancel: threading.Event
    ) -> OcrResult: ...
