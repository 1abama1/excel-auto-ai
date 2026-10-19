"""Document photo processing and per-client ownership. Originals are immutable."""

from __future__ import annotations

import hashlib
import shutil
import tempfile
from pathlib import Path

import cv2
import numpy as np

from document_excel.domain.errors import AppError
from document_excel.domain.models import ClientCase, DocumentSide, SourceImage, new_id
from document_excel.infrastructure.preprocessing import detect_document_region, load_image


class DocumentImageProcessor:
    def suggest(self, original: np.ndarray) -> np.ndarray:
        """Suggest a crop/quad; the user must review it before it becomes active."""
        h, w = original.shape[:2]
        scale = min(1, 1300 / max(h, w))
        small = cv2.resize(original, (round(w * scale), round(h * scale)))
        gray = cv2.cvtColor(small, cv2.COLOR_RGB2GRAY)
        edges = cv2.Canny(cv2.GaussianBlur(gray, (5, 5), 0), 40, 120)
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        quads = []
        for contour in contours:
            area = cv2.contourArea(contour)
            if area < small.shape[0] * small.shape[1] * 0.28:
                continue
            polygon = cv2.approxPolyDP(contour, cv2.arcLength(contour, True) * 0.025, True)
            if len(polygon) == 4 and cv2.isContourConvex(polygon):
                quads.append((area, polygon.reshape(4, 2).astype(np.float32) / scale))
        quads.sort(key=lambda item: item[0], reverse=True)
        if quads and (len(quads) == 1 or quads[1][0] < quads[0][0] * 0.65):
            points = quads[0][1]
            ordered = np.array(
                [
                    points[np.argmin(points.sum(1))],
                    points[np.argmin(np.diff(points, axis=1))],
                    points[np.argmax(points.sum(1))],
                    points[np.argmax(np.diff(points, axis=1))],
                ],
                dtype=np.float32,
            )
            width = round(
                max(np.linalg.norm(ordered[1] - ordered[0]), np.linalg.norm(ordered[2] - ordered[3]))
            )
            height = round(
                max(np.linalg.norm(ordered[3] - ordered[0]), np.linalg.norm(ordered[2] - ordered[1]))
            )
            if (
                len(np.unique(ordered, axis=0)) == 4
                and min(width, height) >= 100
                and 1.15 < max(width, height) / min(width, height) < 2.2
            ):
                target = np.array(
                    [[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype=np.float32
                )
                return cv2.warpPerspective(
                    original, cv2.getPerspectiveTransform(ordered, target), (width, height)
                )
        region = detect_document_region(original)
        if region.detected:
            x, y, width, height = region.bbox
            return original[y : y + height, x : x + width].copy()
        return original.copy()

    @staticmethod
    def quality(image: np.ndarray) -> list[str]:
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        sample = cv2.resize(
            gray,
            (
                round(gray.shape[1] * min(1, 1300 / max(gray.shape))),
                round(gray.shape[0] * min(1, 1300 / max(gray.shape))),
            ),
        )
        notes = []
        if cv2.Laplacian(sample, cv2.CV_64F).var() < 45:
            notes.append("Сильное размытие: приблизьте документ, дождитесь фокусировки или переснимите.")
        if sample.mean() < 65:
            notes.append("Недостаточно света: добавьте освещение без бликов.")
        if min(image.shape[:2]) < 600:
            notes.append("Мелкий текст может читаться плохо: увеличьте документ в кадре.")
        if np.mean(sample > 250) > 0.08 and np.std(sample[sample > 235]) < 8:
            notes.append("Возможен блик: измените угол освещения и проверьте, не скрыты ли символы.")
        return notes

    @staticmethod
    def save(image: np.ndarray, path: Path) -> None:
        ok, encoded = cv2.imencode(".png", cv2.cvtColor(image, cv2.COLOR_RGB2BGR))
        if not ok:
            raise AppError("Не удалось подготовить фотографию документа.")
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("xb") as stream:
            stream.write(encoded.tobytes())


class DocumentCaptureController:
    def __init__(self, case_service):
        self.case_service = case_service
        self.storage = tempfile.TemporaryDirectory(prefix="document-excel-capture-")

    def photo(
        self,
        side: DocumentSide,
        processed: np.ndarray,
        *,
        original: np.ndarray | None = None,
        source: SourceImage | None = None,
        retain: Path | None = None,
        case_id: str = "",
    ) -> SourceImage:
        side = DocumentSide(side)
        root = (
            Path(self.storage.name)
            if retain is None
            else retain / ("client-" + hashlib.sha256(case_id.encode()).hexdigest()[:20])
        )
        root.mkdir(parents=True, exist_ok=True)
        token = new_id()
        if source is None:
            original_path = root / f"{token}-original.png"
            if original is None:
                raise AppError("Отсутствует исходный кадр камеры.")
            DocumentImageProcessor.save(original, original_path)
            source = SourceImage(path=original_path, role=side, captured=True)
        elif retain is not None:
            original_path = root / f"{token}-original{source.path.suffix.lower()}"
            shutil.copyfile(source.path, original_path)
            source = source.model_copy(update={"path": original_path})
        processed_path = root / f"{token}-processed.png"
        DocumentImageProcessor.save(processed, processed_path)
        return source.model_copy(
            update={
                "role": side,
                "processed_path": processed_path,
                "processed_confirmed": True,
                "photo_digest": hashlib.sha256(processed_path.read_bytes()).hexdigest(),
            }
        )

    def assign(self, case: ClientCase, source: SourceImage) -> None:
        if source.role not in {DocumentSide.FRONT, DocumentSide.BACK}:
            raise AppError("Выберите лицевую или оборотную сторону.")
        removed = [s for s in case.images if s.role == source.role or s.id == source.id]
        for old in removed:
            self.case_service.remove_evidence(case, old.id)
        case.images = [s for s in case.images if s not in removed] + [source]
        case.invalidate()

    @staticmethod
    def swap(case: ClientCase) -> None:
        for source in case.images:
            if source.role == DocumentSide.FRONT:
                source.role = DocumentSide.BACK
            elif source.role == DocumentSide.BACK:
                source.role = DocumentSide.FRONT
        case.invalidate()

    def cleanup(self) -> None:
        self.storage.cleanup()


def selected_image(source: SourceImage) -> np.ndarray:
    if (
        source.photo_digest
        and hashlib.sha256(source.selected_path.read_bytes()).hexdigest() != source.photo_digest
    ):
        raise AppError("Фотография изменилась после проверки. Добавьте её заново.")
    return load_image(source.selected_path, source.selected_rotation)
