"""ICAO 9303 Parts 3/5 TD1, with explicit checks and no guessed OCR repairs.

References: https://www.icao.int/sites/default/files/publications/DocSeries/9303_p5_cons_en.pdf
and 9303_p3_cons_en.pdf. MRZ and VIZ data remain in memory only.
"""

from __future__ import annotations

import re
import threading
from datetime import date
from itertools import product

import cv2
import numpy as np

from document_excel.domain.errors import Cancelled
from document_excel.domain.models import (
    DATE_FIELDS,
    DocumentSide,
    FieldCandidate,
    FieldStatus,
    MRZReport,
    parse_date,
    split_document_identifier,
)
from document_excel.infrastructure.preprocessing import deskew


def check_digit(value: str) -> str:
    total = 0
    for index, character in enumerate(value):
        if character == "<":
            number = 0
        elif "0" <= character <= "9":
            number = ord(character) - ord("0")
        elif "A" <= character <= "Z":
            number = ord(character) - ord("A") + 10
        else:
            raise ValueError("Недопустимый символ MRZ")
        total += number * (7, 3, 1)[index % 3]
    return str(total % 10)


def find_mrz_lines(text: str) -> list[str]:
    lines = [re.sub(r"\s+", "", line.upper()) for line in text.splitlines()]
    for index, line in enumerate(lines):
        if re.match(r"[IAC][A-Z<][A-Z<]{3}", line) and ("<" in line or len(line) >= 25):
            return lines[index : index + 3]
    # Preserve malformed/short candidates for diagnostics, never pad to 30.
    candidates = [line for line in lines if "<" in line and len(line) >= 15]
    return candidates[-3:]


class TD1Parser:
    def parse(self, lines: list[str], image_id: str = "") -> MRZReport:
        report = MRZReport(image_id=image_id, lines=[re.sub(r"\s+", "", s.upper()) for s in lines])
        if len(report.lines) != 3 or any(len(line) != 30 for line in report.lines):
            report.errors.append("Нужны три строки MRZ ровно по 30 символов. Переснимите оборот.")
            return report
        a, b, c = report.lines
        if not all(re.fullmatch(r"[A-Z0-9<]{30}", line) for line in report.lines):
            report.errors.append("В MRZ есть недопустимые символы; автоматической замены O/0 и I/1 нет.")
            return report
        if not re.fullmatch(r"[IAC][A-Z<][A-Z<]{3}", a[:5]):
            report.errors.append("Не распознан заголовок TD1.")
        document = a[5:14]
        document_check = a[14]
        optional = a[15:]
        if document_check == "<":
            # ICAO long document numbers: continuation, check digit, then filler.
            continuation = optional.split("<", 1)[0]
            if len(continuation) < 2 or not continuation[-1].isdigit() or "<" not in optional:
                report.errors.append("Не распознано продолжение длинного номера TD1.")
            else:
                document += continuation[:-1]
                document_check = continuation[-1]
        report.checks = {
            "document_number": document_check.isdigit() and check_digit(document) == document_check,
            "birth_date": b[6].isdigit() and check_digit(b[:6]) == b[6],
            "expiration_date": b[14].isdigit() and check_digit(b[8:14]) == b[14],
            "composite": b[29].isdigit() and check_digit(a[5:30] + b[:7] + b[8:15] + b[18:29]) == b[29],
        }
        if not all(report.checks.values()):
            report.errors.append("Не сошлись контрольные цифры MRZ. Нужна сверка или повторный OCR.")
        report.segment_status = {
            name: "VALID" if valid else "INVALID" for name, valid in report.checks.items()
        }
        if not (
            b[:6].isdigit()
            and b[8:14].isdigit()
            and b[7] in "MF<"
            and re.fullmatch(r"[A-Z<]{3}", b[15:18])
            and re.fullmatch(r"[A-Z<]{30}", c)
        ):
            report.errors.append("Нарушена структура полей TD1.")
        for value in (b[:6], b[8:14]):
            if not value.isdigit() or not any(
                self.possible_date(value, century) for century in (1900, 2000, 2100)
            ):
                report.errors.append("Некорректная календарная дата MRZ.")
        surname, separator, given = c.partition("<<")
        report.values = {
            "document_code": a[:2],
            "issuing_state": a[2:5],
            "document_number": document.rstrip("<"),
            "birth_date": b[:6],
            "expiration_date": b[8:14],
            "gender": b[7],
            "citizenship": b[15:18],
            "surname_latin": surname.replace("<", " ").strip(),
            "first_name_latin": given.replace("<", " ").strip() if separator else "",
        }
        if c[-1] != "<":
            report.warnings.append("Имя MRZ может быть усечено; проверьте полное печатное написание.")
        # KG plastic specimen: issuer KGZ, ID + seven digits, 14-digit PIN + filler.
        # This is deliberately not a universal interpretation of TD1 optional data.
        if (
            a[2:5] == "KGZ"
            and re.fullmatch(r"ID[0-9]{7}", document)
            and re.fullmatch(r"[0-9]{14}<", optional)
            and a[14].isdigit()
        ):
            report.values["personal_identifier"] = optional[:14]
        return report

    @staticmethod
    def possible_date(value: str, century: int) -> date | None:
        try:
            return date(century + int(value[:2]), int(value[2:4]), int(value[4:6]))
        except ValueError:
            return None


class MRZRecognizer:
    def __init__(self, engine):
        self.engine = engine

    @staticmethod
    def glyphs(image, expected=30, check_spacing=True):
        if image.size == 0:
            return [], np.zeros((0, 0), dtype=np.uint8)
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY) if image.ndim == 3 else image
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        count, _, stats, _ = cv2.connectedComponentsWithStats(mask)
        h, w = gray.shape
        boxes = sorted(
            [
                (x, y, width, height)
                for x, y, width, height, area in stats[1:]
                if height > h * 0.35 and width < w * (1.8 / expected) and area > max(3, h * 0.10)
            ],
            key=lambda b: b[0],
        )
        # Overlapping vertical fragments (for example a broken glyph) are joined;
        # empty columns are never invented or padded to thirty characters.
        merged: list[tuple[int, int, int, int]] = []
        for x, y, width, height in boxes:
            if merged and x <= merged[-1][0] + merged[-1][2]:
                a, b, c, d = merged.pop()
                right = max(a + c, x + width)
                bottom = max(b + d, y + height)
                merged.append((a, min(b, y), right - a, bottom - min(b, y)))
            else:
                merged.append((x, y, width, height))
        if len(merged) != expected:
            return [], mask
        centers = [x + width / 2 for x, _, width, _ in merged]
        spacing = np.diff(centers)
        if check_spacing and np.std(spacing) / max(1, np.mean(spacing)) > 0.24:
            return [], mask
        return merged, mask

    @staticmethod
    def is_filler(glyph):
        height, width = glyph.shape
        if min(height, width) < 3:
            return False
        points = []
        spans = []
        for y in range(round(height * 0.12), round(height * 0.88)):
            columns = np.flatnonzero(glyph[y])
            if len(columns):
                points.append((y / height, float(columns.mean()) / width))
                spans.append((columns[-1] - columns[0] + 1) / width)
        if len(points) < height * 0.5 or np.median(spans) > 0.68:
            return False
        ys, xs = np.asarray(points).T
        if np.std(xs) < 0.08:
            return False
        # A filler is a single thin chevron. K/C/L have extra strokes or a
        # different row profile, so a word-level OCR 'K' is not blindly replaced.
        return bool(np.corrcoef(np.abs(ys - 0.5), xs)[0, 1] > 0.90 and xs.max() - xs.min() > 0.35)

    def recognize_glyph_spans(self, crop, image_id, cancel, line_index):
        method = getattr(self.engine, "recognize_mrz_span", None)
        boxes, mask = self.glyphs(crop)
        if not boxes or not callable(method):
            return None
        filler = [self.is_filler(mask[y : y + h, x : x + w]) for x, y, w, h in boxes]
        output = []
        index = 0
        while index < 30:
            if filler[index]:
                output.append("<")
                index += 1
                continue
            end = index + 1
            while end < 30 and not filler[end]:
                end += 1
            prefix = "".join(output)
            if line_index == 0:
                boundaries = [2, 7, 15, 29, 30]
            elif line_index == 1:
                boundaries = [7, 8, 15, 18, 29, 30]
            else:
                boundaries = [30]
            end = min(end, min(value for value in boundaries if value > index))
            x = min(b[0] for b in boxes[index:end])
            y = min(b[1] for b in boxes[index:end])
            right = max(b[0] + b[2] for b in boxes[index:end])
            bottom = max(b[1] + b[3] for b in boxes[index:end])
            if line_index == 2 or line_index == 0 and end <= 7 or line_index == 1 and index in {7, 15}:
                kind = "letters"
            elif line_index == 1 and index in {0, 8, 29}:
                kind = "digits"
            elif line_index == 0 and prefix[2:7] == "KGZID" and index in {7, 15}:
                kind = "digits"
            else:
                kind = "mixed"
            result = method(
                crop[max(0, y - 2) : bottom + 2, max(0, x - 2) : right + 2],
                image_id,
                cancel,
                kind + "-char" if end - index == 1 else kind,
            )
            text = re.sub(r"\s+", "", result.text)
            if len(text) != end - index:
                characters = []
                for a, b, c, d in boxes[index:end]:
                    value = method(
                        crop[max(0, b - 2) : b + d + 2, max(0, a - 2) : a + c + 2],
                        image_id,
                        cancel,
                        kind + "-char",
                    )
                    char = re.sub(r"\s+", "", value.text)
                    if len(char) != 1:
                        break
                    characters.append(char)
                text = "".join(characters)
            if len(text) != end - index or not re.fullmatch(r"[A-Z0-9]+", text):
                return None
            output.append(text)
            index = end
        return "".join(output)

    @staticmethod
    def line_boxes(card):
        gray = cv2.cvtColor(card, cv2.COLOR_RGB2GRAY) if card.ndim == 3 else card
        h, w = gray.shape
        top = round(h * 0.60)
        # Dark glyphs on varying background, grouped by long horizontal baselines.
        blackhat = cv2.morphologyEx(
            gray[top:], cv2.MORPH_BLACKHAT, cv2.getStructuringElement(cv2.MORPH_RECT, (max(9, w // 40), 5))
        )
        mask = cv2.threshold(blackhat, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
        mask = cv2.dilate(mask, np.ones((3, max(13, w // 35)), np.uint8))
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        boxes = [cv2.boundingRect(c) for c in contours]
        boxes = sorted(
            [
                (
                    max(0, x - 5),
                    top + max(0, y - 4),
                    min(w - max(0, x - 5), width + 10),
                    min(h - top - max(0, y - 4), height + 8),
                )
                for x, y, width, height in boxes
                if width > w * 0.68 and 5 < height < h * 0.13
            ],
            key=lambda b: b[1],
        )
        if len(boxes) < 3:
            return []
        boxes = boxes[-3:]
        centers = [y + height / 2 for _, y, _, height in boxes]
        boundaries = [round((a + b) / 2) for a, b in zip(centers, centers[1:], strict=False)]
        return [
            (
                x,
                max(y, boundaries[index - 1] if index else y),
                width,
                min(y + height, boundaries[index] if index < 2 else y + height)
                - max(y, boundaries[index - 1] if index else y),
            )
            for index, (x, y, width, height) in enumerate(boxes)
        ]

    def recognize(self, card: np.ndarray, image_id: str, cancel: threading.Event) -> MRZReport:
        height, width = card.shape[:2]
        top = round(height * 0.57)
        region = card[top:, :]
        gray = cv2.cvtColor(region, cv2.COLOR_RGB2GRAY) if region.ndim == 3 else region
        aligned = deskew(gray)
        contrast = cv2.createCLAHE(clipLimit=2, tileGridSize=(8, 4)).apply(aligned)
        report = MRZReport(image_id=image_id)
        line_method = getattr(self.engine, "recognize_mrz_line", None)
        boxes = self.line_boxes(card)
        if len(boxes) == 3 and callable(line_method):
            # Neural line recognition usually reads clean TD1 rows directly.
            # Validate the original triple before generating any alternatives.
            originals = []
            if getattr(self.engine, "preferred_input", "") == "color":
                for x, y, w, h in boxes:
                    if cancel.is_set():
                        raise Cancelled()
                    originals.append(
                        re.sub(r"\s+", "", line_method(card[y : y + h, x : x + w], image_id, cancel).text)
                    )
                direct = TD1Parser().parse(originals, image_id)
                if direct.valid:
                    direct.source_bbox = (0, top, width, height - top)
                    return direct
            span_lines = [
                self.recognize_glyph_spans(card[y : y + h, x : x + w], image_id, cancel, index)
                for index, (x, y, w, h) in enumerate(boxes)
            ]
            if all(span_lines):
                attempt = TD1Parser().parse([str(line) for line in span_lines], image_id)
                if attempt.valid:
                    attempt.source_bbox = (0, top, width, height - top)
                    return attempt
            alternatives = []
            restored_lines = set()
            for line_index, (x, y, w, h) in enumerate(boxes):
                crop = card[y : y + h, x : x + w]
                local = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY) if crop.ndim == 3 else crop
                resized = cv2.resize(local, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
                smooth = cv2.bilateralFilter(resized, 5, 40, 40)
                enhanced = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(8, 2)).apply(smooth)
                options = []
                spans = span_lines[line_index]
                if spans:
                    options.append(spans)
                if originals:
                    options.append(originals[line_index])
                for pixels in (
                    *(() if originals else (crop,)),
                    enhanced,
                    cv2.threshold(smooth, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1],
                ):
                    value = line_method(pixels, image_id, cancel)
                    options.append(re.sub(r"\s+", "", value.text))
                raw_method = getattr(self.engine, "recognize_mrz_rawline", None)
                if callable(raw_method):
                    value = raw_method(enhanced, image_id, cancel)
                    options.append(re.sub(r"\s+", "", value.text))
                glyph_boxes, glyph_mask = self.glyphs(crop)
                if glyph_boxes:
                    filler_positions = {
                        index
                        for index, (a, b, c, d) in enumerate(glyph_boxes)
                        if self.is_filler(glyph_mask[b : b + d, a : a + c])
                    }
                    reconstructed = []
                    for option in options:
                        if len(option) == 30:
                            candidate = "".join(
                                "<" if index in filler_positions else char
                                for index, char in enumerate(option)
                            )
                            reconstructed.append(candidate)
                            if candidate != option:
                                restored_lines.add(candidate)
                    # The final composite digit is read from its actual glyph,
                    # never calculated and inserted as if it had been observed.
                    if line_index == 1:
                        span_method = getattr(self.engine, "recognize_mrz_span", None)
                        if callable(span_method):
                            a, b, c, d = glyph_boxes[29]
                            control = span_method(
                                crop[max(0, b - 2) : b + d + 2, max(0, a - 2) : a + c + 2],
                                image_id,
                                cancel,
                                "digits-char",
                            ).text.strip()
                            if re.fullmatch("[0-9]", control):
                                reconstructed.extend(value[:29] + control for value in list(reconstructed))
                    options = reconstructed + options
                alternatives.append(list(dict.fromkeys(options)))
            for lines in product(*alternatives):
                attempt = TD1Parser().parse(list(lines), image_id)
                if (
                    attempt.valid
                    or sum(attempt.checks.values()) > sum(report.checks.values())
                    or not report.lines
                ):
                    report = attempt
                if report.valid:
                    break
            if report.valid:
                if any(line in restored_lines for line in report.lines):
                    report.warnings.append(
                        "Заполнители MRZ распознаны по форме символов. Сверьте строки с фотографией."
                    )
                report.source_bbox = (
                    boxes[0][0],
                    boxes[0][1],
                    width - boxes[0][0],
                    boxes[-1][1] + boxes[-1][3] - boxes[0][1],
                )
                return report
        # Distinct MRZ preprocessing, no ordinary-address extraction or dictionaries.
        for array in (contrast, cv2.threshold(contrast, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]):
            if cancel.is_set():
                raise Cancelled()
            result = self.engine.recognize_mrz(array, image_id, cancel)
            attempt = TD1Parser().parse(find_mrz_lines(result.text), image_id)
            if (
                attempt.valid
                or sum(attempt.checks.values()) > sum(report.checks.values())
                or not report.lines
            ):
                report = attempt
            if report.valid:
                break
        # Deskew changes geometry; the entire original MRZ search region is honest evidence.
        report.source_bbox = (0, top, width, height - top)
        return report


def mrz_candidates(report: MRZReport, visual: list[FieldCandidate]) -> list[FieldCandidate]:
    if not report.valid:
        return []
    output = []
    values = report.values
    document = values["document_number"]
    split = split_document_identifier(document)
    fields = {"passport_number": split[1] if split else document}
    if split:
        fields["passport_type"] = split[0]
    for name in ("surname", "first_name"):
        latin = values.get(name + "_latin", "")
        printed_names = [
            c.normalized_value
            for c in visual
            if c.field_name == name
            and c.source_method == "VISUAL_OCR"
            and c.source_side == DocumentSide.FRONT
            and re.fullmatch(r"[A-Z -]+", c.normalized_value)
        ]
        if any(
            latin != value and re.sub(r"[ -]", "", latin) == re.sub(r"[ -]", "", value)
            for value in printed_names
        ):
            # A TD1 filler represents a separator, not a printed hyphen.
            # Do not turn this equivalent spelling into a false conflict.
            continue
        # A possibly truncated MRZ name cannot contradict the complete VIZ name.
        if latin and not any("усечено" in warning for warning in report.warnings):
            fields[name] = latin
    for name in ("personal_identifier", "gender", "citizenship"):
        value = values.get(name, "")
        if value and value != "<":
            if name == "citizenship":
                value = {"KGZ": "КЫРГЫЗ РЕСПУБЛИКАСЫ"}.get(value, value)
            if name == "gender":
                value = {"M": "М", "F": "Ж"}.get(value, value)
            fields[name] = value
    for name in DATE_FIELDS & values.keys():
        raw = values[name]
        printed = []
        for candidate in visual:
            if candidate.field_name == name and candidate.validation_status != FieldStatus.INVALID:
                try:
                    printed.append(parse_date(candidate.normalized_value))
                except ValueError:
                    pass
        matches = {d for d in printed if d.strftime("%y%m%d") == raw}
        if len(matches) == 1:
            fields[name] = next(iter(matches)).isoformat()
        elif printed and not matches:
            report.conflicts.append(name)
        else:
            report.warnings.append(
                f"{name}: столетие MRZ требует сверки с полной датой; год не восстановлен."
            )
    for name, value in fields.items():
        output.append(
            FieldCandidate(
                field_name=name,
                raw_value=value,
                normalized_value=value,
                confidence=0.96,
                source_image_id=report.image_id,
                source_bbox=report.source_bbox,
                extraction_method="MRZ-TD1-checked",
                source_method="MRZ",
                validation_status=FieldStatus.REVIEW,
                variant="mrz",
                source_side=DocumentSide.BACK,
                image_version=report.image_version,
                original_latin=value if name in {"surname", "first_name"} else None,
            )
        )
    return output
