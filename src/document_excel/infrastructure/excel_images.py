"""Embedded document photos with cell geometry and explicit replacement consent."""

from __future__ import annotations

import hashlib
import io
import math
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, cast
from xml.dom import Node, minidom

from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, SpreadsheetDrawing, TwoCellAnchor
from openpyxl.utils import range_boundaries
from openpyxl.xml.functions import tostring
from PIL import Image, ImageFont

from document_excel.domain.errors import AppError
from document_excel.domain.models import ClientCase, DocumentSide
from document_excel.infrastructure.document_images import selected_image
from document_excel.infrastructure.xlsx_package import (
    MAIN,
    REL,
    PackageEditor,
    child,
    document,
    elements,
    encoded,
    rel_path,
    resolve,
)

DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
ART = "http://schemas.openxmlformats.org/drawingml/2006/main"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
TARGETS = {DocumentSide.FRONT: "A1:E24", DocumentSide.BACK: "F1:K24"}


def fresh_document(namespace, root):
    return minidom.getDOMImplementation().createDocument(namespace, root, None)


def sha(path: Path) -> str:
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError as exc:
        raise AppError("Фотография недоступна. Добавьте её заново.") from exc


def resolve_photo_sheet(book, configured: str) -> str:
    key = " ".join(configured.casefold().split())
    matches = [s for s in book.sheetnames if " ".join(s.casefold().split()) == key]
    if len(matches) != 1:
        raise AppError(
            f"Не найден лист фотографий «{configured}». Укажите его в настройках расположения полей."
        )
    return matches[0]


class SheetGeometry:
    def __init__(self, sheet):
        self.sheet = sheet
        self.mdw = 7
        font = sheet.parent._named_styles[0].font
        names = {
            "calibri": "calibri.ttf",
            "arial": "arial.ttf",
            "times new roman": "times.ttf",
            "courier new": "cour.ttf",
        }
        filename = names.get((font.name or "").casefold())
        if filename:
            path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / filename
            try:
                native = ImageFont.truetype(str(path), (font.sz or 11) * 96 / 72)
                self.mdw = max(1, round(max(native.getlength(str(i)) for i in range(10))))
            except OSError:
                pass

    def column(self, index: int) -> int:
        dimension = next(
            (d for d in self.sheet.column_dimensions.values() if (d.min or 0) <= index + 1 <= (d.max or 0)),
            None,
        )
        if dimension and dimension.hidden:
            return 0
        width = (
            dimension.width if dimension is not None else self.sheet.sheet_format.defaultColWidth or 9.140625
        )
        # SpreadsheetML width already includes its cell padding. Adding another
        # five pixels distorts TwoCellAnchor proportions in native Excel.
        pixels = math.floor(((256 * width + math.floor(128 / self.mdw)) / 256) * self.mdw)
        return pixels * 9525

    def row(self, index: int) -> int:
        dimension = self.sheet.row_dimensions.get(index + 1)
        if dimension and dimension.hidden:
            return 0
        height = (
            dimension.height
            if dimension and dimension.height is not None
            else self.sheet.sheet_format.defaultRowHeight or 15
        )
        return round(height * 12700)

    def position(self, col: int, row: int, x=0, y=0):
        return sum(self.column(i) for i in range(col)) + x, sum(self.row(i) for i in range(row)) + y

    def bounds(self, area: str):
        c1, r1, c2, r2 = range_boundaries(area)
        assert c1 is not None and r1 is not None and c2 is not None and r2 is not None
        return (*self.position(c1 - 1, r1 - 1), *self.position(c2, r2))

    def marker(self, x: int, y: int) -> AnchorMarker:
        col = row = 0
        while col < 16383 and x >= self.column(col):
            x -= self.column(col)
            col += 1
        while row < 1048575 and y >= self.row(row):
            y -= self.row(row)
            row += 1
        return AnchorMarker(col=col, colOff=x, row=row, rowOff=y)

    def contain(self, area: str, width: int, height: int) -> TwoCellAnchor:
        left, top, right, bottom = self.bounds(area)
        padding = 4 * 9525
        available_w, available_h = right - left - 2 * padding, bottom - top - 2 * padding
        if min(available_w, available_h) <= 0:
            raise AppError(f"Диапазон {area} скрыт или слишком мал для фотографии.")
        scale = min(available_w / width, available_h / height)
        w, h = max(1, round(width * scale)), max(1, round(height * scale))
        x, y = left + (right - left - w) // 2, top + (bottom - top - h) // 2
        return TwoCellAnchor(editAs="absolute", _from=self.marker(x, y), to=self.marker(x + w, y + h))

    def anchor_bounds(self, node):
        def marker(name):
            marks = elements(node, name)
            if not marks:
                return None
            values = {
                n.localName: int(n.firstChild.data)
                for n in marks[0].childNodes
                if n.nodeType == Node.ELEMENT_NODE
            }
            return self.position(
                values["col"], values["row"], values.get("colOff", 0), values.get("rowOff", 0)
            )

        start = marker("from")
        if start is None:
            positions = elements(node, "pos")
            if not positions:
                raise AppError("Неизвестная привязка существующего рисунка на листе фотографий.")
            start = (int(positions[0].getAttribute("x")), int(positions[0].getAttribute("y")))
        end = marker("to")
        if end is None:
            sizes = elements(node, "ext")
            if not sizes:
                raise AppError("Не определён размер существующего рисунка.")
            end = (start[0] + int(sizes[0].getAttribute("cx")), start[1] + int(sizes[0].getAttribute("cy")))
        return (*start, *end)


@dataclass
class PhotoInsertion:
    role: DocumentSide
    area: str
    anchor: TwoCellAnchor
    data: bytes = field(repr=False)
    inputs: dict[Path, str] = field(repr=False)
    name: str = ""


@dataclass
class PhotoExportPlan:
    sheet: str
    case_tag: str
    insertions: list[PhotoInsertion]
    remove_ids: set[str] = field(default_factory=set)
    conflicts: list[str] = field(default_factory=list)
    conflict_previews: list[bytes] = field(default_factory=list, repr=False)
    exception_used: bool = False


class ExcelImageEmbedder:
    @staticmethod
    def drawing(editor: PackageEditor, sheet: str):
        path = editor.sheets[sheet]
        doc = document(editor.parts[path])
        refs = doc.getElementsByTagNameNS(MAIN, "drawing")
        rels_path = rel_path(path)
        rels = document(editor.parts[rels_path]) if rels_path in editor.parts else None
        if not refs:
            return None, None
        rid = refs[0].getAttributeNS(REL, "id")
        links = [] if rels is None else elements(rels.documentElement, "Relationship")
        target = next((n for n in links if n.getAttribute("Id") == rid), None)
        if target is None or target.getAttribute("TargetMode") == "External":
            raise AppError("Не удалось прочитать встроенные рисунки листа фотографий.")
        drawing_path = resolve(path, target.getAttribute("Target"))
        return drawing_path, document(editor.parts[drawing_path])

    def plan(self, editor: PackageEditor, book, case: ClientCase, profile) -> PhotoExportPlan | None:
        if not profile.photo_sheet:
            return None
        sheet_name = resolve_photo_sheet(book, profile.photo_sheet)
        sheet = book[sheet_name]
        if sheet.protection.sheet:
            raise AppError("Лист фотографий защищён. Используйте разрешённую копию шаблона.")
        tag = hashlib.sha256(case.id.encode()).hexdigest()[:20]
        geometry = SheetGeometry(sheet)
        insertions = []
        for side, area in TARGETS.items():
            sources = [s for s in case.images if s.role == side]
            if len(sources) > 1:
                raise AppError(
                    "Выбрано несколько фотографий одной стороны. Уточните стороны на шаге загрузки."
                )
            if not sources:
                if profile.photos_required and not (profile.allow_photo_exception and case.photo_exception):
                    raise AppError(
                        "Для договора нужны лицевая и оборотная фотографии. Добавьте их или укажите разрешённое исключение."
                    )
                continue
            source = sources[0]
            inputs = {p: sha(p) for p in {source.path, source.selected_path}}
            array = selected_image(source)
            stream = io.BytesIO()
            Image.fromarray(array).save(stream, format="PNG")
            anchor = geometry.contain(area, array.shape[1], array.shape[0])
            insertions.append(
                PhotoInsertion(
                    side, area, anchor, stream.getvalue(), inputs, f"DocumentExcel|{tag}|{side.value}"
                )
            )
        plan = PhotoExportPlan(sheet_name, tag, insertions)
        plan.exception_used = bool(
            profile.photos_required
            and len(insertions) < 2
            and profile.allow_photo_exception
            and case.photo_exception
        )
        drawing_path, drawing = self.drawing(editor, sheet_name)
        if drawing is not None:
            for node in drawing.documentElement.childNodes:
                if node.nodeType != Node.ELEMENT_NODE:
                    continue
                if node.localName not in {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"}:
                    continue
                bounds = geometry.anchor_bounds(node)
                intersects = any(
                    bounds[0] < b[2] and b[0] < bounds[2] and bounds[1] < b[3] and b[1] < bounds[3]
                    for b in (geometry.bounds(a) for a in TARGETS.values())
                )
                if not intersects:
                    continue
                properties = node.getElementsByTagNameNS(DRAW, "cNvPr")
                if not properties or not node.getElementsByTagNameNS(DRAW, "pic"):
                    raise AppError(
                        "В области фотографий есть диаграмма или другой объект. Переместите его в копии шаблона перед экспортом."
                    )
                name = properties[0].getAttribute("name")
                identifier = properties[0].getAttribute("id")
                if name.startswith("DocumentExcel|"):
                    if not name.startswith(f"DocumentExcel|{tag}|"):
                        raise AppError(
                            "В области фотографий есть снимки другого клиента. Выберите исходный чистый шаблон."
                        )
                    # Missing sides must not retain stale photographs on repeat export.
                    plan.remove_ids.add(identifier)
                elif insertions or profile.photos_required:
                    plan.remove_ids.add(identifier)
                    plan.conflicts.append(f"Рисунок {identifier} · пересекает A1:K24")
                    blips = node.getElementsByTagNameNS(ART, "blip")
                    links_path = rel_path(drawing_path)
                    links = document(editor.parts[links_path]) if links_path in editor.parts else None
                    rid = blips[0].getAttributeNS(REL, "embed") if blips else ""
                    link = (
                        next(
                            (
                                n
                                for n in elements(links.documentElement, "Relationship")
                                if n.getAttribute("Id") == rid
                            ),
                            None,
                        )
                        if links
                        else None
                    )
                    media = (
                        resolve(drawing_path, link.getAttribute("Target"))
                        if link is not None and link.getAttribute("TargetMode") != "External"
                        else ""
                    )
                    plan.conflict_previews.append(editor.parts.get(media, b""))
        return plan

    def apply(self, editor: PackageEditor, plan: PhotoExportPlan, replace_images: bool = False):
        if plan.conflicts and not replace_images:
            raise AppError("Подтвердите замену существующих изображений в предпросмотре Excel.")
        for photo in plan.insertions:
            if any(sha(p) != expected for p, expected in photo.inputs.items()):
                raise AppError("Фотография изменилась после предпросмотра. Повторите проверку.")
        if not plan.insertions and not plan.remove_ids:
            return
        drawing_path, doc = self.drawing(editor, plan.sheet)
        if drawing_path is None:
            drawing_path = editor.unique_part("xl/drawings", "drawing", ".xml")
            doc = fresh_document(DRAW, "wsDr")
            doc.documentElement.setAttribute("xmlns", DRAW)
            path = editor.sheets[plan.sheet]
            worksheet = document(editor.parts[path])
            rp = rel_path(path)
            rels = (
                document(editor.parts[rp])
                if rp in editor.parts
                else fresh_document(PACKAGE_REL, "Relationships")
            )
            rels.documentElement.setAttribute("xmlns", PACKAGE_REL)
            rid = "rIdDocumentExcelPhotos"
            ids = {n.getAttribute("Id") for n in elements(rels.documentElement, "Relationship")}
            while rid in ids:
                rid += "x"
            link = child(rels.documentElement, "Relationship")
            for key, value in {
                "Id": rid,
                "Type": REL + "/drawing",
                "Target": "../drawings/" + Path(drawing_path).name,
            }.items():
                link.setAttribute(key, value)
            rels.documentElement.appendChild(link)
            ref = child(worksheet.documentElement, "drawing")
            ref.setAttribute("xmlns:r", REL)
            ref.setAttributeNS(REL, "r:id", rid)
            root = worksheet.documentElement
            following = next(
                (
                    n
                    for n in root.childNodes
                    if n.nodeType == Node.ELEMENT_NODE
                    and n.localName
                    in {
                        "legacyDrawing",
                        "legacyDrawingHF",
                        "picture",
                        "oleObjects",
                        "controls",
                        "webPublishItems",
                        "tableParts",
                        "extLst",
                    }
                ),
                None,
            )
            root.insertBefore(ref, following)
            editor.parts[path], editor.parts[rp] = encoded(worksheet), encoded(rels)
            override = child(editor.types.documentElement, "Override")
            override.setAttribute("PartName", "/" + drawing_path)
            override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")
            editor.types.documentElement.appendChild(override)
        rp = rel_path(drawing_path)
        rels = (
            document(editor.parts[rp]) if rp in editor.parts else fresh_document(PACKAGE_REL, "Relationships")
        )
        rels.documentElement.setAttribute("xmlns", PACKAGE_REL)
        for node in list(doc.documentElement.childNodes):
            if node.nodeType != Node.ELEMENT_NODE:
                continue
            props = node.getElementsByTagNameNS(DRAW, "cNvPr")
            if props and props[0].getAttribute("id") in plan.remove_ids:
                doc.documentElement.removeChild(node)
        used_relations = {n.getAttributeNS(REL, "embed") for n in doc.getElementsByTagNameNS(ART, "blip")}
        for link in list(elements(rels.documentElement, "Relationship")):
            if link.getAttribute("Type") == REL + "/image" and link.getAttribute("Id") not in used_relations:
                rels.documentElement.removeChild(link)
        max_id = max([int(n.getAttribute("id")) for n in doc.getElementsByTagNameNS(DRAW, "cNvPr")] + [0])
        for index, photo in enumerate(plan.insertions, max_id + 1):
            image = ExcelImage(io.BytesIO(photo.data))
            image.anchor = photo.anchor
            generator = SpreadsheetDrawing()
            generator.images.append(image)
            generated = document(tostring(cast(Any, generator)._write()))
            anchor = next(n for n in generated.documentElement.childNodes if n.nodeType == Node.ELEMENT_NODE)
            anchor.setAttribute("xmlns", DRAW)
            anchor.setAttribute("xmlns:r", REL)
            anchor.setAttribute("xmlns:a", ART)
            properties = anchor.getElementsByTagNameNS(DRAW, "cNvPr")[0]
            properties.setAttribute("id", str(index))
            properties.setAttribute("name", photo.name)
            rid = f"rIdDocumentExcel{index}"
            while rid in {n.getAttribute("Id") for n in elements(rels.documentElement, "Relationship")}:
                rid += "x"
            anchor.getElementsByTagNameNS(ART, "blip")[0].setAttributeNS(REL, "r:embed", rid)
            doc.documentElement.appendChild(doc.importNode(anchor, True))
            media = f"xl/media/document-excel-{plan.case_tag}-{photo.role.value.lower()}.png"
            editor.parts[media] = photo.data
            link = child(rels.documentElement, "Relationship")
            for key, value in {
                "Id": rid,
                "Type": REL + "/image",
                "Target": "../media/" + Path(media).name,
            }.items():
                link.setAttribute(key, value)
            rels.documentElement.appendChild(link)
        editor.parts[drawing_path], editor.parts[rp] = encoded(doc), encoded(rels)
        # Delete only unreferenced media owned by this app/client. Original
        # logos and media shared with another drawing are always preserved.
        for media in list(editor.parts):
            if not media.startswith(f"xl/media/document-excel-{plan.case_tag}-"):
                continue
            filename = Path(media).name.encode()
            if not any(filename in data for name, data in editor.parts.items() if name.endswith(".rels")):
                editor.parts.pop(media)
        if not any(
            n.getAttribute("Extension").casefold() == "png"
            for n in elements(editor.types.documentElement, "Default")
        ):
            node = child(editor.types.documentElement, "Default")
            node.setAttribute("Extension", "png")
            node.setAttribute("ContentType", "image/png")
            editor.types.documentElement.insertBefore(node, editor.types.documentElement.firstChild)

    def verify(self, editor: PackageEditor, plan: PhotoExportPlan):
        drawing_path, doc = self.drawing(editor, plan.sheet)
        names = (
            []
            if doc is None
            else [
                n.getAttribute("name")
                for n in doc.getElementsByTagNameNS(DRAW, "cNvPr")
                if n.getAttribute("name").startswith(f"DocumentExcel|{plan.case_tag}|")
            ]
        )
        if sorted(names) != sorted(p.name for p in plan.insertions):
            raise AppError("Проверка встроенных фотографий не пройдена.")
        for photo in plan.insertions:
            media = f"xl/media/document-excel-{plan.case_tag}-{photo.role.value.lower()}.png"
            if editor.parts.get(media) != photo.data:
                raise AppError("Проверка данных фотографии в Excel не пройдена.")
            prop = next(
                n for n in doc.getElementsByTagNameNS(DRAW, "cNvPr") if n.getAttribute("name") == photo.name
            )
            anchor = prop.parentNode.parentNode.parentNode
            for tag, expected in [("from", photo.anchor._from), ("to", photo.anchor.to)]:
                markers = elements(anchor, tag)
                if len(markers) != 1:
                    raise AppError("Не удалось проверить расположение фотографии.")
                values = {
                    n.localName: int(n.firstChild.data)
                    for n in markers[0].childNodes
                    if n.nodeType == Node.ELEMENT_NODE
                }
                if any(
                    values.get(key, 0) != getattr(expected, key) for key in ("col", "row", "colOff", "rowOff")
                ):
                    raise AppError("Проверка границ фотографии не пройдена.")
            blip = anchor.getElementsByTagNameNS(ART, "blip")[0]
            links = document(editor.parts[rel_path(drawing_path)])
            link = next(
                (
                    n
                    for n in elements(links.documentElement, "Relationship")
                    if n.getAttribute("Id") == blip.getAttributeNS(REL, "embed")
                ),
                None,
            )
            if (
                link is None
                or link.getAttribute("TargetMode") == "External"
                or resolve(drawing_path, link.getAttribute("Target")) != media
            ):
                raise AppError("Фотография не встроена в книгу корректно.")
