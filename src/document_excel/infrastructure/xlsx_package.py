"""Edit XLSX cell parts without round-tripping drawings or other package objects.

openpyxl remains the reader/validator. Original ZIP members are copied unchanged
unless they contain planned cells, new-page metadata or calculation metadata.
No external relationship is followed or refreshed.
"""

from __future__ import annotations

import posixpath
import re
from collections import defaultdict
from datetime import date
from pathlib import Path
from uuid import uuid4
from xml.dom import Node, minidom
from zipfile import ZipFile

from openpyxl.formula import Tokenizer
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import coordinate_to_tuple, quote_sheetname
from openpyxl.utils.datetime import to_excel

from document_excel.domain.errors import AppError

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def elements(parent, name):
    return [n for n in parent.childNodes if n.nodeType == Node.ELEMENT_NODE and n.localName == name]


def document(data):
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise AppError("Небезопасные XML-объявления в Excel.")
    return minidom.parseString(data)


def child(parent, name):
    prefix = parent.prefix + ":" if parent.prefix else ""
    return parent.ownerDocument.createElementNS(parent.namespaceURI, prefix + name)


def encoded(doc):
    return doc.toxml(encoding="utf-8")


def rel_path(part):
    folder, name = posixpath.split(part)
    return f"{folder}/_rels/{name}.rels"


def resolve(part, target):
    result = posixpath.normpath(posixpath.join(posixpath.dirname(part), target.lstrip("/")))
    if target.startswith("/"):
        result = target.lstrip("/")
    if result.startswith("../") or ":" in result:
        raise AppError("Некорректная внутренняя связь Excel.")
    return result


def remap_formula(formula: str, names: dict[str, str]) -> str:
    """Remap sheet operands in cloned pages without touching string literals."""
    tokens = Tokenizer(formula).items
    for token in tokens:
        if token.type == "OPERAND" and token.subtype == "RANGE" and "!" in token.value:
            sheet, address = token.value.rsplit("!", 1)
            original = sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet
            if original in names:
                token.value = quote_sheetname(names[original]) + "!" + address
    return "=" + "".join(t.value for t in tokens)


class PackageEditor:
    def __init__(self, source: Path):
        with ZipFile(source) as archive:
            self.infos = archive.infolist()
            self.parts = {i.filename: archive.read(i) for i in self.infos}
            self.comment = archive.comment
        self.workbook = document(self.parts["xl/workbook.xml"])
        self.rels = document(self.parts["xl/_rels/workbook.xml.rels"])
        self.types = document(self.parts["[Content_Types].xml"])
        relationships = {
            n.getAttribute("Id"): resolve("xl/workbook.xml", n.getAttribute("Target"))
            for n in self.rels.documentElement.childNodes
            if n.nodeType == Node.ELEMENT_NODE
        }
        self.sheet_nodes = self.workbook.getElementsByTagNameNS(MAIN, "sheet")
        self.sheets = {
            n.getAttribute("name"): relationships[n.getAttributeNS(REL, "id")] for n in self.sheet_nodes
        }
        self.styles = document(self.parts["xl/styles.xml"])
        self.styles_changed = False
        self.personal_shared_strings: set[str] = set()

    def unique_part(self, folder, stem, suffix):
        index = 1
        while f"{folder}/{stem}{index}{suffix}" in self.parts:
            index += 1
        return f"{folder}/{stem}{index}{suffix}"

    def override(self, original, target):
        for item in list(self.types.documentElement.childNodes):
            if item.nodeType == Node.ELEMENT_NODE and item.getAttribute("PartName") == "/" + original:
                clone = item.cloneNode(True)
                clone.setAttribute("PartName", "/" + target)
                self.types.documentElement.appendChild(clone)
                return
        raise AppError("В шаблоне отсутствует описание части Excel.")

    def clone_sheet(self, original, name, names):
        source = self.sheets[original]
        target = self.unique_part("xl/worksheets", "sheet", ".xml")
        doc = document(self.parts[source])
        root = doc.documentElement
        # A new page needs its own revision identity and no duplicate VBA code name.
        for attr in list(root.attributes.values()):
            if attr.localName == "uid":
                root.setAttribute(attr.name, "{" + str(uuid4()).upper() + "}")
        for prop in doc.getElementsByTagNameNS(MAIN, "sheetPr"):
            if prop.hasAttribute("codeName"):
                prop.removeAttribute("codeName")
        for view in doc.getElementsByTagNameNS(MAIN, "sheetView"):
            view.setAttribute("tabSelected", "0")
        for formula in doc.getElementsByTagNameNS(MAIN, "f"):
            text = "".join(n.data for n in formula.childNodes if n.nodeType == Node.TEXT_NODE)
            if text:
                for node in list(formula.childNodes):
                    formula.removeChild(node)
                formula.appendChild(doc.createTextNode(remap_formula("=" + text, names)[1:]))
        self.parts[target] = encoded(doc)
        source_rels = rel_path(source)
        if source_rels in self.parts:
            rels = document(self.parts[source_rels])
            for item in elements(rels.documentElement, "Relationship"):
                kind = item.getAttribute("Type").rsplit("/", 1)[-1]
                if item.getAttribute("TargetMode") == "External" or kind == "printerSettings":
                    continue
                if kind != "drawing":
                    raise AppError(
                        "Копирование страницы с таблицами, комментариями или особыми связями не поддерживается."
                    )
                drawing = resolve(source, item.getAttribute("Target"))
                drawing_copy = self.unique_part("xl/drawings", "drawing", ".xml")
                drawing_rels = rel_path(drawing)
                if drawing_rels in self.parts:
                    links = document(self.parts[drawing_rels])
                    if any(
                        n.getAttribute("Type").rsplit("/", 1)[-1] != "image"
                        for n in elements(links.documentElement, "Relationship")
                    ):
                        raise AppError(
                            "Дополнительная страница с диаграммами или особыми рисунками не поддерживается."
                        )
                    self.parts[rel_path(drawing_copy)] = self.parts[drawing_rels]
                self.parts[drawing_copy] = self.parts[drawing]
                self.override(drawing, drawing_copy)
                item.setAttribute("Target", posixpath.relpath(drawing_copy, posixpath.dirname(target)))
            self.parts[rel_path(target)] = encoded(rels)
        self.override(source, target)
        parent = self.sheet_nodes[0].parentNode
        node = self.sheet_nodes[list(self.sheets).index(original)].cloneNode(True)
        next_id = max(int(n.getAttribute("sheetId")) for n in elements(parent, "sheet")) + 1
        ids = {n.getAttribute("Id") for n in elements(self.rels.documentElement, "Relationship")}
        rid = f"rIdDocumentExcel{next_id}"
        while rid in ids:
            rid += "x"
        node.setAttribute("name", name)
        node.setAttribute("sheetId", str(next_id))
        node.setAttributeNS(REL, node.getAttributeNodeNS(REL, "id").name, rid)
        parent.appendChild(node)
        relation = child(self.rels.documentElement, "Relationship")
        relation.setAttribute("Id", rid)
        relation.setAttribute("Type", REL + "/worksheet")
        relation.setAttribute("Target", posixpath.relpath(target, "xl"))
        self.rels.documentElement.appendChild(relation)
        original_index = list(self.sheets).index(original)
        for defined in list(self.workbook.getElementsByTagNameNS(MAIN, "definedName")):
            if defined.getAttribute("localSheetId") == str(original_index):
                clone = defined.cloneNode(True)
                clone.setAttribute("localSheetId", str(len(self.sheets)))
                for text in clone.childNodes:
                    if text.nodeType == Node.TEXT_NODE:
                        text.data = remap_formula("=" + text.data, names)[1:]
                defined.parentNode.appendChild(clone)
        self.sheets[name] = target

    def date_style(self, cell):
        if is_date_format(cell.number_format):
            return cell.style_id
        xfs = self.styles.getElementsByTagNameNS(MAIN, "cellXfs")[0]
        items = elements(xfs, "xf")
        new = items[cell.style_id].cloneNode(True)
        new.setAttribute("numFmtId", "14")
        new.setAttribute("applyNumberFormat", "1")
        xfs.appendChild(new)
        xfs.setAttribute("count", str(len(items) + 1))
        self.styles_changed = True
        return len(items)

    def cells(self, sheet, changes, book):
        path = self.sheets[sheet]
        doc = document(self.parts[path])
        data = doc.getElementsByTagNameNS(MAIN, "sheetData")[0]
        rows = {int(n.getAttribute("r")): n for n in elements(data, "row")}
        cells = {n.getAttribute("r"): n for row in rows.values() for n in elements(row, "c")}
        # Shared/array formulas must not be partly replaced; their dependants need
        # a dedicated formula-aware profile rather than a silently broken group.
        targets = {c.cell for c in changes}
        groups = defaultdict(set)
        for address, cell in cells.items():
            for formula in elements(cell, "f"):
                if formula.getAttribute("t") == "shared":
                    groups[formula.getAttribute("si")].add(address)
                elif formula.getAttribute("ref") and address in targets:
                    raise AppError("Нельзя частично заменить формулу массива Excel.")
        if any(group & targets and not group <= targets for group in groups.values()):
            raise AppError("Нельзя частично заменить общую формулу Excel.")
        for change in changes:
            cell = cells.get(change.cell)
            if cell is not None and cell.getAttribute("t") == "s" and getattr(change, "field_name", ""):
                for value in elements(cell, "v"):
                    self.personal_shared_strings.add(
                        "".join(n.data for n in value.childNodes if n.nodeType == Node.TEXT_NODE)
                    )
            row_number, column = coordinate_to_tuple(change.cell)
            if row_number not in rows:
                row = child(data, "row")
                row.setAttribute("r", str(row_number))
                following = next((n for r, n in sorted(rows.items()) if r > row_number), None)
                data.insertBefore(row, following)
                rows[row_number] = row
            row = rows[row_number]
            if cell is None:
                cell = child(row, "c")
                cell.setAttribute("r", change.cell)
                following = next(
                    (n for n in elements(row, "c") if coordinate_to_tuple(n.getAttribute("r"))[1] > column),
                    None,
                )
                row.insertBefore(cell, following)
            for node in list(cell.childNodes):
                if node.nodeType == Node.ELEMENT_NODE and node.localName in {"f", "v", "is"}:
                    cell.removeChild(node)
            if cell.hasAttribute("t"):
                cell.removeAttribute("t")
            value = change.new
            if value is None or value == "":
                continue
            if isinstance(value, str) and not change.formula:
                if len(value) > 32767 or re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", value):
                    raise AppError("Текст для Excel слишком длинный или содержит недопустимые символы.")
                cell.setAttribute("t", "inlineStr")
                inline = child(cell, "is")
                text = child(cell, "t")
                text.setAttribute("xml:space", "preserve")
                text.appendChild(doc.createTextNode(value))
                inline.appendChild(text)
                cell.insertBefore(inline, cell.firstChild)
                continue
            if isinstance(value, date):
                cell.setAttribute("s", str(self.date_style(book[sheet][change.cell])))
                value = to_excel(value, book.epoch)
            node = child(cell, "f" if change.formula else "v")
            node.appendChild(doc.createTextNode(str(value)[1:] if change.formula else str(value)))
            cell.insertBefore(node, cell.firstChild)
        self.parts[path] = encoded(doc)

    def clear_formula_caches(self, targets):
        """Drop stale identity caches in mapped cells; keep formula text and all other cells."""
        by_sheet: dict[str, set[str]] = {}
        for sheet, cell in targets:
            by_sheet.setdefault(sheet, set()).add(cell)
        for sheet, cells in by_sheet.items():
            path = self.sheets[sheet]
            doc = document(self.parts[path])
            for cell in doc.getElementsByTagNameNS(MAIN, "c"):
                if cell.getAttribute("r") in cells and elements(cell, "f"):
                    for value in list(elements(cell, "v")):
                        cell.removeChild(value)
            self.parts[path] = encoded(doc)

    def scrub_personal_shared_strings(self):
        """Clear obsolete personal entries without renumbering shared strings.

        Entries still referenced anywhere in the workbook remain unchanged;
        unrelated strings, formulas and their indices are preserved.
        """
        path = "xl/sharedStrings.xml"
        if not self.personal_shared_strings or path not in self.parts:
            return
        used = set()
        references = 0
        for worksheet in self.sheets.values():
            doc = document(self.parts[worksheet])
            for cell in doc.getElementsByTagNameNS(MAIN, "c"):
                if cell.getAttribute("t") == "s":
                    references += 1
                    for value in elements(cell, "v"):
                        used.add("".join(n.data for n in value.childNodes if n.nodeType == Node.TEXT_NODE))
        obsolete = self.personal_shared_strings - used
        if not obsolete:
            return
        doc = document(self.parts[path])
        doc.documentElement.setAttribute("count", str(references))
        for index, entry in enumerate(doc.getElementsByTagNameNS(MAIN, "si")):
            if str(index) in obsolete:
                for node in list(entry.childNodes):
                    entry.removeChild(node)
                entry.appendChild(child(entry, "t"))
        self.parts[path] = encoded(doc)

    def save(self, destination):
        # Excel rebuilds dependencies after edited formulas and new pages. Keep
        # original formulas/caches elsewhere, including external-link metadata.
        for relation in list(elements(self.rels.documentElement, "Relationship")):
            if relation.getAttribute("Type").endswith("/calcChain"):
                path = resolve("xl/workbook.xml", relation.getAttribute("Target"))
                self.parts.pop(path, None)
                relation.parentNode.removeChild(relation)
                for item in list(elements(self.types.documentElement, "Override")):
                    if item.getAttribute("PartName") == "/" + path:
                        item.parentNode.removeChild(item)
        calc = self.workbook.getElementsByTagNameNS(MAIN, "calcPr")
        if calc:
            calc[0].setAttribute("fullCalcOnLoad", "1")
            calc[0].setAttribute("forceFullCalc", "1")
        else:
            node = child(self.workbook.documentElement, "calcPr")
            node.setAttribute("fullCalcOnLoad", "1")
            root = self.workbook.documentElement
            following = next(
                (
                    n
                    for n in root.childNodes
                    if n.nodeType == Node.ELEMENT_NODE
                    and n.localName
                    in {
                        "oleSize",
                        "custWorkbookViews",
                        "pivotCaches",
                        "smartTagPr",
                        "smartTagTypes",
                        "webPublishing",
                        "fileRecoveryPr",
                        "webPublishObjects",
                        "extLst",
                    }
                ),
                None,
            )
            root.insertBefore(node, following)
        self.parts["xl/workbook.xml"] = encoded(self.workbook)
        self.parts["xl/_rels/workbook.xml.rels"] = encoded(self.rels)
        self.parts["[Content_Types].xml"] = encoded(self.types)
        if self.styles_changed:
            self.parts["xl/styles.xml"] = encoded(self.styles)
        with ZipFile(destination, "w") as archive:
            archive.comment = self.comment
            written = set()
            for info in self.infos:
                if info.filename in self.parts:
                    archive.writestr(info, self.parts[info.filename])
                    written.add(info.filename)
            for name in self.parts.keys() - written:
                archive.writestr(name, self.parts[name], compress_type=8)
        with ZipFile(destination) as archive:
            if archive.testzip() is not None or any(
                archive.read(n) != data for n, data in self.parts.items()
            ):
                raise AppError("Проверка сохранности частей XLSX не пройдена.")
