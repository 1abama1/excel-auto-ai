"""Filled contracts are copied, never regenerated. All identities are synthetic."""

from datetime import date
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image

from document_excel.domain.errors import AppError, OutputExistsError
from document_excel.domain.models import (
    Address,
    ClientCase,
    ClientData,
    DocumentSide,
    FieldEvidence,
    FieldStatus,
    MissingValuePolicy,
    RegistrationStatus,
    SourceImage,
    ToolItem,
    WorkflowMode,
)
from document_excel.domain.person_name import PersonNameFormatter
from document_excel.infrastructure.contract_filename import ContractFilenameGenerator
from document_excel.infrastructure.excel import (
    Change,
    TemplateRegistry,
    WorkbookService,
    create_demo_template,
    file_hash,
    legacy_profile,
)


@pytest.fixture
def filled_contract(tmp_path):
    path = tmp_path / "filled.xlsx"
    create_demo_template(path)
    book = load_workbook(path)
    values = {
        "A18": "ПРЕЖНИЙ КЛИЕНТ ТЕСТОВИЧ",
        "E19": "ID",
        "G19": "0000001",
        "J19": "МКК",
        "K19": "001234",
        "M19": date(2019, 2, 1),
        "C20": "Прежняя область",
        "J20": "Прежняя улица, 1",
        "C22": "00123456789012",
        "A23": date(1980, 1, 1),
        "A12": "Существующий перфоратор",
        "A21": "=C20",
        "I21": "=J20",
        "D7": "+996 555 000001",
        "K22": "Прежний объект",
    }
    for cell, value in values.items():
        book["Дог."][cell] = value
    for cell, value in {"B20": "Перфоратор", "I20": 2, "N20": 350, "P20": 5000}.items():
        book["Пр №1"][cell] = value
    book["Акт прием"]["N21"] = 25000
    photos = book.create_sheet("Пасп Доп Данные")
    photos["M30"] = "Логотип вне паспортных областей"
    for index, anchor in enumerate(("A1", "F1", "M1")):
        image = tmp_path / f"old-{index}.png"
        Image.new("RGB", (100, 60), (index * 80, 20, 20)).save(image)
        photos.add_image(ExcelImage(image), anchor)
    book.save(path)
    book.close()
    service = WorkbookService()
    _, profile = service.describe(path, legacy_profile(), strict=True)
    case = ClientCase(
        workflow_mode=WorkflowMode.EXISTING_CONTRACT,
        template_path=path,
        workbook_profile=profile.model_dump(mode="json"),
        data=ClientData(
            full_name="ИВАНОВ ИВАН ИВАНОВИЧ",
            personal_identifier="00987654321098",
            passport_type="ID",
            passport_number="0000002",
            passport_issued_by="МКК",
            passport_department_code="005678",
            passport_issued_date=date(2021, 3, 2),
            birth_date=date(1990, 4, 3),
            document_type="kg_physical_id",
            registration_address=Address(region="Новая область", street="Новая улица, 2"),
        ),
        registration_status=RegistrationStatus.MANUALLY_ENTERED,
    )
    # Even accidentally populated commercial data must not enter this plan.
    case.contract.tools = [ToolItem(name="НЕ ЗАПИСЫВАТЬ", price_per_day=1)]
    case.contract.deposit_amount = 1
    case.contract.purchase_price = 1
    case.contract.actual_return_date = date(2026, 1, 1)
    for side, size in ((DocumentSide.FRONT, (800, 500)), (DocumentSide.BACK, (500, 800))):
        image = tmp_path / f"new-{side}.png"
        Image.new("RGB", size, "blue").save(image)
        case.images.append(SourceImage(path=image, role=side, case_id=case.id))
    case.confirmed_digest = case.digest()
    return path, case, profile


def approve(plan):
    plan.image_replace_confirmed = True
    plan.clear_missing_confirmed = True
    plan.preserved_confirmed = True
    if plan.best_effort:
        plan.select_missing_policy(MissingValuePolicy.CLEAR_OLD_PERSONAL if plan.identity_conflict else MissingValuePolicy.PRESERVE_EXISTING)
    return plan


def test_existing_contract_updates_only_personal_cells_and_current_photos(filled_contract, tmp_path):
    source, case, profile = filled_contract
    digest = file_hash(source)
    service = WorkbookService()
    plan = service.plan(source, case, profile)
    assert plan.existing_contract and plan.identity_conflict
    assert len(plan.photos.conflicts) == 2
    assert {c.field_name for c in plan.preserved} >= {"phones", "object_address"}
    assert not plan.new_sheets
    assert all(c.sheet == "Дог." and not c.formula for c in plan.changes)
    output = service.write(approve(plan), tmp_path / "copy.xlsx", case, profile)
    assert file_hash(source) == digest
    book = load_workbook(output)
    old = load_workbook(source)
    for cell, expected in {
        "A18": "Иванов Иван Иванович",
        "E19": "ID",
        "G19": "0000002",
        "J19": "МКК",
        "K19": "005678",
        "C22": "00987654321098",
        "C20": "Новая область",
        "J20": "Новая улица, 2",
    }.items():
        assert book["Дог."][cell].value == expected
    assert book["Дог."]["A23"].value.date() == date(1990, 4, 3)
    assert book["Дог."]["M19"].value.date() == date(2021, 3, 2)
    for sheet, cells in {
        "Дог.": ["A12", "D7", "K22", "A21", "I21"],
        "Пр №1": ["B20", "I20", "N20", "P20"],
        "Акт прием": ["N21"],
        "Акт расч": ["L21"],
    }.items():
        for cell in cells:
            assert book[sheet][cell].value == old[sheet][cell].value
    assert book.sheetnames == old.sheetnames
    assert len(book[profile.photo_sheet]._images) == 3  # two passports + unrelated logo
    assert str(book["Дог."].print_area) == str(old["Дог."].print_area)
    book.close()
    old.close()
    with ZipFile(source) as before, ZipFile(output) as after:
        # Commercial worksheets and original logo are preserved byte for byte.
        for part in (
            "xl/worksheets/sheet2.xml",
            "xl/worksheets/sheet3.xml",
            "xl/worksheets/sheet4.xml",
            "xl/media/image3.png",
        ):
            assert before.read(part) == after.read(part)
        assert "xl/media/image1.png" not in after.namelist()
        assert "xl/media/image2.png" not in after.namelist()


def test_missing_previous_personal_fields_require_explicit_clear(filled_contract, tmp_path):
    source, case, profile = filled_contract
    case.data.passport_department_code = None
    case.confirmed_digest = case.digest()
    service = WorkbookService()
    plan = service.plan(source, case, profile)
    assert plan.missing_personal
    plan.image_replace_confirmed = plan.preserved_confirmed = True
    with pytest.raises(AppError, match="очистку"):
        service.write(plan, tmp_path / "blocked.xlsx", case, profile)
    plan.clear_missing_confirmed = True
    output = service.write(plan, tmp_path / "cleared.xlsx", case, profile)
    book = load_workbook(output)
    assert book["Дог."]["K19"].value is None
    book.close()


def test_existing_identifiers_remain_text_even_with_numeric_custom_setting(filled_contract, tmp_path):
    source, case, profile = filled_contract
    for mapping in profile.field_mappings:
        if mapping.field in {"personal_identifier", "passport_number", "passport_department_code"}:
            mapping.value_type = "number"
    case.workbook_profile = profile.model_dump(mode="json")
    case.confirmed_digest = case.digest()
    service = WorkbookService()
    output = service.write(
        approve(service.plan(source, case, profile)), tmp_path / "text-identifiers.xlsx", case, profile
    )
    book = load_workbook(output)
    for address in ("C22", "G19", "K19"):
        assert book["Дог."][address].data_type == "s"
        assert book["Дог."][address].value.startswith("00")
    book.close()


def test_preserved_contacts_require_confirmation(filled_contract, tmp_path):
    source, case, profile = filled_contract
    service = WorkbookService()
    plan = approve(service.plan(source, case, profile))
    plan.preserved_confirmed = False
    with pytest.raises(AppError, match="сохраняемые"):
        service.write(plan, tmp_path / "blocked.xlsx", case, profile)


def test_explicit_contact_clear_preserves_address_formulas(filled_contract, tmp_path):
    source, case, profile = filled_contract
    service = WorkbookService()
    plan = approve(service.plan(source, case, profile))
    plan.clear_preserved_contacts()
    output = service.write(plan, tmp_path / "clear-contacts.xlsx", case, profile)
    book = load_workbook(output)
    assert book["Дог."]["D7"].value is None and book["Дог."]["K22"].value is None
    assert book["Дог."]["A21"].value == "=C20"
    book.close()


@pytest.mark.parametrize("retain_reference", [False, True])
def test_obsolete_personal_shared_strings_are_scrubbed_without_breaking_references(
    filled_contract, tmp_path, retain_reference
):
    from xml.sax.saxutils import escape

    from document_excel.infrastructure.xlsx_package import MAIN, child, document, encoded

    source, case, profile = filled_contract
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    old = load_workbook(source)
    values = [old["Дог."][cell].value for cell in ("A18", "C22", "G19")]
    old.close()
    doc = document(parts["xl/worksheets/sheet1.xml"])
    for cell in doc.getElementsByTagNameNS(MAIN, "c"):
        address = cell.getAttribute("r")
        if address not in ("A18", "C22", "G19"):
            continue
        for node in list(cell.childNodes):
            cell.removeChild(node)
        cell.setAttribute("t", "s")
        value = child(cell, "v")
        value.appendChild(doc.createTextNode(str(("A18", "C22", "G19").index(address))))
        cell.appendChild(value)
    if retain_reference:
        # A note elsewhere deliberately uses the same entry: it must survive.
        for cell in doc.getElementsByTagNameNS(MAIN, "c"):
            if cell.getAttribute("r") == "P30":
                for node in list(cell.childNodes):
                    cell.removeChild(node)
                cell.setAttribute("t", "s")
                value = child(cell, "v")
                value.appendChild(doc.createTextNode("0"))
                cell.appendChild(value)
    parts["xl/worksheets/sheet1.xml"] = encoded(doc)
    parts["xl/sharedStrings.xml"] = (
        f'<sst xmlns="{MAIN}" count="3" uniqueCount="3">'
        + "".join(f"<si><t>{escape(str(value))}</t></si>" for value in values)
        + "</sst>"
    ).encode("utf8")
    doc = document(parts["[Content_Types].xml"])
    node = child(doc.documentElement, "Override")
    node.setAttribute("PartName", "/xl/sharedStrings.xml")
    node.setAttribute(
        "ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    )
    doc.documentElement.appendChild(node)
    parts["[Content_Types].xml"] = encoded(doc)
    doc = document(parts["xl/_rels/workbook.xml.rels"])
    node = child(doc.documentElement, "Relationship")
    for key, value in {
        "Id": "rIdSST",
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        "Target": "sharedStrings.xml",
    }.items():
        node.setAttribute(key, value)
    doc.documentElement.appendChild(node)
    parts["xl/_rels/workbook.xml.rels"] = encoded(doc)
    with ZipFile(source, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    service = WorkbookService()
    output = service.write(
        approve(service.plan(source, case, profile)), tmp_path / "shared-copy.xlsx", case, profile
    )
    with ZipFile(output) as archive:
        sst = archive.read("xl/sharedStrings.xml").decode("utf8")
        assert (values[0] in sst) == retain_reference
        assert values[1] not in sst and values[2] not in sst
    book = load_workbook(output)
    assert book["Дог."]["A18"].value == "Иванов Иван Иванович"
    if retain_reference:
        assert book["Дог."]["P30"].value == values[0]
    book.close()


@pytest.mark.parametrize("mutation", ["financial", "new_sheet", "new_value", "photo", "formula"])
def test_mutated_existing_plan_cannot_expand_allowlist(filled_contract, tmp_path, mutation):
    source, case, profile = filled_contract
    service = WorkbookService()
    plan = approve(service.plan(source, case, profile))
    if mutation == "financial":
        plan.changes.append(Change("Пр №1", "N20", 350, 1, "attack"))
    elif mutation == "new_sheet":
        plan.new_sheets.append(("Пр №1", "Unexpected"))
    elif mutation == "new_value":
        plan.changes[0].new = "НЕ ТОТ КЛИЕНТ"
    elif mutation == "photo":
        plan.photos.insertions[0].data = b"unexpected"
    else:
        plan.changes[0].formula = True
    with pytest.raises(AppError, match="неожиданное"):
        service.write(plan, tmp_path / "blocked.xlsx", case, profile)
    assert not (tmp_path / "blocked.xlsx").exists()


def test_formula_and_unverified_coordinates_are_blocked(filled_contract):
    source, case, profile = filled_contract
    mapping = next(m for m in profile.field_mappings if m.field == "full_name")
    mapping.cell = "N20"
    case.workbook_profile = profile.model_dump(mode="json")
    case.confirmed_digest = case.digest()
    with pytest.raises(AppError, match="неразрешённая"):
        WorkbookService().plan(source, case, profile)


def test_replacing_another_clients_app_photos_requires_confirmation(filled_contract, tmp_path):
    source, case, profile = filled_contract
    service = WorkbookService()
    first = service.write(
        approve(service.plan(source, case, profile)), tmp_path / "first.xlsx", case, profile
    )
    case.template_path = first
    case.id = "new-synthetic-case"
    case.data.full_name = "ПЕТРОВ ПЁТР ТЕСТОВИЧ"
    case.confirmed_digest = case.digest()
    plan = approve(service.plan(first, case, profile))
    assert len(plan.photos.conflicts) == 2
    plan.image_replace_confirmed = False
    with pytest.raises(AppError, match="изображений"):
        service.write(plan, tmp_path / "blocked.xlsx", case, profile)
    plan.image_replace_confirmed = True
    second = service.write(plan, tmp_path / "second.xlsx", case, profile)
    with ZipFile(second) as archive:
        assert len([n for n in archive.namelist() if n.startswith("xl/media/document-excel-")]) == 2


def test_unknown_layout_never_uses_filename_and_saved_mapping_is_bound(tmp_path):
    source = tmp_path / "lermontov.xlsx"
    book = Workbook()
    book.active.title = "Дог."
    book.active["P30"] = "User layout"
    book.save(source)
    book.close()
    service = WorkbookService()
    _, unknown = service.describe(source, legacy_profile(), strict=True)
    assert not unknown.field_mappings
    custom = legacy_profile()
    custom.field_mappings = [m for m in custom.field_mappings if m.sheet == "Дог."]
    bound = service.bind_profile(source, custom)
    registry = TemplateRegistry(tmp_path / "profiles")
    registry.save(bound)
    _, selected = service.describe(source, unknown, strict=True, registry=registry)
    assert selected == bound
    book = load_workbook(source)
    book["Дог."].merge_cells("A2:C3")
    book.save(source)
    book.close()
    _, selected = service.describe(source, bound, strict=True, registry=registry)
    assert not selected.field_mappings


def test_same_size_different_contract_cannot_inherit_user_mapping(tmp_path):
    path = tmp_path / "custom.xlsx"
    book = Workbook()
    book.active.title = "Дог."
    book.active["A1"] = "Бланк клиента"
    book.active["P30"] = "Печать"
    book.save(path)
    book.close()
    service = WorkbookService()
    custom = legacy_profile()
    custom.field_mappings = [m for m in custom.field_mappings if m.sheet == "Дог."]
    bound = service.bind_profile(path, custom)
    registry = TemplateRegistry(tmp_path / "profiles")
    registry.save(bound)
    book = load_workbook(path)
    book["Дог."]["A1"] = "Совершенно другой бланк"
    book.save(path)
    book.close()
    _, selected = service.describe(path, bound, strict=True, registry=registry)
    assert not selected.field_mappings


def test_custom_fingerprint_does_not_create_cells_outside_used_range(tmp_path):
    from document_excel.infrastructure.excel import workbook_structure

    book = Workbook()
    book.active.title = "Дог."
    book.active["A10"] = "Существующая область"
    profile = legacy_profile()
    before = (book.active.max_row, book.active.max_column, len(book.active._cells))
    first = workbook_structure(book, profile)
    assert (book.active.max_row, book.active.max_column, len(book.active._cells)) == before
    assert workbook_structure(book, profile) == first
    source = tmp_path / "small.xlsx"
    book.save(source)
    book.close()
    with pytest.raises(AppError, match="вне существующей области"):
        WorkbookService().bind_profile(source, profile)


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("ИВАНОВ ИВАН ИВАНОВИЧ", "Иванов Иван Иванович"),
        (" СИДОРОВ-ПЕТРОВ  ИВАН ", "Сидоров-Петров Иван"),
        ("ӨМҮРБЕКОВ ҮСӨН ҢУР", "Өмүрбеков Үсөн Ңур"),
        ("O'NEILL D’ARTAGNAN", "O'Neill D’Artagnan"),
        ("McDonald Jean-Luc", "McDonald Jean-Luc"),
    ],
)
def test_person_name_formatter(raw, expected):
    assert PersonNameFormatter.format(raw) == expected


def test_manual_case_and_original_ocr_are_retained():
    case = ClientCase(data=ClientData(full_name="ИВАНОВ ИВАН"))
    assert case.formatted_name == "Иванов Иван"
    assert case.data.full_name == "ИВАНОВ ИВАН"
    case.data.full_name = "де ЛаКРУЗ Иван"
    case.evidence["full_name"] = FieldEvidence(source="ручной ввод", status=FieldStatus.CONFIRMED)
    assert case.formatted_name == "де ЛаКРУЗ Иван"


def test_filename_uses_clock_at_each_save_and_safe_name(tmp_path):
    today = [date(2026, 10, 19)]
    generator = ContractFilenameGenerator(lambda: today[0])
    expected = "Иванов Иван Иванович - 19.10.2026.xlsx"
    assert generator.generate("ИВАНОВ ИВАН ИВАНОВИЧ") == expected
    assert generator.generate('ИВАНОВ<>:"/\\|?* ИВАН') == "Иванов Иван - 19.10.2026.xlsx"
    today[0] = date(2026, 10, 20)
    assert generator.generate("ИВАНОВ ИВАН").endswith("20.10.2026.xlsx")
    path = tmp_path / expected
    path.touch()
    assert generator.available(path).name == expected.replace(".xlsx", " (2).xlsx")
    generator.available(path).touch()
    assert generator.available(path).name == expected.replace(".xlsx", " (3).xlsx")


@pytest.mark.parametrize("name", ["", "  ", "00000000000000", "<>/\\"])
def test_empty_filename_uses_neutral_fallback(name):
    assert ContractFilenameGenerator(lambda: date(2026, 10, 20)).generate(name) == "Договор - 20.10.2026.xlsx"


@pytest.mark.parametrize(
    "pattern",
    ["{personal_identifier}", "{full_name.__class__} - {date}", "{full_name!r} {date}", "{full_name}"],
)
def test_filename_pattern_cannot_include_identifiers(pattern):
    with pytest.raises(AppError):
        ContractFilenameGenerator().generate("Иванов Иван", pattern=pattern)


def test_exclusive_publication_race_keeps_existing_file(filled_contract, tmp_path, monkeypatch):
    source, case, profile = filled_contract
    service = WorkbookService()
    target = tmp_path / "race.xlsx"
    plan = approve(service.plan(source, case, profile))

    def race(_source, destination):
        destination.write_bytes(b"another complete file")
        raise FileExistsError

    monkeypatch.setattr("document_excel.infrastructure.excel.os.link", race)
    with pytest.raises(OutputExistsError):
        service.write(plan, target, case, profile)
    assert target.read_bytes() == b"another complete file"
    assert not list(tmp_path.glob(".document-excel-*"))


def test_gui_contract_passport_review_save_without_advanced_settings(
    filled_contract, qtbot, tmp_path, monkeypatch
):
    from types import SimpleNamespace

    from PySide6.QtWidgets import QDialog, QFileDialog, QWidget

    from document_excel.application.workflows import RecognitionOutput
    from document_excel.domain.models import FieldCandidate
    from document_excel.infrastructure.storage import SettingsRepository
    from document_excel.presentation import case_actions
    from document_excel.presentation.components import ConfirmDialog
    from document_excel.presentation.dialogs import CopySaveDialog
    from document_excel.presentation.main_window import MainWindow

    source, synthetic, _ = filled_contract
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "settings"))
    qtbot.addWidget(window)
    errors = []
    monkeypatch.setattr(window, "error", errors.append)
    window.show()
    assert not window.primary_button.isEnabled()
    assert sum(button.isVisible() for button in window.step_indicator.buttons) == 3
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(source), ""))
    window.contract_button.click()
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
    assert "Можно продолжить" in window.contract_status.text()
    assert not window.mapping_table.isVisible()
    window.case_actions.add_images([image.path for image in synthetic.images])
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
    case = window.current_case
    for image, role in zip(case.images, (DocumentSide.FRONT, DocumentSide.BACK), strict=True):
        image.role = role
    candidates = []
    values = {
        "full_name": "ИВАНОВ ИВАН ИВАНОВИЧ",
        "personal_identifier": "00987654321098",
        "passport_type": "ID",
        "passport_number": "0000002",
        "passport_issued_by": "МКК",
        "passport_department_code": "005678",
        "passport_issued_date": "02.03.2021",
        "birth_date": "03.04.1990",
        "registration_address": "Новая область, Новая улица, дом 2",
    }
    for index, image in enumerate(case.images):
        for name, value in values.items():
            if index and name not in {"personal_identifier", "passport_number"}:
                continue
            candidates.append(
                FieldCandidate(
                    field_name=name,
                    raw_value=value,
                    normalized_value=value,
                    confidence=0.99,
                    source_image_id=image.id,
                    validation_status=FieldStatus.REVIEW,
                )
            )
    output = RecognitionOutput(candidates, "kg_id", [], True, 0.1)
    monkeypatch.setattr(
        case_actions, "RecognitionService", lambda *_: SimpleNamespace(run=lambda *_, **__: output)
    )
    window.refresh_form()
    assert window.primary_button.isEnabled()
    assert (
        next(
            card for card in window.findChildren(QWidget) if card.property("documentSide") == "FRONT"
        ).height()
        >= 130
    )
    window.primary_button.click()
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
    assert window.current_step == 2
    window.resize(900, 700)
    qtbot.waitUntil(lambda: window.review_host_stack.currentIndex() == 1)
    window.resize(1366, 900)
    qtbot.waitUntil(lambda: window.review_host_stack.currentIndex() == 0)
    assert window.viewer_panel.isVisible() and window.review_form.isVisible()
    assert window.review_form.fields["full_name"].editor.text() == "Иванов Иван Иванович"
    monkeypatch.setattr(ConfirmDialog, "ask", lambda *args: True)

    today = [date(2026, 10, 19)]
    window.excel_actions.filenames.clock = lambda: today[0]

    def accept_export(dialog):
        today[0] = date(2026, 10, 20)  # Midnight while reviewing the copy.
        dialog.accept()
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(CopySaveDialog, "exec", accept_export)
    window.primary_button.click()
    qtbot.waitUntil(lambda: case.output_path is not None or bool(errors), timeout=15000)
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
    assert not errors
    assert case.output_path.name == "Иванов Иван Иванович 20.10.2026.xlsx"
    assert case.template_path == case.output_path and not source.exists()
    qtbot.waitUntil(lambda: window.current_step == 4, timeout=3000)
    window.case_actions.new_case()
    assert not window.current_case.images and not window.current_case.data.full_name
    assert not window.current_case.template_path and not window.current_case.output_path
    assert not window.current_case.confirmed


def test_gui_reselect_contract_resets_confirmation_keeps_client(
    filled_contract, qtbot, tmp_path, monkeypatch
):
    from shutil import copy2

    from PySide6.QtWidgets import QFileDialog

    from document_excel.infrastructure.storage import SettingsRepository
    from document_excel.presentation.main_window import MainWindow

    source, case, _ = filled_contract
    other = tmp_path / "other.xlsx"
    copy2(source, other)
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "settings"))
    qtbot.addWidget(window)
    window.batch.cases = [case]
    window.select_case(0)
    before = case.data.model_dump()
    photos = [image.id for image in case.images]
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(other), ""))
    window.excel_actions.choose_template()
    qtbot.waitUntil(lambda: not window.background_workers, timeout=10000)
    assert case.template_path == other and not case.confirmed
    assert case.data.model_dump() == before
    assert [image.id for image in case.images] == photos


def test_delete_last_client_keeps_existing_contract_mode(qtbot, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QInputDialog, QMessageBox

    from document_excel.infrastructure.storage import SettingsRepository
    from document_excel.presentation.main_window import MainWindow

    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "settings"))
    qtbot.addWidget(window)
    monkeypatch.setattr(QInputDialog, "getItem", lambda *args: ("Удалить запись из памяти", True))
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Yes)
    window.case_actions.manage_case()
    assert window.current_case.workflow_mode == WorkflowMode.BEST_EFFORT_UPDATE
    assert not window.current_case.template_path and not window.can_recognize()


def test_queue_checks_each_clients_own_profile(filled_contract, qtbot, tmp_path, monkeypatch):
    from document_excel.infrastructure.storage import SettingsRepository
    from document_excel.presentation.main_window import MainWindow

    _, case, profile = filled_contract
    profile.pin_length = 6
    case.data.personal_identifier = "000123"
    case.workbook_profile = profile.model_dump(mode="json")
    case.confirmed_digest = case.digest()
    monkeypatch.setattr(MainWindow, "diagnose", lambda self: None)
    window = MainWindow(SettingsRepository(tmp_path / "settings"))
    qtbot.addWidget(window)
    assert window.profile.pin_length == 14
    window.batch_export_queue = [case]
    window.batch_export_total = 1
    window.batch_export_folder = tmp_path
    planned = []
    monkeypatch.setattr(window, "run_worker", lambda *args, **kwargs: planned.append(args[0]))
    window.excel_actions.next_batch_export()
    assert len(planned) == 1


@pytest.mark.skipif(not Path("assets/lermontov.xlsx").exists(), reason="Local supplied workbook")
@pytest.mark.parametrize("mode", [WorkflowMode.EXISTING_CONTRACT, WorkflowMode.BEST_EFFORT_UPDATE])
def test_supplied_filled_contract_preserves_all_commercial_worksheets(filled_contract, tmp_path, mode):
    source = Path("assets/lermontov.xlsx").resolve()
    _, case, profile = filled_contract
    case.workflow_mode = mode
    service = WorkbookService()
    _, profile = service.describe(source, profile, strict=True)
    case.template_path = source
    case.workbook_profile = profile.model_dump(mode="json")
    case.confirmed_digest = case.digest()
    digest = file_hash(source)
    plan = approve(service.plan(source, case, profile))
    output = service.write(plan, tmp_path / "supplied-existing.xlsx", case, profile)
    assert file_hash(source) == digest
    assert not plan.new_sheets and all(c.sheet == "Дог." for c in plan.changes)
    with ZipFile(source) as before, ZipFile(output) as after:
        unmodified = all(
            before.read(name) == after.read(name)
            for name in before.namelist()
            if name.startswith("xl/worksheets/sheet") and name != "xl/worksheets/sheet1.xml"
        )
    assert unmodified
