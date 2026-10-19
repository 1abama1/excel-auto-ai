from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from document_excel.domain.models import Address, ContractData, RegistrationStatus, ToolItem, parse_date
from document_excel.infrastructure.excel import (
    EXPORT_FIELDS,
    CellMapping,
    ExportPlan,
    TemplateProfile,
    legacy_profile,
)
from document_excel.infrastructure.preprocessing import ImagePreprocessingService
from document_excel.infrastructure.storage import Settings
from document_excel.presentation.widgets import PreviewModel, configure_table


def buttons(dialog: QDialog, layout, save_text: str = "Сохранить") -> QDialogButtonBox:
    box = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
    box.button(QDialogButtonBox.StandardButton.Save).setText(save_text)
    box.button(QDialogButtonBox.StandardButton.Cancel).setText("Отмена")
    box.accepted.connect(dialog.accept)
    box.rejected.connect(dialog.reject)
    layout.addWidget(box)
    return box


class AddressDialog(QDialog):
    def __init__(self, address: Address, registration: bool = True, parent=None) -> None:
        super().__init__(parent)
        self.registration = registration
        self.setWindowTitle("Адрес регистрации" if registration else "Фактический адрес проживания")
        self.resize(560, 510)
        layout = QVBoxLayout(self)
        hint = QLabel(
            "На удостоверении нет адреса регистрации. Укажите прописку для договора.\nМожно добавить отдельный документ с пропиской."
            if registration
            and parent
            and hasattr(parent, "current_case")
            and parent.current_case.data.document_type == "kg_physical_id"
            and parent.current_case.registration_status == RegistrationStatus.NOT_PRESENT
            else "Адрес регистрации не найден. Заполните его вручную.\nМожно отменить ввод и добавить дополнительное фото."
            if not address.display() and registration
            else "Сверьте адрес с документом. Структурированные поля можно исправить по отдельности."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        form = QFormLayout()
        self.entries = {}
        for key, label in [
            ("country", "Страна"),
            ("region", "Регион"),
            ("city", "Город / населённый пункт"),
            ("district", "Район"),
            ("street", "Улица"),
            ("house", "Дом"),
            ("apartment", "Квартира"),
            ("extra", "Дополнительно"),
        ]:
            edit = QLineEdit(getattr(address, key))
            edit.setMaxLength(500)
            self.entries[key] = edit
            form.addRow(label, edit)
        self.raw = QLineEdit(address.raw)
        self.raw.setMaxLength(2000)
        form.addRow("Строка из документа", self.raw)
        layout.addLayout(form)
        self.absent = QCheckBox("Адрес действительно отсутствует в исходном документе")
        self.absent.setVisible(registration)
        layout.addWidget(self.absent)
        self.preview = QLabel()
        self.preview.setWordWrap(True)
        self.preview.setObjectName("addressPreview")
        layout.addWidget(self.preview)
        self.error = QLabel()
        self.error.setStyleSheet("color:#b42318")
        self.error.setWordWrap(True)
        layout.addWidget(self.error)
        for entry in [*self.entries.values(), self.raw]:
            entry.textChanged.connect(self.update_preview)
        self.absent.toggled.connect(self.update_preview)
        buttons(self, layout, "Подтвердить адрес")
        if registration and parent and hasattr(parent, "case_actions"):
            add_photo = QPushButton("Добавить фото прописки")

            def choose_photo():
                self.reject()
                parent.case_actions.choose_images()

            add_photo.clicked.connect(choose_photo)
            layout.addWidget(add_photo)
        self.update_preview()

    def address(self) -> Address:
        if self.absent.isChecked():
            return Address()
        return Address(
            **{key: entry.text().strip() for key, entry in self.entries.items()}, raw=self.raw.text().strip()
        )

    def update_preview(self):
        self.preview.setText("Итог: " + (self.address().display() or "адрес не указан"))

    def accept(self):
        address = self.address()
        if not self.absent.isChecked():
            structured = any(entry.text().strip() for entry in self.entries.values())
            if structured and (not (address.region or address.city) or not (address.street or address.extra)):
                self.error.setText(
                    "Укажите регион или населённый пункт, а также улицу или описание местоположения."
                )
                return
            if not structured and len(address.raw) < 10:
                self.error.setText("Введите адрес полностью или укажите его отсутствие в документе.")
                return
        super().accept()


class SettingsDialog(QDialog):
    def __init__(self, settings: Settings, parent=None) -> None:
        super().__init__(parent)
        self.initial = settings
        self.setWindowTitle("Настройки · локальная обработка")
        self.resize(680, 670)
        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        body = QWidget()
        form = QFormLayout(body)
        self.engine = QComboBox()
        self.engine.addItem("RapidOCR · PP-OCRv5 / ONNX CPU (основной)", "rapidocr")
        self.engine.addItem("PaddleOCR · PP-OCRv5 CPU (сравнение / резерв)", "paddle")
        self.engine.addItem("Tesseract · резервный, выбирается вручную", "tesseract")
        self.engine.setCurrentIndex(self.engine.findData(settings.ocr_engine))
        form.addRow("OCR-движок", self.engine)
        self.device = QLineEdit("cpu")
        self.device.setReadOnly(True)
        form.addRow("Устройство · без NVIDIA / CUDA", self.device)
        self.cpu_threads = QSpinBox()
        self.cpu_threads.setRange(1, 8)
        self.cpu_threads.setValue(settings.paddle_cpu_threads)
        form.addRow("Потоки CPU для OCR", self.cpu_threads)
        self.edits: dict[str, QLineEdit] = {}
        for key, title, file in [
            ("rapid_model_directory", "Локальные модели PP-OCRv5 ONNX", False),
            ("paddle_model_directory", "Локальные модели Paddle (резерв)", False),
            ("tesseract_path", "Tesseract executable", True),
            ("tessdata_dir", "Папка traineddata", False),
            ("output_directory", "Папка результатов", False),
            ("image_directory", "Папка копий фото", False),
            ("template_directory", "Папка шаблонов", False),
            ("temp_directory", "Временная папка OCR", False),
        ]:
            row = QWidget()
            box = QHBoxLayout(row)
            box.setContentsMargins(0, 0, 0, 0)
            edit = QLineEdit(getattr(settings, key))
            edit.setPlaceholderText(
                "Автоматически" if file or key == "tessdata_dir" else "По выбору пользователя"
            )
            self.edits[key] = edit
            browse = QPushButton("…")
            browse.setFixedWidth(38)
            browse.clicked.connect(lambda _checked=False, e=edit, f=file: self.browse(e, f))
            box.addWidget(edit)
            box.addWidget(browse)
            form.addRow(title, row)
        self.languages = QLineEdit("+".join(settings.languages))
        self.languages.setPlaceholderText("rus+eng; киргизский: kir при наличии модели")
        form.addRow("Tesseract: языки (коды)", self.languages)
        self.psm = QSpinBox()
        self.psm.setRange(3, 13)
        self.psm.setValue(settings.psm)
        form.addRow("Tesseract: сегментация PSM", self.psm)
        self.timeout = QSpinBox()
        self.timeout.setRange(5, 120)
        self.timeout.setValue(settings.timeout)
        self.timeout.setSuffix(" сек.")
        form.addRow("Timeout одного прохода", self.timeout)
        self.mode = QComboBox()
        self.mode.addItems(ImagePreprocessingService.MODES)
        self.mode.setCurrentText(settings.preprocessing)
        form.addRow("Подготовка изображения", self.mode)
        self.parallel = QSpinBox()
        self.parallel.setRange(1, 2)
        self.parallel.setValue(settings.max_parallel)
        form.addRow("Фоновые задачи (OCR последовательно)", self.parallel)
        self.theme = QComboBox()
        self.theme.addItem("Светлая", "light")
        self.theme.addItem("Тёмная", "dark")
        self.theme.setCurrentIndex(1 if settings.theme == "dark" else 0)
        form.addRow("Тема", self.theme)
        self.scale = QSpinBox()
        self.scale.setRange(80, 150)
        self.scale.setValue(settings.ui_scale)
        self.scale.setSuffix(" %")
        form.addRow("Размер шрифта", self.scale)
        self.history = QCheckBox("Время и результат операции; без ФИО, ПИН и путей")
        self.history.setChecked(settings.history_enabled)
        form.addRow("Локальная история", self.history)
        self.copy = QCheckBox("Копировать фото по нажатию «Сохранить копии фото»")
        self.copy.setChecked(settings.copy_images)
        form.addRow("Хранение фото", self.copy)
        note = QLabel(
            "Документы обрабатываются локально на CPU. RapidOCR и Paddle используют установленные модели PP-OCRv5 без интернета, NVIDIA, CUDA и cuDNN.\nОтмена нейронного OCR останавливает процесс; Tesseract: после текущего прохода, не позднее timeout."
        )
        note.setWordWrap(True)
        form.addRow(note)
        scroll.setWidget(body)
        outer.addWidget(scroll)
        buttons(self, outer)

    def browse(self, edit: QLineEdit, file: bool):
        value = (
            QFileDialog.getOpenFileName(self, "Tesseract", edit.text())[0]
            if file
            else QFileDialog.getExistingDirectory(self, "Папка", edit.text())
        )
        if value:
            edit.setText(value)

    def value(self) -> Settings:
        values = {key: edit.text().strip() for key, edit in self.edits.items()}
        return self.initial.model_copy(
            update={
                **values,
                "ocr_engine": self.engine.currentData(),
                "paddle_device": self.device.text().strip(),
                "paddle_cpu_threads": self.cpu_threads.value(),
                "languages": [s.strip() for s in self.languages.text().split("+") if s.strip()],
                "psm": self.psm.value(),
                "timeout": self.timeout.value(),
                "preprocessing": self.mode.currentText(),
                "auto_ocr": False,
                "max_parallel": self.parallel.value(),
                "theme": self.theme.currentData(),
                "ui_scale": self.scale.value(),
                "history_enabled": self.history.isChecked(),
                "copy_images": self.copy.isChecked(),
            }
        )


class MappingDialog(QDialog):
    def __init__(self, profile: TemplateProfile, sheets: list[str], parent=None) -> None:
        super().__init__(parent)
        self.profile = profile
        self.sheets = sheets
        self.result_profile = profile
        self.setWindowTitle("Соответствие полей и ячеек Excel")
        self.resize(1040, 700)
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Выберите лист и ведущую ячейку объединения. Формулы защищены; разрешение замены задаётся для каждой строки.\nПустые поля по умолчанию сохраняют текущие значения Excel. Проверьте профиль на копии своего шаблона."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        form = QFormLayout()
        self.name = QLineEdit(profile.name)
        self.required = QLineEdit(", ".join(profile.required_fields))
        self.pin = QSpinBox()
        self.pin.setRange(6, 30)
        self.pin.setValue(profile.pin_length)
        self.identifier_label = QLineEdit(profile.identifier_label)
        self.required_sheets = QLineEdit("; ".join(profile.required_sheets))
        self.optional_sheets = QLineEdit("; ".join(profile.optional_sheets))
        self.phone = QComboBox()
        self.phone.addItems(["combined", "separate"])
        self.phone.setCurrentText(profile.phone_mode)
        self.absence = QCheckBox("Разрешить отсутствие регистрации после явного указания пользователя")
        self.absence.setChecked(profile.allow_missing_registration)
        form.addRow("Название", self.name)
        form.addRow("Обязательные поля (через запятую)", self.required)
        form.addRow("Обязательные листы (через ;)", self.required_sheets)
        form.addRow("Необязательные листы (через ;)", self.optional_sheets)
        form.addRow("Длина ПИН", self.pin)
        form.addRow("Метка ПИН / ИНН", self.identifier_label)
        form.addRow("Телефоны: D7 вместе / D7 + G7", self.phone)
        form.addRow(self.absence)
        layout.addLayout(form)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Поле", "Лист", "Ячейка", "Тип", "Заменять формулу", "Пропускать пустые"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)
        actions = QHBoxLayout()
        for title, callback in [
            ("+ Строка", self.add_row),
            ("Удалить строку", self.remove_row),
            ("Карта D7 вместе", lambda: self.preset("combined")),
            ("Карта D7 / G7", lambda: self.preset("separate")),
        ]:
            button = QPushButton(title)
            button.clicked.connect(callback)
            actions.addWidget(button)
        layout.addLayout(actions)
        advanced = QHBoxLayout()
        self.tool_sheet = QLineEdit(profile.tool_sheet or "")
        self.tool_row = QSpinBox()
        self.tool_row.setRange(1, 100000)
        self.tool_row.setValue(profile.tool_first_row)
        self.tool_step = QSpinBox()
        self.tool_step.setRange(1, 100)
        self.tool_step.setValue(profile.tool_row_step)
        self.tool_inventory = QLineEdit(profile.tool_inventory_column or "")
        self.acceptance = QLineEdit(profile.acceptance_sheet or "")
        self.formula = QCheckBox("Возврат: H20/H21 очистить, L21 = G21-F21")
        self.formula.setChecked(profile.enable_return_formula)
        for label, control in [
            ("Лист инструментов", self.tool_sheet),
            ("Первая строка", self.tool_row),
            ("Лист акта приёма", self.acceptance),
        ]:
            advanced.addWidget(QLabel(label))
            advanced.addWidget(control)
        layout.addLayout(advanced)
        tool_options = QFormLayout()
        tool_options.addRow("Шаг строк инструментов", self.tool_step)
        tool_options.addRow(
            "Колонка заводского номера в приложении (пусто — отдельный акт)", self.tool_inventory
        )
        if profile.tool_inventory_sheet:
            tool_options.addRow(
                QLabel(
                    f"Заводские номера: {profile.tool_inventory_sheet}!{profile.tool_inventory_target_column}{profile.tool_inventory_first_row} и следующие строки. Акты копируются вместе с приложением."
                )
            )
        layout.addLayout(tool_options)
        photo_options = QFormLayout()
        self.photo_sheet = QComboBox()
        self.photo_sheet.setEditable(True)
        self.photo_sheet.addItem("Не вставлять фотографии", "")
        self.photo_sheet.addItems(sheets)
        self.photo_sheet.setCurrentText(profile.photo_sheet or "Не вставлять фотографии")
        self.photos_required = QCheckBox("Для этого шаблона нужны обе стороны документа")
        self.photos_required.setChecked(profile.photos_required)
        self.allow_photo_exception = QCheckBox("Разрешить явное исключение без отсутствующих фотографий")
        self.allow_photo_exception.setChecked(profile.allow_photo_exception)
        photo_options.addRow("Лист фотографий · A1:E24 / F1:K24", self.photo_sheet)
        photo_options.addRow(self.photos_required)
        photo_options.addRow(self.allow_photo_exception)
        layout.addLayout(photo_options)
        layout.addWidget(self.formula)
        self.error = QLabel()
        self.error.setWordWrap(True)
        self.error.setStyleSheet("color:#b42318")
        layout.addWidget(self.error)
        buttons(self, layout, "Применить профиль")
        for mapping in profile.field_mappings:
            self.add_row(mapping)

    def add_row(self, mapping=None):
        mapping = (
            mapping
            if isinstance(mapping, CellMapping)
            else CellMapping(field="full_name", sheet=self.sheets[0] if self.sheets else "Дог.", cell="A1")
        )
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col, choices, value in [
            (0, sorted(EXPORT_FIELDS), mapping.field),
            (1, self.sheets, mapping.sheet),
            (3, ["text", "date", "number"], mapping.value_type),
        ]:
            combo = QComboBox()
            combo.setEditable(col == 1)
            combo.addItems(choices)
            combo.setCurrentText(value)
            self.table.setCellWidget(row, col, combo)
        self.table.setItem(row, 2, QTableWidgetItem(mapping.cell))
        for col, checked in [(4, mapping.allow_formula_replace), (5, mapping.skip_empty)]:
            item = QTableWidgetItem()
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
            self.table.setItem(row, col, item)

    def remove_row(self):
        if self.table.currentRow() >= 0:
            self.table.removeRow(self.table.currentRow())

    def preset(self, mode):
        self.table.setRowCount(0)
        self.phone.setCurrentText(mode)
        for mapping in legacy_profile(mode).field_mappings:
            self.add_row(mapping)

    def accept(self):
        try:
            mappings = []
            for row in range(self.table.rowCount()):
                mappings.append(
                    CellMapping(
                        field=self.table.cellWidget(row, 0).currentText(),
                        sheet=self.table.cellWidget(row, 1).currentText(),
                        cell=self.table.item(row, 2).text(),
                        value_type=self.table.cellWidget(row, 3).currentText(),
                        allow_formula_replace=self.table.item(row, 4).checkState() == Qt.CheckState.Checked,
                        skip_empty=self.table.item(row, 5).checkState() == Qt.CheckState.Checked,
                    )
                )
            values = self.profile.model_dump()
            values.update(
                name=self.name.text().strip(),
                field_mappings=[m.model_dump() for m in mappings],
                required_fields=[s.strip() for s in self.required.text().split(",") if s.strip()],
                required_sheets=[s.strip() for s in self.required_sheets.text().split(";") if s.strip()],
                optional_sheets=[s.strip() for s in self.optional_sheets.text().split(";") if s.strip()],
                phone_mode=self.phone.currentText(),
                pin_length=self.pin.value(),
                identifier_label=self.identifier_label.text().strip() or "ПИН / ИНН",
                allow_missing_registration=self.absence.isChecked(),
                tool_sheet=self.tool_sheet.text().strip() or None,
                tool_first_row=self.tool_row.value(),
                tool_row_step=self.tool_step.value(),
                tool_inventory_column=self.tool_inventory.text().strip().upper() or None,
                acceptance_sheet=self.acceptance.text().strip() or None,
                enable_return_formula=self.formula.isChecked(),
                photo_sheet=None
                if self.photo_sheet.currentText() == "Не вставлять фотографии"
                else self.photo_sheet.currentText().strip() or None,
                photos_required=self.photos_required.isChecked(),
                allow_photo_exception=self.allow_photo_exception.isChecked(),
            )
            self.result_profile = TemplateProfile.model_validate(values)
        except (ValueError, AttributeError) as exc:
            self.error.setText("Проверьте mapping: " + str(exc)[:500])
            return
        super().accept()


class ExportDialog(QDialog):
    def __init__(self, plan: ExportPlan, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр Excel · подтверждение записи")
        self.resize(1100, 620)
        layout = QVBoxLayout(self)
        title = QLabel(
            f"Будет записано ячеек: {len(plan.changes)}. Новых листов: {len(plan.new_sheets)}.\nИсходный Excel сохраняется; результат записывается в новый файл."
        )
        title.setWordWrap(True)
        layout.addWidget(title)
        self.table = QTableView()
        self.model = PreviewModel(plan)
        self.table.setModel(self.model)
        configure_table(self.table)
        self.table.setColumnWidth(0, 130)
        self.table.setColumnWidth(1, 65)
        self.table.setColumnWidth(2, 170)
        self.table.setColumnWidth(3, 240)
        layout.addWidget(self.table)
        self.plan = plan
        self.image_confirm = QCheckBox("Разрешаю убрать показанные существующие рисунки из области A1:K24")
        self.image_confirm.setChecked(not (plan.photos and plan.photos.conflicts))
        self.image_confirm.setVisible(bool(plan.photos and plan.photos.conflicts))
        if plan.photos is not None:
            from PySide6.QtGui import QPixmap

            photos = QHBoxLayout()
            for photo in plan.photos.insertions:
                column = QVBoxLayout()
                column.addWidget(
                    QLabel(
                        f"{'Лицевая' if photo.role.value == 'FRONT' else 'Оборотная'} · {plan.photos.sheet}!{photo.area}"
                    )
                )
                image = QLabel()
                pixmap = QPixmap()
                pixmap.loadFromData(photo.data)
                image.setPixmap(
                    pixmap.scaled(
                        160,
                        80,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                )
                column.addWidget(image)
                photos.addLayout(column)
            if not plan.photos.insertions:
                photos.addWidget(QLabel("Фотографии отсутствуют · сохранение с явно выбранным исключением"))
            layout.addLayout(photos)
            if plan.photos.exception_used:
                layout.addWidget(QLabel("Явное исключение: отсутствующие фотографии не будут вставлены."))
            if plan.photos.conflicts:
                previous = QHBoxLayout()
                for label, data in zip(plan.photos.conflicts, plan.photos.conflict_previews, strict=True):
                    column = QVBoxLayout()
                    title = QLabel("Будет заменён: " + label)
                    title.setWordWrap(True)
                    column.addWidget(title)
                    image = QLabel("Просмотрите исходный рисунок в Excel, если превью не поддерживается.")
                    pixmap = QPixmap()
                    if pixmap.loadFromData(data):
                        image.setPixmap(pixmap.scaled(130, 70, Qt.AspectRatioMode.KeepAspectRatio))
                    column.addWidget(image)
                    previous.addLayout(column)
                layout.addLayout(previous)
        layout.addWidget(self.image_confirm)
        self.confirm = QCheckBox("Я проверил(а) значения и фотографии и разрешаю запись")
        layout.addWidget(self.confirm)
        box = buttons(self, layout, "Выбрать файл и сохранить")
        self.save_button = box.button(QDialogButtonBox.StandardButton.Save)
        self.save_button.setEnabled(False)
        self.confirm.toggled.connect(self.update_save)
        self.image_confirm.toggled.connect(self.update_save)

    def update_save(self):
        self.save_button.setEnabled(self.confirm.isChecked() and self.image_confirm.isChecked())

    def accept(self):
        if self.confirm.isChecked() and self.image_confirm.isChecked():
            self.plan.image_replace_confirmed = bool(self.plan.photos and self.plan.photos.conflicts)
            super().accept()


class ContractDialog(QDialog):
    def __init__(self, contract: ContractData, parent=None) -> None:
        super().__init__(parent)
        self.value = contract
        self.setWindowTitle("Инструменты и данные договора · ручной ввод")
        self.resize(850, 560)
        layout = QVBoxLayout(self)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Инструмент", "Количество", "Инвентарный №", "Цена / день"])
        self.table.horizontalHeader().setStretchLastSection(True)
        for tool in contract.tools:
            self.add_row(tool)
        layout.addWidget(self.table)
        row = QHBoxLayout()
        add = QPushButton("+ Инструмент")
        add.clicked.connect(self.add_row)
        remove = QPushButton("Удалить")
        remove.clicked.connect(
            lambda: self.table.removeRow(self.table.currentRow()) if self.table.currentRow() >= 0 else None
        )
        row.addWidget(add)
        row.addWidget(remove)
        layout.addLayout(row)
        form = QFormLayout()
        self.fields = {}
        for key, title in [
            ("deposit_amount", "Залог"),
            ("purchase_price", "Стоимость приобретения"),
            ("actual_return_date", "Дата возврата · ДД.ММ.ГГГГ"),
            ("actual_return_time", "Время возврата · ЧЧ:ММ"),
        ]:
            value = getattr(contract, key)
            edit = QLineEdit(
                value.strftime("%d.%m.%Y")
                if hasattr(value, "strftime")
                else str(value)
                if value is not None
                else ""
            )
            self.fields[key] = edit
            form.addRow(title, edit)
        layout.addLayout(form)
        self.error = QLabel()
        self.error.setWordWrap(True)
        layout.addWidget(self.error)
        buttons(self, layout)

    def add_row(self, tool=None):
        tool = tool if isinstance(tool, ToolItem) else ToolItem(name="")
        row = self.table.rowCount()
        self.table.insertRow(row)
        for column, value in enumerate([tool.name, tool.quantity, tool.inventory_number, tool.price_per_day]):
            self.table.setItem(row, column, QTableWidgetItem(str(value)))

    def accept(self):
        try:
            tools = []
            for row in range(self.table.rowCount()):
                values = [self.table.item(row, col).text().strip() for col in range(4)]
                if not values[0]:
                    raise ValueError("Укажите название каждого инструмента или удалите пустую строку.")
                tools.append(
                    ToolItem(
                        name=values[0],
                        quantity=int(values[1]),
                        inventory_number=values[2],
                        price_per_day=float(values[3].replace(",", ".")),
                    )
                )
            values = {key: edit.text().strip() or None for key, edit in self.fields.items()}
            if values["actual_return_date"]:
                values["actual_return_date"] = parse_date(values["actual_return_date"])
            self.value = ContractData(tools=tools, **values)
        except (ValueError, AttributeError) as exc:
            self.error.setText("Проверьте данные: " + str(exc)[:400])
            return
        super().accept()
