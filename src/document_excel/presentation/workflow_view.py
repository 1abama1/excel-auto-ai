"""View construction for the document workflow, separate from window state and actions."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QMenu,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableView,
    QTableWidget,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from document_excel.presentation.components import (
    DocumentCard,
    EmptyState,
    IconButton,
    LoadingState,
    Notification,
    PrimaryButton,
    ReviewForm,
    SecondaryButton,
    StepIndicator,
    icon,
)
from document_excel.presentation.widgets import ImageViewer, configure_table


def _label(text: str = "", name: str = "muted", wrap: bool = False) -> QLabel:
    label = QLabel(text)
    label.setObjectName(name)
    label.setWordWrap(wrap)
    label.setTextFormat(Qt.TextFormat.PlainText)
    return label


def _page(title: str, subtitle: str) -> tuple[QWidget, QVBoxLayout]:
    page = QWidget()
    box = QVBoxLayout(page)
    box.setContentsMargins(30, 26, 30, 24)
    box.setSpacing(18)
    box.addWidget(_label(title, "title"))
    box.addWidget(_label(subtitle, wrap=True))
    return page, box


def _table(headers: list[str]) -> QTableWidget:
    table = QTableWidget(0, len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.verticalHeader().setVisible(False)
    table.verticalHeader().setDefaultSectionSize(52)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setAlternatingRowColors(False)
    table.setShowGrid(False)
    table.setWordWrap(False)
    table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    return table


def build_ui(w) -> None:
    root = QWidget()
    root.setObjectName("appRoot")
    root_row = QHBoxLayout(root)
    root_row.setContentsMargins(0, 0, 0, 0)
    root_row.setSpacing(0)
    _sidebar(w, root_row)
    w.pages = QStackedWidget()
    root_row.addWidget(w.pages, 1)
    _workflow(w)
    _templates(w)
    _history(w)
    _queue(w)
    _legacy_controls(w)
    w.setCentralWidget(root)
    update_responsive(w, w.width())


def _sidebar(w, root_row: QHBoxLayout) -> None:
    w.sidebar = QWidget()
    w.sidebar.setObjectName("sidebar")
    w.sidebar.setFixedWidth(216)
    box = QVBoxLayout(w.sidebar)
    box.setContentsMargins(16, 24, 16, 20)
    box.setSpacing(7)
    brand_row = QHBoxLayout()
    brand_mark = _label("Д", "brandMark")
    brand_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
    brand_mark.setFixedSize(36, 36)
    brand_row.addWidget(brand_mark)
    w.brand_label = _label("Документ", "brand")
    brand_row.addWidget(w.brand_label, 1)
    box.addLayout(brand_row)
    w.brand_subtitle = _label("Подготовка договоров", "caption")
    box.addWidget(w.brand_subtitle)
    box.addSpacing(25)
    w.nav_buttons = {}
    for index, text, glyph in [
        (0, "Документы", "document"),
        (3, "Очередь", "queue"),
        (1, "Excel-шаблоны", "templates"),
        (2, "История", "history"),
    ]:
        button = QPushButton(text)
        button.setObjectName("navButton")
        button.setIcon(icon(glyph))
        button.setCheckable(True)
        button.setChecked(index == 0)
        button.setMinimumHeight(44)
        button.setToolTip(text)
        button.setProperty("fullText", text)
        button.clicked.connect(lambda _checked=False, page=index: w.navigate(page))
        w.nav_buttons[index] = button
        box.addWidget(button)
    w.queue_badge = _label("1 клиент в работе", "caption")
    w.queue_badge.setContentsMargins(12, 3, 0, 0)
    box.addWidget(w.queue_badge)
    box.addStretch()
    w.new_case_button = SecondaryButton("Новый клиент", w.case_actions.new_case, "plus")
    w.new_case_button.setToolTip("Новый клиент · Ctrl+N")
    box.addWidget(w.new_case_button)
    box.addSpacing(16)
    w.settings_button = SecondaryButton("Настройки", w.show_settings, "settings")
    w.settings_button.setObjectName("navButton")
    box.addWidget(w.settings_button)
    w.privacy_label = _label("Обработка на компьютере\nДокументы остаются у вас", "caption", True)
    w.privacy_label.setContentsMargins(8, 12, 0, 0)
    box.addWidget(w.privacy_label)
    root_row.addWidget(w.sidebar)


def _workflow(w) -> None:
    page = QWidget()
    box = QVBoxLayout(page)
    box.setContentsMargins(30, 24, 30, 20)
    box.setSpacing(14)
    header = QHBoxLayout()
    headings = QVBoxLayout()
    headings.setSpacing(5)
    w.title_label = _label("Подготовка договора", "title")
    w.subtitle_label = _label("Добавьте фотографии удостоверения личности", wrap=True)
    headings.addWidget(w.title_label)
    headings.addWidget(w.subtitle_label)
    header.addLayout(headings, 1)
    w.case_title = _label("Клиент 1", "caseTitle")
    header.addWidget(w.case_title)
    help_button = IconButton("info", "Помощь и диагностика")
    help_menu = QMenu(help_button)
    help_menu.addAction(
        "Как подготовить договор",
        lambda: w.message(
            "Добавьте обе стороны документа, распознайте текст и проверьте данные. "
            "Затем выберите Excel-шаблон и сохраните готовый договор."
        ),
    )
    help_menu.addAction("Диагностика OCR", w.show_diagnostics)
    help_button.setMenu(help_menu)
    header.addWidget(help_button)
    box.addLayout(header)
    w.step_indicator = StepIndicator()
    w.step_indicator.selected.connect(w.go_step)
    box.addWidget(w.step_indicator)
    w.notice = Notification()
    box.addWidget(w.notice)
    w.step_stack = QStackedWidget()
    w.step_stack.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    _upload(w)
    _recognition(w)
    _review(w)
    _excel(w)
    _success(w)
    box.addWidget(w.step_stack, 1)
    footer = QFrame()
    footer.setObjectName("workflowFooter")
    footer_row = QHBoxLayout(footer)
    footer_row.setContentsMargins(0, 14, 0, 0)
    w.back_button = SecondaryButton(
        "Назад", lambda: w.go_step(0 if w.current_step == 2 else max(0, w.current_step - 1))
    )
    w.back_button.setIcon(icon("back"))
    footer_row.addWidget(w.back_button)
    w.footer_note = _label("Шаг 1 из 5 · Добавьте фотографии", "caption", True)
    footer_row.addWidget(w.footer_note, 1)
    w.primary_button = PrimaryButton("Распознать документы", w.primary_action)
    w.primary_button.setMinimumWidth(214)
    footer_row.addWidget(w.primary_button)
    box.addWidget(footer)
    w.pages.addWidget(page)


def _upload(w) -> None:
    page = QWidget()
    box = QVBoxLayout(page)
    box.setContentsMargins(0, 0, 0, 0)
    box.setSpacing(16)
    w.dropzone = EmptyState(
        "Добавьте фотографии документа", "Перетащите изображения сюда или выберите их на компьютере"
    )
    w.dropzone.files_dropped.connect(w.case_actions.add_images)
    w.add_button = SecondaryButton("Выбрать фотографии", w.case_actions.choose_images, "plus")
    w.add_button.setToolTip("Открыть фотографии · Ctrl+O")
    w.folder_button = SecondaryButton("Выбрать папку", w.case_actions.choose_folder, "folder")
    w.dropzone.actions.addWidget(w.add_button)
    w.camera_button = SecondaryButton("Сделать фото", w.case_actions.capture_documents)
    w.dropzone.actions.addWidget(w.camera_button)
    w.dropzone.actions.addWidget(w.folder_button)
    w.formats_label = _label(
        "JPG, PNG, WEBP, BMP, TIFF · Можно добавить обе стороны удостоверения", "caption", True
    )
    w.formats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    w.dropzone.layout().addWidget(w.formats_label)
    box.addWidget(w.dropzone, 1)
    w.upload_summary = _label(
        "Лицевая и оборотная стороны одного человека — одна запись клиента", "caption", True
    )
    box.addWidget(w.upload_summary)
    w.upload_document_type = QComboBox()
    for label, value in [
        ("Определить тип автоматически", "unknown"),
        ("Пластиковая ID-карта КР", "kg_physical_id"),
        ("Цифровое удостоверение КР", "kg_digital_id"),
        ("Кыргызстан · ID / паспорт", "kg_id"),
        ("Другой документ", "manual"),
    ]:
        w.upload_document_type.addItem(label, value)
    w.upload_document_type.currentIndexChanged.connect(w.case_actions.change_document_type)
    box.addWidget(w.upload_document_type)
    w.thumbnail_scroll = QScrollArea()
    w.thumbnail_scroll.setObjectName("thumbnailScroll")
    w.thumbnail_scroll.setWidgetResizable(True)
    w.thumbnail_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    w.thumbnail_scroll.setMinimumHeight(258)
    w.thumbnail_scroll.setMaximumHeight(270)
    thumbnails = QWidget()
    w.thumbnail_layout = QHBoxLayout(thumbnails)
    w.thumbnail_layout.setContentsMargins(0, 0, 0, 8)
    w.thumbnail_layout.setSpacing(14)
    w.thumbnail_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
    w.thumbnail_scroll.setWidget(thumbnails)
    w.thumbnail_scroll.hide()
    box.addWidget(w.thumbnail_scroll)
    w.step_stack.addWidget(page)


def _recognition(w) -> None:
    page = QWidget()
    box = QVBoxLayout(page)
    box.setContentsMargins(32, 24, 32, 24)
    box.setSpacing(18)
    box.addStretch()
    card = LoadingState(
        "Распознаём документ…", "Это займёт немного времени. Фотографии обрабатываются на вашем компьютере."
    )
    card.setMaximumWidth(640)
    card.setMinimumWidth(520)
    w.loading_title = card.heading
    w.substage_label = card.stage
    w.progress = card.progress
    w.loading_details = card.details
    cancel_row = QHBoxLayout()
    cancel_row.addStretch()
    w.cancel_button = SecondaryButton("Отменить распознавание", w.cancel_work)
    w.cancel_button.setToolTip("Отменить текущую операцию · Esc")
    w.cancel_button.setEnabled(False)
    cancel_row.addWidget(w.cancel_button)
    w.cancel_button.hide()
    card.box.addLayout(cancel_row)
    box.addWidget(card, 0, Qt.AlignmentFlag.AlignHCenter)
    box.addStretch()
    w.step_stack.addWidget(page)


def _review(w) -> None:
    page = QWidget()
    box = QVBoxLayout(page)
    box.setContentsMargins(0, 0, 0, 0)
    box.setSpacing(10)
    summary_row = QHBoxLayout()
    w.review_summary = _label("Данные извлечены. Проверьте выделенные поля.", "reviewSummary", True)
    summary_row.addWidget(w.review_summary, 1)
    w.issues_button = SecondaryButton("Показать ошибки", w.show_first_issue)
    w.issues_button.setObjectName("textButton")
    summary_row.addWidget(w.issues_button)
    box.addLayout(summary_row)
    w.review_host_stack = QStackedWidget()
    w.review_splitter = QSplitter(Qt.Orientation.Horizontal)
    w.review_splitter.setChildrenCollapsible(False)
    w.review_tabs = QTabWidget()
    w.review_tabs.setObjectName("reviewTabs")
    w.review_host_stack.addWidget(w.review_splitter)
    w.review_host_stack.addWidget(w.review_tabs)
    w.viewer_panel = QFrame()
    w.viewer_panel.setObjectName("viewerPanel")
    image_box = QVBoxLayout(w.viewer_panel)
    image_box.setContentsMargins(12, 12, 12, 10)
    image_box.setSpacing(8)
    image_header = QHBoxLayout()
    w.image_list = QComboBox()
    w.image_list.setMinimumWidth(100)
    w.image_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    w.image_list.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
    w.image_list.currentIndexChanged.connect(w.select_image)
    image_header.addWidget(w.image_list, 1)
    w.view_mode = QComboBox()
    w.view_mode.addItem("Документ", "processed")
    w.view_mode.addItem("Оригинал", "original")
    w.view_mode.addItem("Улучшенное", "enhanced")
    w.view_mode.setMaximumWidth(135)
    w.view_mode.currentIndexChanged.connect(lambda: _display_mode(w))
    w.display_mode = w.view_mode
    image_header.addWidget(w.view_mode)
    photo_menu = IconButton("more", "Действия с фотографией")
    photo_menu.clicked.connect(
        lambda: w.photo_menu(
            w.image_list.currentIndex(), photo_menu.mapToGlobal(photo_menu.rect().bottomLeft())
        )
    )
    image_header.addWidget(photo_menu)
    image_box.addLayout(image_header)
    w.viewer = ImageViewer()
    w.viewer.zoom_changed.connect(w.update_zoom)
    w.viewer.setMinimumSize(260, 200)
    w.viewer.files_dropped.connect(w.case_actions.add_images)
    w.viewer.region_selected.connect(w.case_actions.ocr_region)
    image_box.addWidget(w.viewer, 1)
    tools = QHBoxLayout()
    tools.setSpacing(4)
    for name, tooltip, method in [
        ("fit", "Вписать документ · Ctrl+0", "fit"),
        ("width", "Вписать по ширине", "fit_width"),
        ("minus", "Уменьшить · Ctrl+Minus", "zoom_out"),
        ("plus", "Увеличить · Ctrl+Plus", "zoom_in"),
    ]:
        tools.addWidget(
            IconButton(name, tooltip, lambda _checked=False, target=method: _viewer_action(w, target))
        )
    w.zoom_label = _label("100%", "caption")
    w.zoom_label.setMinimumWidth(46)
    w.zoom_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    tools.addWidget(w.zoom_label)
    tools.addStretch()
    tools.addWidget(IconButton("rotate", "Повернуть на 90°", w.case_actions.rotate_image))
    tools.addWidget(IconButton("region", "Выделить область для повторного OCR", w.viewer.select_region))
    tools.addWidget(
        IconButton("source", "Показать исходное изображение", lambda: w.view_mode.setCurrentIndex(1))
    )
    image_box.addLayout(tools)
    w.photo_info = _label("Колёсико — масштаб · Перетаскивание — перемещение", "caption", True)
    image_box.addWidget(w.photo_info)
    w.data_panel = QFrame()
    w.data_panel.setObjectName("dataPanel")
    data_box = QVBoxLayout(w.data_panel)
    data_box.setContentsMargins(0, 0, 0, 0)
    data_box.setSpacing(0)
    data_header = QWidget()
    heading = QVBoxLayout(data_header)
    heading.setContentsMargins(22, 16, 22, 12)
    row = QHBoxLayout()
    row.addWidget(_label("Данные клиента", "sectionTitle"), 1)
    w.completeness = _label("", "caption")
    row.addWidget(w.completeness)
    heading.addLayout(row)
    w.document_type = QComboBox()
    for label, value in [
        ("Тип не определён", "unknown"),
        ("Пластиковая ID-карта КР", "kg_physical_id"),
        ("Цифровое удостоверение КР", "kg_digital_id"),
        ("Кыргызстан · ID / паспорт", "kg_id"),
        ("Регистрация", "registration"),
        ("Другой документ · ручная проверка", "manual"),
    ]:
        w.document_type.addItem(label, value)
    w.document_type.currentIndexChanged.connect(w.case_actions.change_document_type)
    heading.addWidget(w.document_type)
    w.mrz_status = QPushButton("MRZ не проверялась")
    w.mrz_status.clicked.connect(w.case_actions.show_mrz)
    w.mrz_status.hide()
    heading.addWidget(w.mrz_status)
    w.identity_status = QPushButton("Проверить принадлежность фотографий")
    w.identity_status.clicked.connect(w.case_actions.review_identity)
    w.identity_status.hide()
    heading.addWidget(w.identity_status)
    w.side_review_button = SecondaryButton("Проверить стороны", w.case_actions.review_sides)
    w.side_review_button.hide()
    heading.addWidget(w.side_review_button)
    w.registration_notice = _label(
        "На удостоверении нет адреса регистрации. Укажите прописку для договора.", "caption", True
    )
    w.registration_notice.hide()
    heading.addWidget(w.registration_notice)
    w.registration_actions = QWidget()
    address_actions = QHBoxLayout(w.registration_actions)
    address_actions.setContentsMargins(0, 0, 0, 0)
    address_actions.addWidget(SecondaryButton("Ввести адрес", lambda: w.case_actions.address_dialog(True)))
    address_actions.addWidget(SecondaryButton("Добавить документ с пропиской", w.case_actions.choose_images))
    w.registration_actions.hide()
    heading.addWidget(w.registration_actions)
    data_box.addWidget(data_header)
    w.review_form = ReviewForm()
    w.review_form.value_edited.connect(w.apply_field_edit)
    w.review_form.focus_requested.connect(w.focus_review_field)
    w.review_form.details_requested.connect(w.show_field_details)
    w.review_form.address_requested.connect(
        lambda name: w.case_actions.address_dialog(name == "registration_address")
    )
    data_box.addWidget(w.review_form, 1)
    extras = QHBoxLayout()
    extras.setContentsMargins(20, 8, 20, 12)
    extras.addWidget(SecondaryButton("Инструменты и договор", w.case_actions.contract_dialog))
    extras.addStretch()
    more = IconButton("more", "Дополнительные действия")
    menu = QMenu(more)
    menu.addAction("Исходный текст OCR", w.show_raw_ocr)
    menu.addAction("Распознать ещё раз", w.case_actions.start_ocr)
    menu.addAction("Сохранить копии фото", w.excel_actions.save_photos)
    more.setMenu(menu)
    extras.addWidget(more)
    data_box.addLayout(extras)
    w.review_splitter.addWidget(w.viewer_panel)
    w.review_splitter.addWidget(w.data_panel)
    w.review_splitter.setStretchFactor(0, 3)
    w.review_splitter.setStretchFactor(1, 2)
    w.review_splitter.setSizes([640, 440])
    box.addWidget(w.review_host_stack, 1)
    w.step_stack.addWidget(page)


def _excel(w) -> None:
    page = QScrollArea()
    page.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    page.setWidgetResizable(True)
    page.setFrameShape(QFrame.Shape.NoFrame)
    content = QWidget()
    page.setWidget(content)
    box = QVBoxLayout(content)
    box.setContentsMargins(0, 0, 0, 0)
    box.setSpacing(16)
    card = DocumentCard("Шаблон договора")
    card.box.setContentsMargins(20, 14, 20, 14)
    card.box.setSpacing(8)
    card.heading.hide()
    w.template_label = _label("Excel-шаблон ещё не выбран", "templateName", True)
    card.box.addWidget(w.template_label)
    w.template_meta = _label(
        "Добавьте файл XLSX, чтобы увидеть его листы и необходимые поля.", "caption", True
    )
    card.box.addWidget(w.template_meta)
    w.template_meta.setMinimumHeight(42)
    w.template_meta.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
    w.profile_label = _label("", "caption", True)
    card.box.addWidget(w.profile_label)
    actions = QHBoxLayout()
    w.template_button = SecondaryButton("Выбрать Excel-шаблон", w.excel_actions.choose_template, "folder")
    actions.addWidget(w.template_button)
    actions.addWidget(SecondaryButton("Расположение полей", w.excel_actions.edit_mapping, "templates"))
    actions.addStretch()
    card.box.addLayout(actions)
    box.addWidget(card)
    from document_excel.presentation.photo_panel import DocumentPhotoPanel

    w.photo_panel = DocumentPhotoPanel()
    w.photo_panel.review_requested.connect(w.case_actions.review_document_side)
    w.photo_panel.exception_changed.connect(w.excel_actions.photo_exception)
    box.addWidget(w.photo_panel)
    preview = DocumentCard("Поля договора")
    w.mapping_table = _table(["Поле", "Лист", "Ячейка"])
    w.mapping_table.verticalHeader().setDefaultSectionSize(38)
    w.mapping_table.setMinimumHeight(165)
    preview.box.addWidget(w.mapping_table, 1)
    box.addWidget(preview, 1)
    w.excel_warning = _label("", "fieldError", True)
    w.excel_warning.hide()
    box.addWidget(w.excel_warning)
    w.step_stack.addWidget(page)


def _success(w) -> None:
    page = QWidget()
    box = QVBoxLayout(page)
    box.setContentsMargins(24, 20, 24, 20)
    box.setSpacing(18)
    box.addStretch()
    mark = QLabel()
    mark.setPixmap(icon("success", "#15803D", 64).pixmap(64, 64))
    mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
    box.addWidget(mark)
    title = _label("Договор успешно создан", "title")
    title.setAlignment(Qt.AlignmentFlag.AlignCenter)
    box.addWidget(title)
    w.result_label = _label("", "resultPath", True)
    w.result_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
    w.result_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    box.addWidget(w.result_label)
    actions = QHBoxLayout()
    actions.addStretch()
    w.open_result_button = SecondaryButton("Открыть Excel", w.excel_actions.open_result, "document")
    w.open_folder_button = SecondaryButton("Открыть папку", w.excel_actions.open_result_folder, "folder")
    actions.addWidget(w.open_result_button)
    actions.addWidget(w.open_folder_button)
    actions.addStretch()
    box.addLayout(actions)
    box.addStretch()
    w.step_stack.addWidget(page)


def _templates(w) -> None:
    page, box = _page(
        "Excel-шаблоны", "Настройте расположение полей один раз и используйте его для следующих договоров."
    )
    card = DocumentCard("Шаблон и поля")
    card.box.addWidget(SecondaryButton("Выбрать Excel-шаблон", w.excel_actions.choose_template, "folder"))
    card.box.addWidget(
        SecondaryButton("Изменить расположение полей", w.excel_actions.edit_mapping, "templates")
    )
    card.box.addWidget(SecondaryButton("Создать новый профиль", w.excel_actions.new_mapping, "plus"))
    box.addWidget(card)
    transfer = DocumentCard(
        "Перенос настроек", "Сохраните настройки полей или загрузите их с другого компьютера."
    )
    row = QHBoxLayout()
    row.addWidget(SecondaryButton("Импорт профиля", w.excel_actions.import_profile))
    row.addWidget(SecondaryButton("Экспорт профиля", w.excel_actions.export_profile))
    row.addStretch()
    transfer.box.addLayout(row)
    box.addWidget(transfer)
    demo = SecondaryButton("Создать демонстрационный шаблон", w.excel_actions.demo_template)
    demo.setObjectName("textButton")
    box.addWidget(demo, 0, Qt.AlignmentFlag.AlignLeft)
    box.addWidget(
        _label("Демонстрационный файл предназначен для проверки работы приложения.", "caption", True)
    )
    box.addStretch()
    w.pages.addWidget(page)


def _history(w) -> None:
    page, box = _page(
        "История", "Только время, операция и результат. Данные документов здесь не сохраняются."
    )
    w.history_text = QTextEdit()
    w.history_text.setReadOnly(True)
    w.history_text.setPlaceholderText("Выполненные операции появятся здесь")
    box.addWidget(w.history_text, 1)
    row = QHBoxLayout()
    row.addStretch()
    row.addWidget(SecondaryButton("Очистить историю", w.clear_history))
    box.addLayout(row)
    w.pages.addWidget(page)


def _queue(w) -> None:
    page, box = _page(
        "Очередь клиентов",
        "Обрабатывайте документы по порядку. Дважды нажмите на клиента, чтобы открыть его данные.",
    )
    actions = QHBoxLayout()
    actions.addWidget(SecondaryButton("Новый клиент", w.case_actions.new_case, "plus"))
    w.batch_ocr_button = SecondaryButton(
        "Распознать очередь", lambda: w.case_actions.start_ocr(all_cases=True), "scan"
    )
    actions.addWidget(w.batch_ocr_button)
    w.batch_export_button = SecondaryButton(
        "Сохранить готовые договоры", w.excel_actions.export_batch, "folder"
    )
    actions.addWidget(w.batch_export_button)
    actions.addStretch()
    menu_button = IconButton("more", "Действия с клиентом")
    menu = QMenu(menu_button)
    menu.addAction("Переименовать", w.case_actions.rename_case)
    menu.addAction("Объединить / удалить", w.case_actions.manage_case)
    menu_button.setMenu(menu)
    actions.addWidget(menu_button)
    box.addLayout(actions)
    w.queue_table = _table(["Клиент", "Фото", "Статус", "Проверить", "Excel"])
    w.queue_table.cellClicked.connect(lambda row, _column: _select_queue_case(w, row))
    w.queue_table.cellDoubleClicked.connect(lambda row, _column: _open_case(w, row))
    box.addWidget(w.queue_table, 1)
    w.pages.addWidget(page)


def _legacy_controls(w) -> None:
    """Keep the existing table model and action adapters without exposing a duplicate form."""
    w.legacy_host = QWidget(w)
    w.legacy_host.hide()
    w.fields = QTableView(w.legacy_host)
    configure_table(w.fields)
    w.fields.clicked.connect(w.case_actions.focus_field)
    w.fields.doubleClicked.connect(w.case_actions.edit_address_field)
    w.alternatives = QComboBox(w.legacy_host)
    w.field_detail = QLabel(w.legacy_host)
    w.cases = QListWidget(w.legacy_host)
    w.cases.currentRowChanged.connect(w.select_case)
    w.cases.itemDoubleClicked.connect(w.case_actions.rename_case)
    for name, callback in [
        ("ocr_button", w.case_actions.start_ocr),
        ("confirm_button", w.case_actions.confirm_data),
        ("export_button", w.excel_actions.export_current),
        ("merge_button", w.case_actions.manage_case),
    ]:
        button = QPushButton(w.legacy_host)
        button.clicked.connect(callback)
        setattr(w, name, button)


def _open_case(w, row: int) -> None:
    w.select_case(row)
    w.navigate(0)


def _select_queue_case(w, row: int) -> None:
    w.select_case(row)
    w.pages.setCurrentIndex(3)


def _viewer_action(w, name: str) -> None:
    if name in {"zoom_in", "zoom_out"}:
        factor = 1.2 if name == "zoom_in" else 1 / 1.2
        w.viewer.zoom(factor)
    else:
        getattr(w.viewer, name)()


def _display_mode(w) -> None:
    w.viewer.set_mode(w.view_mode.currentIndex())
    if w.current_case.images and any(s.processed_confirmed for s in w.current_case.images):
        w.select_image(w.image_list.currentIndex())


def update_responsive(w, width: int) -> None:
    """Switch document/data to tabs on small desktops without rebuilding their state."""
    compact = width < 1180
    w.sidebar.setFixedWidth(76 if compact else 216)
    w.brand_label.setVisible(not compact)
    w.brand_subtitle.setVisible(not compact)
    w.privacy_label.setVisible(not compact)
    w.queue_badge.setVisible(not compact)
    for button in w.nav_buttons.values():
        button.setText("" if compact else button.property("fullText"))
    w.settings_button.setText("" if compact else "Настройки")
    w.new_case_button.setText("" if compact else "Новый клиент")
    if compact and w.review_host_stack.currentIndex() != 1:
        w.review_tabs.addTab(w.viewer_panel, "Документ")
        w.review_tabs.addTab(w.data_panel, "Данные клиента")
        w.review_host_stack.setCurrentIndex(1)
    elif not compact and w.review_host_stack.currentIndex() != 0:
        while w.review_tabs.count():
            w.review_tabs.removeTab(0)
        w.review_splitter.addWidget(w.viewer_panel)
        w.review_splitter.addWidget(w.data_panel)
        w.review_host_stack.setCurrentIndex(0)
        w.review_splitter.setSizes([640, 440])
