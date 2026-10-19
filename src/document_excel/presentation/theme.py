"""Shared visual tokens and application theme for the desktop workflow."""

from __future__ import annotations

from dataclasses import dataclass
from importlib.resources import files

from PySide6.QtGui import QColor, QFont, QFontDatabase, QPalette
from PySide6.QtWidgets import QApplication

from document_excel.infrastructure.storage import Settings


@dataclass(frozen=True)
class ThemeTokens:
    background: str = "#F8FAFC"
    surface: str = "#FFFFFF"
    text: str = "#0F172A"
    muted: str = "#64748B"
    border: str = "#E2E8F0"
    accent: str = "#2563EB"
    accent_hover: str = "#1D4ED8"
    accent_soft: str = "#EFF6FF"
    success: str = "#15803D"
    success_soft: str = "#F0FDF4"
    warning: str = "#B45309"
    warning_soft: str = "#FFFBEB"
    error: str = "#DC2626"
    error_soft: str = "#FEF2F2"
    viewer: str = "#EDF1F6"


LIGHT = ThemeTokens()
DARK = ThemeTokens(
    background="#111827",
    surface="#192334",
    text="#F1F5F9",
    muted="#A6B4C8",
    border="#334155",
    accent="#60A5FA",
    accent_hover="#3B82F6",
    accent_soft="#1E3556",
    success="#86EFAC",
    success_soft="#163729",
    warning="#FCD34D",
    warning_soft="#3F321C",
    error="#FCA5A5",
    error_soft="#42262D",
    viewer="#101927",
)


def apply_theme(app: QApplication, settings: Settings) -> None:
    t = DARK if settings.theme == "dark" else LIGHT
    accent_foreground = "#0F172A" if settings.theme == "dark" else "#FFFFFF"
    app.setStyle("Fusion")
    families = QFontDatabase.families()
    family = next((name for name in ("Inter", "Segoe UI", "Noto Sans") if name in families), "Noto Sans")
    if family not in families:
        for weight in ("Regular", "Bold"):
            QFontDatabase.addApplicationFont(
                str(files("document_excel").joinpath(f"resources/fonts/NotoSans-{weight}.ttf"))
            )
    scale = settings.ui_scale / 100
    app.setFont(QFont(family, round(10 * scale)))
    palette = QPalette()
    for role, value in [
        (QPalette.ColorRole.Window, t.background),
        (QPalette.ColorRole.Base, t.surface),
        (QPalette.ColorRole.AlternateBase, t.background),
        (QPalette.ColorRole.Text, t.text),
        (QPalette.ColorRole.WindowText, t.text),
        (QPalette.ColorRole.Button, t.surface),
        (QPalette.ColorRole.ButtonText, t.text),
        (QPalette.ColorRole.Highlight, t.accent),
        (QPalette.ColorRole.HighlightedText, accent_foreground),
        (QPalette.ColorRole.PlaceholderText, t.muted),
    ]:
        palette.setColor(role, QColor(value))
    app.setPalette(palette)
    app.setStyleSheet(f"""
        QWidget {{ color:{t.text}; font-size:{round(14 * scale)}px; }}
        QMainWindow, QDialog, QWidget#appRoot {{ background:{t.background}; }}
        QLabel {{ background:transparent; border:0; }}
        QLabel#title {{ font-size:{round(26 * scale)}px; font-weight:600; }}
        QLabel#sectionTitle {{ font-size:{round(18 * scale)}px; font-weight:600; }}
        QLabel#muted, QLabel#subtitle {{ color:{t.muted}; }}
        QLabel#caption {{ color:{t.muted}; font-size:{round(12 * scale)}px; }}
        QLabel#brand {{ font-size:{round(19 * scale)}px; font-weight:600; }}
        QLabel#brandMark {{ background:{t.accent_soft}; color:{t.accent}; border-radius:9px; font-size:23px; font-weight:700; }}
        QLabel#caseTitle {{ background:{t.surface}; color:{t.muted}; border:1px solid {t.border}; border-radius:8px; padding:8px 12px; }}
        QWidget#sidebar {{ background:{t.surface}; border-right:1px solid {t.border}; }}
        QPushButton {{ background:{t.surface}; color:{t.text}; border:1px solid {t.border}; border-radius:8px; padding:7px 12px; min-height:20px; }}
        QPushButton:hover {{ background:{t.background}; border-color:{t.muted}; }}
        QPushButton:pressed {{ background:{t.accent_soft}; }}
        QPushButton:focus {{ border:1px solid {t.accent}; }}
        QPushButton:disabled {{ color:{t.muted}; background:{t.background}; border-color:{t.border}; }}
        QPushButton#primary {{ background:{t.accent}; border-color:{t.accent}; color:{accent_foreground}; font-weight:600; padding:9px 20px; }}
        QPushButton#primary:hover {{ background:{t.accent_hover}; border-color:{t.accent_hover}; }}
        QPushButton#primary:disabled {{ background:{t.border}; border-color:{t.border}; color:{t.muted}; }}
        QPushButton#navButton {{ border:0; background:transparent; text-align:left; padding:10px 12px; }}
        QPushButton#navButton:hover {{ background:{t.background}; }}
        QPushButton#navButton:checked {{ background:{t.accent_soft}; color:{t.accent}; font-weight:600; }}
        QPushButton#iconButton {{ padding:3px; border:0; background:transparent; border-radius:6px; }}
        QPushButton#iconButton:hover {{ background:{t.accent_soft}; }}
        QPushButton#textButton {{ background:transparent; border:0; color:{t.accent}; padding:6px; }}
        QPushButton#step {{ background:transparent; color:{t.muted}; border:0; padding:6px 10px; font-size:{round(12 * scale)}px; }}
        QPushButton#step[state="active"] {{ background:{t.accent_soft}; color:{t.accent}; font-weight:600; }}
        QPushButton#step[state="complete"] {{ color:{t.success}; }}
        QPushButton#step:disabled {{ background:transparent; color:{t.muted}; }}
        QPushButton#step[state="active"]:disabled {{ background:{t.accent_soft}; color:{t.accent}; font-weight:600; }}
        QFrame#card, QFrame#dataPanel {{ background:{t.surface}; border:1px solid {t.border}; border-radius:10px; }}
        QFrame#viewerPanel {{ background:{t.viewer}; border:1px solid {t.border}; border-radius:10px; }}
        QFrame#workflowFooter {{ border:0; border-top:1px solid {t.border}; }}
        QFrame#dropZone {{ background:{t.surface}; border:2px dashed {t.border}; border-radius:12px; }}
        QFrame#dropZone[dragOver="true"] {{ background:{t.accent_soft}; border-color:{t.accent}; }}
        QFrame#photoThumbnail {{ background:{t.surface}; border:1px solid {t.border}; border-radius:10px; }}
        QPushButton#thumbnailImage {{ background:{t.viewer}; border:0; padding:0; }}
        QFrame#notification, QLabel#notice, QLabel#addressPreview {{ background:{t.accent_soft}; border:1px solid {t.border}; border-radius:8px; }}
        QLabel#reviewSummary {{ color:{t.muted}; font-size:{round(13 * scale)}px; }}
        QLabel#fieldLabel {{ color:{t.text}; font-weight:600; font-size:{round(13 * scale)}px; }}
        QLabel#fieldError {{ color:{t.error}; font-size:{round(12 * scale)}px; }}
        QLabel#loadingStage {{ color:{t.accent}; font-weight:600; padding-top:12px; }}
        QLabel#templateName {{ font-size:{round(17 * scale)}px; font-weight:600; }}
        QLabel#resultPath {{ color:{t.muted}; font-size:{round(15 * scale)}px; padding:8px; }}
        QLabel#statusIndicator {{ border-radius:5px; padding:3px 7px; font-size:{round(11 * scale)}px; }}
        QLabel#statusIndicator[tone="muted"] {{ background:{t.background}; color:{t.muted}; }}
        QLabel#statusIndicator[tone="success"] {{ background:{t.success_soft}; color:{t.success}; }}
        QLabel#statusIndicator[tone="warning"] {{ background:{t.warning_soft}; color:{t.warning}; }}
        QLabel#statusIndicator[tone="error"] {{ background:{t.error_soft}; color:{t.error}; }}
        QLabel#statusIndicator[tone="info"] {{ background:{t.accent_soft}; color:{t.accent}; }}
        QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{ background:{t.surface}; border:1px solid {t.border}; border-radius:7px; padding:7px 10px; min-height:20px; selection-background-color:{t.accent}; }}
        QLineEdit:hover, QComboBox:hover {{ border-color:{t.muted}; }}
        QLineEdit:focus, QComboBox:focus {{ border-color:{t.accent}; }}
        QLineEdit[invalid="true"] {{ border-color:{t.error}; background:{t.error_soft}; }}
        QLineEdit:read-only {{ background:{t.background}; }}
        QComboBox::drop-down {{ border:0; width:24px; }}
        QComboBox QAbstractItemView {{ background:{t.surface}; border:1px solid {t.border}; padding:4px; selection-background-color:{t.accent_soft}; selection-color:{t.text}; }}
        QComboBox#representation {{ font-size:{round(11 * scale)}px; min-height:16px; padding:2px 6px; }}
        QScrollArea {{ background:transparent; border:0; }}
        QWidget#reviewBody, QScrollArea#reviewScroll {{ background:{t.surface}; }}
        QScrollArea#thumbnailScroll, QScrollArea#thumbnailScroll > QWidget > QWidget {{ background:transparent; border:0; }}
        QPushButton#sectionToggle {{ background:transparent; border:0; border-bottom:1px solid {t.border}; border-radius:0; text-align:left; padding:3px 0 12px 0; font-size:{round(16 * scale)}px; font-weight:600; }}
        QGraphicsView {{ background:{t.viewer}; border:0; border-radius:8px; }}
        QListWidget, QTableView, QTextEdit {{ background:{t.surface}; border:1px solid {t.border}; border-radius:9px; padding:4px; selection-background-color:{t.accent_soft}; selection-color:{t.text}; }}
        QListWidget::item {{ padding:12px 8px; }}
        QTableView::item {{ border-bottom:1px solid {t.border}; padding:8px; }}
        QHeaderView::section {{ background:{t.background}; color:{t.muted}; padding:11px 8px; border:0; border-bottom:1px solid {t.border}; font-size:{round(12 * scale)}px; font-weight:600; }}
        QProgressBar {{ background:{t.border}; border:0; border-radius:4px; text-align:center; }}
        QProgressBar::chunk {{ background:{t.accent}; border-radius:4px; }}
        QSplitter::handle {{ background:{t.background}; width:16px; }}
        QTabWidget::pane {{ background:{t.surface}; border:0; }}
        QTabBar::tab {{ background:transparent; color:{t.muted}; padding:10px 18px; border-bottom:2px solid transparent; }}
        QTabBar::tab:selected {{ color:{t.accent}; border-bottom-color:{t.accent}; }}
        QMenu {{ background:{t.surface}; border:1px solid {t.border}; border-radius:8px; padding:6px; }}
        QMenu::item {{ padding:9px 20px; border-radius:5px; }}
        QMenu::item:selected {{ background:{t.accent_soft}; color:{t.text}; }}
        QMenu::separator {{ height:1px; background:{t.border}; margin:5px 8px; }}
        QToolTip {{ background:{t.text}; color:{t.surface}; border:0; padding:7px; border-radius:4px; }}
        QStatusBar {{ color:{t.muted}; background:{t.background}; border-top:1px solid {t.border}; font-size:{round(12 * scale)}px; }}
        QScrollBar:vertical {{ border:0; background:transparent; width:10px; margin:2px; }}
        QScrollBar::handle:vertical {{ background:{t.border}; border-radius:4px; min-height:30px; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background:transparent; }}
    """)
