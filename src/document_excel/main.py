from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QMessageBox

from document_excel.infrastructure.storage import SettingsRepository, configure_logging
from document_excel.presentation.main_window import MainWindow
from document_excel.presentation.theme import apply_theme


def main() -> int:
    parser = argparse.ArgumentParser(description="Локальное OCR → Excel")
    parser.add_argument("--smoke-test", type=Path, help="Show the real window, save a screenshot and quit")
    parser.add_argument("--self-test", type=Path, help="Run a synthetic image → OCR → Excel check and quit")
    parser.add_argument(
        "--physical-self-test", type=Path, help="Run synthetic plastic TD1 → MRZ → Excel and quit"
    )
    args, qt_args = parser.parse_known_args()
    if args.self_test or args.physical_self_test:
        from document_excel.domain.errors import AppError
        from document_excel.infrastructure.verification import run_self_test

        output = args.physical_self_test or args.self_test
        output.mkdir(parents=True, exist_ok=True)
        try:
            if args.physical_self_test:
                from document_excel.infrastructure.physical_verification import run_physical_self_test

                report = run_physical_self_test(output)
            else:
                report = run_self_test(output)
        except (AppError, OSError) as exc:
            report = {
                "ok": False,
                "error_type": type(exc).__name__,
                "message": str(exc) if isinstance(exc, AppError) else "Нет доступа к файлам проверки",
            }
        (output / ("physical-self-test.json" if args.physical_self_test else "self-test.json")).write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return 0 if report["ok"] else 1
    app = QApplication(sys.argv)
    app.setApplicationName("Document Excel")
    app.setOrganizationName("DocumentExcel")
    if args.smoke_test:
        args.smoke_test.mkdir(parents=True, exist_ok=True)
    repository = SettingsRepository(args.smoke_test / "app-data" if args.smoke_test else None)
    try:
        configure_logging(repository.root)
    except OSError:
        QMessageBox.warning(
            None, "Диагностика", "Папка журнала недоступна. Приложение продолжит работу без записи журнала."
        )
    settings = repository.load()
    if settings.temp_directory:
        try:
            with tempfile.TemporaryFile(dir=settings.temp_directory):
                pass
            tempfile.tempdir = settings.temp_directory
        except OSError:
            QMessageBox.warning(
                None,
                "Временная папка",
                "Настроенная папка OCR недоступна. Используется системная временная папка.",
            )
    apply_theme(app, settings)
    window = MainWindow(repository)
    window.show()
    if args.smoke_test:

        def capture():
            window.grab().save(str(args.smoke_test / "window.png"))
            (args.smoke_test / "smoke.json").write_text(
                json.dumps(
                    {
                        "window_visible": window.isVisible(),
                        "field_rows": window.field_model.rowCount(),
                        "profile": window.profile.id,
                        "diagnostic": window.notice.text(),
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            window.close()
            QTimer.singleShot(500, app.quit)

        QTimer.singleShot(2000, capture)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
