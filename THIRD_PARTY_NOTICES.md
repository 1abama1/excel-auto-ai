# Зависимости

CPU-поставка использует RapidOCR 3.8.1 (Apache-2.0), ONNX Runtime 1.23.2 CPU (MIT) и совместимые ONNX-модели PP-OCRv5 (Apache-2.0). Источники: https://github.com/RapidAI/RapidOCR и https://github.com/microsoft/onnxruntime. Модели закреплены SHA-256; установочные URL — в scripts/prepare_rapid_models.py. Лицензии новых зависимостей находятся в licenses/cpu-dependencies, фактические версии — в requirements-locked.txt. NVIDIA, CUDA и cuDNN не входят в требования этой поставки.

Полные тексты [лицензии RapidOCR 3.8.1](https://github.com/RapidAI/RapidOCR/blob/v3.8.1/LICENSE), [лицензии ONNX Runtime 1.23.2](https://github.com/microsoft/onnxruntime/blob/v1.23.2/LICENSE) и его ThirdPartyNotices сохранены в `licenses/cpu`.

Новая Paddle-поставка включает PaddleOCR 3.7.0, PaddlePaddle 3.3.1, PaddleX 3.7.2 и официальные inference-модели PP-OCRv5 (Apache-2.0); источники: https://github.com/PaddlePaddle/PaddleOCR, https://github.com/PaddlePaddle/Paddle, https://github.com/PaddlePaddle/PaddleX. Контрольные суммы и исходные URL моделей находятся в models/paddle/manifest.json. Фактически установленные версии — в requirements-locked.txt, обнаруженные лицензионные тексты пакетов — в licenses/. При распространении сохраняйте эти файлы и лицензии CPython/Qt и транзитивных зависимостей.

PySide6-Essentials / Shiboken / Qt: условия и исходные проекты — https://doc.qt.io/qtforpython-6/licenses.html. Qt в сборке остаётся в виде заменяемых DLL. Скрипт сборки копирует найденные лицензионные тексты установленных дистрибутивов в LICENSES.

Noto Sans: Copyright The Noto Project Authors, SIL Open Font License 1.1. Полный текст включён в `resources/fonts/LICENSE.txt`. Источник: https://github.com/notofonts/noto-fonts.

Tesseract и модели поставляются отдельно: https://github.com/tesseract-ocr/tesseract и https://github.com/tesseract-ocr/tessdata_best. Локальная папка tools не включается в PyInstaller-пакет.

Прочие компоненты: openpyxl, Pydantic, NumPy, OpenCV, Pillow, pytesseract, их транзитивные зависимости и CPython. Сохраняйте их license/notice при распространении. Версии — в uv.lock.
