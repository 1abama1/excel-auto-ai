# Дополнительный выпуск одним EXE — 20.10.2026

**`dist-onefile/DocumentExcel-current.exe`** — текущая версия приложения в том же формате, что прежний `dist-onefile/DocumentExcel.exe`. Предыдущий EXE сохранён. Размер нового файла — **367 609 098 байт** (350,58 MiB), SHA256: `0868967727f5be9076cd139a59488c95926c74fb9db09eee8e4fbfcc0febf119`. Для переноса достаточно одного EXE; при первой подготовке автоматически устанавливаются закреплённые Paddle CPU wheels. Подробности: [ONEFILE_WINDOWS](docs/ONEFILE_WINDOWS.md).

```powershell
dist-renamed\DocumentExcel\runtime\python.exe scripts/build_onefile.py --base dist-renamed/DocumentExcel --output dist-onefile/DocumentExcel-current.exe --work build/onefile-current
```

Включены текущие исходники с обновлением выбранного XLSX и переименованием по ФИО и дате. `Launch.cmd` продолжает использовать ранее проверенную папочную поставку; дополнительный EXE запускается непосредственно.

# Обновление: тот же файл с именем по ФИО и дате — 20.10.2026

По последнему уточнению выбранный XLSX обновляется и переименовывается в **`ФИО DD.MM.YYYY.xlsx`**. Папка сохраняется, постоянная копия не создаётся. При отсутствии имени используется `Договор DD.MM.YYYY.xlsx`; коллизии получают `(2)`, `(3)` без перезаписи другого файла. Повторное сохранение не нумерует текущий файл заново. Дата берётся в момент сохранения.

GUI показывает новое имя в единственной сводке, а после успеха переключает путь текущего документа на новое имя. Проверенный временный XLSX готовится заранее; затем исходный файл перемещается без перезаписи занятого имени и заменяется проверенным результатом. При ошибке записи выполняется возврат прежнего имени. Если возврат тоже заблокирован, программа сообщает путь, где остался документ. Переименование и замена — две файловые операции; единая атомарность обеих не заявляется.

Запуск актуальной Windows x64 CPU-поставки: **`dist-renamed/DocumentExcel/DocumentExcel.cmd`**. `Launch.cmd` использует её через `current-build.txt`. Ниже — история предыдущих вариантов поведения.

Проверенная команда сборки актуального выпуска:

```powershell
$env:PYTHONPATH = "$PWD/src"
$env:PYTHONUTF8 = '1'
dist-inplace\DocumentExcel\runtime\python.exe scripts/package_cpu_windows.py --base dist-inplace/DocumentExcel --output dist-renamed/DocumentExcel --packages dist-inplace/DocumentExcel/runtime/Lib/site-packages --models models/rapidocr --release inplace-rename-2026-10-20
```

# Текущая поставка с обновлением выбранного XLSX — 20.10.2026

Запуск: **`dist-inplace/DocumentExcel/DocumentExcel.cmd`**. Windows x64, Python 3.12, RapidOCR/PaddleOCR CPU и локальные модели включены. Переносится вся папка. `Launch.cmd` выбирает её через `current-build.txt`.

```powershell
$env:PYTHONPATH = "$PWD/src"
dist-simplified\DocumentExcel\runtime\python.exe scripts/package_cpu_windows.py --base dist-simplified/DocumentExcel --output dist-inplace/DocumentExcel --packages dist-simplified/DocumentExcel/runtime/Lib/site-packages --models models/rapidocr --release inplace-update-2026-10-20
```

Результат записывается в выбранный пользователем Excel после проверки временного XLSX. Имя и папка не меняются, новая копия договора не создаётся. Сборщик требует новый каталог; прежние поставки остаются доступными. Текущий отчёт — в начале TEST_REPORT.md.

Ниже — история предыдущих поставок.

# Текущая упрощённая поставка — 20.10.2026

Windows x64 CPU: **`dist-simplified/DocumentExcel/DocumentExcel.cmd`**. Переносится вся папка `DocumentExcel`; Python 3.12, Qt, RapidOCR/PaddleOCR CPU и локальные модели включены. Установка и сеть при обычном запуске не требуются.

Воспроизводимая сборка из имеющегося проверенного runtime:

```powershell
$env:PYTHONPATH = "$PWD/src"
dist-contract\DocumentExcel\runtime\python.exe scripts/package_cpu_windows.py --base dist-contract/DocumentExcel --output dist-simplified/DocumentExcel --packages dist-contract/DocumentExcel/runtime/Lib/site-packages --models models/rapidocr --release best-effort-2026-10-20
```

Сборщик требует новый каталог результата. Старые поставки не заменяются. `Launch.cmd` выбирает новую поставку через `current-build.txt`. Проверки текущего выпуска записаны в начале TEST_REPORT.md. Это автономная папка с командой запуска, без подписи и установщика; проверка на отдельной чистой Windows 11 не выполнялась.

Ниже — исторические варианты упаковки; они не обозначают текущую версию.

# Сборка

## Текущий готовый договор — 20.10.2026

Автономная CPU-поставка: `dist-contract/DocumentExcel/DocumentExcel.cmd`.
Python 3.12, Qt, RapidOCR/PaddleOCR CPU и проверенные модели включены;
переносится вся папка. Пользовательские договоры и реальные фото не включены.
`Launch.cmd` выбирает эту поставку через `current-build.txt`.

Сборка из существующего локального проверенного runtime:

```powershell
dist-digital\DocumentExcel\runtime\python.exe -X utf8 scripts/package_cpu_windows.py --base dist-digital/DocumentExcel --output dist-contract/DocumentExcel --packages dist-digital/DocumentExcel/runtime/Lib/site-packages --models models/rapidocr --release existing-contract-2026-10-20
```

Для запуска скрипта сборки source `src` должен быть в sys.path (как при
обычной установке проекта). Точный проверенный запуск и результаты записаны
в TEST_REPORT.md. Ниже — история других вариантов упаковки.

Новая поставка одним EXE: `dist-onefile/DocumentExcel.exe`. Сборщик
`scripts/build_onefile.py` упаковывает проверенный Python 3.12 runtime,
приложение, RapidOCR и модели; C# launcher автоматически устанавливает
закреплённые CPU wheels PaddleOCR и проверенные модели при первом запуске.
Инструкция и параметры изолированной проверки: [ONEFILE_WINDOWS](docs/ONEFILE_WINDOWS.md).
Для переноса на другой компьютер достаточно одного EXE.

Обновление PaddleOCR: автономная Windows CPU-поставка использует Python 3.12, PaddleOCR 3.7.0, PaddlePaddle 3.3.1 и три проверенные локальные модели. Установка: `scripts/setup_paddle.ps1`; упаковка: `scripts/package_paddle_windows.py --output dist-paddle/DocumentExcel --models models/paddle` из Paddle-окружения. Выход: `DocumentExcel.cmd` + `runtime` + `app` + `models`, переносится вся папка. Подробности: [PADDLE_MIGRATION](docs/PADDLE_MIGRATION.md). Ниже сохранена инструкция прежней PyInstaller/Tesseract-сборки.

Проверяемая платформа: Windows 10 build 19045 x64, CPython 3.14.0. Зависимости зафиксированы в `uv.lock`.

```powershell
uv sync --frozen --extra dev
powershell -ExecutionPolicy Bypass -File scripts/check.ps1
powershell -ExecutionPolicy Bypass -File scripts/build_windows.ps1
```

Выход: `dist/DocumentExcel/DocumentExcel.exe`. Копируется вся папка DocumentExcel с `_internal`. Python включён; ресурсы ищутся через пакет и не зависят от текущей папки. Настройки — в пользовательском каталоге. Для локального окружения `Launch.cmd` передаёт путь к `tools/tesseract`.

Если программа уже открыта, сборка в её папку блокируется до удаления файлов. Используйте `-OutputDirectory dist-update`, чтобы собрать рядом. После успешной сборки `current-build.txt` указывает новый EXE для `Launch.cmd`. В обновлении разделения типа/номера используется `dist-update/DocumentExcel/DocumentExcel.exe`; открытая старая программа не закрывается автоматически.

Текущее обновление пластиковой ID-карты собрано с `-OutputDirectory dist-state`: `dist-state/DocumentExcel/DocumentExcel.exe`. Предыдущие сборки не заменяются. Проверка `--self-test` включает прежний OCR и XLSX с двумя фотографиями. Новый `--physical-self-test C:\папка\проверки` создаёт две синтетические пластиковые стороны, выполняет настоящий OCR, проверяет отдельные официальные формы имени, номер/ПИН, дату, MRZ и отсутствие регистрации, добавляет синтетический ручной адрес и сохраняет XLSX с двумя фотографиями. Отчёт — `physical-self-test.json`. Пользовательские изображения не читаются; аппаратная камера не используется. Нужны системные обычный и моноширинный шрифты (Windows Arial/Courier New либо Linux DejaVu).

Windows-сборка очищает сторонние каталоги из PATH на время анализа DLL. Для Qt используется системная ICU Windows 10, а несовместимая одноимённая библиотека из внешних PDF-инструментов не включается. Исходное PATH после сборки восстанавливается.

В EXE-пакете находятся Qt Widgets, OpenCV, Pillow, openpyxl, Pydantic, интерфейс, значок, Noto Sans, профили и демонстрационный XLSX. Tesseract, traineddata и установщик **не включаются**. Установите OCR отдельно либо выберите путь в настройках. Это стратегия A, а не полностью автономный OCR-пакет.

Smoke test показывает окно, сохраняет его изображение/JSON и закрывается:

```powershell
$env:DOCUMENT_EXCEL_TESSERACT = 'C:\путь\tesseract.exe'
dist\DocumentExcel\DocumentExcel.exe --smoke-test C:\папка\проверки
```

`QT_QPA_PLATFORM=offscreen` позволяет выполнить эту проверку без видимого окна. `scripts/smoke_gui.py` рендерит настоящие виджеты с синтетическим документом в 1366×768; `--dark` проверяет другую тему.

`scripts/visual_redesign.py` выполняет настоящий OCR двух синтетических сторон, снимает 12 состояний, проверяет четыре логических размера окна и создаёт XLSX через безопасный экспорт. `--scale 1.25`, `1.5`, `2` задают DPI Qt; `--skip-ocr` пригоден для повторной проверки компоновки. Полный фактический отчёт и ограничения приведены в TEST_REPORT.md.

Полная проверка собранного пакета: `DocumentExcel.exe --self-test C:\папка\проверки`. Она создаёт исключительно синтетические изображения, выполняет настоящий OCR, сверяет шесть полей и дополнительное чтение ID, готовит демо-шаблон с фотолистом, записывает и повторно открывает XLSX с двумя встроенными изображениями. Отчёт — `self-test.json`, результат — отдельный synthetic-result XLSX. Пользовательские документы в этом режиме не читаются. Камера этим режимом не проверяется.

Камера использует уже установленный OpenCV: PySide6-Essentials не содержит Qt Multimedia, поэтому дополнительные Qt-пакеты и зависимости не добавлялись. Захват и поиск работают в фоновых потоках; Windows использует DirectShow/MSMF, Linux — V4L2. Отдельная аппаратная проверка: `.venv\Scripts\python scripts/check_camera_hardware.py` — кадр остаётся в памяти, отчёт содержит только разрешение и состояние освобождения устройства. Результаты текущего компьютера указаны в TEST_REPORT.md; Linux аппаратно не проверялся.

Linux-артефакт собирается на Linux: `sh scripts/build_linux.sh`. Могут потребоваться системные Qt xcb/Wayland-библиотеки и Tesseract. В Windows-среде Linux-сборка не выполнялась. EXE не подписан сертификатом выпуска; MSI, SmartScreen и чистая Windows 11 не проверены.

Перед распространением сохраняйте THIRD_PARTY_NOTICES и LICENSES. Qt остаётся набором заменяемых динамических библиотек. Политика лицензирования: [Qt for Python](https://doc.qt.io/qtforpython-6/licenses.html). Сборка предназначена для локальной проверки и использования; это не подписанный коммерческий установщик.
