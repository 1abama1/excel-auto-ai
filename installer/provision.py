"""First-run component installation; application OCR never accesses the network."""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile
import time
import urllib.request
from pathlib import Path

REQUIRED = ("inference.json", "inference.pdiparams", "inference.yml")
NETWORK_HOSTS = {"files.pythonhosted.org", "paddle-model-ecology.bj.bcebos.com"}


def sha(path):
    with Path(path).open("rb") as source:
        return hashlib.file_digest(source, "sha256").hexdigest()


def emit(percent, message):
    print(f"PROGRESS|{percent}|{message}", flush=True)


class Provisioner:
    def __init__(self, root, *, offline=False, cancel_file=None):
        self.root = Path(root).resolve()
        self.cache = self.root / "downloads"
        self.cache.mkdir(exist_ok=True)
        self.offline, self.cancel_file = offline, cancel_file
        self.requests = 0

    def check_cancel(self):
        if self.cancel_file and self.cancel_file.exists():
            raise InterruptedError("Установка отменена")

    def download(self, url, name, digest=None, expected_size=None, percent=45):
        from urllib.parse import urlparse

        if urlparse(url).scheme != "https" or urlparse(url).hostname not in NETWORK_HOSTS:
            raise ValueError("Недопустимый источник компонентов")
        if Path(name).name != name:
            raise ValueError("Недопустимое имя компонента")
        target = self.cache / name
        if target.is_file() and (not digest or sha(target) == digest):
            return target
        if self.offline:
            raise OSError("Для первой установки PaddleOCR нужен интернет. Готовые компоненты ещё не скачаны.")
        part = target.with_suffix(target.suffix + ".part")
        for attempt in range(3):
            self.check_cancel()
            try:
                self.requests += 1
                request = urllib.request.Request(url, headers={"User-Agent": "DocumentExcel-Setup/1.0"})
                with urllib.request.urlopen(request, timeout=30) as response, part.open("wb") as output:
                    total = int(response.headers.get("Content-Length", expected_size or 0))
                    downloaded, tick = 0, 0.0
                    while chunk := response.read(262144):
                        self.check_cancel()
                        downloaded += len(chunk)
                        if downloaded > (expected_size or 300_000_000) + 1048576:
                            raise ValueError("Размер компонента превышает ожидаемый")
                        output.write(chunk)
                        if time.monotonic() - tick > 0.2:
                            emit(
                                percent,
                                f"Скачивание {name}: {downloaded / 1048576:.1f} / {total / 1048576:.1f} МБ",
                            )
                            tick = time.monotonic()
                if expected_size and part.stat().st_size != expected_size:
                    raise ValueError("Компонент скачан не полностью")
                if digest and sha(part) != digest:
                    raise ValueError("Контрольная сумма компонента не совпала")
                os.replace(part, target)
                return target
            except InterruptedError:
                raise
            except (OSError, ValueError) as exc:
                part.unlink(missing_ok=True)
                if attempt == 2:
                    raise OSError(
                        "Не удалось скачать проверенный компонент. Проверьте интернет и повторите запуск EXE."
                    ) from exc
                time.sleep(0.5)
        raise AssertionError("unreachable")

    def command(self, arguments, label):
        emit(62, label)
        log = self.root / "install-components.log"
        with log.open("a", encoding="utf8") as output:
            child = subprocess.Popen(
                [sys.executable, *arguments],
                cwd=self.root,
                stdout=output,
                stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
            )
            try:
                deadline = time.monotonic() + 600
                while child.poll() is None:
                    self.check_cancel()
                    if time.monotonic() > deadline:
                        raise TimeoutError("Установка компонентов превысила десять минут")
                    time.sleep(0.1)
                if child.returncode:
                    raise RuntimeError(
                        "Установка CPU-компонентов не завершена. Подробности: install-components.log"
                    )
            finally:
                if child.poll() is None:
                    child.kill()
                    child.wait()

    def install_packages(self, packages):
        needed = []
        for package in packages:
            try:
                version = importlib.metadata.version(package["name"])
            except importlib.metadata.PackageNotFoundError:
                version = ""
            if (
                version != package["version"]
                or not (self.root / "runtime/Lib/site-packages" / package["module"] / "__init__.py").exists()
            ):
                needed.append(package)
        if not needed:
            return
        wheels = [self.download(p["url"], p["filename"], p["sha256"], p["size"]) for p in needed]
        if not importlib.util.find_spec("pip"):
            self.command(["-m", "ensurepip", "--upgrade"], "Подготовка установщика Python")
        # Every transitive dependency is already in the tested core bundle.
        # pip receives verified local wheels only, and cannot resolve or upgrade.
        self.command(
            [
                "-m",
                "pip",
                "install",
                "--no-index",
                "--no-deps",
                "--no-cache-dir",
                "--disable-pip-version-check",
                "--force-reinstall",
                *map(str, wheels),
            ],
            "Установка PaddleOCR / PaddlePaddle CPU",
        )
        importlib.invalidate_caches()

    def extract_model(self, archive, target, expected):
        # Only three regular, pinned files are read from each archive. No
        # member paths, links, devices or arbitrary extraction are accepted.
        with tempfile.TemporaryDirectory(prefix="model-", dir=target.parent) as temporary:
            stage = Path(temporary)
            with tarfile.open(archive) as bundle:
                members = bundle.getmembers()
                if len(members) > 100 or sum(m.size for m in members) > 250_000_000:
                    raise ValueError("Неверный архив модели")
                for name in REQUIRED:
                    matches = [m for m in members if Path(m.name).name == name]
                    if len(matches) != 1 or not matches[0].isfile():
                        raise ValueError("Неверный состав модели")
                    source = bundle.extractfile(matches[0])
                    if source is None:
                        raise ValueError("Модель не читается")
                    with source, (stage / name).open("wb") as output:
                        shutil.copyfileobj(source, output)
                    if sha(stage / name) != expected[name]:
                        raise ValueError("Контрольная сумма модели не совпала")
            target.mkdir(exist_ok=True)
            for name in REQUIRED:
                os.replace(stage / name, target / name)

    def models(self, manifest):
        root = self.root / "models/paddle"
        root.mkdir(parents=True, exist_ok=True)
        for index, (name, expected) in enumerate(manifest["models"].items()):
            self.check_cancel()
            target = root / name
            if all((target / n).is_file() and sha(target / n) == expected[n] for n in REQUIRED):
                continue
            emit(70 + index * 8, "Подготовка модели " + name)
            archive = self.download(expected["source"], name + "_infer.tar", percent=70 + index * 8)
            try:
                self.extract_model(archive, target, expected)
            except (ValueError, tarfile.TarError):
                archive.unlink(missing_ok=True)
                raise ValueError(
                    "Модель повреждена. Повторите запуск для скачивания исправной копии."
                ) from None
        destination = root / "manifest.json"
        destination.write_text(json.dumps(manifest, indent=2), encoding="utf8")

    def configure(self):
        from document_excel.infrastructure.storage import SettingsRepository

        repo = SettingsRepository()
        settings = repo.load()
        for key, folder in (("rapid_model_directory", "rapidocr"), ("paddle_model_directory", "paddle")):
            previous = getattr(settings, key)
            if not previous or not (Path(previous) / "manifest.json").is_file():
                setattr(settings, key, str(self.root / "models" / folder))
        template = self.root / "templates/lermontov.xlsx"
        if not settings.default_template or not Path(settings.default_template).is_file():
            settings.default_template = str(template)
            settings.default_profile = "lermontov-combined"
        settings.auto_ocr = False
        repo.save(settings)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--offline", action="store_true")
    parser.add_argument("--cancel-file", type=Path)
    args = parser.parse_args()
    worker = Provisioner(args.root, offline=args.offline, cancel_file=args.cancel_file)
    expected = json.loads((args.root / "installer/components.json").read_text(encoding="utf8"))
    result = {"ok": False, "network_requests": 0}
    try:
        worker.install_packages(expected["packages"])
        worker.models(expected["paddle_models"])
        emit(96, "Проверка локальных OCR-компонентов")
        # Real imports run in a child, so the launcher itself keeps no model
        # memory while the application starts its background CPU process.
        critical = [
            args.root / "runtime/Lib/site-packages" / p["module"] / "__init__.py"
            for p in expected["packages"]
        ]
        critical.append(args.root / "runtime/Lib/site-packages/paddle/base/libpaddle.pyd")
        stamps = {str(p.relative_to(args.root)): [p.stat().st_size, p.stat().st_mtime_ns] for p in critical}
        ready = args.root / ".components-ready.json"
        previous = json.loads(ready.read_text(encoding="utf8")) if ready.is_file() else {}
        if previous != stamps:
            worker.command(
                [
                    "-c",
                    "import os,socket; os.environ['PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK']='True'; socket.socket.connect=lambda *a,**k: (_ for _ in ()).throw(OSError('Offline verification')); socket.create_connection=socket.socket.connect; import paddle, paddleocr, paddlex, onnxruntime; assert not paddle.is_compiled_with_cuda(); assert 'CPUExecutionProvider' in onnxruntime.get_available_providers()",
                ],
                "Проверка PaddleOCR CPU",
            )
            ready.write_text(json.dumps(stamps), encoding="utf8")
        worker.configure()
        result.update(
            ok=True, versions={p["name"]: importlib.metadata.version(p["name"]) for p in expected["packages"]}
        )
        emit(100, "Готово. Запуск приложения")
        return 0
    except (OSError, ValueError, RuntimeError, tarfile.TarError) as exc:
        result["error"] = str(exc)
        print("ERROR|" + str(exc), flush=True)
        return 2
    finally:
        result["network_requests"] = worker.requests
        (args.root / "bootstrap-result.json").write_text(json.dumps(result, indent=2), encoding="utf8")


if __name__ == "__main__":
    raise SystemExit(main())
