"""Installer contracts: pinned downloads, safe models, offline reuse and paths."""

import hashlib
import importlib.util
import io
import json
import tarfile
from pathlib import Path

import pytest

spec = importlib.util.spec_from_file_location(
    "onefile_provision", Path(__file__).resolve().parents[1] / "installer/provision.py"
)
provision = importlib.util.module_from_spec(spec)
spec.loader.exec_module(provision)


def test_verified_cached_package_does_not_access_network(tmp_path, monkeypatch):
    worker = provision.Provisioner(tmp_path, offline=True)
    data = b"synthetic wheel bytes"
    (worker.cache / "test.whl").write_bytes(data)
    monkeypatch.setattr(
        provision.urllib.request, "urlopen", lambda *_a, **_k: pytest.fail("Network on cached installation")
    )
    result = worker.download(
        "https://files.pythonhosted.org/test.whl", "test.whl", hashlib.sha256(data).hexdigest()
    )
    assert result.read_bytes() == data and worker.requests == 0


def test_corrupt_cached_wheel_is_not_accepted_offline(tmp_path):
    worker = provision.Provisioner(tmp_path, offline=True)
    (worker.cache / "test.whl").write_bytes(b"corrupt")
    with pytest.raises(OSError, match="интернет"):
        worker.download("https://files.pythonhosted.org/test.whl", "test.whl", "0" * 64)


def test_first_download_is_verified_and_then_reused_offline(tmp_path, monkeypatch):
    data = b"synthetic package for the first installation"

    def response(*_args, **_kwargs):
        stream = io.BytesIO(data)
        stream.headers = {"Content-Length": str(len(data))}
        return stream

    monkeypatch.setattr(provision.urllib.request, "urlopen", response)
    worker = provision.Provisioner(tmp_path)
    digest = hashlib.sha256(data).hexdigest()
    result = worker.download("https://files.pythonhosted.org/test.whl", "test.whl", digest, len(data))
    assert result.read_bytes() == data and worker.requests == 1
    assert not list(worker.cache.glob("*.part"))
    worker.offline = True
    assert worker.download("https://files.pythonhosted.org/test.whl", "test.whl", digest) == result
    assert worker.requests == 1


def test_failed_checksum_never_publishes_or_keeps_partial_package(tmp_path, monkeypatch):
    def response(*_args, **_kwargs):
        stream = io.BytesIO(b"corrupted download")
        stream.headers = {}
        return stream

    monkeypatch.setattr(provision.urllib.request, "urlopen", response)
    monkeypatch.setattr(provision.time, "sleep", lambda *_: None)
    worker = provision.Provisioner(tmp_path)
    with pytest.raises(OSError, match="проверенный компонент"):
        worker.download("https://files.pythonhosted.org/test.whl", "test.whl", "0" * 64)
    assert worker.requests == 3 and not list(worker.cache.iterdir())


@pytest.mark.parametrize(
    "url,name",
    [
        ("http://files.pythonhosted.org/a", "a"),
        ("https://untrusted.invalid/a", "a"),
        ("https://files.pythonhosted.org/a", "../outside"),
    ],
)
def test_untrusted_download_source_or_path_is_rejected(tmp_path, url, name):
    with pytest.raises(ValueError):
        provision.Provisioner(tmp_path).download(url, name)


def model_archive(path, *, corrupt=False, link=False):
    hashes = {}
    with tarfile.open(path, "w") as output:
        for name in provision.REQUIRED:
            data = ("synthetic " + name).encode()
            hashes[name] = hashlib.sha256(data).hexdigest()
            item = tarfile.TarInfo("model/" + name)
            item.size = len(data)
            if link:
                item.type, item.linkname, item.size = tarfile.SYMTYPE, "../outside", 0
                output.addfile(item)
            else:
                output.addfile(item, io.BytesIO(b"x" * len(data) if corrupt else data))
        # Malicious unrelated member is never extracted, even if present.
        item = tarfile.TarInfo("../../outside.txt")
        item.size = 3
        output.addfile(item, io.BytesIO(b"bad"))
    return hashes


@pytest.mark.parametrize("bad", ["checksum", "link"])
def test_invalid_model_never_publishes_partial_files(tmp_path, bad):
    archive = tmp_path / "model.tar"
    expected = model_archive(archive, corrupt=bad == "checksum", link=bad == "link")
    target = tmp_path / "model"
    with pytest.raises(ValueError):
        provision.Provisioner(tmp_path).extract_model(archive, target, expected)
    assert not target.exists() and not (tmp_path.parent / "outside.txt").exists()


def test_only_verified_regular_model_files_are_installed(tmp_path):
    archive = tmp_path / "model.tar"
    expected = model_archive(archive)
    target = tmp_path / "model"
    provision.Provisioner(tmp_path).extract_model(archive, target, expected)
    assert {p.name for p in target.iterdir()} == set(provision.REQUIRED)
    assert all(provision.sha(target / name) == digest for name, digest in expected.items())


def test_clean_profile_uses_installed_models_template_and_manual_ocr(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCUMENT_EXCEL_DATA_DIR", str(tmp_path / "user-data"))
    (tmp_path / "templates").mkdir()
    (tmp_path / "templates/lermontov.xlsx").write_bytes(b"synthetic")
    worker = provision.Provisioner(tmp_path)
    worker.configure()
    data = json.loads((tmp_path / "user-data/settings.json").read_text(encoding="utf8"))
    assert data["auto_ocr"] is False and data["ocr_engine"] == "rapidocr"
    assert Path(data["paddle_model_directory"]) == tmp_path / "models/paddle"
    assert Path(data["default_template"]) == tmp_path / "templates/lermontov.xlsx"
