from datetime import date

import pytest

from document_excel.domain.models import Address, ClientCase, normalize, parse_date, validate_field
from document_excel.infrastructure.excel import readiness
from document_excel.infrastructure.storage import SettingsRepository, collect_images, safe_filename


@pytest.mark.parametrize("text", ["01.02.2000", "1/2/2000", "2000-02-01", "01-02-2000"])
def test_date_formats(text):
    assert parse_date(text) == date(2000, 2, 1)


@pytest.mark.parametrize("text", ["31.02.2000", "29.02.2001", "01.13.2000", "1.2.00", "tomorrow"])
def test_invalid_date(text):
    with pytest.raises(ValueError):
        parse_date(text)


def test_pin_is_string_and_does_not_guess():
    assert normalize("personal_identifier", "0012 3456789012") == "00123456789012"
    assert validate_field("personal_identifier", "00123456789012") is None
    assert validate_field("personal_identifier", "O0123456789012")
    assert normalize("passport_number", "aO 001") == "AO001"


@pytest.mark.parametrize(
    "name, expected",
    [
        ("CON", "document"),
        ("../x:y?", "_x_y_"),
        ("ok.xlsx", "ok.xlsx"),
        ("...", "document"),
        ("LPT1.txt", "document"),
    ],
)
def test_filenames(name, expected):
    assert safe_filename(name) == expected


def test_address_separation():
    address = Address(region="Регион", city="Город", street="Улица", house="001", apartment="2")
    assert address.display() == "Регион, Город, Улица, д. 001, кв. 2"
    assert address.legacy_street() == "Город, Улица, д. 001, кв. 2"


def test_confirmation_is_bound_to_data(case, profile):
    assert not readiness(case, profile)
    case.data.passport_number = "AA0000002"
    assert not case.confirmed
    assert readiness(case, profile)


def test_no_registration_living_address_substitution(case, profile):
    case.data.registration_address = Address()
    case.data.living_address = Address(raw="Город, улица, дом")
    assert any("регистрации" in error for error in readiness(case, profile))


def test_settings_corruption_is_visible(tmp_path):
    repo = SettingsRepository(tmp_path)
    repo.path.write_text("not json", encoding="utf-8")
    assert repo.load().languages == ["rus", "eng"]
    assert repo.warning


def test_settings_roundtrip(tmp_path):
    repo = SettingsRepository(tmp_path)
    settings = repo.load()
    settings.timeout = 42
    repo.save(settings)
    assert repo.load().timeout == 42


def test_folder_and_duplicate_collection(image_factory):
    path = image_factory()
    values, errors = collect_images([path, path.parent, path])
    assert values == [path.resolve()]
    assert not errors


def test_defaults_have_no_sensitive_cache():
    case = ClientCase()
    assert case.data.personal_identifier is None
    assert case.data.whatsapp_phone is None
