from __future__ import annotations

import os
from pathlib import Path

import pytest
from PIL import Image, ImageDraw, ImageFont

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from document_excel.domain.models import Address, ClientCase, ClientData, RegistrationStatus
from document_excel.infrastructure.excel import create_demo_template, legacy_profile


@pytest.fixture
def profile():
    return legacy_profile()


@pytest.fixture
def case():
    value = ClientCase(
        data=ClientData(
            full_name="SYNTHETIC TEST",
            personal_identifier="00123456789012",
            passport_type="ID",
            passport_number="AA0000001",
            registration_address=Address(region="Тестовая область", street="ул. Пример, 1"),
        ),
        registration_status=RegistrationStatus.MANUALLY_ENTERED,
    )
    value.confirmed_digest = value.digest()
    return value


@pytest.fixture
def template(tmp_path):
    path = tmp_path / "template.xlsx"
    create_demo_template(path)
    return path


@pytest.fixture
def image_factory(tmp_path):
    def make(lines=None, name="synthetic.png", rotation=0, low_contrast=False):
        image = Image.new("RGB", (1600, 1100), "white" if not low_contrast else (190, 190, 190))
        font_path = next(
            (
                path
                for path in [
                    Path("C:/Windows/Fonts/arial.ttf"),
                    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
                ]
                if path.exists()
            ),
            None,
        )
        font = ImageFont.truetype(str(font_path), 36) if font_path else ImageFont.load_default(size=36)
        draw = ImageDraw.Draw(image)
        lines = (
            lines
            if lines is not None
            else [
                "KYRGYZ REPUBLIC PASSPORT",
                "FULL NAME: SYNTHETIC TEST",
                "SURNAME: TEST",
                "GIVEN NAME: SYNTHETIC",
                "DATE OF BIRTH: 01.02.1990",
                "PERSONAL NUMBER: 00123456789012",
                "PASSPORT NO: AA0000001",
                "DATE OF ISSUE: 03.04.2020",
                "REGISTRATION ADDRESS: TEST REGION, TEST STREET 1",
            ]
        )
        for i, text in enumerate(lines):
            draw.text((65, 65 + i * 95), text, fill=(135, 135, 135) if low_contrast else "black", font=font)
        if rotation:
            image = image.rotate(rotation, expand=True)
        path = tmp_path / name
        image.save(path)
        return path

    return make
