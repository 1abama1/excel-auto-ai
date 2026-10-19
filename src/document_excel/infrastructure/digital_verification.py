"""Generated mobile screenshots for repeatable public OCR/Excel checks."""

from importlib.resources import files
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def synthetic_digital_images(directory: Path, *, small=False, overlay=False, wrapped_names=False):
    directory.mkdir(parents=True, exist_ok=True)
    resource = files("document_excel").joinpath("resources/fonts/NotoSans-Regular.ttf")
    label_font = ImageFont.truetype(str(resource), 12)
    value_font = ImageFont.truetype(str(resource), 18)
    controls = {
        "surname_cyrillic": "СИНТЕТИКОВ",
        "surname_latin": "SINTETIKOV",
        "first_name_cyrillic": "ПРИМЕР",
        "first_name_latin": "PRIMER",
        "patronymic_cyrillic": "ТЕСТОВИЧ",
        "full_name": "СИНТЕТИКОВ ПРИМЕР ТЕСТОВИЧ",
        "birth_date": "1990-11-07",
        "personal_identifier": "00123456789012",
        "document_number": "ID0000001",
        "passport_issued_by": "МКК",
        "passport_department_code": "012345",
        "passport_issued_date": "2024-11-22",
        "expiration_date": "2034-11-22",
        "gender": "MALE",
        "citizenship": "KGZ",
    }

    def base(page):
        image = Image.new("RGB", (720, 1600), "#bedffc")
        draw = ImageDraw.Draw(image)
        draw.text((35, 20), "10:20       4G       70%", font=value_font, fill="black")
        draw.text((40, 80), "Идентификационная карта", font=value_font, fill="black")
        draw.rounded_rectangle((45, 300, 675, 1000), radius=40, fill="white")
        draw.text((75, 324), "КЫРГЫЗ РЕСПУБЛИКАСЫ", font=label_font, fill="#16637c")
        draw.text((75, 343), "THE KYRGYZ REPUBLIC", font=label_font, fill="#16637c")
        draw.text((365, 324), "ИДЕНТИФИКАЦИЯЛЫК КАРТА", font=label_font, fill="#16637c")
        draw.text((365, 343), "IDENTITY CARD", font=label_font, fill="#16637c")
        draw.text((640, 960), str(page), font=value_font, fill="black")
        draw.rounded_rectangle((0, 1410, 720, 1590), radius=35, fill="#f4f6f8")
        for row in range(23):
            for col in range(23):
                if (row * 17 + col * 31 + row * col) % 7 < 3:
                    draw.rectangle(
                        (290 + col * 6, 1360 + row * 6, 295 + col * 6, 1365 + row * 6), fill="#145b9c"
                    )
        draw.text((80, 1540), "Потяните для просмотра QR кода", font=value_font, fill="black")
        return image, draw

    front, draw = base(1)
    draw.rounded_rectangle((80, 415, 235, 680), radius=15, fill="#d6dfe8")
    draw.text((102, 520), "ПРИМЕР", font=value_font, fill="#6e7e8e")
    rows = [
        (430, "Аты / Имя / Name", "ПРИМЕР / PRIMER"),
        (490, "Фамилиясы / Фамилия / Surname", "СИНТЕТИКОВ / SINTETIKOV"),
        (550, "Атасынын аты / Отчество / Patronymic", "ТЕСТОВИЧ"),
        (610, "Жынысы / Пол / Sex", "Э / М"),
        (670, "Жарандыгы / Гражданство / Citizenship", "Кыргыз Республикасы"),
        (730, "Улуту / Национальность / Nationality", "ТЕСТОВАЯ НАЦИОНАЛЬНОСТЬ"),
    ]
    for y, label, value in rows:
        draw.text((270, y), label, font=label_font, fill="#343434")
        if wrapped_names and y in {430, 490}:
            a, b = value.split(" / ")
            draw.text((270, y + 18), a, font=value_font, fill="black")
            draw.text((270, y + 38), b, font=label_font, fill="black")
        else:
            draw.text((270, y + 23), value, font=value_font, fill="black")
    draw.text((270, 810), "Туулган күнү / Дата рождения", font=label_font, fill="#343434")
    draw.text((270, 829), "/ Date of birth", font=label_font, fill="#343434")
    draw.text((270, 855), "07.11.1990", font=value_font, fill="black")
    if overlay:
        draw.rounded_rectangle((648, 380, 710, 810), radius=20, fill="#f9fafb", outline="#d1d5db", width=2)
        draw.text((665, 390), "...", font=value_font, fill="black")
    back, draw = base(2)
    for y, label, translation, value in [
        (410, "Берген мекеме / Орган выдачи / Authority", None, "МКК 012345"),
        (485, "Берилген күнү / Дата выдачи", "/ Date of issue", "22.11.2024"),
        (575, "Жеке номери / Персональный номер", "/ Personal number", "00123456789012"),
        (665, "Документтин № / № Документа", "/ Document №", "ID0000001"),
        (755, "Колдонуу мөөнөтү / Срок действия", "/ Date of expiry", "22.11.2034"),
    ]:
        draw.text((85, y), label, font=label_font, fill="#343434")
        if translation:
            draw.text((85, y + 19), translation, font=label_font, fill="#343434")
        draw.text((85, y + (46 if translation else 24)), value, font=value_font, fill="black")
    draw.text((85, 855), "Катталган жери / Адрес прописки / Address", font=label_font, fill="#343434")
    draw.text((85, 880), "Кыргызская Республика, г. Тестовый, Тестовый район,", font=value_font, fill="black")
    draw.text((85, 910), "улица Примерная, дом 7", font=value_font, fill="black")
    paths = [directory / "synthetic-digital-front.png", directory / "synthetic-digital-back.png"]
    for image, path in zip((front, back), paths, strict=True):
        if small:
            image = image.resize((540, 1200), Image.Resampling.LANCZOS)
        image.save(path)
    return paths, controls
