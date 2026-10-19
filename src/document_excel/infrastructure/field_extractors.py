"""Field-specific parsing; never a search for arbitrary words in the document."""

import re

from document_excel.domain.models import normalize, parse_date, validate_field

# Only optically identical glyphs, with an explicitly Cyrillic field context.
# Y/S/R/I are deliberately absent: they do not identify a unique KG letter.
_CYRILLIC_HOMOGLYPHS = str.maketrans("ABCEHKMOPTXaceopx", "АВСЕНКМОРТХасеорх")


def cyrillic_field_value(text: str) -> str:
    return text.translate(_CYRILLIC_HOMOGLYPHS)


class NamePartExtractor:
    field_name = ""

    def value(self, text, script):
        from document_excel.infrastructure.extraction import label_matches, name_script

        value = " ".join(text.split())
        if script == "cyrillic":
            value = cyrillic_field_value(value)
        if label_matches(value) or validate_field(self.field_name, value):
            return None
        if name_script(value) != script:
            return None
        return value


class SurnameExtractor(NamePartExtractor):
    field_name = "surname"


class FirstNameExtractor(NamePartExtractor):
    field_name = "first_name"


class PatronymicExtractor(NamePartExtractor):
    field_name = "patronymic"


class DateExtractor:
    fields = {"birth_date", "passport_issued_date", "expiration_date"}

    def value(self, field, text):
        if field not in self.fields:
            raise ValueError("Неподходящее поле даты")
        value = " ".join(text.split()).strip()
        if re.fullmatch(r"[0-9]{8}", value):
            value = value[:2] + "." + value[2:4] + "." + value[4:]
        # A whole date must be read. Multiple dates or adjacent field text are
        # not resolved by taking the first date-like fragment.
        return parse_date(value).isoformat()


class IssuingAuthorityExtractor:
    def value(self, text):
        value = " ".join(text.split()).upper()
        match = re.fullmatch(r"([А-ЯЁҮӨҢA-Z]{2,6})\s*([0-9]{6})", value)
        return match[1] + " " + match[2] if match else None


class GenderExtractor:
    def value(self, text):
        result = normalize("gender", text)
        return result if not validate_field("gender", result) else None


class CitizenshipExtractor:
    def value(self, text):
        result = normalize("citizenship", text)
        return result if not validate_field("citizenship", result) else None
