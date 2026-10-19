"""Conservative splitting of explicitly printed digital address fragments."""

import re

from document_excel.domain.models import Address


def digital_address_from_text(raw: str) -> Address:
    address = Address(raw=raw)
    extra = []
    # OCR often drops a comma between country/city or street/house. Explicit
    # printed prefixes still delimit fragments; the source string stays intact.
    fragments = re.sub(
        r"(?<!\w)(?=(?:(?:г\.|город\s|с\.|село\s|айыл\s|ул\.|улица\s|көчөсү\s|street\s)\s*\S|(?:дом|д\.|үй|кв\.|квартира|батир)\s*\d))",
        ";",
        raw,
        flags=re.I,
    )
    for part in re.split(r"[,;]", fragments):
        part = part.strip()
        if not part:
            continue
        field = (
            "country"
            if re.search(r"республик|republic", part, re.I)
            else "region"
            if re.search(r"обл(?:асть|\.)|облусу", part, re.I)
            else "city"
            if re.match(r"(?:г\.|город|с\.|село|айыл)\s*\S", part, re.I)
            else "district"
            if re.search(r"район|р-н|району", part, re.I)
            else "street"
            if re.match(r"(?:ул\.|улица|көчөсү|street)\s*\S", part, re.I)
            else "house"
            if re.match(r"(?:дом|д\.|үй)\s*\S", part, re.I)
            else "apartment"
            if re.match(r"(?:кв\.|квартира|батир)\s*\S", part, re.I)
            else ""
        )
        if field and not getattr(address, field):
            value = (
                re.sub(r"^(?:дом|д\.|үй|кв\.|квартира|батир)\s*", "", part, flags=re.I)
                if field in {"house", "apartment"}
                else part
            )
            setattr(address, field, value)
        else:
            extra.append(part)
    address.extra = ", ".join(extra)
    return address
