"""Names are generated at save time; publication remains WorkbookService's job."""

import re
from collections.abc import Callable
from datetime import date, datetime
from pathlib import Path
from string import Formatter

from document_excel.domain.errors import AppError
from document_excel.domain.person_name import PersonNameFormatter

DEFAULT_FILENAME_PATTERN = "{full_name} - {date}"


def local_save_date() -> date:
    return datetime.now().astimezone().date()


class ContractFilenameGenerator:
    def __init__(self, clock: Callable[[], date] = local_save_date):
        self.clock = clock

    def generate(self, name: str, *, preserve: bool = False, pattern: str = DEFAULT_FILENAME_PATTERN) -> str:
        name = PersonNameFormatter.format(name, preserve=preserve)
        if not name or not any(c.isalpha() for c in name):
            name = "Договор"
        try:
            if len(pattern) > 100:
                raise ValueError
            fields = list(Formatter().parse(pattern))
            if {field for _, field, _, _ in fields if field} != {"full_name", "date"} or any(
                spec or conversion for _, _, spec, conversion in fields
            ):
                raise ValueError
            text = pattern.format(full_name=name[:120].rstrip(), date=self.clock().strftime("%d.%m.%Y"))
        except (KeyError, ValueError, IndexError, AttributeError) as exc:
            raise AppError("Шаблон имени должен содержать только {full_name} и {date}.") from exc
        # Strip Windows path syntax and controls; never accept a path from a name.
        text = re.sub(r'[<>:"/\\|?*\x00-\x1f\x7f]', "", text)
        text = " ".join(text.split()).strip(" .")[:180].rstrip(" .")
        if not text or re.fullmatch(r"(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?", text):
            raise AppError("Укажите безопасное имя файла.")
        return text + ".xlsx"

    @staticmethod
    def available(path: Path, *, current: Path | None = None) -> Path:
        """Suggestion only. os.link still arbitrates concurrent publications."""
        candidate = path
        number = 2
        while candidate.exists() and (current is None or candidate.resolve() != current.resolve()):
            candidate = path.with_name(f"{path.stem} ({number}){path.suffix}")
            number += 1
        return candidate
