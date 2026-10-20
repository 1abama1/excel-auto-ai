"""Presentation of official names, independent of OCR normalization."""

import re


class PersonNameFormatter:
    @staticmethod
    def format(value: str | None, *, preserve: bool = False) -> str:
        value = " ".join((value or "").split())
        if preserve:
            return value

        # Keep mixed case (McDonald, deLaCruz) intact. Each hyphen/apostrophe
        # component gets its own initial; Unicode includes Ү, Ө and Ң.
        def component(match: re.Match) -> str:
            word = match.group()
            return word[0].upper() + word[1:].lower() if word.isupper() or word.islower() else word

        return re.sub(r"[^\W\d_]+", component, value, flags=re.UNICODE)
