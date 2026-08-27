from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol


@dataclass(frozen=True)
class AutomotiveTerm:
    source: str
    target: str
    language_pair: str = "zh-CN>en"
    note: str = ""


class TranslationProvider(Protocol):
    """Future offline translation interface; no model is bundled in phase 1."""

    def translate(self, text: str, *, source_language: str, target_language: str) -> str:
        ...


class LocalTermDictionary:
    def __init__(self, terms: list[AutomotiveTerm] | None = None):
        self.terms = terms or []

    def lookup(self, text: str, language_pair: str = "zh-CN>en") -> str | None:
        normalized = text.strip().casefold()
        for term in self.terms:
            if term.language_pair == language_pair and term.source.strip().casefold() == normalized:
                return term.target
        return None
