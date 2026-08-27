from __future__ import annotations

import hashlib
import math
import re
import unicodedata
from typing import Any


PLACEHOLDERS = {"", "-", "—", "–", "n/a", "na", "none", "null", "无"}


def clean_text(value: Any) -> str | None:
    if value is None:
        return None
    text = unicodedata.normalize("NFKC", str(value)).strip()
    text = re.sub(r"[ \t]+", " ", text)
    text = text.replace("×", "*").replace("✕", "*").replace("（", "(").replace("）", ")")
    if text.casefold() in PLACEHOLDERS:
        return None
    return text


def identifier(value: Any) -> str | None:
    text = clean_text(value)
    if text is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return text


def normalize_header(value: Any) -> str:
    text = clean_text(value) or ""
    text = text.casefold().replace("\n", "")
    return re.sub(r"[\s_\-—–/\\()（）\[\]【】.:：]", "", text)


def fingerprint(headers: list[str]) -> str:
    normalized = "|".join(normalize_header(h) for h in headers)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:20]


def as_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return None if math.isnan(number) or math.isinf(number) else number
    text = clean_text(value)
    if text is None:
        return None
    match = re.search(r"[-+]?\d+(?:[.,]\d+)?", text.replace(",", ""))
    return float(match.group(0)) if match else None


def display_number(value: float | None) -> int | float | None:
    if value is None:
        return None
    return int(value) if value.is_integer() else round(value, 12)


def combine_unique(parts: list[Any]) -> str | None:
    result: list[str] = []
    keys: list[str] = []
    for value in parts:
        text = clean_text(value)
        if not text:
            continue
        key = re.sub(r"\s+", "", text).casefold()
        if any(key == existing or key in existing for existing in keys):
            continue
        if any(existing in key for existing in keys):
            result = [old for old, old_key in zip(result, keys) if old_key not in key]
            keys = [old_key for old_key in keys if old_key not in key]
        result.append(text)
        keys.append(key)
    return " ".join(result) if result else None


def column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def split_cell_reference(reference: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", reference.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {reference}")
    col = 0
    for char in match.group(1):
        col = col * 26 + ord(char) - 64
    return int(match.group(2)), col
