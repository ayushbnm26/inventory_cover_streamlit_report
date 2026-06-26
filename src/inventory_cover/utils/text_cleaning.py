"""Text and header normalization helpers."""

from __future__ import annotations

import re
from typing import Any


def normalize_header(value: Any) -> str:
    """Normalize a header for resilient matching."""

    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def is_blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def clean_text(value: Any, number_format: str = "") -> str:
    """Return a stable text value, preserving leading zeroes when possible."""

    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return _apply_zero_format(str(value), number_format)
    if isinstance(value, float):
        if value.is_integer():
            return _apply_zero_format(str(int(value)), number_format)
        return str(value).strip()
    return str(value).strip()


def _apply_zero_format(text: str, number_format: str) -> str:
    width = zero_format_width(number_format)
    if width > len(text) and text.isdigit():
        return text.zfill(width)
    return text


def zero_format_width(number_format: str) -> int:
    """Infer a fixed-width zero format such as 0000000000000."""

    if not number_format:
        return 0
    cleaned = number_format.strip()
    if re.fullmatch(r"0+", cleaned):
        return len(cleaned)
    return 0
