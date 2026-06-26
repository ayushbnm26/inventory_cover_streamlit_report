"""Numeric parsing helpers for Amazon export fields."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import math
import re


@dataclass(frozen=True)
class NumericParseResult:
    value: int | float | None
    raw_value: Any
    ok: bool
    was_blank: bool
    currency: str | None = None
    detail: str = ""


def parse_number(value: Any) -> NumericParseResult:
    """Parse quantities and costs without crashing on bad source values."""

    if value is None or str(value).strip() == "":
        return NumericParseResult(
            value=None,
            raw_value=value,
            ok=True,
            was_blank=True,
            detail="blank",
        )
    if isinstance(value, bool):
        return NumericParseResult(
            value=None,
            raw_value=value,
            ok=False,
            was_blank=False,
            detail="boolean is not a valid number",
        )
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return NumericParseResult(
                value=None,
                raw_value=value,
                ok=False,
                was_blank=False,
                detail="number is not finite",
            )
        return NumericParseResult(
            value=_simplify_number(float(value)),
            raw_value=value,
            ok=True,
            was_blank=False,
        )

    text = str(value).strip()
    currency = _extract_currency(text)
    cleaned = re.sub(r"(?i)\binr\b", "", text)
    cleaned = cleaned.replace("\u20b9", "")
    cleaned = cleaned.replace(",", "").strip()
    negative = cleaned.startswith("(") and cleaned.endswith(")")
    if negative:
        cleaned = cleaned[1:-1].strip()

    if not re.fullmatch(r"[+-]?\d+(\.\d+)?", cleaned):
        return NumericParseResult(
            value=None,
            raw_value=value,
            ok=False,
            was_blank=False,
            currency=currency,
            detail=f"could not parse numeric value '{text}'",
        )

    parsed = float(cleaned)
    if negative:
        parsed = -parsed
    return NumericParseResult(
        value=_simplify_number(parsed),
        raw_value=value,
        ok=True,
        was_blank=False,
        currency=currency,
    )


def numbers_differ(left: int | float | None, right: int | float | None, tolerance: float = 0.0001) -> bool:
    if left is None or right is None:
        return False
    return abs(float(left) - float(right)) > tolerance


def _simplify_number(value: float) -> int | float:
    if value.is_integer():
        return int(value)
    return value


def _extract_currency(text: str) -> str | None:
    upper = text.upper()
    if "INR" in upper or "\u20b9" in text:
        return "INR"
    return None
