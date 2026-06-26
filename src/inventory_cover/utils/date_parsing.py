"""Column-specific date parsing for Amazon PO Items exports."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from openpyxl.utils.datetime import from_excel


WINDOW_DATE_FORMATS: tuple[str, ...] = ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y")
EXPECTED_DATE_FORMATS: tuple[str, ...] = ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y")
ISO_DATE_FORMATS: tuple[str, ...] = ("%Y-%m-%d", "%Y/%m/%d")
B2B_DISPATCH_DATE_FORMATS: tuple[str, ...] = (
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%Y-%m-%d",
    "%d-%m-%y",
    "%d/%m/%y",
)
VENDOR_CENTRAL_DATE_FORMATS: tuple[str, ...] = (
    "%d/%m/%y",
    "%d-%m-%y",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%Y-%m-%d",
    "%Y-%m-%d %H:%M:%S",
)


@dataclass(frozen=True)
class DateParseResult:
    value: date | str | None
    raw_value: Any
    ok: bool
    was_blank: bool
    detail: str = ""


def parse_po_date(value: Any, field_name: str) -> DateParseResult:
    """Parse a PO date using the policy for its specific source column."""

    if value is None or str(value).strip() == "":
        return DateParseResult(value=None, raw_value=value, ok=True, was_blank=True, detail="blank")
    if isinstance(value, datetime):
        return DateParseResult(value=value.date(), raw_value=value, ok=True, was_blank=False)
    if isinstance(value, date):
        return DateParseResult(value=value, raw_value=value, ok=True, was_blank=False)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            parsed = from_excel(value)
            if isinstance(parsed, datetime):
                parsed = parsed.date()
            if isinstance(parsed, date):
                return DateParseResult(value=parsed, raw_value=value, ok=True, was_blank=False)
        except (ValueError, TypeError, OverflowError) as exc:
            return DateParseResult(
                value=value,
                raw_value=value,
                ok=False,
                was_blank=False,
                detail=f"invalid Excel serial date: {exc}",
            )

    text = str(value).strip()
    formats = _formats_for_field(field_name)
    for fmt in formats + ISO_DATE_FORMATS:
        try:
            return DateParseResult(
                value=datetime.strptime(text, fmt).date(),
                raw_value=value,
                ok=True,
                was_blank=False,
            )
        except ValueError:
            continue

    return DateParseResult(
        value=text,
        raw_value=value,
        ok=False,
        was_blank=False,
        detail=f"could not parse '{text}' using {field_name} date policy",
    )


def parse_b2b_dispatch_date(value: Any) -> DateParseResult:
    """Parse Dispatch Tracker dates with a strict day-first policy."""

    if value is None or str(value).strip() == "":
        return DateParseResult(value=None, raw_value=value, ok=True, was_blank=True, detail="blank")
    if isinstance(value, datetime):
        return DateParseResult(value=value.date(), raw_value=value, ok=True, was_blank=False)
    if isinstance(value, date):
        return DateParseResult(value=value, raw_value=value, ok=True, was_blank=False)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            parsed = from_excel(value)
            if isinstance(parsed, datetime):
                parsed = parsed.date()
            if isinstance(parsed, date):
                return DateParseResult(value=parsed, raw_value=value, ok=True, was_blank=False)
        except (ValueError, TypeError, OverflowError) as exc:
            return DateParseResult(
                value=value,
                raw_value=value,
                ok=False,
                was_blank=False,
                detail=f"invalid Excel serial date: {exc}",
            )

    text = str(value).strip()
    if _looks_like_multi_date(text):
        return DateParseResult(
            value=text,
            raw_value=value,
            ok=False,
            was_blank=False,
            detail="multiple dates found; ambiguous Dispatch Date not guessed",
        )
    if not _looks_like_supported_b2b_date(text):
        return DateParseResult(
            value=text,
            raw_value=value,
            ok=False,
            was_blank=False,
            detail=f"unsupported Dispatch Date format '{text}'",
        )

    for fmt in B2B_DISPATCH_DATE_FORMATS:
        try:
            return DateParseResult(
                value=datetime.strptime(text, fmt).date(),
                raw_value=value,
                ok=True,
                was_blank=False,
            )
        except ValueError:
            continue

    return DateParseResult(
        value=text,
        raw_value=value,
        ok=False,
        was_blank=False,
        detail=f"invalid Dispatch Date '{text}'",
    )


def parse_vendor_central_date(value: Any) -> DateParseResult:
    """Parse Vendor Central dates with a day-first policy."""

    if value is None or str(value).strip() == "":
        return DateParseResult(value=None, raw_value=value, ok=True, was_blank=True, detail="blank")
    if isinstance(value, datetime):
        return DateParseResult(value=value.date(), raw_value=value, ok=True, was_blank=False)
    if isinstance(value, date):
        return DateParseResult(value=value, raw_value=value, ok=True, was_blank=False)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            parsed = from_excel(value)
            if isinstance(parsed, datetime):
                parsed = parsed.date()
            if isinstance(parsed, date):
                return DateParseResult(value=parsed, raw_value=value, ok=True, was_blank=False)
        except (ValueError, TypeError, OverflowError) as exc:
            return DateParseResult(
                value=value,
                raw_value=value,
                ok=False,
                was_blank=False,
                detail=f"invalid Excel serial date: {exc}",
            )

    text = str(value).strip()
    for fmt in VENDOR_CENTRAL_DATE_FORMATS:
        try:
            return DateParseResult(
                value=datetime.strptime(text, fmt).date(),
                raw_value=value,
                ok=True,
                was_blank=False,
            )
        except ValueError:
            continue

    return DateParseResult(
        value=text,
        raw_value=value,
        ok=False,
        was_blank=False,
        detail=f"could not parse Vendor Central date '{text}'",
    )


def _looks_like_multi_date(text: str) -> bool:
    return "&" in text or "," in text or " and " in text.lower()


def _looks_like_supported_b2b_date(text: str) -> bool:
    import re

    return bool(
        re.fullmatch(r"\d{1,2}[-/]\d{1,2}[-/]\d{2}|\d{1,2}[-/]\d{1,2}[-/]\d{4}", text)
        or re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", text)
    )


def _formats_for_field(field_name: str) -> tuple[str, ...]:
    if field_name in {"Window Start", "Window End"}:
        return WINDOW_DATE_FORMATS
    if field_name == "Expected Date":
        return EXPECTED_DATE_FORMATS
    raise ValueError(f"No date parsing policy configured for {field_name}")
