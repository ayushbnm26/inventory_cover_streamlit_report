from __future__ import annotations

from datetime import date

from openpyxl.utils.datetime import to_excel

from inventory_cover.utils.date_parsing import parse_po_date


def test_window_start_uses_day_first_policy() -> None:
    result = parse_po_date("22/6/2026", "Window Start")

    assert result.ok
    assert result.value == date(2026, 6, 22)


def test_window_end_uses_day_first_policy() -> None:
    result = parse_po_date("13/7/2026", "Window End")

    assert result.ok
    assert result.value == date(2026, 7, 13)


def test_expected_date_uses_month_first_policy() -> None:
    result = parse_po_date("06/08/2026", "Expected Date")

    assert result.ok
    assert result.value == date(2026, 6, 8)


def test_expected_date_ambiguous_month_first_policy() -> None:
    result = parse_po_date("06/01/2026", "Expected Date")

    assert result.ok
    assert result.value == date(2026, 6, 1)


def test_excel_serial_date_parsing() -> None:
    serial = to_excel(date(2026, 6, 22))

    result = parse_po_date(serial, "Window Start")

    assert result.ok
    assert result.value == date(2026, 6, 22)
