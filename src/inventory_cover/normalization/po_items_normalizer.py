"""Normalize raw Amazon PO Items rows into the standard master schema."""

from __future__ import annotations

from typing import Any

from inventory_cover.schemas import (
    COST_FIELDS,
    DATE_FIELDS,
    MASTER_HEADERS,
    NUMERIC_FIELDS,
    QUANTITY_FIELDS,
    REQUIRED_FIELDS,
    TEXT_FIELDS,
    CellValue,
    NormalizedPoItemRow,
    RawPoItemRow,
    ValidationIssue,
)
from inventory_cover.utils.date_parsing import parse_po_date
from inventory_cover.utils.numbers import NumericParseResult, numbers_differ, parse_number
from inventory_cover.utils.text_cleaning import clean_text


def normalize_po_item_row(
    raw_row: RawPoItemRow,
    run_id: str,
    default_currency: str = "INR",
) -> NormalizedPoItemRow:
    """Normalize one raw PO Items row and attach row-level validation issues."""

    data: dict[str, Any] = {header: None for header in MASTER_HEADERS}
    data.update(
        {
            "Run ID": run_id,
            "Source File": raw_row.source_file,
            "Source Sheet": raw_row.source_sheet,
            "Source Row": raw_row.source_row,
        }
    )
    issues: list[ValidationIssue] = []
    parsed_numbers: dict[str, NumericParseResult] = {}

    for field, cell in raw_row.values.items():
        if field in TEXT_FIELDS:
            data[field] = clean_text(cell.value, cell.number_format)
        elif field in DATE_FIELDS:
            parsed = parse_po_date(cell.value, field)
            data[field] = parsed.value
            if not parsed.ok:
                issues.append(
                    _issue(
                        "WARNING",
                        "DATE_PARSE_FAILED",
                        raw_row,
                        field,
                        cell.value,
                        parsed.detail,
                        "Kept raw date value in output.",
                    )
                )
        elif field in NUMERIC_FIELDS:
            parsed_number = parse_number(cell.value)
            parsed_numbers[field] = parsed_number
            data[field] = parsed_number.value if parsed_number.ok else cell.value
            if not parsed_number.ok:
                issues.append(
                    _issue(
                        "ERROR",
                        "NUMERIC_PARSE_FAILED",
                        raw_row,
                        field,
                        cell.value,
                        parsed_number.detail,
                        "Kept raw value in output; excluded from derived calculations.",
                    )
                )

    _add_missing_identifier_warnings(raw_row, data, issues)
    _add_currency_fields(data, parsed_numbers, default_currency)
    _add_open_po_fields(data, parsed_numbers, raw_row, issues)

    row = NormalizedPoItemRow(data=data, issues=issues)
    row.refresh_validation_status()
    return row


def _add_missing_identifier_warnings(
    raw_row: RawPoItemRow,
    data: dict[str, Any],
    issues: list[ValidationIssue],
) -> None:
    for field in ("PO", "Vendor Code", "ASIN"):
        if not data.get(field):
            issues.append(
                _issue(
                    "WARNING",
                    f"MISSING_{field.upper().replace(' ', '_')}",
                    raw_row,
                    field,
                    None,
                    f"{field} is blank on a non-empty source row.",
                    "Row kept for traceability.",
                )
            )


def _add_currency_fields(
    data: dict[str, Any],
    parsed_numbers: dict[str, NumericParseResult],
    default_currency: str,
) -> None:
    data["Unit Cost Currency"] = parsed_numbers.get("Unit Cost", NumericParseResult(None, None, True, True)).currency
    data["Total Cost Currency"] = parsed_numbers.get("Total Cost", NumericParseResult(None, None, True, True)).currency
    if data["Unit Cost"] is not None and not data["Unit Cost Currency"]:
        data["Unit Cost Currency"] = default_currency
    if data["Total Cost"] is not None and not data["Total Cost Currency"]:
        data["Total Cost Currency"] = default_currency


def _add_open_po_fields(
    data: dict[str, Any],
    parsed_numbers: dict[str, NumericParseResult],
    raw_row: RawPoItemRow,
    issues: list[ValidationIssue],
) -> None:
    accepted = _usable_number(parsed_numbers.get("Quantity Accepted"))
    received = _usable_number(parsed_numbers.get("Quantity Received"))
    outstanding = _usable_number(parsed_numbers.get("Quantity Outstanding"))
    unit_cost = _usable_number(parsed_numbers.get("Unit Cost"))

    received_normalized = 0 if received is None else received
    outstanding_normalized = 0 if outstanding is None else outstanding
    data["Quantity Received Normalized"] = received_normalized
    data["Quantity Outstanding Normalized"] = outstanding_normalized
    data["Open PO Qty - Source"] = outstanding_normalized

    derived = None
    if accepted is not None:
        derived = max(accepted - received_normalized, 0)
    data["Open PO Qty - Derived"] = derived

    outstanding_result = parsed_numbers.get("Quantity Outstanding")
    outstanding_present = outstanding_result is not None and not outstanding_result.was_blank and outstanding_result.ok
    data["Open PO Qty - Final"] = outstanding if outstanding_present else derived

    if accepted is not None and received is not None and received > accepted:
        issues.append(
            _issue(
                "WARNING",
                "RECEIVED_EXCEEDS_ACCEPTED",
                raw_row,
                "Quantity Received",
                raw_row.values.get("Quantity Received", CellValue(None)).value,
                "Quantity Received is greater than Quantity Accepted.",
                "Row kept; open PO derived quantity floored at zero.",
            )
        )

    if outstanding_present and derived is not None and numbers_differ(outstanding, derived):
        issues.append(
            _issue(
                "WARNING",
                "OUTSTANDING_DIFFERS_FROM_DERIVED",
                raw_row,
                "Quantity Outstanding",
                raw_row.values.get("Quantity Outstanding", CellValue(None)).value,
                f"Source outstanding {outstanding} differs from accepted minus received {derived}.",
                "Used source outstanding for final open PO quantity.",
            )
        )

    final_qty = _usable_data_number(data.get("Open PO Qty - Final"))
    data["Open PO Value - Final"] = final_qty * unit_cost if final_qty is not None and unit_cost is not None else None


def _usable_number(result: NumericParseResult | None) -> int | float | None:
    if result is None or not result.ok:
        return None
    return result.value


def _usable_data_number(value: Any) -> int | float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    return None


def _issue(
    severity: str,
    issue_type: str,
    raw_row: RawPoItemRow,
    field_name: str,
    raw_value: Any,
    detail: str,
    action: str,
) -> ValidationIssue:
    po = clean_text(raw_row.values.get("PO", CellValue(None)).value)
    asin = clean_text(raw_row.values.get("ASIN", CellValue(None)).value)
    return ValidationIssue(
        severity=severity,
        issue_type=issue_type,
        source_file=raw_row.source_file,
        source_sheet=raw_row.source_sheet,
        source_row=raw_row.source_row,
        po=po,
        asin=asin,
        field_name=field_name,
        raw_value=raw_value,
        issue_detail=detail,
        action_taken=action,
    )
