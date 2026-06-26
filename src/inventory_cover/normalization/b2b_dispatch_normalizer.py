"""Normalize B2B Dispatch Tracker rows into the backend audit schema."""

from __future__ import annotations

from datetime import date
from typing import Any

from inventory_cover.b2b_dispatch_schemas import (
    B2BCellValue,
    B2B_MASTER_HEADERS,
    B2B_NUMERIC_FIELDS,
    B2B_RECOMMENDED_FIELDS,
    B2B_SOURCE_FIELDS,
    B2BValidationIssue,
    B2BRowProcessingResult,
    NormalizedB2BDispatchRow,
    RawB2BDispatchRow,
)
from inventory_cover.utils.date_parsing import parse_b2b_dispatch_date
from inventory_cover.utils.numbers import NumericParseResult, parse_number
from inventory_cover.utils.text_cleaning import clean_text


def normalize_b2b_dispatch_row(
    raw_row: RawB2BDispatchRow,
    run_id: str,
    lookback_start: date,
    lookback_end: date,
    value_difference_tolerance: float = 1.0,
) -> B2BRowProcessingResult:
    """Normalize one raw dispatch row and decide whether it is written."""

    data: dict[str, Any] = {header: None for header in B2B_MASTER_HEADERS}
    data.update(
        {
            "Run ID": run_id,
            "Source File": raw_row.source_file,
            "Source Sheet": raw_row.source_sheet,
            "Source Row": raw_row.source_row,
            "Source Channel": raw_row.source_channel,
            "Included In Lookback Window": False,
        }
    )
    issues: list[B2BValidationIssue] = []

    po = _text(raw_row, "PO")
    asin = _text(raw_row, "ASIN")
    invoice_no = _text(raw_row, "Invoice No")
    data["PO"] = po
    data["ASIN"] = asin
    data["Invoice No"] = invoice_no

    if not po:
        issues.append(_issue("ERROR", "MISSING_PO", raw_row, run_id, "PO", _raw(raw_row, "PO"), "PO is blank.", "Row rejected."))
    if not asin:
        issues.append(
            _issue("ERROR", "MISSING_ASIN", raw_row, run_id, "ASIN", _raw(raw_row, "ASIN"), "ASIN is blank.", "Row rejected.")
        )

    dispatch_qty_result = parse_number(_raw(raw_row, "Dispatch Qty"))
    if dispatch_qty_result.was_blank:
        issues.append(
            _issue(
                "ERROR",
                "MISSING_DISPATCH_QTY",
                raw_row,
                run_id,
                "Dispatch Qty",
                dispatch_qty_result.raw_value,
                "Dispatch Qty is blank.",
                "Row rejected.",
            )
        )
    elif not dispatch_qty_result.ok:
        issues.append(
            _issue(
                "ERROR",
                "INVALID_DISPATCH_QTY",
                raw_row,
                run_id,
                "Dispatch Qty",
                dispatch_qty_result.raw_value,
                dispatch_qty_result.detail,
                "Row rejected.",
            )
        )
    else:
        data["Dispatch Qty"] = dispatch_qty_result.value

    dispatch_date_result = parse_b2b_dispatch_date(_raw(raw_row, "Dispatch Date"))
    has_valid_dispatch_date = False
    dispatch_date: date | None = None
    if dispatch_date_result.was_blank:
        issues.append(
            _issue(
                "ERROR",
                "MISSING_DISPATCH_DATE",
                raw_row,
                run_id,
                "Dispatch Date",
                dispatch_date_result.raw_value,
                "Dispatch Date is blank.",
                "Row rejected.",
            )
        )
    elif not dispatch_date_result.ok or not isinstance(dispatch_date_result.value, date):
        issues.append(
            _issue(
                "ERROR",
                "INVALID_DISPATCH_DATE",
                raw_row,
                run_id,
                "Dispatch Date",
                dispatch_date_result.raw_value,
                dispatch_date_result.detail,
                "Row rejected.",
            )
        )
    else:
        has_valid_dispatch_date = True
        dispatch_date = dispatch_date_result.value
        data["Dispatch Date"] = dispatch_date

    if _has_error(issues):
        return B2BRowProcessingResult(
            normalized_row=None,
            issues=issues,
            has_valid_dispatch_date=has_valid_dispatch_date,
            included_in_window=False,
            rejected=True,
            excluded_outside_window=False,
        )

    if dispatch_date is None or not (lookback_start <= dispatch_date <= lookback_end):
        outside_issue = _issue(
            "INFO",
            "OUTSIDE_LOOKBACK_WINDOW",
            raw_row,
            run_id,
            "Dispatch Date",
            dispatch_date,
            f"Dispatch Date is outside {lookback_start.isoformat()} through {lookback_end.isoformat()}.",
            "Row excluded from master output.",
        )
        return B2BRowProcessingResult(
            normalized_row=None,
            issues=[outside_issue],
            has_valid_dispatch_date=has_valid_dispatch_date,
            included_in_window=False,
            rejected=False,
            excluded_outside_window=True,
        )

    data["Included In Lookback Window"] = True
    parsed_numbers = _populate_source_fields(data, raw_row, run_id, issues)
    _validate_po_asin_key(data, raw_row, run_id, issues)
    _add_recommended_field_warnings(data, raw_row, run_id, issues)
    _add_dispatch_value_fields(data, raw_row, run_id, issues, parsed_numbers, value_difference_tolerance)

    row = NormalizedB2BDispatchRow(data=data, issues=issues)
    row.refresh_validation_status()
    return B2BRowProcessingResult(
        normalized_row=row,
        issues=issues,
        has_valid_dispatch_date=has_valid_dispatch_date,
        included_in_window=True,
        rejected=False,
        excluded_outside_window=False,
    )


def _populate_source_fields(
    data: dict[str, Any],
    raw_row: RawB2BDispatchRow,
    run_id: str,
    issues: list[B2BValidationIssue],
) -> dict[str, NumericParseResult]:
    parsed_numbers: dict[str, NumericParseResult] = {}
    for field in B2B_SOURCE_FIELDS:
        cell = raw_row.values.get(field, B2BCellValue(None))
        if field in {"PO", "ASIN", "Invoice No", "Dispatch Qty", "Dispatch Date"}:
            continue
        if field in {"Appointment ID", "Ship To Location", "PO ASIN Key", "Model Number", "Dispatch Location"}:
            data[field] = clean_text(cell.value, cell.number_format)
        elif field == "PO Date":
            parsed = parse_b2b_dispatch_date(cell.value)
            if parsed.was_blank:
                data[field] = None
            elif parsed.ok:
                data[field] = parsed.value
            else:
                data[field] = parsed.value
                issues.append(
                    _issue(
                        "WARNING",
                        "INVALID_PO_DATE",
                        raw_row,
                        run_id,
                        field,
                        cell.value,
                        parsed.detail,
                        "Row kept; PO Date left as source text.",
                    )
                )
        elif field in B2B_NUMERIC_FIELDS:
            parsed_number = parse_number(cell.value)
            parsed_numbers[field] = parsed_number
            if parsed_number.ok:
                data[field] = parsed_number.value
            else:
                data[field] = None
                issues.append(
                    _issue(
                        "WARNING",
                        f"INVALID_{_issue_suffix(field)}",
                        raw_row,
                        run_id,
                        field,
                        cell.value,
                        parsed_number.detail,
                        "Row kept; field excluded from numeric calculations.",
                    )
                )
    return parsed_numbers


def _validate_po_asin_key(
    data: dict[str, Any],
    raw_row: RawB2BDispatchRow,
    run_id: str,
    issues: list[B2BValidationIssue],
) -> None:
    po = str(data.get("PO") or "")
    asin = str(data.get("ASIN") or "")
    expected_key = f"{po}{asin}"
    source_key = str(data.get("PO ASIN Key") or "").strip()
    if not source_key:
        data["PO ASIN Key"] = expected_key
        issues.append(
            _issue(
                "WARNING",
                "PO_ASIN_KEY_DERIVED",
                raw_row,
                run_id,
                "PO ASIN Key",
                _raw(raw_row, "PO ASIN Key"),
                "PO ASIN Key is blank.",
                "Derived as PO + ASIN.",
            )
        )
        return
    if source_key != expected_key:
        issues.append(
            _issue(
                "WARNING",
                "PO_ASIN_KEY_MISMATCH",
                raw_row,
                run_id,
                "PO ASIN Key",
                source_key,
                f"PO ASIN Key does not equal derived key {expected_key}.",
                "Source value kept.",
            )
        )


def _add_recommended_field_warnings(
    data: dict[str, Any],
    raw_row: RawB2BDispatchRow,
    run_id: str,
    issues: list[B2BValidationIssue],
) -> None:
    already_reported = {issue.field_name for issue in issues}
    for field in B2B_RECOMMENDED_FIELDS:
        if field in already_reported:
            continue
        value = data.get(field)
        if value is None or str(value).strip() == "":
            issues.append(
                _issue(
                    "WARNING",
                    f"MISSING_{_issue_suffix(field)}",
                    raw_row,
                    run_id,
                    field,
                    _raw(raw_row, field),
                    f"{field} is blank.",
                    "Row kept with warning.",
                )
            )


def _add_dispatch_value_fields(
    data: dict[str, Any],
    raw_row: RawB2BDispatchRow,
    run_id: str,
    issues: list[B2BValidationIssue],
    parsed_numbers: dict[str, NumericParseResult],
    tolerance: float,
) -> None:
    dispatch_qty = _number_or_none(data.get("Dispatch Qty"))
    unit_value = _number_or_none(data.get("Unit Value"))
    source_value = _number_or_none(data.get("Dispatch Value Source"))
    if dispatch_qty is not None and unit_value is not None:
        data["Dispatch Value Derived"] = dispatch_qty * unit_value
    derived = _number_or_none(data.get("Dispatch Value Derived"))
    if source_value is not None and derived is not None:
        difference = source_value - derived
        data["Dispatch Value Difference"] = difference
        if abs(float(difference)) > tolerance:
            issues.append(
                _issue(
                    "WARNING",
                    "DISPATCH_VALUE_MISMATCH",
                    raw_row,
                    run_id,
                    "Dispatch Value Source",
                    _raw(raw_row, "Dispatch Value Source"),
                    f"Source value differs from Dispatch Qty * Unit Value by {difference}.",
                    "Row kept; difference recorded.",
                )
            )


def _text(raw_row: RawB2BDispatchRow, field: str) -> str:
    cell = raw_row.values.get(field, B2BCellValue(None))
    return clean_text(cell.value, cell.number_format)


def _raw(raw_row: RawB2BDispatchRow, field: str) -> Any:
    return raw_row.values.get(field, B2BCellValue(None)).value


def _has_error(issues: list[B2BValidationIssue]) -> bool:
    return any(issue.severity.upper() == "ERROR" for issue in issues)


def _number_or_none(value: Any) -> int | float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    return None


def _issue(
    severity: str,
    issue_type: str,
    raw_row: RawB2BDispatchRow,
    run_id: str,
    field_name: str,
    raw_value: Any,
    detail: str,
    action: str,
) -> B2BValidationIssue:
    return B2BValidationIssue(
        run_id=run_id,
        severity=severity,
        issue_type=issue_type,
        source_file=raw_row.source_file,
        source_sheet=raw_row.source_sheet,
        source_row=raw_row.source_row,
        source_channel=raw_row.source_channel,
        po=_text(raw_row, "PO"),
        asin=_text(raw_row, "ASIN"),
        invoice_no=_text(raw_row, "Invoice No"),
        field_name=field_name,
        raw_value=raw_value,
        issue_detail=detail,
        action_taken=action,
    )


def _issue_suffix(field_name: str) -> str:
    return field_name.upper().replace(" ", "_").replace("-", "_")
