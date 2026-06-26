"""Validation helpers for Vendor Central sales and inventory rows."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
import json
from typing import Any

from inventory_cover.sales_inventory_schemas import (
    NormalizedSalesInventoryRow,
    SalesInventoryDuplicateRecord,
    SalesInventoryValidationIssue,
)


DUPLICATE_KEY_FIELDS: tuple[str, ...] = (
    "Report Type",
    "Viewing Range Start",
    "Viewing Range End",
    "ASIN",
    "Child Vendor Code",
    "Model Number",
)

TRACEABILITY_FIELDS: frozenset[str] = frozenset(
    {
        "Run ID",
        "Source File",
        "Source Sheet",
        "Source Row",
        "Row Validation Status",
        "Row Validation Notes",
    }
)


def attach_sales_inventory_duplicate_findings(
    rows: list[NormalizedSalesInventoryRow],
    dedupe_exact_rows: bool = False,
) -> tuple[list[NormalizedSalesInventoryRow], list[SalesInventoryValidationIssue], list[SalesInventoryDuplicateRecord]]:
    """Flag duplicate rows and optionally drop only exact normalized duplicates."""

    business_groups: dict[str, list[NormalizedSalesInventoryRow]] = defaultdict(list)
    exact_groups: dict[str, list[NormalizedSalesInventoryRow]] = defaultdict(list)
    for row in rows:
        row.duplicate_key = duplicate_key(row)
        row.exact_fingerprint = exact_fingerprint(row)
        business_groups[row.duplicate_key].append(row)
        exact_groups[row.exact_fingerprint].append(row)

    duplicate_keys = [key for key, group in business_groups.items() if len(group) > 1]
    group_ids = {key: f"SI-DUP-{index:04d}" for index, key in enumerate(duplicate_keys, start=1)}

    exact_drop_ids: set[int] = set()
    if dedupe_exact_rows:
        for group in exact_groups.values():
            if len(group) <= 1:
                continue
            for row in group[1:]:
                exact_drop_ids.add(id(row))

    kept_rows: list[NormalizedSalesInventoryRow] = []
    duplicate_issues: list[SalesInventoryValidationIssue] = []
    duplicate_records: list[SalesInventoryDuplicateRecord] = []

    for row in rows:
        group = business_groups[row.duplicate_key]
        should_drop = id(row) in exact_drop_ids
        if len(group) == 1 and not should_drop:
            kept_rows.append(row)
            continue

        action = "DROPPED_BY_DEDUPE" if should_drop else "KEPT"
        issue = SalesInventoryValidationIssue(
            run_id=str(row.data.get("Run ID") or ""),
            report_type=str(row.data.get("Report Type") or row.report_type),
            severity="WARNING",
            issue_type="DUPLICATE_ROW",
            source_file=str(row.data.get("Source File") or ""),
            source_sheet=str(row.data.get("Source Sheet") or ""),
            source_row=int(row.data.get("Source Row") or 0),
            asin=str(row.data.get("ASIN") or ""),
            child_vendor_code=str(row.data.get("Child Vendor Code") or ""),
            model_number=str(row.data.get("Model Number") or ""),
            field_name="DUPLICATE_KEY",
            raw_value=row.duplicate_key,
            issue_detail=(
                "Duplicate key found; exact repeated normalized row dropped by configuration."
                if should_drop
                else "Duplicate key found; row kept because it is not an exact duplicate or dedupe is disabled."
            ),
            action_taken=action,
        )
        row.issues.append(issue)
        row.refresh_validation_status()
        duplicate_issues.append(issue)
        duplicate_records.append(
            SalesInventoryDuplicateRecord(
                run_id=str(row.data.get("Run ID") or ""),
                report_type=str(row.data.get("Report Type") or row.report_type),
                duplicate_group_id=group_ids.get(row.duplicate_key, "SI-EXACT-DUP"),
                duplicate_key=row.duplicate_key,
                source_file=str(row.data.get("Source File") or ""),
                source_sheet=str(row.data.get("Source Sheet") or ""),
                source_row=int(row.data.get("Source Row") or 0),
                action_taken=action,
            )
        )
        if should_drop:
            row.dropped = True
        else:
            kept_rows.append(row)

    return kept_rows, duplicate_issues, duplicate_records


def duplicate_key(row: NormalizedSalesInventoryRow) -> str:
    values = {field: _normalize_for_key(row.data.get(field)) for field in DUPLICATE_KEY_FIELDS}
    return json.dumps(values, sort_keys=True, separators=(",", ":"))


def exact_fingerprint(row: NormalizedSalesInventoryRow) -> str:
    values = {
        key: _normalize_for_key(value)
        for key, value in row.data.items()
        if key not in TRACEABILITY_FIELDS
    }
    return json.dumps(values, sort_keys=True, separators=(",", ":"))


def _normalize_for_key(value: Any) -> str:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if value is None:
        return ""
    return str(value).strip()
