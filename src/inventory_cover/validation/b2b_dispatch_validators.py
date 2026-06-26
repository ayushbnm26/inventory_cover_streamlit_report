"""B2B Dispatch Tracker validation helpers."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
import json
from typing import Any

from inventory_cover.b2b_dispatch_schemas import (
    B2B_DUPLICATE_KEY_FIELDS,
    B2BDuplicateRecord,
    B2BValidationIssue,
    NormalizedB2BDispatchRow,
)


def attach_b2b_duplicate_findings(
    rows: list[NormalizedB2BDispatchRow],
    dedupe_exact_rows: bool = False,
) -> tuple[list[NormalizedB2BDispatchRow], list[B2BValidationIssue], list[B2BDuplicateRecord]]:
    """Flag duplicate dispatch rows and optionally drop later copies."""

    groups: dict[str, list[NormalizedB2BDispatchRow]] = defaultdict(list)
    for row in rows:
        row.duplicate_key = duplicate_key(row)
        groups[row.duplicate_key].append(row)

    duplicate_keys = [key for key, group in groups.items() if len(group) > 1]
    group_ids = {key: f"B2B-DUP-{index:04d}" for index, key in enumerate(duplicate_keys, start=1)}

    kept_rows: list[NormalizedB2BDispatchRow] = []
    duplicate_issues: list[B2BValidationIssue] = []
    duplicate_records: list[B2BDuplicateRecord] = []

    for row in rows:
        group = groups[row.duplicate_key]
        if len(group) == 1:
            kept_rows.append(row)
            continue

        position = group.index(row)
        should_drop = dedupe_exact_rows and position > 0
        action = "DROPPED_BY_DEDUPE" if should_drop else "KEPT"
        issue = B2BValidationIssue(
            run_id=str(row.data.get("Run ID") or ""),
            severity="WARNING",
            issue_type="DUPLICATE_DISPATCH_ROW",
            source_file=str(row.data.get("Source File") or ""),
            source_sheet=str(row.data.get("Source Sheet") or ""),
            source_row=int(row.data.get("Source Row") or 0),
            source_channel=str(row.data.get("Source Channel") or ""),
            po=str(row.data.get("PO") or ""),
            asin=str(row.data.get("ASIN") or ""),
            invoice_no=str(row.data.get("Invoice No") or ""),
            field_name="DUPLICATE_KEY",
            raw_value=row.duplicate_key,
            issue_detail=(
                "Duplicate dispatch row found; later copies dropped by configuration."
                if dedupe_exact_rows
                else "Duplicate dispatch row found; row kept because dedupe is disabled."
            ),
            action_taken=action,
        )
        row.issues.append(issue)
        row.refresh_validation_status()
        duplicate_issues.append(issue)
        duplicate_records.append(
            B2BDuplicateRecord(
                run_id=str(row.data.get("Run ID") or ""),
                duplicate_group_id=group_ids[row.duplicate_key],
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


def duplicate_key(row: NormalizedB2BDispatchRow) -> str:
    values = {field: _normalize_for_key(row.data.get(field)) for field in B2B_DUPLICATE_KEY_FIELDS}
    return json.dumps(values, sort_keys=True, separators=(",", ":"))


def _normalize_for_key(value: Any) -> str:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if value is None:
        return ""
    return str(value).strip()
