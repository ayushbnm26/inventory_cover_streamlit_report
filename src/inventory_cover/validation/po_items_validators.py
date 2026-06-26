"""PO Items row validation and duplicate handling."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
import hashlib
import json
from typing import Any

from inventory_cover.schemas import (
    DUPLICATE_FINGERPRINT_FIELDS,
    DuplicateRecord,
    NormalizedPoItemRow,
    ValidationIssue,
)


def attach_duplicate_findings(
    rows: list[NormalizedPoItemRow],
    dedupe_exact_rows: bool = False,
) -> tuple[list[NormalizedPoItemRow], list[ValidationIssue], list[DuplicateRecord]]:
    """Flag exact business-row duplicates and optionally drop later copies."""

    groups: dict[str, list[NormalizedPoItemRow]] = defaultdict(list)
    for row in rows:
        row.fingerprint = fingerprint_row(row)
        groups[row.fingerprint].append(row)

    duplicate_issues: list[ValidationIssue] = []
    duplicate_records: list[DuplicateRecord] = []
    kept_rows: list[NormalizedPoItemRow] = []
    duplicate_fingerprints = [fingerprint for fingerprint, group in groups.items() if len(group) > 1]
    group_ids = {
        fingerprint: f"DUP-{index:04d}"
        for index, fingerprint in enumerate(duplicate_fingerprints, start=1)
    }

    for row in rows:
        group = groups[row.fingerprint]
        if len(group) == 1:
            kept_rows.append(row)
            continue

        group_id = group_ids[row.fingerprint]
        position = group.index(row)
        should_drop = dedupe_exact_rows and position > 0
        action = "DROPPED_BY_DEDUPE" if should_drop else "KEPT"
        detail = (
            "Exact duplicate business row found; later copies dropped by configuration."
            if dedupe_exact_rows
            else "Exact duplicate business row found; row kept because dedupe is disabled."
        )
        issue = ValidationIssue(
            severity="WARNING",
            issue_type="DUPLICATE_EXACT_ROW",
            source_file=str(row.data.get("Source File") or ""),
            source_sheet=str(row.data.get("Source Sheet") or ""),
            source_row=int(row.data.get("Source Row") or 0),
            po=str(row.data.get("PO") or ""),
            asin=str(row.data.get("ASIN") or ""),
            field_name="FULL_ROW",
            raw_value=row.fingerprint,
            issue_detail=detail,
            action_taken=action,
        )
        row.issues.append(issue)
        row.refresh_validation_status()
        duplicate_issues.append(issue)
        duplicate_records.append(
            DuplicateRecord(
                duplicate_type="EXACT_BUSINESS_ROW",
                duplicate_group_id=group_id,
                source_file=str(row.data.get("Source File") or ""),
                source_row=int(row.data.get("Source Row") or 0),
                po=str(row.data.get("PO") or ""),
                asin=str(row.data.get("ASIN") or ""),
                full_row_fingerprint=row.fingerprint,
                action_taken=action,
            )
        )
        if should_drop:
            row.dropped = True
        else:
            kept_rows.append(row)

    return kept_rows, duplicate_issues, duplicate_records


def fingerprint_row(row: NormalizedPoItemRow) -> str:
    values = {field: _normalize_for_fingerprint(row.data.get(field)) for field in DUPLICATE_FINGERPRINT_FIELDS}
    payload = json.dumps(values, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _normalize_for_fingerprint(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if value is None:
        return ""
    return str(value).strip()
