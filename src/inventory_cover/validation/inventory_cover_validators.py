"""Freshness checks and Source_Summary construction for the cover engine."""

from __future__ import annotations

from datetime import date
from typing import Any

from inventory_cover.io.inventory_cover_input_loader import LoadedSheet
from inventory_cover.normalization.inventory_cover_aggregator import SourceStats
from inventory_cover.inventory_cover_schemas import (
    SOURCE_ASIN_MASTER,
    SOURCE_B2B,
    SOURCE_INVENTORY,
    SOURCE_PO,
    SOURCE_SALES,
    InventoryCoverValidationIssue,
    SourceSummaryRecord,
)


def build_source_summaries(
    run_id: str,
    run_timestamp: str,
    config: Any,
    sheets: dict[str, LoadedSheet],
    stats: dict[str, SourceStats],
    copied_paths: dict[str, str],
    run_date: date,
) -> tuple[list[SourceSummaryRecord], list[InventoryCoverValidationIssue]]:
    issues: list[InventoryCoverValidationIssue] = []
    summaries: list[SourceSummaryRecord] = []

    inventory_date = stats[SOURCE_INVENTORY].report_updated_date or stats[SOURCE_INVENTORY].period_end
    sales_end = stats[SOURCE_SALES].period_end

    for source_type in (SOURCE_PO, SOURCE_B2B, SOURCE_SALES, SOURCE_INVENTORY, SOURCE_ASIN_MASTER):
        sheet = sheets[source_type]
        stat = stats[source_type]
        warnings = list(sheet.warnings) + list(stat.warnings)
        freshness, fresh_warnings = _freshness_for(
            source_type, config, stat, sales_end, inventory_date, run_date
        )
        warnings.extend(fresh_warnings)

        summaries.append(
            SourceSummaryRecord(
                run_id=run_id,
                run_timestamp=run_timestamp,
                source_type=source_type,
                source_latest_path=str(sheet.path),
                copied_run_path=copied_paths.get(source_type, ""),
                workbook_exists=sheet.exists,
                sheet_used=sheet.sheet_name,
                rows_read=stat.rows_read,
                rows_accepted=stat.rows_read,
                rows_used=stat.rows_used,
                report_period_start=stat.period_start,
                report_period_end=stat.period_end,
                report_updated_date=stat.report_updated_date,
                freshness_status=freshness,
                warnings="; ".join(w for w in warnings if w),
            )
        )

        for warning in fresh_warnings:
            issues.append(
                InventoryCoverValidationIssue(
                    run_id=run_id,
                    severity="WARNING",
                    issue_type="SOURCE_FRESHNESS",
                    source_type=source_type,
                    source_file=sheet.path.name,
                    source_sheet=sheet.sheet_name,
                    issue_detail=warning,
                    action_taken="Run continued (use --strict-freshness to fail instead).",
                )
            )
        if not sheet.exists and source_type != SOURCE_ASIN_MASTER:
            issues.append(
                InventoryCoverValidationIssue(
                    run_id=run_id,
                    severity="WARNING",
                    issue_type="SOURCE_MISSING",
                    source_type=source_type,
                    source_file=sheet.path.name,
                    issue_detail=f"{source_type} backend workbook not found.",
                    action_taken="Source treated as empty; affected metrics default to zero.",
                )
            )

    return summaries, issues


def _freshness_for(
    source_type: str,
    config: Any,
    stat: SourceStats,
    sales_end: date | None,
    inventory_date: date | None,
    run_date: date,
) -> tuple[str, list[str]]:
    warnings: list[str] = []

    if source_type == SOURCE_SALES:
        if stat.period_end is None:
            warnings.append("No sales period end detected.")
            return "UNKNOWN", warnings
        if inventory_date is not None:
            gap = (inventory_date - stat.period_end).days
            if gap > config.sales_staleness_days:
                warnings.append(
                    f"Sales period end {stat.period_end} is {gap} days older than inventory date "
                    f"{inventory_date} (threshold {config.sales_staleness_days})."
                )
                return "STALE", warnings
        return "FRESH", warnings

    if source_type == SOURCE_INVENTORY:
        if inventory_date is None:
            warnings.append("No inventory report date detected.")
            return "UNKNOWN", warnings
        gap = (run_date - inventory_date).days
        if gap > config.inventory_staleness_days:
            warnings.append(
                f"Inventory date {inventory_date} is {gap} days older than run date {run_date} "
                f"(threshold {config.inventory_staleness_days})."
            )
            return "STALE", warnings
        return "FRESH", warnings

    if source_type == SOURCE_B2B:
        if stat.rows_used == 0:
            warnings.append("No dispatch rows in the latest lookback window.")
            return "EMPTY", warnings
        return "FRESH", warnings

    if source_type == SOURCE_PO:
        if stat.period_end is None:
            warnings.append("No reliable PO date column available; PO freshness not date-filtered.")
            return "NO DATE", warnings
        return "FRESH", warnings

    if source_type == SOURCE_ASIN_MASTER:
        if stat.rows_used == 0:
            return "NOT AVAILABLE", warnings
        return "AVAILABLE", warnings

    return "UNKNOWN", warnings


def strict_freshness_errors(summaries: list[SourceSummaryRecord]) -> list[str]:
    """Return blocking messages for stale required sources under strict mode."""

    errors: list[str] = []
    for summary in summaries:
        if summary.source_type in (SOURCE_SALES, SOURCE_INVENTORY) and summary.freshness_status == "STALE":
            errors.append(f"{summary.source_type} source is stale: {summary.warnings}")
    return errors
