"""Orchestration for the Final Inventory Cover Calculation Engine.

This engine is loosely coupled to the source pipelines: it consumes their latest
backend workbook artifacts through a stable sheet/column interface contract and
never mutates the source files.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
import time
from typing import Any

from inventory_cover.config import InventoryCoverPipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError
from inventory_cover.io.inventory_cover_input_loader import LoadedSheet, load_backend_sheet
from inventory_cover.logging_utils import setup_run_logger, write_json_file
from inventory_cover.normalization.inventory_cover_aggregator import aggregate_products
from inventory_cover.calculations.inventory_cover_builder import compute_product
from inventory_cover.reports.inventory_cover_excel_writer import (
    build_formula_guide_rows,
    write_backend_workbook,
    write_team_workbook,
)
from inventory_cover.inventory_cover_schemas import (
    B2B_SHEET,
    INVENTORY_SHEET,
    PIPELINE_NAME,
    PO_SHEET,
    SALES_SHEET,
    SOURCE_ASIN_MASTER,
    SOURCE_B2B,
    SOURCE_INVENTORY,
    SOURCE_PO,
    SOURCE_SALES,
    InventoryCoverPipelineRunResult,
    InventoryCoverValidationIssue,
)
from inventory_cover.validation.inventory_cover_validators import (
    build_source_summaries,
    strict_freshness_errors,
)


class InventoryCoverPipeline:
    """Coordinates loading, consolidation, calculation, and report writing."""

    def __init__(self, config: InventoryCoverPipelineConfig):
        self.config = config.resolved()

    def run(self) -> InventoryCoverPipelineRunResult:
        run_id = _create_run_id()
        run_dir = self.config.run_root / run_id
        paths = _create_run_dirs(run_dir)
        run_dir = paths["run"]
        log_file = paths["logs"] / "inventory_cover_pipeline.log"
        logger = setup_run_logger(log_file, self.config.log_level, logger_name="inventory_cover.inventory_cover")
        start_time = datetime.now()
        run_timestamp = start_time.isoformat(timespec="seconds")

        team_output_file = paths["outputs"] / f"Inventory_Cover_Report_{run_id}.xlsx"
        backend_output_file = paths["outputs"] / f"Inventory_Cover_Backend_Audit_{run_id}.xlsx"
        team_latest_file = self.config.processed_dir / "latest" / "Inventory_Cover_Report_latest.xlsx"
        backend_latest_file = self.config.processed_dir / "latest" / "Inventory_Cover_Backend_Audit_latest.xlsx"
        metadata_file = paths["metadata"] / "run_metadata.json"
        validation_file = paths["validation"] / "inventory_cover_validation_issues.json"

        logger.info("%s started. run_id=%s", PIPELINE_NAME, run_id)
        logger.info(
            "Config: po=%s b2b=%s sales=%s inventory=%s asin_master=%s window=%s default_target=%s "
            "blank_policy=%s strict_freshness=%s",
            self.config.po_backend_path,
            self.config.b2b_backend_path,
            self.config.sales_backend_path,
            self.config.inventory_backend_path,
            self.config.asin_master_path,
            self.config.sales_window_days,
            self.config.default_target_doh,
            self.config.blank_numeric_policy,
            self.config.strict_freshness,
        )

        source_paths = {
            SOURCE_SALES: (self.config.sales_backend_path, SALES_SHEET),
            SOURCE_INVENTORY: (self.config.inventory_backend_path, INVENTORY_SHEET),
            SOURCE_B2B: (self.config.b2b_backend_path, B2B_SHEET),
            SOURCE_PO: (self.config.po_backend_path, PO_SHEET),
            SOURCE_ASIN_MASTER: (self.config.asin_master_path, "ASIN_Master"),
        }

        sheets: dict[str, LoadedSheet] = {}
        copied_paths: dict[str, str] = {}
        for source_type, (path, sheet_name) in source_paths.items():
            sheets[source_type] = load_backend_sheet(source_type, path, sheet_name)
            copied = _copy_input(path, paths["inputs"])
            copied_paths[source_type] = str(copied) if copied else ""
            for warning in sheets[source_type].warnings:
                logger.warning("%s: %s", source_type, warning)

        try:
            aggregation = aggregate_products(
                run_id=run_id,
                sales=sheets[SOURCE_SALES],
                inventory=sheets[SOURCE_INVENTORY],
                b2b=sheets[SOURCE_B2B],
                po=sheets[SOURCE_PO],
                asin_master=sheets[SOURCE_ASIN_MASTER],
                config=self.config,
            )
            products = aggregation.products
            issues: list[InventoryCoverValidationIssue] = list(aggregation.issues)

            if not products:
                raise CatastrophicPipelineError(
                    "No products could be built from any source. Check that Sales/Inventory backend "
                    "artifacts exist and contain rows."
                )

            for product in products:
                compute_product(product, self.config)

            summaries, freshness_issues = build_source_summaries(
                run_id=run_id,
                run_timestamp=run_timestamp,
                config=self.config,
                sheets=sheets,
                stats=aggregation.stats,
                copied_paths=copied_paths,
                run_date=start_time.date(),
            )
            issues.extend(freshness_issues)

            if self.config.strict_freshness:
                blocking = strict_freshness_errors(summaries)
                if blocking:
                    raise CatastrophicPipelineError("Strict freshness failed: " + "; ".join(blocking))

            guide_rows = build_formula_guide_rows(
                run_id=run_id,
                run_timestamp=run_timestamp,
                config=self.config,
                summaries=summaries,
                product_count=len(products),
            )

            metadata_kv = _metadata_pairs(run_id, run_timestamp, self.config, sheets, products, summaries, issues)

            write_team_workbook(team_output_file, run_id, products, summaries, guide_rows, self.config)
            _copy_latest_or_warn(team_output_file, team_latest_file, run_id, "Team workbook", issues, logger)

            write_backend_workbook(
                backend_output_file, run_id, products, aggregation.traces, summaries, issues,
                metadata_kv, guide_rows, self.config,
            )
            _copy_latest_or_warn(
                backend_output_file, backend_latest_file, run_id, "Backend workbook", issues, logger
            )

            write_json_file(validation_file, {"issues": [issue.as_json() for issue in issues]})

            warning_count = sum(1 for issue in issues if issue.severity.upper() == "WARNING")
            end_time = datetime.now()
            metadata = {
                "run_id": run_id,
                "pipeline_name": PIPELINE_NAME,
                "start_time": start_time.isoformat(),
                "end_time": end_time.isoformat(),
                "duration_seconds": round((end_time - start_time).total_seconds(), 3),
                "status": "SUCCESS",
                "product_count": len(products),
                "validation_issue_count": len(issues),
                "warning_count": warning_count,
                "team_output_file": str(team_output_file),
                "backend_output_file": str(backend_output_file),
                "team_latest_file": str(team_latest_file),
                "backend_latest_file": str(backend_latest_file),
                "validation_file": str(validation_file),
                "log_file": str(log_file),
                "sources": {key: pair[0].__str__() for key, pair in source_paths.items()},
                "source_summaries": [summary.as_json() for summary in summaries],
                "bucket_counts": _bucket_counts(products),
            }
            write_json_file(metadata_file, metadata)
            logger.info(
                "%s completed. products=%s issues=%s warnings=%s",
                PIPELINE_NAME, len(products), len(issues), warning_count,
            )

            return InventoryCoverPipelineRunResult(
                run_id=run_id,
                run_dir=run_dir,
                team_output_file=team_output_file,
                team_latest_file=team_latest_file,
                backend_output_file=backend_output_file,
                backend_latest_file=backend_latest_file,
                metadata_file=metadata_file,
                validation_file=validation_file,
                log_file=log_file,
                product_count=len(products),
                validation_issue_count=len(issues),
                warning_count=warning_count,
            )
        except CatastrophicPipelineError as exc:
            logger.error("%s failed: %s", PIPELINE_NAME, exc)
            end_time = datetime.now()
            write_json_file(metadata_file, {
                "run_id": run_id,
                "pipeline_name": PIPELINE_NAME,
                "status": "FAILED",
                "start_time": start_time.isoformat(),
                "end_time": end_time.isoformat(),
                "error": str(exc),
            })
            raise


def _metadata_pairs(
    run_id: str,
    run_timestamp: str,
    config: InventoryCoverPipelineConfig,
    sheets: dict[str, LoadedSheet],
    products: list,
    summaries: list,
    issues: list,
) -> list[tuple[str, Any]]:
    pairs: list[tuple[str, Any]] = [
        ("Run ID", run_id),
        ("Run Timestamp", run_timestamp),
        ("Pipeline", PIPELINE_NAME),
        ("Sales window days", config.sales_window_days),
        ("Default target DOH", config.default_target_doh),
        ("Blank numeric policy", config.blank_numeric_policy),
        ("Strict freshness", config.strict_freshness),
        ("Products", len(products)),
        ("Validation issues", len(issues)),
    ]
    for summary in summaries:
        pairs.append((f"{summary.source_type} workbook exists", "Yes" if summary.workbook_exists else "No"))
        pairs.append((f"{summary.source_type} rows used", summary.rows_used))
    for bucket, count in _bucket_counts(products).items():
        pairs.append((f"Bucket: {bucket}", count))
    return pairs


def _bucket_counts(products: list) -> dict[str, int]:
    counts: dict[str, int] = {}
    for product in products:
        counts[product.cover_bucket] = counts.get(product.cover_bucket, 0) + 1
    return counts


def _create_run_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _create_run_dirs(run_dir: Path) -> dict[str, Path]:
    if run_dir.exists():
        time.sleep(1.05)
        run_dir = run_dir.parent / _create_run_id()
    paths = {
        "run": run_dir,
        "inputs": run_dir / "inputs" / "inventory_cover",
        "outputs": run_dir / "outputs" / "inventory_cover",
        "logs": run_dir / "logs",
        "validation": run_dir / "validation",
        "metadata": run_dir / "metadata",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    return paths


def _copy_input(source: Path, destination: Path) -> Path | None:
    if not source.exists():
        return None
    destination.mkdir(parents=True, exist_ok=True)
    target = destination / source.name
    shutil.copy2(source, target)
    return target


def _copy_latest(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temp_latest = target.with_suffix(target.suffix + ".tmp")
    shutil.copy2(source, temp_latest)
    temp_latest.replace(target)


def _copy_latest_or_warn(
    source: Path,
    target: Path,
    run_id: str,
    artifact_name: str,
    issues: list[InventoryCoverValidationIssue],
    logger: Any,
) -> bool:
    try:
        _copy_latest(source, target)
        logger.info("%s latest updated: %s", artifact_name, target)
        return True
    except PermissionError as exc:
        temp_latest = target.with_suffix(target.suffix + ".tmp")
        if temp_latest.exists():
            temp_latest.unlink(missing_ok=True)
        detail = f"{artifact_name} latest could not be replaced, likely because the file is open: {target}"
        issues.append(
            InventoryCoverValidationIssue(
                run_id=run_id,
                severity="WARNING",
                issue_type="LATEST_COPY_FAILED",
                field_name=artifact_name,
                raw_value=str(target),
                issue_detail=f"{detail}. {exc}",
                action_taken="Timestamped run output kept; close the latest workbook and rerun.",
            )
        )
        logger.warning("%s", detail)
        return False
