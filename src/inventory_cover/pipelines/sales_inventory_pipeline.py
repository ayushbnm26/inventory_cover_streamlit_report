"""Orchestration for the Vendor Central Sales & Inventory backend pipeline."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
import time
from typing import Any

from inventory_cover.config import SalesInventoryPipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError, FileValidationError
from inventory_cover.io.sales_inventory_excel_io import read_sales_inventory_workbook
from inventory_cover.io.sales_inventory_file_discovery import discover_sales_inventory_files
from inventory_cover.io.sales_inventory_mapping_io import SalesInventoryMappingLookup, read_sales_inventory_mapping_workbook
from inventory_cover.logging_utils import setup_run_logger, write_json_file
from inventory_cover.normalization.sales_inventory_normalizer import normalize_sales_inventory_row
from inventory_cover.reports.sales_inventory_excel_writer import (
    write_inventory_backend_report,
    write_sales_backend_report,
    write_sales_inventory_run_summary,
)
from inventory_cover.sales_inventory_schemas import (
    INVENTORY_REPORT_TYPE,
    PIPELINE_NAME,
    SALES_REPORT_TYPE,
    NormalizedSalesInventoryRow,
    ReportType,
    SalesInventoryDuplicateRecord,
    SalesInventoryFileAuditRecord,
    SalesInventoryMappingAuditRecord,
    SalesInventoryPipelineRunResult,
    SalesInventoryValidationIssue,
)
from inventory_cover.validation.sales_inventory_validators import attach_sales_inventory_duplicate_findings


class SalesInventoryPipeline:
    """Coordinates Vendor Central sales and inventory source ingestion."""

    def __init__(self, config: SalesInventoryPipelineConfig):
        self.config = config.resolved()

    def run(self) -> SalesInventoryPipelineRunResult:
        run_id = _create_run_id()
        run_dir = self.config.run_root / run_id
        paths = _create_run_dirs(run_dir)
        run_dir = paths["run"]
        log_file = paths["logs"] / "sales_inventory_pipeline.log"
        logger = setup_run_logger(
            log_file,
            self.config.log_level,
            logger_name="inventory_cover.sales_inventory",
        )
        start_time = datetime.now()

        output_dir = paths["outputs_sales_inventory"]
        sales_output_file = output_dir / f"Sales_Backend_Audit_{run_id}.xlsx"
        inventory_output_file = output_dir / f"Inventory_Backend_Audit_{run_id}.xlsx"
        summary_output_file = output_dir / f"Sales_Inventory_Run_Summary_{run_id}.xlsx"
        latest_sales_file = self.config.processed_dir / "latest" / "Sales_Backend_Audit_latest.xlsx"
        latest_inventory_file = self.config.processed_dir / "latest" / "Inventory_Backend_Audit_latest.xlsx"
        latest_summary_file = self.config.processed_dir / "latest" / "Sales_Inventory_Run_Summary_latest.xlsx"
        metadata_file = paths["metadata"] / "run_metadata.json"
        validation_file = paths["validation"] / "sales_inventory_validation_issues.json"
        duplicates_file = paths["validation"] / "sales_inventory_duplicates.json"

        metadata: dict[str, Any] = {
            "run_id": run_id,
            "pipeline_name": PIPELINE_NAME,
            "start_time": start_time.isoformat(),
            "status": "STARTED",
            "sales_input_directory": str(self.config.sales_input_dir),
            "inventory_input_directory": str(self.config.inventory_input_dir),
            "mapping_input_directory": str(self.config.mapping_input_dir),
            "run_directory": str(run_dir),
        }

        discovered_sales_files: list[Path] = []
        discovered_inventory_files: list[Path] = []
        discovered_mapping_files: list[Path] = []
        file_audit: list[SalesInventoryFileAuditRecord] = []
        mapping_audit: list[SalesInventoryMappingAuditRecord] = []
        validation_issues: list[SalesInventoryValidationIssue] = []
        duplicate_records: list[SalesInventoryDuplicateRecord] = []
        rows: list[NormalizedSalesInventoryRow] = []
        rows_to_write: list[NormalizedSalesInventoryRow] = []
        mapping_lookup: SalesInventoryMappingLookup | None = None
        sales_output_generated = False
        inventory_output_generated = False

        logger.info("Sales & Inventory pipeline started. run_id=%s", run_id)
        logger.info(
            "Config: sales_input_dir=%s inventory_input_dir=%s run_root=%s processed_dir=%s "
            "mapping_input_dir=%s require_sales=%s require_inventory=%s allow_multiple_sales_files=%s "
            "allow_multiple_inventory_files=%s dedupe_exact_rows=%s",
            self.config.sales_input_dir,
            self.config.inventory_input_dir,
            self.config.run_root,
            self.config.processed_dir,
            self.config.mapping_input_dir,
            self.config.require_sales,
            self.config.require_inventory,
            self.config.allow_multiple_sales_files,
            self.config.allow_multiple_inventory_files,
            self.config.dedupe_exact_rows,
        )

        try:
            discovery = discover_sales_inventory_files(self.config)
            discovered_sales_files = discovery.sales_files
            discovered_inventory_files = discovery.inventory_files
            discovered_mapping_files = discovery.mapping_files
            for warning in discovery.warnings:
                logger.warning(warning)
                validation_issues.append(
                    SalesInventoryValidationIssue(
                        run_id=run_id,
                        report_type="SOURCE",
                        severity="WARNING",
                        issue_type="SOURCE_MISSING_OR_MULTIPLE_ALLOWED",
                        issue_detail=warning,
                        action_taken="Run continued.",
                    )
                )

            if not discovered_sales_files and not discovered_inventory_files:
                raise CatastrophicPipelineError("Both sales and inventory sources are missing; nothing to process.")

            logger.info(
                "Files discovered: sales=%s inventory=%s",
                len(discovered_sales_files),
                len(discovered_inventory_files),
            )

            copied_sales = _copy_inputs(discovered_sales_files, paths["inputs_sales"])
            copied_inventory = _copy_inputs(discovered_inventory_files, paths["inputs_inventory"])
            copied_mapping = _copy_inputs(discovered_mapping_files, paths["inputs_mapping"])

            if discovered_mapping_files:
                mapping_path = discovered_mapping_files[0]
                try:
                    mapping_result = read_sales_inventory_mapping_workbook(
                        mapping_path,
                        run_id=run_id,
                        copied_run_path=copied_mapping.get(mapping_path),
                    )
                    mapping_lookup = mapping_result.lookup
                    mapping_audit.append(mapping_result.audit)
                    logger.info(
                        "Mapping workbook loaded: file=%s rows_loaded=%s unique_asin_keys=%s unique_sku_keys=%s",
                        mapping_path.name,
                        mapping_result.audit.rows_loaded,
                        mapping_result.audit.unique_asin_keys,
                        mapping_result.audit.unique_sku_keys,
                    )
                except FileValidationError as exc:
                    mapping_audit.append(
                        SalesInventoryMappingAuditRecord(
                            run_id=run_id,
                            mapping_file=mapping_path.name,
                            full_path=str(mapping_path),
                            copied_run_path=str(copied_mapping.get(mapping_path, "")),
                            status="FAILED",
                            notes=str(exc),
                        )
                    )
                    validation_issues.append(
                        SalesInventoryValidationIssue(
                            run_id=run_id,
                            report_type="SOURCE",
                            severity="WARNING",
                            issue_type="REFERENCE_MAPPING_FAILED",
                            source_file=mapping_path.name,
                            issue_detail=str(exc),
                            action_taken="Mapping enrichment disabled; sales/inventory processing continued.",
                        )
                    )
                    logger.warning("Mapping workbook could not be used; enrichment disabled: %s", exc)

            for source_path in discovered_sales_files:
                read_rows = _process_source_file(
                    source_path=source_path,
                    report_type=SALES_REPORT_TYPE,
                    run_id=run_id,
                    copied_path=copied_sales.get(source_path),
                    config=self.config,
                    file_audit=file_audit,
                    validation_issues=validation_issues,
                    mapping_lookup=mapping_lookup,
                    logger=logger,
                )
                rows.extend(read_rows)

            for source_path in discovered_inventory_files:
                read_rows = _process_source_file(
                    source_path=source_path,
                    report_type=INVENTORY_REPORT_TYPE,
                    run_id=run_id,
                    copied_path=copied_inventory.get(source_path),
                    config=self.config,
                    file_audit=file_audit,
                    validation_issues=validation_issues,
                    mapping_lookup=mapping_lookup,
                    logger=logger,
                )
                rows.extend(read_rows)

            if self.config.require_sales and not any(row.report_type == SALES_REPORT_TYPE for row in rows):
                raise CatastrophicPipelineError("Sales was required but zero usable sales rows were produced.")
            if self.config.require_inventory and not any(row.report_type == INVENTORY_REPORT_TYPE for row in rows):
                raise CatastrophicPipelineError("Inventory was required but zero usable inventory rows were produced.")
            if not rows:
                raise CatastrophicPipelineError("Zero usable rows after processing discovered sources.")

            rows_to_write, duplicate_issues, duplicate_records = attach_sales_inventory_duplicate_findings(
                rows,
                dedupe_exact_rows=self.config.dedupe_exact_rows,
            )
            validation_issues.extend(duplicate_issues)
            _annotate_dedupe_drops(file_audit, rows, rows_to_write)

            if not rows_to_write:
                raise CatastrophicPipelineError("All rows were removed by dedupe; output not created.")

            sales_rows = [row for row in rows_to_write if row.report_type == SALES_REPORT_TYPE]
            inventory_rows = [row for row in rows_to_write if row.report_type == INVENTORY_REPORT_TYPE]

            summary = _build_run_summary(
                run_id=run_id,
                start_time=start_time,
                config=self.config,
                output_folder=output_dir,
                discovered_sales_files=discovered_sales_files,
                discovered_inventory_files=discovered_inventory_files,
                file_audit=file_audit,
                validation_issues=validation_issues,
                duplicate_count=len(duplicate_records),
                sales_output_file=sales_output_file if sales_rows else None,
                inventory_output_file=inventory_output_file if inventory_rows else None,
                latest_sales_file=latest_sales_file if sales_rows else None,
                latest_inventory_file=latest_inventory_file if inventory_rows else None,
                metadata_file=metadata_file,
                log_file=log_file,
            )
            source_status_rows = _build_source_status_rows(
                config=self.config,
                discovered_sales_files=discovered_sales_files,
                discovered_inventory_files=discovered_inventory_files,
                file_audit=file_audit,
                sales_output_file=sales_output_file if sales_rows else None,
                inventory_output_file=inventory_output_file if inventory_rows else None,
                latest_sales_file=latest_sales_file if sales_rows else None,
                latest_inventory_file=latest_inventory_file if inventory_rows else None,
            )

            sales_audit = [record for record in file_audit if record.report_type == SALES_REPORT_TYPE]
            inventory_audit = [record for record in file_audit if record.report_type == INVENTORY_REPORT_TYPE]
            sales_issues = _issues_for_report(validation_issues, SALES_REPORT_TYPE)
            inventory_issues = _issues_for_report(validation_issues, INVENTORY_REPORT_TYPE)
            sales_duplicates = [record for record in duplicate_records if record.report_type == SALES_REPORT_TYPE]
            inventory_duplicates = [record for record in duplicate_records if record.report_type == INVENTORY_REPORT_TYPE]

            write_json_file(validation_file, {"issues": [issue.as_json() for issue in validation_issues]})
            write_json_file(duplicates_file, {"duplicates": [record.as_json() for record in duplicate_records]})
            logger.info("Validation JSON written: %s", validation_file)
            logger.info("Duplicate JSON written: %s", duplicates_file)

            if sales_rows:
                write_sales_backend_report(
                    output_path=sales_output_file,
                    rows=sales_rows,
                    run_summary=summary,
                    file_audit=sales_audit,
                    validation_issues=sales_issues,
                    duplicates=sales_duplicates,
                    mapping_audit=mapping_audit,
                )
                _copy_latest_or_warn(
                    source=sales_output_file,
                    target=latest_sales_file,
                    run_id=run_id,
                    artifact_name="Latest sales backend workbook",
                    validation_issues=validation_issues,
                    logger=logger,
                )
                sales_output_generated = True
                logger.info("Sales backend workbook written: %s", sales_output_file)
            else:
                logger.warning("Sales backend workbook not generated; no sales rows were accepted.")

            if inventory_rows:
                write_inventory_backend_report(
                    output_path=inventory_output_file,
                    rows=inventory_rows,
                    run_summary=summary,
                    file_audit=inventory_audit,
                    validation_issues=inventory_issues,
                    duplicates=inventory_duplicates,
                    mapping_audit=mapping_audit,
                )
                _copy_latest_or_warn(
                    source=inventory_output_file,
                    target=latest_inventory_file,
                    run_id=run_id,
                    artifact_name="Latest inventory backend workbook",
                    validation_issues=validation_issues,
                    logger=logger,
                )
                inventory_output_generated = True
                logger.info("Inventory backend workbook written: %s", inventory_output_file)
            else:
                logger.warning("Inventory backend workbook not generated; no inventory rows were accepted.")

            write_sales_inventory_run_summary(
                output_path=summary_output_file,
                run_summary=summary,
                source_status_rows=source_status_rows,
                validation_issues=validation_issues,
                mapping_audit=mapping_audit,
            )
            _copy_latest_or_warn(
                source=summary_output_file,
                target=latest_summary_file,
                run_id=run_id,
                artifact_name="Latest run summary workbook",
                validation_issues=validation_issues,
                logger=logger,
            )
            logger.info("Combined run summary workbook written: %s", summary_output_file)

            write_json_file(validation_file, {"issues": [issue.as_json() for issue in validation_issues]})

            end_time = datetime.now()
            metadata.update(
                {
                    "status": "SUCCESS",
                    "end_time": end_time.isoformat(),
                    "duration_seconds": round((end_time - start_time).total_seconds(), 3),
                    "sales_files_discovered": [str(path) for path in discovered_sales_files],
                    "inventory_files_discovered": [str(path) for path in discovered_inventory_files],
                    "mapping_files_discovered": [str(path) for path in discovered_mapping_files],
                    "sales_files_processed_successfully": _processed_count(file_audit, SALES_REPORT_TYPE),
                    "inventory_files_processed_successfully": _processed_count(file_audit, INVENTORY_REPORT_TYPE),
                    "sales_rows_scanned": _rows_scanned(file_audit, SALES_REPORT_TYPE),
                    "sales_rows_written": len(sales_rows),
                    "inventory_rows_scanned": _rows_scanned(file_audit, INVENTORY_REPORT_TYPE),
                    "inventory_rows_written": len(inventory_rows),
                    "warning_count": _count_warnings(validation_issues),
                    "error_count": _count_errors(validation_issues),
                    "duplicate_count": len(duplicate_records),
                    "sales_output_file": str(sales_output_file) if sales_output_generated else None,
                    "inventory_output_file": str(inventory_output_file) if inventory_output_generated else None,
                    "run_summary_file": str(summary_output_file),
                    "latest_sales_backend_file": str(latest_sales_file) if sales_output_generated else None,
                    "latest_inventory_backend_file": str(latest_inventory_file) if inventory_output_generated else None,
                    "latest_run_summary_file": str(latest_summary_file),
                    "metadata_file": str(metadata_file),
                    "log_file": str(log_file),
                    "validation_file": str(validation_file),
                    "duplicates_file": str(duplicates_file),
                    "file_audit": [record.as_json() for record in file_audit],
                    "mapping_audit": [record.as_json() for record in mapping_audit],
                    "duplicates": [record.as_json() for record in duplicate_records],
                }
            )
            write_json_file(metadata_file, metadata)
            logger.info("Run metadata written: %s", metadata_file)
            logger.info(
                "Sales & Inventory pipeline completed. sales_rows=%s inventory_rows=%s issues=%s duplicates=%s",
                len(sales_rows),
                len(inventory_rows),
                len(validation_issues),
                len(duplicate_records),
            )

            return SalesInventoryPipelineRunResult(
                run_id=run_id,
                run_dir=run_dir,
                sales_output_file=sales_output_file if sales_output_generated else None,
                inventory_output_file=inventory_output_file if inventory_output_generated else None,
                summary_output_file=summary_output_file,
                latest_sales_backend_file=latest_sales_file if sales_output_generated else None,
                latest_inventory_backend_file=latest_inventory_file if inventory_output_generated else None,
                latest_run_summary_file=latest_summary_file,
                metadata_file=metadata_file,
                log_file=log_file,
                sales_rows_written=len(sales_rows),
                inventory_rows_written=len(inventory_rows),
                validation_issue_count=len(validation_issues),
                duplicate_count=len(duplicate_records),
            )
        except CatastrophicPipelineError as exc:
            logger.error("Sales & Inventory pipeline failed: %s", exc)
            _write_failure_artifacts(
                metadata=metadata,
                metadata_file=metadata_file,
                validation_file=validation_file,
                duplicates_file=duplicates_file,
                validation_issues=validation_issues,
                duplicate_records=duplicate_records,
                file_audit=file_audit,
                discovered_sales_files=discovered_sales_files,
                discovered_inventory_files=discovered_inventory_files,
                discovered_mapping_files=discovered_mapping_files,
                start_time=start_time,
                error=str(exc),
            )
            raise


def _process_source_file(
    source_path: Path,
    report_type: ReportType,
    run_id: str,
    copied_path: Path | None,
    config: SalesInventoryPipelineConfig,
    file_audit: list[SalesInventoryFileAuditRecord],
    validation_issues: list[SalesInventoryValidationIssue],
    mapping_lookup: SalesInventoryMappingLookup | None,
    logger: Any,
) -> list[NormalizedSalesInventoryRow]:
    audit = SalesInventoryFileAuditRecord(
        run_id=run_id,
        report_type=report_type,
        source_file=source_path.name,
        full_path=str(source_path),
        copied_run_path=str(copied_path or ""),
    )
    file_audit.append(audit)
    try:
        read_result = read_sales_inventory_workbook(source_path, config, report_type)
    except FileValidationError as exc:
        audit.status = "FAILED"
        audit.notes = str(exc)
        validation_issues.append(
            SalesInventoryValidationIssue(
                run_id=run_id,
                report_type=report_type,
                severity="ERROR",
                issue_type="FILE_VALIDATION_FAILED",
                source_file=source_path.name,
                issue_detail=str(exc),
                action_taken="Run failed because the source file could not be safely processed.",
            )
        )
        raise CatastrophicPipelineError(f"{report_type} source file failed validation: {source_path.name}: {exc}") from exc

    audit.source_sheet = read_result.sheet_name
    audit.header_row_found = read_result.header_row
    audit.rows_scanned = len(read_result.rows)
    audit.rows_blank_skipped = read_result.rows_blank_skipped
    audit.expected_columns_count = len(read_result.expected_columns)
    audit.found_expected_columns_count = len(read_result.found_expected_columns)
    audit.missing_expected_columns = ", ".join(read_result.missing_expected_columns)
    audit.extra_source_columns = ", ".join(read_result.extra_source_columns)

    if read_result.missing_expected_columns:
        validation_issues.append(
            SalesInventoryValidationIssue(
                run_id=run_id,
                report_type=report_type,
                severity="WARNING",
                issue_type="MISSING_EXPECTED_COLUMNS",
                source_file=source_path.name,
                source_sheet=read_result.sheet_name,
                field_name="Columns",
                raw_value=", ".join(read_result.missing_expected_columns),
                issue_detail="Some expected columns were not present in the source export.",
                action_taken="Missing columns were written as blank values.",
            )
        )
    if read_result.extra_source_columns:
        validation_issues.append(
            SalesInventoryValidationIssue(
                run_id=run_id,
                report_type=report_type,
                severity="INFO",
                issue_type="EXTRA_SOURCE_COLUMNS",
                source_file=source_path.name,
                source_sheet=read_result.sheet_name,
                field_name="Columns",
                raw_value=", ".join(read_result.extra_source_columns),
                issue_detail="Unknown source columns were present.",
                action_taken="Unknown columns were ignored after being logged.",
            )
        )

    rows: list[NormalizedSalesInventoryRow] = []
    rejected = 0
    for raw_row in read_result.rows:
        result = normalize_sales_inventory_row(raw_row, run_id, mapping_lookup=mapping_lookup)
        validation_issues.extend(result.issues)
        if result.rejected:
            rejected += 1
        if result.normalized_row is not None:
            rows.append(result.normalized_row)

    audit.rows_accepted = len(rows)
    audit.rows_rejected = rejected
    issue_count = sum(
        1
        for issue in validation_issues
        if issue.source_file == source_path.name and issue.report_type == report_type
    )
    if rejected:
        audit.status = "SUCCESS_WITH_REJECTIONS"
        audit.notes = f"{rejected} row(s) rejected; {issue_count} issue(s) logged."
    elif issue_count:
        audit.status = "SUCCESS_WITH_WARNINGS"
        audit.notes = f"{issue_count} issue(s) logged."
    else:
        audit.status = "SUCCESS"
        audit.notes = ""

    logger.info(
        "Processed %s file=%s sheet=%s header_row=%s rows_scanned=%s rows_accepted=%s rows_rejected=%s",
        report_type,
        source_path.name,
        audit.source_sheet,
        audit.header_row_found,
        audit.rows_scanned,
        audit.rows_accepted,
        audit.rows_rejected,
    )
    return rows


def _build_run_summary(
    run_id: str,
    start_time: datetime,
    config: SalesInventoryPipelineConfig,
    output_folder: Path,
    discovered_sales_files: list[Path],
    discovered_inventory_files: list[Path],
    file_audit: list[SalesInventoryFileAuditRecord],
    validation_issues: list[SalesInventoryValidationIssue],
    duplicate_count: int,
    sales_output_file: Path | None,
    inventory_output_file: Path | None,
    latest_sales_file: Path | None,
    latest_inventory_file: Path | None,
    metadata_file: Path,
    log_file: Path,
) -> dict[str, Any]:
    status = "SUCCESS_WITH_WARNINGS" if _count_warnings(validation_issues) else "SUCCESS"
    return {
        "Run ID": run_id,
        "Run timestamp": start_time.isoformat(timespec="seconds"),
        "Pipeline name": PIPELINE_NAME,
        "Input sales folder": str(config.sales_input_dir),
        "Input inventory folder": str(config.inventory_input_dir),
        "Output folder": str(output_folder),
        "Sales files discovered": len(discovered_sales_files),
        "Inventory files discovered": len(discovered_inventory_files),
        "Sales files processed successfully": _processed_count(file_audit, SALES_REPORT_TYPE),
        "Inventory files processed successfully": _processed_count(file_audit, INVENTORY_REPORT_TYPE),
        "Sales files skipped/failed": _failed_count(file_audit, SALES_REPORT_TYPE),
        "Inventory files skipped/failed": _failed_count(file_audit, INVENTORY_REPORT_TYPE),
        "Sales rows scanned": _rows_scanned(file_audit, SALES_REPORT_TYPE),
        "Sales rows accepted": _rows_accepted(file_audit, SALES_REPORT_TYPE),
        "Sales rows rejected": _rows_rejected(file_audit, SALES_REPORT_TYPE),
        "Inventory rows scanned": _rows_scanned(file_audit, INVENTORY_REPORT_TYPE),
        "Inventory rows accepted": _rows_accepted(file_audit, INVENTORY_REPORT_TYPE),
        "Inventory rows rejected": _rows_rejected(file_audit, INVENTORY_REPORT_TYPE),
        "Warning count": _count_warnings(validation_issues),
        "Error count": _count_errors(validation_issues),
        "Duplicate count": duplicate_count,
        "Sales output file": str(sales_output_file) if sales_output_file else "NOT GENERATED",
        "Inventory output file": str(inventory_output_file) if inventory_output_file else "NOT GENERATED",
        "Latest sales backend file": str(latest_sales_file) if latest_sales_file else "NOT GENERATED",
        "Latest inventory backend file": str(latest_inventory_file) if latest_inventory_file else "NOT GENERATED",
        "Metadata file": str(metadata_file),
        "Log file": str(log_file),
        "Status": status,
    }


def _build_source_status_rows(
    config: SalesInventoryPipelineConfig,
    discovered_sales_files: list[Path],
    discovered_inventory_files: list[Path],
    file_audit: list[SalesInventoryFileAuditRecord],
    sales_output_file: Path | None,
    inventory_output_file: Path | None,
    latest_sales_file: Path | None,
    latest_inventory_file: Path | None,
) -> list[dict[str, Any]]:
    return [
        {
            "Report Type": SALES_REPORT_TYPE,
            "Input Folder": str(config.sales_input_dir),
            "Files Discovered": len(discovered_sales_files),
            "Files Processed Successfully": _processed_count(file_audit, SALES_REPORT_TYPE),
            "Files Skipped/Failed": _failed_count(file_audit, SALES_REPORT_TYPE),
            "Rows Scanned": _rows_scanned(file_audit, SALES_REPORT_TYPE),
            "Rows Accepted": _rows_accepted(file_audit, SALES_REPORT_TYPE),
            "Rows Rejected": _rows_rejected(file_audit, SALES_REPORT_TYPE),
            "Output File": str(sales_output_file) if sales_output_file else "NOT GENERATED",
            "Latest Backend File": str(latest_sales_file) if latest_sales_file else "NOT GENERATED",
            "Status": "PROCESSED" if sales_output_file else "NOT GENERATED",
        },
        {
            "Report Type": INVENTORY_REPORT_TYPE,
            "Input Folder": str(config.inventory_input_dir),
            "Files Discovered": len(discovered_inventory_files),
            "Files Processed Successfully": _processed_count(file_audit, INVENTORY_REPORT_TYPE),
            "Files Skipped/Failed": _failed_count(file_audit, INVENTORY_REPORT_TYPE),
            "Rows Scanned": _rows_scanned(file_audit, INVENTORY_REPORT_TYPE),
            "Rows Accepted": _rows_accepted(file_audit, INVENTORY_REPORT_TYPE),
            "Rows Rejected": _rows_rejected(file_audit, INVENTORY_REPORT_TYPE),
            "Output File": str(inventory_output_file) if inventory_output_file else "NOT GENERATED",
            "Latest Backend File": str(latest_inventory_file) if latest_inventory_file else "NOT GENERATED",
            "Status": "PROCESSED" if inventory_output_file else "NOT GENERATED",
        },
    ]


def _issues_for_report(
    validation_issues: list[SalesInventoryValidationIssue],
    report_type: ReportType,
) -> list[SalesInventoryValidationIssue]:
    return [issue for issue in validation_issues if issue.report_type in {report_type, "SOURCE"}]


def _write_failure_artifacts(
    metadata: dict[str, Any],
    metadata_file: Path,
    validation_file: Path,
    duplicates_file: Path,
    validation_issues: list[SalesInventoryValidationIssue],
    duplicate_records: list[SalesInventoryDuplicateRecord],
    file_audit: list[SalesInventoryFileAuditRecord],
    discovered_sales_files: list[Path],
    discovered_inventory_files: list[Path],
    discovered_mapping_files: list[Path],
    start_time: datetime,
    error: str,
) -> None:
    end_time = datetime.now()
    metadata.update(
        {
            "status": "FAILED",
            "end_time": end_time.isoformat(),
            "duration_seconds": round((end_time - start_time).total_seconds(), 3),
            "error": error,
            "sales_files_discovered": [str(path) for path in discovered_sales_files],
            "inventory_files_discovered": [str(path) for path in discovered_inventory_files],
            "mapping_files_discovered": [str(path) for path in discovered_mapping_files],
            "warning_count": _count_warnings(validation_issues),
            "error_count": _count_errors(validation_issues),
            "duplicate_count": len(duplicate_records),
            "file_audit": [record.as_json() for record in file_audit],
        }
    )
    write_json_file(validation_file, {"issues": [issue.as_json() for issue in validation_issues]})
    write_json_file(duplicates_file, {"duplicates": [record.as_json() for record in duplicate_records]})
    write_json_file(metadata_file, metadata)


def _create_run_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _create_run_dirs(run_dir: Path) -> dict[str, Path]:
    if run_dir.exists():
        time.sleep(1.05)
        run_dir = run_dir.parent / _create_run_id()
    paths = {
        "run": run_dir,
        "inputs_sales": run_dir / "inputs" / "sales_inventory" / "sales",
        "inputs_inventory": run_dir / "inputs" / "sales_inventory" / "inventory",
        "inputs_mapping": run_dir / "inputs" / "sales_inventory" / "mapping",
        "outputs_sales_inventory": run_dir / "outputs" / "sales_inventory",
        "logs": run_dir / "logs",
        "validation": run_dir / "validation",
        "metadata": run_dir / "metadata",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    return paths


def _copy_inputs(files: list[Path], destination: Path) -> dict[Path, Path]:
    copied: dict[Path, Path] = {}
    destination.mkdir(parents=True, exist_ok=True)
    for path in files:
        target = destination / path.name
        shutil.copy2(path, target)
        copied[path] = target
    return copied


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
    validation_issues: list[SalesInventoryValidationIssue],
    logger: Any,
) -> bool:
    try:
        _copy_latest(source, target)
        logger.info("%s updated: %s", artifact_name, target)
        return True
    except PermissionError as exc:
        temp_latest = target.with_suffix(target.suffix + ".tmp")
        if temp_latest.exists():
            temp_latest.unlink(missing_ok=True)
        detail = f"{artifact_name} could not be replaced, likely because the file is open: {target}"
        validation_issues.append(
            SalesInventoryValidationIssue(
                run_id=run_id,
                report_type="SOURCE",
                severity="WARNING",
                issue_type="LATEST_COPY_FAILED",
                field_name=artifact_name,
                raw_value=str(target),
                issue_detail=f"{detail}. {exc}",
                action_taken="Timestamped run output was kept; close the latest workbook and rerun to refresh latest copy.",
            )
        )
        logger.warning("%s", detail)
        return False


def _annotate_dedupe_drops(
    file_audit: list[SalesInventoryFileAuditRecord],
    all_rows: list[NormalizedSalesInventoryRow],
    kept_rows: list[NormalizedSalesInventoryRow],
) -> None:
    kept_ids = {id(row) for row in kept_rows}
    dropped_counts: dict[tuple[str, str], int] = {}
    for row in all_rows:
        if id(row) in kept_ids:
            continue
        key = (str(row.data.get("Report Type") or ""), str(row.data.get("Source File") or ""))
        dropped_counts[key] = dropped_counts.get(key, 0) + 1
    for record in file_audit:
        dropped = dropped_counts.get((record.report_type, record.source_file), 0)
        if not dropped:
            continue
        record.rows_rejected += dropped
        record.rows_accepted = max(record.rows_accepted - dropped, 0)
        note = f"{dropped} exact duplicate row(s) dropped by dedupe."
        record.notes = f"{record.notes} {note}".strip()


def _processed_count(file_audit: list[SalesInventoryFileAuditRecord], report_type: ReportType) -> int:
    return sum(1 for record in file_audit if record.report_type == report_type and record.status.startswith("SUCCESS"))


def _failed_count(file_audit: list[SalesInventoryFileAuditRecord], report_type: ReportType) -> int:
    return sum(1 for record in file_audit if record.report_type == report_type and record.status == "FAILED")


def _rows_scanned(file_audit: list[SalesInventoryFileAuditRecord], report_type: ReportType) -> int:
    return sum(record.rows_scanned for record in file_audit if record.report_type == report_type)


def _rows_accepted(file_audit: list[SalesInventoryFileAuditRecord], report_type: ReportType) -> int:
    return sum(record.rows_accepted for record in file_audit if record.report_type == report_type)


def _rows_rejected(file_audit: list[SalesInventoryFileAuditRecord], report_type: ReportType) -> int:
    return sum(record.rows_rejected for record in file_audit if record.report_type == report_type)


def _count_warnings(issues: list[SalesInventoryValidationIssue]) -> int:
    return sum(1 for issue in issues if issue.severity.upper() == "WARNING")


def _count_errors(issues: list[SalesInventoryValidationIssue]) -> int:
    return sum(1 for issue in issues if issue.severity.upper() == "ERROR")
