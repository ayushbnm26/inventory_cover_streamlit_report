"""Production-grade orchestration for Amazon PO Items consolidation."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
import time
from typing import Any

from inventory_cover.config import PipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError, FileValidationError
from inventory_cover.io.excel_io import read_po_items_workbook
from inventory_cover.io.file_discovery import discover_po_item_files
from inventory_cover.logging_utils import setup_run_logger, write_json_file
from inventory_cover.normalization.po_items_normalizer import normalize_po_item_row
from inventory_cover.reports.po_items_excel_writer import write_po_items_report, write_po_items_team_workbook
from inventory_cover.schemas import (
    FileAuditRecord,
    NormalizedPoItemRow,
    PipelineRunResult,
    ValidationIssue,
)
from inventory_cover.validation.po_items_validators import attach_duplicate_findings


class PoItemsPipeline:
    """Orchestrates discovery, reading, normalization, validation, and reporting."""

    def __init__(self, config: PipelineConfig):
        self.config = config.resolved()

    def run(self) -> PipelineRunResult:
        run_id = _create_run_id()
        run_dir = self.config.run_root / run_id
        paths = _create_run_dirs(run_dir)
        run_dir = paths["run"]
        log_file = paths["logs"] / "po_items_pipeline.log"
        logger = setup_run_logger(log_file, self.config.log_level)
        start_time = datetime.now()

        logger.info("PO Items pipeline started. run_id=%s", run_id)
        logger.info("Input directory: %s", self.config.input_dir)

        metadata: dict[str, Any] = {
            "run_id": run_id,
            "start_time": start_time.isoformat(),
            "input_directory": str(self.config.input_dir),
            "run_directory": str(run_dir),
            "status": "STARTED",
        }

        discovered_files: list[Path] = []
        file_audit: list[FileAuditRecord] = []
        validation_issues: list[ValidationIssue] = []
        duplicate_records = []
        rows: list[NormalizedPoItemRow] = []
        discovery_warnings: list[str] = []
        output_file = paths["outputs_po"] / f"PO_Items_Team_Workbook_{run_id}.xlsx"
        backend_output_file = paths["outputs_po"] / f"PO_Items_Backend_Audit_{run_id}.xlsx"
        latest_file = self.config.processed_dir / "latest" / "PO_Items_Team_Workbook_latest.xlsx"
        backend_latest_file = self.config.processed_dir / "latest" / "PO_Items_Backend_Audit_latest.xlsx"
        metadata_file = paths["metadata"] / "run_metadata.json"

        try:
            discovered_files, discovery_warnings = discover_po_item_files(self.config)
            for warning in discovery_warnings:
                logger.warning(warning)
            logger.info("Files discovered: %s", len(discovered_files))

            copied_paths = _copy_inputs(discovered_files, paths["inputs_po"])

            for source_path in discovered_files:
                audit = FileAuditRecord(
                    file_name=source_path.name,
                    full_path=str(source_path),
                    copied_run_path=str(copied_paths.get(source_path, "")),
                )
                file_audit.append(audit)
                try:
                    read_result = read_po_items_workbook(source_path, self.config)
                    audit.sheet_used = read_result.sheet_name
                    audit.header_row_found = read_result.header_row
                    audit.rows_read = len(read_result.rows)
                    normalized_rows = [
                        normalize_po_item_row(raw_row, run_id, self.config.default_currency)
                        for raw_row in read_result.rows
                    ]
                    rows.extend(normalized_rows)
                    audit.rows_accepted = len(normalized_rows)
                    issue_count = sum(len(row.issues) for row in normalized_rows)
                    audit.status = "SUCCESS_WITH_WARNINGS" if issue_count else "SUCCESS"
                    if issue_count:
                        audit.notes = f"{issue_count} row-level validation issue(s)."
                    logger.info(
                        "Processed %s: rows_read=%s rows_accepted=%s issues=%s",
                        source_path.name,
                        audit.rows_read,
                        audit.rows_accepted,
                        issue_count,
                    )
                except FileValidationError as exc:
                    audit.status = "FAILED"
                    audit.notes = str(exc)
                    validation_issues.append(
                        ValidationIssue(
                            severity="ERROR",
                            issue_type="FILE_SKIPPED",
                            source_file=source_path.name,
                            issue_detail=str(exc),
                            action_taken="File skipped; run continued with other valid files.",
                        )
                    )
                    logger.error("Skipped %s: %s", source_path.name, exc)
                except Exception as exc:
                    audit.status = "FAILED"
                    audit.notes = f"Unexpected file processing error: {exc}"
                    validation_issues.append(
                        ValidationIssue(
                            severity="ERROR",
                            issue_type="FILE_PROCESSING_ERROR",
                            source_file=source_path.name,
                            issue_detail=str(exc),
                            action_taken="File skipped; run continued with other valid files.",
                        )
                    )
                    logger.exception("Unexpected error while processing %s", source_path.name)

            if not rows:
                raise CatastrophicPipelineError("Zero valid rows after processing; output not created.")

            validation_issues.extend(issue for row in rows for issue in row.issues)
            rows_to_write, duplicate_issues, duplicate_records = attach_duplicate_findings(
                rows,
                dedupe_exact_rows=self.config.dedupe_exact_rows,
            )
            validation_issues.extend(duplicate_issues)
            _apply_dedupe_rejections(file_audit, rows, rows_to_write)

            if not rows_to_write:
                raise CatastrophicPipelineError("All rows were removed by dedupe; output not created.")

            summary = _build_run_summary(
                run_id=run_id,
                start_time=start_time,
                input_folder=self.config.input_dir,
                output_folder=paths["outputs_po"],
                discovered_files=discovered_files,
                file_audit=file_audit,
                rows_to_write=rows_to_write,
                validation_issues=validation_issues,
                duplicate_count=len(duplicate_records),
                output_file=output_file,
                backend_output_file=backend_output_file,
                discovery_warning_count=len(discovery_warnings),
            )

            write_po_items_report(
                output_path=backend_output_file,
                rows=rows_to_write,
                run_summary=summary,
                file_audit=file_audit,
                validation_issues=validation_issues,
                duplicates=duplicate_records,
            )
            logger.info("Backend audit workbook written: %s", backend_output_file)

            write_po_items_team_workbook(output_path=output_file, rows=rows_to_write)
            logger.info("Team workbook written: %s", output_file)

            _copy_latest_or_warn(
                source=output_file,
                target=latest_file,
                artifact_name="Latest team workbook",
                validation_issues=validation_issues,
                logger=logger,
            )
            _copy_latest_or_warn(
                source=backend_output_file,
                target=backend_latest_file,
                artifact_name="Latest backend audit workbook",
                validation_issues=validation_issues,
                logger=logger,
            )

            end_time = datetime.now()
            metadata.update(
                {
                    "status": "SUCCESS",
                    "end_time": end_time.isoformat(),
                    "duration_seconds": round((end_time - start_time).total_seconds(), 3),
                    "files_discovered": [str(path) for path in discovered_files],
                    "files_processed_successfully": sum(
                        1 for record in file_audit if record.status.startswith("SUCCESS")
                    ),
                    "files_skipped_or_failed": sum(1 for record in file_audit if record.status == "FAILED"),
                    "rows_read": sum(record.rows_read for record in file_audit),
                    "rows_written": len(rows_to_write),
                    "rejected_rows": sum(record.rows_rejected for record in file_audit),
                    "warning_count": _count_warnings(validation_issues) + len(discovery_warnings),
                    "error_count": _count_errors(validation_issues),
                    "duplicate_count": len(duplicate_records),
                    "output_file": str(output_file),
                    "latest_file": str(latest_file),
                    "team_workbook": str(output_file),
                    "latest_team_workbook": str(latest_file),
                    "backend_audit_workbook": str(backend_output_file),
                    "latest_backend_audit_workbook": str(backend_latest_file),
                    "log_file": str(log_file),
                    "validation_summary": {
                        "issues": len(validation_issues),
                        "warnings": _count_warnings(validation_issues),
                        "errors": _count_errors(validation_issues),
                    },
                    "file_audit": [record.as_json() for record in file_audit],
                    "duplicates": [record.as_json() for record in duplicate_records],
                }
            )
            write_json_file(
                paths["validation"] / "validation_issues.json",
                {"issues": [i.as_json() for i in validation_issues]},
            )
            write_json_file(
                paths["validation"] / "duplicates.json",
                {"duplicates": [d.as_json() for d in duplicate_records]},
            )
            write_json_file(metadata_file, metadata)
            logger.info("Run metadata written: %s", metadata_file)
            logger.info("PO Items pipeline completed successfully.")

            return PipelineRunResult(
                run_id=run_id,
                run_dir=run_dir,
                output_file=output_file,
                latest_file=latest_file,
                backend_output_file=backend_output_file,
                backend_latest_file=backend_latest_file,
                metadata_file=metadata_file,
                log_file=log_file,
                rows_written=len(rows_to_write),
                validation_issue_count=len(validation_issues),
                duplicate_count=len(duplicate_records),
            )
        except CatastrophicPipelineError as exc:
            end_time = datetime.now()
            metadata.update(
                {
                    "status": "FAILED",
                    "end_time": end_time.isoformat(),
                    "duration_seconds": round((end_time - start_time).total_seconds(), 3),
                    "error": str(exc),
                    "files_discovered": [str(path) for path in discovered_files],
                    "file_audit": [record.as_json() for record in file_audit],
                }
            )
            write_json_file(metadata_file, metadata)
            logger.error("PO Items pipeline failed: %s", exc)
            raise


def _create_run_id() -> str:
    base = datetime.now().strftime("%Y%m%d_%H%M%S")
    return base


def _create_run_dirs(run_dir: Path) -> dict[str, Path]:
    if run_dir.exists():
        time.sleep(1.05)
        run_dir = run_dir.parent / _create_run_id()
    paths = {
        "run": run_dir,
        "inputs_po": run_dir / "inputs" / "po_items",
        "outputs_po": run_dir / "outputs" / "po_items",
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
    artifact_name: str,
    validation_issues: list[ValidationIssue],
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
            ValidationIssue(
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


def _apply_dedupe_rejections(
    audit_records: list[FileAuditRecord],
    all_rows: list[NormalizedPoItemRow],
    kept_rows: list[NormalizedPoItemRow],
) -> None:
    kept_ids = {id(row) for row in kept_rows}
    rejected_by_file: dict[str, int] = {}
    for row in all_rows:
        if id(row) not in kept_ids:
            file_name = str(row.data.get("Source File") or "")
            rejected_by_file[file_name] = rejected_by_file.get(file_name, 0) + 1
    for record in audit_records:
        rejected = rejected_by_file.get(record.file_name, 0)
        if rejected:
            record.rows_rejected += rejected
            record.rows_accepted = max(record.rows_accepted - rejected, 0)
            note = f"{rejected} duplicate row(s) dropped by dedupe."
            record.notes = f"{record.notes} {note}".strip()


def _build_run_summary(
    run_id: str,
    start_time: datetime,
    input_folder: Path,
    output_folder: Path,
    discovered_files: list[Path],
    file_audit: list[FileAuditRecord],
    rows_to_write: list[NormalizedPoItemRow],
    validation_issues: list[ValidationIssue],
    duplicate_count: int,
    output_file: Path,
    backend_output_file: Path,
    discovery_warning_count: int,
) -> dict[str, Any]:
    return {
        "Run ID": run_id,
        "Run timestamp": start_time.isoformat(timespec="seconds"),
        "Input folder": str(input_folder),
        "Output folder": str(output_folder),
        "Files discovered": len(discovered_files),
        "Files processed successfully": sum(1 for record in file_audit if record.status.startswith("SUCCESS")),
        "Files skipped/failed": sum(1 for record in file_audit if record.status == "FAILED"),
        "Total source rows": sum(record.rows_read for record in file_audit),
        "Valid rows written": len(rows_to_write),
        "Warning count": _count_warnings(validation_issues) + discovery_warning_count,
        "Error count": _count_errors(validation_issues),
        "Duplicate count": duplicate_count,
        "Output file name": output_file.name,
        "Team workbook file name": output_file.name,
        "Backend audit workbook file name": backend_output_file.name,
    }


def _count_warnings(issues: list[ValidationIssue]) -> int:
    return sum(1 for issue in issues if issue.severity.upper() == "WARNING")


def _count_errors(issues: list[ValidationIssue]) -> int:
    return sum(1 for issue in issues if issue.severity.upper() == "ERROR")
