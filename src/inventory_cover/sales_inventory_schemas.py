"""Schemas and constants for Vendor Central Sales & Inventory ingestion."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Literal

from inventory_cover.schemas import json_safe


ReportType = Literal["SALES", "INVENTORY"]

SALES_REPORT_TYPE: ReportType = "SALES"
INVENTORY_REPORT_TYPE: ReportType = "INVENTORY"
PIPELINE_NAME = "Vendor Central Sales & Inventory Pipeline"

COMMON_SOURCE_COLUMNS: tuple[str, ...] = (
    "ASIN",
    "Child Vendor Code",
    "Product Title",
    "Brand Code",
    "Brand",
    "Category",
    "Subcategory",
    "Parent ASIN",
    "UPC",
    "EAN",
    "ISBN",
    "Model Number",
    "Store Code",
    "MSRP",
    "Binding",
    "Colour",
    "Release Date",
    "Replenishment Code",
)

SALES_SOURCE_COLUMNS: tuple[str, ...] = COMMON_SOURCE_COLUMNS + (
    "Shipped Revenue",
    "Shipped COGS",
    "Shipped Units",
    "Customer Returns",
    "Confirmed Units",
    "Sales Discount",
    "Contra-COGS",
    "Net PPM %",
    "ASIN Confirmation %",
)

INVENTORY_SOURCE_COLUMNS: tuple[str, ...] = COMMON_SOURCE_COLUMNS + (
    "Vendor Confirmation %",
    "Net Received",
    "Net Received Units",
    "Open Purchase Order Quantity",
    "Receive Fill %",
    "Overall Vendor Lead Time (days)",
    "Aged 90+ Days Sellable Inventory",
    "Aged 90+ Days Sellable Units",
    "Sellable On-Hand Inventory",
    "Sellable On Hand Units",
    "Unsellable On-Hand Inventory",
    "Unsellable On-Hand Units",
    "Confirmed Units",
    "Sales Discount",
    "Contra-COGS",
    "In Transit Quantity",
    "Sellable In Transit Units",
    "Unsellable In Transit Units",
)

SALES_MINIMUM_HEADERS: tuple[str, ...] = (
    "ASIN",
    "Child Vendor Code",
    "Product Title",
    "Model Number",
    "Shipped Units",
)

INVENTORY_MINIMUM_HEADERS: tuple[str, ...] = (
    "ASIN",
    "Child Vendor Code",
    "Product Title",
    "Model Number",
    "Sellable On Hand Units",
)

TEXT_COLUMNS: frozenset[str] = frozenset(
    {
        "ASIN",
        "Child Vendor Code",
        "Product Title",
        "Brand Code",
        "Brand",
        "Category",
        "Subcategory",
        "Parent ASIN",
        "UPC",
        "EAN",
        "ISBN",
        "Model Number",
        "Store Code",
        "Binding",
        "Colour",
        "Replenishment Code",
    }
)

DATE_COLUMNS: frozenset[str] = frozenset(
    {
        "Release Date",
        "Viewing Range Start",
        "Viewing Range End",
        "Report Updated Date",
    }
)

SALES_NUMERIC_COLUMNS: frozenset[str] = frozenset(
    {
        "MSRP",
        "Shipped Revenue",
        "Shipped COGS",
        "Shipped Units",
        "Customer Returns",
        "Confirmed Units",
        "Sales Discount",
        "Contra-COGS",
        "Net PPM %",
        "ASIN Confirmation %",
    }
)

INVENTORY_NUMERIC_COLUMNS: frozenset[str] = frozenset(
    {
        "MSRP",
        "Vendor Confirmation %",
        "Net Received",
        "Net Received Units",
        "Open Purchase Order Quantity",
        "Receive Fill %",
        "Overall Vendor Lead Time (days)",
        "Aged 90+ Days Sellable Inventory",
        "Aged 90+ Days Sellable Units",
        "Sellable On-Hand Inventory",
        "Sellable On Hand Units",
        "Unsellable On-Hand Inventory",
        "Unsellable On-Hand Units",
        "Confirmed Units",
        "Sales Discount",
        "Contra-COGS",
        "In Transit Quantity",
        "Sellable In Transit Units",
        "Unsellable In Transit Units",
    }
)

INVENTORY_ZERO_DEFAULT_NUMERIC_COLUMNS: frozenset[str] = frozenset(
    {
        "Sellable On-Hand Inventory",
        "Sellable On Hand Units",
        "In Transit Quantity",
        "Sellable In Transit Units",
        "Unsellable In Transit Units",
    }
)

SALES_MASTER_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Report Type",
    "Programme",
    "Distributor View",
    "View By",
    "Country",
    "Currency",
    "Viewing Range Start",
    "Viewing Range End",
    "Report Updated Date",
    "ASIN",
    "Child Vendor Code",
    "Product Title",
    "Brand Code",
    "Brand",
    "Category",
    "Subcategory",
    "Parent ASIN",
    "UPC",
    "EAN",
    "ISBN",
    "Model Number",
    "Store Code",
    "MSRP",
    "Binding",
    "Colour",
    "Release Date",
    "Replenishment Code",
    "Shipped Revenue",
    "Shipped COGS",
    "Shipped Units",
    "Customer Returns",
    "Confirmed Units",
    "Sales Discount",
    "Contra-COGS",
    "Net PPM %",
    "ASIN Confirmation %",
    "Row Validation Status",
    "Row Validation Notes",
)

INVENTORY_MASTER_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Report Type",
    "Programme",
    "Distributor View",
    "View By",
    "Country",
    "Currency",
    "Viewing Range Start",
    "Viewing Range End",
    "Report Updated Date",
    "ASIN",
    "Child Vendor Code",
    "Product Title",
    "Brand Code",
    "Brand",
    "Category",
    "Subcategory",
    "Parent ASIN",
    "UPC",
    "EAN",
    "ISBN",
    "Model Number",
    "Store Code",
    "MSRP",
    "Binding",
    "Colour",
    "Release Date",
    "Replenishment Code",
    "Vendor Confirmation %",
    "Net Received",
    "Net Received Units",
    "Open Purchase Order Quantity",
    "Receive Fill %",
    "Overall Vendor Lead Time (days)",
    "Aged 90+ Days Sellable Inventory",
    "Aged 90+ Days Sellable Units",
    "Sellable On-Hand Inventory",
    "Sellable On Hand Units",
    "Unsellable On-Hand Inventory",
    "Unsellable On-Hand Units",
    "Confirmed Units",
    "Sales Discount",
    "Contra-COGS",
    "In Transit Quantity",
    "Sellable In Transit Units",
    "Unsellable In Transit Units",
    "Row Validation Status",
    "Row Validation Notes",
)

RUN_SUMMARY_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Run timestamp",
    "Pipeline name",
    "Input sales folder",
    "Input inventory folder",
    "Output folder",
    "Sales files discovered",
    "Inventory files discovered",
    "Sales files processed successfully",
    "Inventory files processed successfully",
    "Sales files skipped/failed",
    "Inventory files skipped/failed",
    "Sales rows scanned",
    "Sales rows accepted",
    "Sales rows rejected",
    "Inventory rows scanned",
    "Inventory rows accepted",
    "Inventory rows rejected",
    "Warning count",
    "Error count",
    "Duplicate count",
    "Sales output file",
    "Inventory output file",
    "Latest sales backend file",
    "Latest inventory backend file",
    "Metadata file",
    "Log file",
    "Status",
)

FILE_AUDIT_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Report Type",
    "Source File",
    "Full Path",
    "Copied Run Path",
    "Source Sheet",
    "Header Row Found",
    "Rows Scanned",
    "Rows Accepted",
    "Rows Rejected",
    "Rows Blank Skipped",
    "Expected Columns Count",
    "Found Expected Columns Count",
    "Missing Expected Columns",
    "Extra Source Columns",
    "Status",
    "Notes",
)

VALIDATION_ISSUE_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Report Type",
    "Severity",
    "Issue Type",
    "Source File",
    "Source Sheet",
    "Source Row",
    "ASIN",
    "Child Vendor Code",
    "Model Number",
    "Field Name",
    "Raw Value",
    "Issue Detail",
    "Action Taken",
)

DUPLICATE_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Report Type",
    "Duplicate Group ID",
    "Duplicate Key",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Action Taken",
)

PROCESSING_GUIDE_HEADERS: tuple[str, ...] = (
    "Section",
    "Topic",
    "Explanation",
    "Operational Meaning",
)

MAPPING_AUDIT_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Mapping File",
    "Full Path",
    "Copied Run Path",
    "Source Sheet",
    "Header Row Found",
    "Rows Scanned",
    "Rows Loaded",
    "Unique ASIN Keys",
    "Unique SKU Keys",
    "Ambiguous ASIN Keys",
    "Ambiguous SKU Keys",
    "Status",
    "Notes",
)

SOURCE_STATUS_HEADERS: tuple[str, ...] = (
    "Report Type",
    "Input Folder",
    "Files Discovered",
    "Files Processed Successfully",
    "Files Skipped/Failed",
    "Rows Scanned",
    "Rows Accepted",
    "Rows Rejected",
    "Output File",
    "Latest Backend File",
    "Status",
)


@dataclass(frozen=True)
class SalesInventoryCellValue:
    value: Any
    number_format: str = ""
    data_type: str = ""


@dataclass(frozen=True)
class RawSalesInventoryRow:
    report_type: ReportType
    source_file: str
    source_path: Path
    source_sheet: str
    source_row: int
    values: dict[str, SalesInventoryCellValue]
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class SalesInventoryWorkbookReadResult:
    source_path: Path
    report_type: ReportType
    sheet_name: str
    header_row: int
    rows: list[RawSalesInventoryRow]
    metadata: dict[str, Any]
    raw_metadata_text: str
    expected_columns: tuple[str, ...]
    found_expected_columns: tuple[str, ...]
    missing_expected_columns: tuple[str, ...]
    extra_source_columns: tuple[str, ...]
    rows_blank_skipped: int = 0


@dataclass
class SalesInventoryValidationIssue:
    run_id: str
    report_type: str
    severity: str
    issue_type: str
    source_file: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    asin: str = ""
    child_vendor_code: str = ""
    model_number: str = ""
    field_name: str = ""
    raw_value: Any = None
    issue_detail: str = ""
    action_taken: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.report_type,
            self.severity,
            self.issue_type,
            self.source_file,
            self.source_sheet,
            self.source_row,
            self.asin,
            self.child_vendor_code,
            self.model_number,
            self.field_name,
            self.raw_value,
            self.issue_detail,
            self.action_taken,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "report_type": self.report_type,
            "severity": self.severity,
            "issue_type": self.issue_type,
            "source_file": self.source_file,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "asin": self.asin,
            "child_vendor_code": self.child_vendor_code,
            "model_number": self.model_number,
            "field_name": self.field_name,
            "raw_value": json_safe(self.raw_value),
            "issue_detail": self.issue_detail,
            "action_taken": self.action_taken,
        }


@dataclass
class NormalizedSalesInventoryRow:
    report_type: ReportType
    data: dict[str, Any]
    issues: list[SalesInventoryValidationIssue] = field(default_factory=list)
    duplicate_key: str = ""
    exact_fingerprint: str = ""
    dropped: bool = False

    def refresh_validation_status(self) -> None:
        if not self.issues:
            self.data["Row Validation Status"] = "OK"
            self.data["Row Validation Notes"] = ""
            return
        severities = {issue.severity.upper() for issue in self.issues}
        if "ERROR" in severities:
            status = "ERROR"
        elif "WARNING" in severities:
            status = "WARNING"
        else:
            status = "INFO"
        notes = "; ".join(dict.fromkeys(issue.issue_type for issue in self.issues))
        self.data["Row Validation Status"] = status
        self.data["Row Validation Notes"] = notes

    def as_master_row(self) -> list[Any]:
        self.refresh_validation_status()
        headers = SALES_MASTER_HEADERS if self.report_type == SALES_REPORT_TYPE else INVENTORY_MASTER_HEADERS
        return [self.data.get(header) for header in headers]


@dataclass
class SalesInventoryFileAuditRecord:
    run_id: str
    report_type: str
    source_file: str
    full_path: str
    copied_run_path: str = ""
    source_sheet: str = ""
    header_row_found: int | None = None
    rows_scanned: int = 0
    rows_accepted: int = 0
    rows_rejected: int = 0
    rows_blank_skipped: int = 0
    expected_columns_count: int = 0
    found_expected_columns_count: int = 0
    missing_expected_columns: str = ""
    extra_source_columns: str = ""
    status: str = "PENDING"
    notes: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.report_type,
            self.source_file,
            self.full_path,
            self.copied_run_path,
            self.source_sheet,
            self.header_row_found,
            self.rows_scanned,
            self.rows_accepted,
            self.rows_rejected,
            self.rows_blank_skipped,
            self.expected_columns_count,
            self.found_expected_columns_count,
            self.missing_expected_columns,
            self.extra_source_columns,
            self.status,
            self.notes,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "report_type": self.report_type,
            "source_file": self.source_file,
            "full_path": self.full_path,
            "copied_run_path": self.copied_run_path,
            "source_sheet": self.source_sheet,
            "header_row_found": self.header_row_found,
            "rows_scanned": self.rows_scanned,
            "rows_accepted": self.rows_accepted,
            "rows_rejected": self.rows_rejected,
            "rows_blank_skipped": self.rows_blank_skipped,
            "expected_columns_count": self.expected_columns_count,
            "found_expected_columns_count": self.found_expected_columns_count,
            "missing_expected_columns": self.missing_expected_columns,
            "extra_source_columns": self.extra_source_columns,
            "status": self.status,
            "notes": self.notes,
        }


@dataclass
class SalesInventoryDuplicateRecord:
    run_id: str
    report_type: str
    duplicate_group_id: str
    duplicate_key: str
    source_file: str
    source_sheet: str
    source_row: int
    action_taken: str

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.report_type,
            self.duplicate_group_id,
            self.duplicate_key,
            self.source_file,
            self.source_sheet,
            self.source_row,
            self.action_taken,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "report_type": self.report_type,
            "duplicate_group_id": self.duplicate_group_id,
            "duplicate_key": self.duplicate_key,
            "source_file": self.source_file,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "action_taken": self.action_taken,
        }


@dataclass
class SalesInventoryMappingAuditRecord:
    run_id: str
    mapping_file: str
    full_path: str
    copied_run_path: str = ""
    source_sheet: str = ""
    header_row_found: int | None = None
    rows_scanned: int = 0
    rows_loaded: int = 0
    unique_asin_keys: int = 0
    unique_sku_keys: int = 0
    ambiguous_asin_keys: int = 0
    ambiguous_sku_keys: int = 0
    status: str = "PENDING"
    notes: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.mapping_file,
            self.full_path,
            self.copied_run_path,
            self.source_sheet,
            self.header_row_found,
            self.rows_scanned,
            self.rows_loaded,
            self.unique_asin_keys,
            self.unique_sku_keys,
            self.ambiguous_asin_keys,
            self.ambiguous_sku_keys,
            self.status,
            self.notes,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "mapping_file": self.mapping_file,
            "full_path": self.full_path,
            "copied_run_path": self.copied_run_path,
            "source_sheet": self.source_sheet,
            "header_row_found": self.header_row_found,
            "rows_scanned": self.rows_scanned,
            "rows_loaded": self.rows_loaded,
            "unique_asin_keys": self.unique_asin_keys,
            "unique_sku_keys": self.unique_sku_keys,
            "ambiguous_asin_keys": self.ambiguous_asin_keys,
            "ambiguous_sku_keys": self.ambiguous_sku_keys,
            "status": self.status,
            "notes": self.notes,
        }


@dataclass(frozen=True)
class SalesInventoryPipelineRunResult:
    run_id: str
    run_dir: Path
    sales_output_file: Path | None
    inventory_output_file: Path | None
    summary_output_file: Path
    latest_sales_backend_file: Path | None
    latest_inventory_backend_file: Path | None
    latest_run_summary_file: Path
    metadata_file: Path
    log_file: Path
    sales_rows_written: int
    inventory_rows_written: int
    validation_issue_count: int
    duplicate_count: int
