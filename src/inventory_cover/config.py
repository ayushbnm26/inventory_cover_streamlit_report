"""Configuration defaults for local reporting pipelines."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Mapping


PROJECT_ROOT = Path(__file__).resolve().parents[2]
GDRIVE_SERVICE_ACCOUNT_DEFAULT_PATH = Path("secrets") / "google_drive" / "service_account.json"
GDRIVE_SERVICE_ACCOUNT_COMPATIBILITY_PATHS = (
    Path("secrets") / "google" / "google_drive" / "service_account.json",
)
GDRIVE_OAUTH_CREDENTIALS_DEFAULT_PATH = Path("secrets") / "google_oauth" / "credentials.json"
GDRIVE_OAUTH_TOKEN_DEFAULT_PATH = Path("secrets") / "google_oauth" / "drive_token.json"


@dataclass(frozen=True)
class PipelineConfig:
    """Runtime configuration for the PO Items consolidation pipeline."""

    project_root: Path = PROJECT_ROOT
    input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "po_items"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "po_items"
    min_files: int = 2
    max_files: int = 10
    allow_single_file: bool = False
    allow_more_than_max_files: bool = False
    dedupe_exact_rows: bool = False
    log_level: str = "INFO"
    header_scan_rows: int = 30
    default_currency: str = "INR"

    def resolved(self) -> "PipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return PipelineConfig(
            project_root=Path(self.project_root).resolve(),
            input_dir=resolve(self.input_dir),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            min_files=self.min_files,
            max_files=self.max_files,
            allow_single_file=self.allow_single_file,
            allow_more_than_max_files=self.allow_more_than_max_files,
            dedupe_exact_rows=self.dedupe_exact_rows,
            log_level=self.log_level,
            header_scan_rows=self.header_scan_rows,
            default_currency=self.default_currency,
        )


@dataclass(frozen=True)
class B2BDispatchPipelineConfig:
    """Runtime configuration for the B2B Dispatch Tracker pipeline."""

    project_root: Path = PROJECT_ROOT
    input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "b2b_dispatch"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "b2b_dispatch"
    lookback_days: int = 2
    as_of_date: date | None = None
    source_mode: str = "excel"
    allow_multiple_files: bool = False
    allow_missing_target_sheets: bool = False
    dedupe_exact_rows: bool = False
    log_level: str = "INFO"
    header_scan_rows: int = 30
    value_difference_tolerance: float = 1.0
    google_spreadsheet_id: str = ""
    google_credentials_path: Path = PROJECT_ROOT / "secrets" / "google_oauth" / "credentials.json"
    google_token_path: Path = PROJECT_ROOT / "secrets" / "google_oauth" / "token.json"
    google_readonly_scope: tuple[str, ...] = ("https://www.googleapis.com/auth/spreadsheets.readonly",)
    google_values_max_rows: int = 20000
    google_values_max_column: str = "AZ"

    def resolved(self) -> "B2BDispatchPipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return B2BDispatchPipelineConfig(
            project_root=Path(self.project_root).resolve(),
            input_dir=resolve(self.input_dir),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            lookback_days=self.lookback_days,
            as_of_date=self.as_of_date,
            source_mode=self.source_mode,
            allow_multiple_files=self.allow_multiple_files,
            allow_missing_target_sheets=self.allow_missing_target_sheets,
            dedupe_exact_rows=self.dedupe_exact_rows,
            log_level=self.log_level,
            header_scan_rows=self.header_scan_rows,
            value_difference_tolerance=self.value_difference_tolerance,
            google_spreadsheet_id=self.google_spreadsheet_id,
            google_credentials_path=resolve(self.google_credentials_path),
            google_token_path=resolve(self.google_token_path),
            google_readonly_scope=tuple(self.google_readonly_scope),
            google_values_max_rows=self.google_values_max_rows,
            google_values_max_column=self.google_values_max_column,
        )


@dataclass(frozen=True)
class SalesInventoryPipelineConfig:
    """Runtime configuration for the Vendor Central Sales & Inventory pipeline."""

    project_root: Path = PROJECT_ROOT
    sales_input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "sales"
    inventory_input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "inventory"
    mapping_input_dir: Path = PROJECT_ROOT / "data" / "reference" / "sales_inventory_mapping"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "sales_inventory"
    require_sales: bool = False
    require_inventory: bool = False
    allow_multiple_sales_files: bool = False
    allow_multiple_inventory_files: bool = False
    dedupe_exact_rows: bool = False
    log_level: str = "INFO"
    header_scan_rows: int = 20

    def resolved(self) -> "SalesInventoryPipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return SalesInventoryPipelineConfig(
            project_root=Path(self.project_root).resolve(),
            sales_input_dir=resolve(self.sales_input_dir),
            inventory_input_dir=resolve(self.inventory_input_dir),
            mapping_input_dir=resolve(self.mapping_input_dir),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            require_sales=self.require_sales,
            require_inventory=self.require_inventory,
            allow_multiple_sales_files=self.allow_multiple_sales_files,
            allow_multiple_inventory_files=self.allow_multiple_inventory_files,
            dedupe_exact_rows=self.dedupe_exact_rows,
            log_level=self.log_level,
            header_scan_rows=self.header_scan_rows,
        )


@dataclass(frozen=True)
class InventoryCoverPipelineConfig:
    """Runtime configuration for the Final Inventory Cover Calculation Engine.

    This engine consumes the latest backend artifacts produced by the source
    pipelines (PO Items, B2B Dispatch, Sales & Inventory). It does not depend on
    the internal reading logic of those pipelines; the stable backend workbook
    sheet and column headers are the interface contract.
    """

    project_root: Path = PROJECT_ROOT
    po_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "po_items" / "latest" / "PO_Items_Backend_Audit_latest.xlsx"
    )
    b2b_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "b2b_dispatch" / "latest" / "B2B_Dispatch_Backend_Audit_latest.xlsx"
    )
    sales_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "sales_inventory" / "latest" / "Sales_Backend_Audit_latest.xlsx"
    )
    inventory_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "sales_inventory" / "latest" / "Inventory_Backend_Audit_latest.xlsx"
    )
    asin_master_path: Path = PROJECT_ROOT / "data" / "reference" / "master_data" / "ASIN_Master.xlsx"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "inventory_cover"
    sales_window_days: int = 30
    default_target_doh: float = 30.0
    blank_numeric_policy: str = "zero_for_calculation"
    strict_freshness: bool = False
    sales_staleness_days: int = 45
    inventory_staleness_days: int = 7
    log_level: str = "INFO"

    def resolved(self) -> "InventoryCoverPipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return InventoryCoverPipelineConfig(
            project_root=Path(self.project_root).resolve(),
            po_backend_path=resolve(self.po_backend_path),
            b2b_backend_path=resolve(self.b2b_backend_path),
            sales_backend_path=resolve(self.sales_backend_path),
            inventory_backend_path=resolve(self.inventory_backend_path),
            asin_master_path=resolve(self.asin_master_path),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            sales_window_days=self.sales_window_days,
            default_target_doh=self.default_target_doh,
            blank_numeric_policy=self.blank_numeric_policy,
            strict_freshness=self.strict_freshness,
            sales_staleness_days=self.sales_staleness_days,
            inventory_staleness_days=self.inventory_staleness_days,
            log_level=self.log_level,
        )


@dataclass(frozen=True)
class GoogleDriveReportConfig:
    """Optional Google Drive latest-report publishing settings."""

    project_root: Path = PROJECT_ROOT
    enabled: bool = False
    auth_mode: str = "service_account"
    folder_id: str = ""
    service_account_json_path: Path = PROJECT_ROOT / GDRIVE_SERVICE_ACCOUNT_DEFAULT_PATH
    service_account_json: str = ""
    oauth_credentials_path: Path = PROJECT_ROOT / GDRIVE_OAUTH_CREDENTIALS_DEFAULT_PATH
    oauth_token_path: Path = PROJECT_ROOT / GDRIVE_OAUTH_TOKEN_DEFAULT_PATH
    report_file_name: str = "Inventory_Cover_Report_latest.xlsx"
    audit_file_name: str = "Inventory_Cover_Backend_Audit_latest.xlsx"
    upload_audit: bool = False
    fail_on_upload_error: bool = False
    scopes: tuple[str, ...] = ("https://www.googleapis.com/auth/drive",)

    def resolved(self) -> "GoogleDriveReportConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        service_account_json_path = resolve(self.service_account_json_path)
        default_service_account_path = resolve(GDRIVE_SERVICE_ACCOUNT_DEFAULT_PATH)
        if service_account_json_path == default_service_account_path and not service_account_json_path.exists():
            for candidate in GDRIVE_SERVICE_ACCOUNT_COMPATIBILITY_PATHS:
                candidate_path = resolve(candidate)
                if candidate_path.is_file():
                    service_account_json_path = candidate_path
                    break

        return GoogleDriveReportConfig(
            project_root=Path(self.project_root).resolve(),
            enabled=self.enabled,
            auth_mode=_normalize_auth_mode(self.auth_mode),
            folder_id=self.folder_id,
            service_account_json_path=service_account_json_path,
            service_account_json=self.service_account_json,
            oauth_credentials_path=resolve(self.oauth_credentials_path),
            oauth_token_path=resolve(self.oauth_token_path),
            report_file_name=self.report_file_name,
            audit_file_name=self.audit_file_name,
            upload_audit=self.upload_audit,
            fail_on_upload_error=self.fail_on_upload_error,
            scopes=tuple(self.scopes),
        )


def google_drive_report_config_from_values(
    values: Mapping[str, str],
    *,
    project_root: Path = PROJECT_ROOT,
) -> GoogleDriveReportConfig:
    """Build Google Drive report config from environment-style key/value pairs."""

    default = GoogleDriveReportConfig(project_root=project_root)
    return GoogleDriveReportConfig(
        project_root=project_root,
        enabled=_parse_config_bool(values.get("GDRIVE_ENABLED"), default=default.enabled),
        auth_mode=_clean_config_value(values.get("GDRIVE_AUTH_MODE")) or default.auth_mode,
        folder_id=_clean_config_value(values.get("GDRIVE_FOLDER_ID")),
        service_account_json_path=Path(
            _clean_config_value(values.get("GDRIVE_SERVICE_ACCOUNT_JSON_PATH"))
            or GDRIVE_SERVICE_ACCOUNT_DEFAULT_PATH
        ),
        service_account_json=_clean_config_value(values.get("GDRIVE_SERVICE_ACCOUNT_JSON")),
        oauth_credentials_path=Path(
            _clean_config_value(values.get("GDRIVE_OAUTH_CREDENTIALS_PATH"))
            or GDRIVE_OAUTH_CREDENTIALS_DEFAULT_PATH
        ),
        oauth_token_path=Path(
            _clean_config_value(values.get("GDRIVE_OAUTH_TOKEN_PATH"))
            or GDRIVE_OAUTH_TOKEN_DEFAULT_PATH
        ),
        report_file_name=(
            _clean_config_value(values.get("GDRIVE_REPORT_FILE_NAME")) or default.report_file_name
        ),
        audit_file_name=(
            _clean_config_value(values.get("GDRIVE_AUDIT_FILE_NAME")) or default.audit_file_name
        ),
        upload_audit=_parse_config_bool(values.get("GDRIVE_UPLOAD_AUDIT"), default=default.upload_audit),
        fail_on_upload_error=_parse_config_bool(
            values.get("GDRIVE_FAIL_ON_UPLOAD_ERROR"),
            default=default.fail_on_upload_error,
        ),
        scopes=default.scopes,
    ).resolved()


def _clean_config_value(value: object) -> str:
    return "" if value is None else str(value).strip()


def _parse_config_bool(value: object, *, default: bool) -> bool:
    text = _clean_config_value(value)
    if not text:
        return default
    lowered = text.lower()
    if lowered in {"1", "true", "yes", "y", "on"}:
        return True
    if lowered in {"0", "false", "no", "n", "off"}:
        return False
    return default


def _normalize_auth_mode(value: object) -> str:
    text = _clean_config_value(value).lower().replace("-", "_")
    if text in {"oauth", "oauth_user", "user_oauth"}:
        return "oauth_user"
    return "service_account"
