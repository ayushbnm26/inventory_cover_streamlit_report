"""File discovery for Vendor Central sales and inventory exports."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from inventory_cover.config import SalesInventoryPipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError
from inventory_cover.sales_inventory_schemas import INVENTORY_REPORT_TYPE, SALES_REPORT_TYPE, ReportType


@dataclass(frozen=True)
class SourceFileDiscovery:
    sales_files: list[Path]
    inventory_files: list[Path]
    mapping_files: list[Path]
    warnings: list[str]


def discover_sales_inventory_files(config: SalesInventoryPipelineConfig) -> SourceFileDiscovery:
    """Discover sales and inventory workbooks using the configured source folders."""

    sales_files, sales_warnings = _discover_source_files(
        folder=config.sales_input_dir,
        report_type=SALES_REPORT_TYPE,
        required=config.require_sales,
        allow_multiple=config.allow_multiple_sales_files,
    )
    inventory_files, inventory_warnings = _discover_source_files(
        folder=config.inventory_input_dir,
        report_type=INVENTORY_REPORT_TYPE,
        required=config.require_inventory,
        allow_multiple=config.allow_multiple_inventory_files,
    )
    mapping_files, mapping_warnings = _discover_optional_mapping_files(config.mapping_input_dir)
    return SourceFileDiscovery(
        sales_files=sales_files,
        inventory_files=inventory_files,
        mapping_files=mapping_files,
        warnings=sales_warnings + inventory_warnings + mapping_warnings,
    )


def _discover_source_files(
    folder: Path,
    report_type: ReportType,
    required: bool,
    allow_multiple: bool,
) -> tuple[list[Path], list[str]]:
    warnings: list[str] = []
    folder = Path(folder)
    label = report_type.lower()
    if not folder.exists():
        message = f"{report_type} input folder does not exist: {folder}"
        if required:
            raise CatastrophicPipelineError(message)
        warnings.append(f"{message}; source treated as missing.")
        return [], warnings
    if not folder.is_dir():
        raise CatastrophicPipelineError(f"{report_type} input path is not a folder: {folder}")

    files = sorted(
        path
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")
    )
    if not files:
        message = f"No .xlsx {label} files found in {folder}"
        if required:
            raise CatastrophicPipelineError(message)
        warnings.append(f"{message}; source treated as missing.")
        return [], warnings
    if len(files) > 1 and not allow_multiple:
        raise CatastrophicPipelineError(
            f"Found {len(files)} {label} Excel files in {folder}. "
            f"Use --allow-multiple-{label}-files only after confirming this is intentional."
        )
    if len(files) > 1:
        warnings.append(
            f"Found {len(files)} {label} files; processing all because multiple-file mode is enabled."
        )
    return files, warnings


def _discover_optional_mapping_files(folder: Path) -> tuple[list[Path], list[str]]:
    folder = Path(folder)
    if not folder.exists():
        return [], []
    if not folder.is_dir():
        return [], [f"Mapping input path is not a folder: {folder}; mapping enrichment disabled."]

    files = sorted(
        path
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")
    )
    if not files:
        return [], []
    if len(files) > 1:
        return [], [
            f"Found {len(files)} mapping workbooks in {folder}; mapping enrichment disabled to avoid ambiguous reference data."
        ]
    return files, []
