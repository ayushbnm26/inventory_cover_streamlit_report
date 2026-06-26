"""Normalize Vendor Central sales and inventory rows into backend schemas."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from inventory_cover.sales_inventory_schemas import (
    INVENTORY_MASTER_HEADERS,
    INVENTORY_NUMERIC_COLUMNS,
    INVENTORY_REPORT_TYPE,
    INVENTORY_ZERO_DEFAULT_NUMERIC_COLUMNS,
    INVENTORY_SOURCE_COLUMNS,
    SALES_MASTER_HEADERS,
    SALES_NUMERIC_COLUMNS,
    SALES_REPORT_TYPE,
    SALES_SOURCE_COLUMNS,
    TEXT_COLUMNS,
    NormalizedSalesInventoryRow,
    RawSalesInventoryRow,
    ReportType,
    SalesInventoryCellValue,
    SalesInventoryValidationIssue,
)
from inventory_cover.io.sales_inventory_mapping_io import SalesInventoryMappingLookup
from inventory_cover.utils.date_parsing import parse_vendor_central_date
from inventory_cover.utils.numbers import parse_number
from inventory_cover.utils.text_cleaning import clean_text


@dataclass(frozen=True)
class SalesInventoryRowProcessingResult:
    normalized_row: NormalizedSalesInventoryRow | None
    issues: list[SalesInventoryValidationIssue]
    rejected: bool


def normalize_sales_inventory_row(
    raw_row: RawSalesInventoryRow,
    run_id: str,
    mapping_lookup: SalesInventoryMappingLookup | None = None,
) -> SalesInventoryRowProcessingResult:
    """Normalize one Vendor Central sales or inventory source row."""

    headers = SALES_MASTER_HEADERS if raw_row.report_type == SALES_REPORT_TYPE else INVENTORY_MASTER_HEADERS
    source_columns = SALES_SOURCE_COLUMNS if raw_row.report_type == SALES_REPORT_TYPE else INVENTORY_SOURCE_COLUMNS
    numeric_columns = SALES_NUMERIC_COLUMNS if raw_row.report_type == SALES_REPORT_TYPE else INVENTORY_NUMERIC_COLUMNS
    data: dict[str, Any] = {header: None for header in headers}
    data.update(
        {
            "Run ID": run_id,
            "Source File": raw_row.source_file,
            "Source Sheet": raw_row.source_sheet,
            "Source Row": raw_row.source_row,
            "Report Type": raw_row.report_type,
            "Programme": raw_row.metadata.get("Programme"),
            "Distributor View": raw_row.metadata.get("Distributor View"),
            "View By": raw_row.metadata.get("View By"),
            "Country": raw_row.metadata.get("Country") or raw_row.metadata.get("Countries"),
            "Currency": raw_row.metadata.get("Currency"),
            "Viewing Range Start": raw_row.metadata.get("Viewing Range Start"),
            "Viewing Range End": raw_row.metadata.get("Viewing Range End"),
            "Report Updated Date": raw_row.metadata.get("Report Updated Date"),
        }
    )
    issues: list[SalesInventoryValidationIssue] = []

    for field in source_columns:
        cell = raw_row.values.get(field, SalesInventoryCellValue(None))
        if field in TEXT_COLUMNS:
            data[field] = clean_text(cell.value, cell.number_format)
        elif field == "Release Date":
            parsed = parse_vendor_central_date(cell.value)
            if parsed.was_blank:
                data[field] = None
            elif parsed.ok:
                data[field] = parsed.value
            else:
                data[field] = parsed.value
                issues.append(
                    _issue(
                        "WARNING",
                        "INVALID_RELEASE_DATE",
                        raw_row,
                        run_id,
                        field,
                        cell.value,
                        parsed.detail,
                        "Row kept; Release Date left as source text.",
                    )
                )
        elif field in numeric_columns:
            parsed_number = parse_number(cell.value)
            if (
                raw_row.report_type == INVENTORY_REPORT_TYPE
                and field in INVENTORY_ZERO_DEFAULT_NUMERIC_COLUMNS
                and parsed_number.was_blank
            ):
                data[field] = 0
                continue
            if parsed_number.ok:
                data[field] = parsed_number.value
                if parsed_number.value is not None and float(parsed_number.value) < 0:
                    issues.append(
                        _issue(
                            "WARNING",
                            f"NEGATIVE_{_issue_suffix(field)}",
                            raw_row,
                            run_id,
                            field,
                            cell.value,
                            f"{field} is negative.",
                            "Row kept; value preserved.",
                        )
                    )
            else:
                data[field] = None
                issues.append(
                    _issue(
                        "WARNING",
                        f"INVALID_{_issue_suffix(field)}",
                        raw_row,
                        run_id,
                        field,
                        cell.value,
                        parsed_number.detail,
                        "Row kept; blank written for parsed numeric value.",
                    )
                )

    _apply_mapping_enrichment(data, raw_row, run_id, issues, mapping_lookup)
    _attach_identifier_issues(data, raw_row, run_id, issues)
    _attach_inventory_quantity_warnings(data, raw_row, run_id, issues)

    if _has_error(issues):
        return SalesInventoryRowProcessingResult(
            normalized_row=None,
            issues=issues,
            rejected=True,
        )

    row = NormalizedSalesInventoryRow(report_type=raw_row.report_type, data=data, issues=issues)
    row.refresh_validation_status()
    return SalesInventoryRowProcessingResult(
        normalized_row=row,
        issues=issues,
        rejected=False,
    )


def _apply_mapping_enrichment(
    data: dict[str, Any],
    raw_row: RawSalesInventoryRow,
    run_id: str,
    issues: list[SalesInventoryValidationIssue],
    mapping_lookup: SalesInventoryMappingLookup | None,
) -> None:
    if mapping_lookup is None:
        return

    asin = str(data.get("ASIN") or "").strip()
    model_number = str(data.get("Model Number") or "").strip()

    if asin and not model_number:
        match = mapping_lookup.model_for_asin(asin)
        if match.found and match.value:
            data["Model Number"] = match.value
            issues.append(
                _issue(
                    "INFO",
                    "MODEL_NUMBER_FILLED_FROM_MAPPING",
                    raw_row,
                    run_id,
                    "Model Number",
                    _raw(raw_row, "Model Number"),
                    f"Model Number was blank and ASIN {asin} matched the mapping workbook.",
                    "Model Number filled from mapping workbook.",
                )
            )
        elif match.status == "AMBIGUOUS":
            issues.append(
                _issue(
                    "WARNING",
                    "MODEL_NUMBER_MAPPING_AMBIGUOUS",
                    raw_row,
                    run_id,
                    "Model Number",
                    _raw(raw_row, "Model Number"),
                    f"ASIN {asin} has multiple mapped SKU values.",
                    "Row kept; Model Number was not guessed.",
                )
            )

    asin = str(data.get("ASIN") or "").strip()
    model_number = str(data.get("Model Number") or "").strip()
    if model_number and not asin:
        match = mapping_lookup.asin_for_model(model_number)
        if match.found and match.value:
            data["ASIN"] = match.value
            issues.append(
                _issue(
                    "INFO",
                    "ASIN_FILLED_FROM_MAPPING",
                    raw_row,
                    run_id,
                    "ASIN",
                    _raw(raw_row, "ASIN"),
                    f"ASIN was blank and Model Number {model_number} matched the mapping workbook.",
                    "ASIN filled from mapping workbook.",
                )
            )
        elif match.status == "AMBIGUOUS":
            issues.append(
                _issue(
                    "WARNING",
                    "ASIN_MAPPING_AMBIGUOUS",
                    raw_row,
                    run_id,
                    "ASIN",
                    _raw(raw_row, "ASIN"),
                    f"Model Number {model_number} has multiple mapped ASIN values.",
                    "Row kept; ASIN was not guessed.",
                )
            )


def _attach_identifier_issues(
    data: dict[str, Any],
    raw_row: RawSalesInventoryRow,
    run_id: str,
    issues: list[SalesInventoryValidationIssue],
) -> None:
    asin = str(data.get("ASIN") or "").strip()
    model_number = str(data.get("Model Number") or "").strip()
    product_title = str(data.get("Product Title") or "").strip()
    child_vendor_code = str(data.get("Child Vendor Code") or "").strip()

    if asin:
        if not child_vendor_code:
            issues.append(
                _issue(
                    "WARNING",
                    "MISSING_CHILD_VENDOR_CODE",
                    raw_row,
                    run_id,
                    "Child Vendor Code",
                    _raw(raw_row, "Child Vendor Code"),
                    "Child Vendor Code is blank.",
                    "Row kept with warning.",
                )
            )
        if not model_number:
            issues.append(
                _issue(
                    "WARNING",
                    "MISSING_MODEL_NUMBER",
                    raw_row,
                    run_id,
                    "Model Number",
                    _raw(raw_row, "Model Number"),
                    "Model Number is blank; ASIN is present.",
                    "Row kept with warning.",
                )
            )
        return

    if model_number and product_title:
        issues.append(
            _issue(
                "WARNING",
                "MISSING_ASIN_IDENTIFIER",
                raw_row,
                run_id,
                "ASIN",
                _raw(raw_row, "ASIN"),
                "ASIN is blank, but Model Number and Product Title are present.",
                "Row kept with warning.",
            )
        )
        return

    issues.append(
        _issue(
            "ERROR",
            "UNUSABLE_ROW",
            raw_row,
            run_id,
            "ASIN",
            _raw(raw_row, "ASIN"),
            "ASIN, Model Number, and Product Title are all blank.",
            "Row rejected.",
        )
    )


def _attach_inventory_quantity_warnings(
    data: dict[str, Any],
    raw_row: RawSalesInventoryRow,
    run_id: str,
    issues: list[SalesInventoryValidationIssue],
) -> None:
    if raw_row.report_type != INVENTORY_REPORT_TYPE:
        return
    for field in ("Sellable On Hand Units", "In Transit Quantity"):
        if data.get(field) is None:
            issues.append(
                _issue(
                    "WARNING",
                    f"MISSING_{_issue_suffix(field)}",
                    raw_row,
                    run_id,
                    field,
                    _raw(raw_row, field),
                    f"{field} is blank.",
                    "Row kept with warning.",
                )
            )


def _issue(
    severity: str,
    issue_type: str,
    raw_row: RawSalesInventoryRow,
    run_id: str,
    field_name: str,
    raw_value: Any,
    detail: str,
    action: str,
) -> SalesInventoryValidationIssue:
    return SalesInventoryValidationIssue(
        run_id=run_id,
        report_type=raw_row.report_type,
        severity=severity,
        issue_type=issue_type,
        source_file=raw_row.source_file,
        source_sheet=raw_row.source_sheet,
        source_row=raw_row.source_row,
        asin=_text(raw_row, "ASIN"),
        child_vendor_code=_text(raw_row, "Child Vendor Code"),
        model_number=_text(raw_row, "Model Number"),
        field_name=field_name,
        raw_value=raw_value,
        issue_detail=detail,
        action_taken=action,
    )


def _text(raw_row: RawSalesInventoryRow, field: str) -> str:
    cell = raw_row.values.get(field, SalesInventoryCellValue(None))
    return clean_text(cell.value, cell.number_format)


def _raw(raw_row: RawSalesInventoryRow, field: str) -> Any:
    return raw_row.values.get(field, SalesInventoryCellValue(None)).value


def _has_error(issues: list[SalesInventoryValidationIssue]) -> bool:
    return any(issue.severity.upper() == "ERROR" for issue in issues)


def _issue_suffix(field_name: str) -> str:
    return (
        field_name.upper()
        .replace(" ", "_")
        .replace("-", "_")
        .replace("+", "PLUS")
        .replace("%", "PCT")
        .replace("(", "")
        .replace(")", "")
    )
