"""Python mirror of the Excel cover formulas.

The team workbook carries live Excel formulas, but the engine also computes the
same values in Python so it can:

* split products into annexure sheets by bucket,
* populate the backend Calculation_Audit sheet,
* derive the Data Quality Flag,
* let tests assert numeric behaviour without an Excel engine.

The Python logic here is kept in lock-step with ``inventory_cover_formulas``.
"""

from __future__ import annotations

from typing import Any

from inventory_cover.calculations import inventory_cover_formulas as F
from inventory_cover.inventory_cover_schemas import (
    COVER_ALERTS,
    NO_SALES_ALERT,
    NO_SALES_LABEL,
    ProductCoverRow,
)


def compute_sales_days(product: ProductCoverRow, window_days: int) -> int:
    if product.sales_period_start is not None and product.sales_period_end is not None:
        actual = (product.sales_period_end - product.sales_period_start).days + 1
        return max(0, min(window_days, actual))
    if (product.sales_units or 0) > 0:
        return window_days
    return 0


def classify_bucket(total_supply_cover: Any) -> str:
    if not isinstance(total_supply_cover, (int, float)):
        return NO_SALES_LABEL
    value = float(total_supply_cover)
    if value < 5:
        return "Critical"
    if value < 15:
        return "High Risk"
    if value < 25:
        return "Watch"
    if value <= 30:
        return "Near Target"
    return "Healthy"


def compute_product(product: ProductCoverRow, config: Any) -> None:
    """Populate all computed fields on the product row in place."""

    window = int(config.sales_window_days)

    product.sales_days_raw = product.sales_days_raw if product.sales_days_raw is not None else None
    product.sales_days = compute_sales_days(product, window)

    drr = (product.sales_units / product.sales_days) if product.sales_days > 0 else 0.0
    product.daily_run_rate = drr

    on_hand = product.on_hand_units
    amz = product.amazon_transit_units
    own = product.own_transit_units
    po = product.open_po_units

    def doh(numerator: float) -> Any:
        return (numerator / drr) if drr > 0 else NO_SALES_LABEL

    product.current_stock_doh = doh(on_hand)
    product.stock_amazon_doh = doh(on_hand + amz)
    product.stock_own_doh = doh(on_hand + own)
    product.total_transit_doh = doh(on_hand + amz + own)
    product.doc_including_open_po = doh(on_hand + po + own)
    product.total_supply_cover_doh = doh(on_hand + po + amz + own)

    product.total_supply_units = on_hand + amz + own + po
    product.gap_to_target_units = max(product.target_doh * drr - product.total_supply_units, 0.0)

    product.cover_bucket = classify_bucket(product.total_supply_cover_doh)
    product.cover_alert = COVER_ALERTS.get(product.cover_bucket, NO_SALES_ALERT)

    product.data_quality_flag = _data_quality_flag(product)


def _data_quality_flag(product: ProductCoverRow) -> str:
    flags: list[str] = []
    if product.daily_run_rate <= 0:
        flags.append(NO_SALES_LABEL)
    if not product.asin:
        flags.append("Missing ASIN")
    from inventory_cover.inventory_cover_schemas import SOURCE_INVENTORY, SOURCE_SALES

    if SOURCE_INVENTORY not in product.source_presence:
        flags.append("Missing Inventory")
    if SOURCE_SALES not in product.source_presence:
        flags.append("Missing Sales")
    if product.aligned_doh_target is None:
        flags.append("Missing Target DOH")
    if product.identifier_conflict_notes:
        flags.append("Identifier Conflict")
    if product.calculation_warning_notes:
        flags.append("Calculation Warning")
    if not flags:
        return "OK"
    # De-duplicate while preserving order.
    seen: list[str] = []
    for flag in flags:
        if flag not in seen:
            seen.append(flag)
    return "; ".join(seen)


def calculation_audit_row(product: ProductCoverRow, run_id: str, config: Any) -> list[Any]:
    return [
        run_id,
        product.product_key,
        product.asin,
        product.model_number,
        product.sales_units_raw,
        product.sales_days_raw,
        product.sales_units,
        product.sales_days,
        F.daily_run_rate_formula(),
        product.on_hand_units,
        product.amazon_transit_units,
        product.own_transit_units,
        product.open_po_units,
        product.target_doh,
        product.total_supply_units,
        F.total_supply_cover_doh_formula(),
        F.cover_bucket_formula(),
        "; ".join(product.calculation_warning_notes),
    ]
