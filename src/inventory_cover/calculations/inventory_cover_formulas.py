"""Excel formula strings for the Inventory Cover report.

These functions own the *exact* Excel formula text that is engraved into the
team workbook cells. They use Excel structured table references (``[@[Column]]``)
so the formulas survive column reordering and remain readable for the team.

Every formula is divide-by-zero safe and returns the ``"No Sales"`` sentinel
where the Daily Run Rate is zero, so the workbook never shows ``#DIV/0!``,
``#VALUE!`` or ``#N/A``.
"""

from __future__ import annotations

from inventory_cover.inventory_cover_schemas import NO_SALES_ALERT, NO_SALES_LABEL


def _ref(column: str) -> str:
    return f"[@[{column}]]"


def sales_days_formula(window_days: int) -> str:
    start = _ref("Sales Period Start")
    end = _ref("Sales Period End")
    units = _ref("Sales Units")
    return (
        f"=IF(AND({start}<>\"\",{end}<>\"\"),"
        f"MIN({window_days},{end}-{start}+1),"
        f"IF(N({units})>0,{window_days},0))"
    )


def daily_run_rate_formula() -> str:
    days = _ref("Sales Days")
    units = _ref("Sales Units")
    return f"=IFERROR(IF({days}>0,N({units})/{days},0),0)"


def target_doh_formula(default_target: float) -> str:
    aligned = _ref("Aligned DOH Target")
    return f"=IF(N({aligned})>0,{aligned},{_num(default_target)})"


def _doh_formula(numerator_columns: tuple[str, ...]) -> str:
    drr = _ref("Daily Run Rate")
    numerator = "+".join(_ref(col) for col in numerator_columns)
    return (
        f"=IFERROR(IF({drr}>0,({numerator})/{drr},\"{NO_SALES_LABEL}\"),\"{NO_SALES_LABEL}\")"
    )


def current_stock_doh_formula() -> str:
    return _doh_formula(("Sellable On Hand Units",))


def stock_amazon_doh_formula() -> str:
    return _doh_formula(("Sellable On Hand Units", "Amazon In-Transit Units"))


def stock_own_doh_formula() -> str:
    return _doh_formula(("Sellable On Hand Units", "Own In-Transit Units"))


def total_transit_doh_formula() -> str:
    return _doh_formula(("Sellable On Hand Units", "Amazon In-Transit Units", "Own In-Transit Units"))


def doc_including_open_po_formula() -> str:
    return _doh_formula(("Sellable On Hand Units", "Open PO Quantity", "Own In-Transit Units"))


def total_supply_cover_doh_formula() -> str:
    return _doh_formula(
        ("Sellable On Hand Units", "Open PO Quantity", "Amazon In-Transit Units", "Own In-Transit Units")
    )


def gap_to_target_units_formula() -> str:
    target = _ref("Target DOH")
    drr = _ref("Daily Run Rate")
    supply = "+".join(
        _ref(col)
        for col in ("Sellable On Hand Units", "Amazon In-Transit Units", "Own In-Transit Units", "Open PO Quantity")
    )
    return f"=IFERROR(MAX({target}*{drr}-({supply}),0),0)"


def cover_bucket_formula() -> str:
    cover = _ref("Total Supply Cover DOH")
    return (
        f"=IF({cover}=\"{NO_SALES_LABEL}\",\"{NO_SALES_LABEL}\","
        f"IF({cover}<5,\"Critical\","
        f"IF({cover}<15,\"High Risk\","
        f"IF({cover}<25,\"Watch\","
        f"IF({cover}<=30,\"Near Target\",\"Healthy\")))))"
    )


def cover_alert_formula() -> str:
    bucket = _ref("Cover Bucket")
    return (
        f"=IF({bucket}=\"{NO_SALES_LABEL}\",\"{NO_SALES_ALERT}\","
        f"IF({bucket}=\"Critical\",\"Immediate action\","
        f"IF({bucket}=\"High Risk\",\"Urgent replenishment\","
        f"IF({bucket}=\"Watch\",\"Plan replenishment\","
        f"IF({bucket}=\"Near Target\",\"Monitor\",\"No immediate action\")))))"
    )


def formula_for(column: str, window_days: int, default_target: float) -> str | None:
    """Return the Excel formula for a column, or None if it is a value column."""

    builders = {
        "Sales Days": lambda: sales_days_formula(window_days),
        "Daily Run Rate": daily_run_rate_formula,
        "Target DOH": lambda: target_doh_formula(default_target),
        "Current Stock DOH": current_stock_doh_formula,
        "Stock + Amazon Transit DOH": stock_amazon_doh_formula,
        "Stock + Own Transit DOH": stock_own_doh_formula,
        "Total Transit DOH": total_transit_doh_formula,
        "DOC Including Open PO": doc_including_open_po_formula,
        "Total Supply Cover DOH": total_supply_cover_doh_formula,
        "Gap to Target Units": gap_to_target_units_formula,
        "Cover Bucket": cover_bucket_formula,
        "Cover Alert": cover_alert_formula,
    }
    builder = builders.get(column)
    return builder() if builder else None


def _num(value: float) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)
