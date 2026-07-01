from __future__ import annotations

from io import BytesIO

from openpyxl import Workbook
import pandas as pd

from streamlit_app import (
    bucket_distribution,
    filter_inventory_dataframe,
    load_inventory_report_from_bytes,
    priority_action_dataframe,
)


def test_load_inventory_report_from_bytes_prefers_official_sheet() -> None:
    workbook_bytes = _workbook_bytes()

    result = load_inventory_report_from_bytes(workbook_bytes)

    assert result.sheet_name == "Inventory_Cover_Report"
    assert result.sheet_names == ("Other", "Inventory_Cover_Report")
    assert list(result.dataframe["ASIN"]) == ["A3", "A2", "A1"]
    assert "Cover Bucket" in result.dataframe.columns


def test_bucket_distribution_uses_report_bucket_labels() -> None:
    df = load_inventory_report_from_bytes(_workbook_bytes()).dataframe

    summary = bucket_distribution(df)

    assert list(summary["Cover Bucket"]) == ["Critical", "High Risk", "No Sales"]
    assert list(summary["SKU Count"]) == [1, 1, 1]


def test_priority_action_table_sorts_urgent_low_cover_high_gap_first() -> None:
    df = load_inventory_report_from_bytes(_workbook_bytes()).dataframe

    priority = priority_action_dataframe(df)

    assert list(priority["ASIN"]) == ["A1", "A2", "A3"]


def test_filter_inventory_dataframe_search_bucket_and_risk_only() -> None:
    df = load_inventory_report_from_bytes(_workbook_bytes()).dataframe

    filtered = filter_inventory_dataframe(
        df,
        search_text="sku-a2",
        buckets=("Critical", "High Risk", "No Sales"),
        risk_only=True,
    )

    assert list(filtered["ASIN"]) == ["A2"]


def test_priority_action_table_handles_missing_bucket_column() -> None:
    priority = priority_action_dataframe(
        pd.DataFrame(
            {
                "ASIN": ["A1"],
                "Total Supply Cover DOH": [5],
                "Gap to Target Units": [10],
            }
        )
    )

    assert list(priority["ASIN"]) == ["A1"]


def _workbook_bytes() -> bytes:
    wb = Workbook()
    other = wb.active
    other.title = "Other"
    other.append(["Notes"])
    other.append(["ignore"])
    ws = wb.create_sheet("Inventory_Cover_Report")
    ws.append(
        [
            "ASIN",
            "Model Number / SKU",
            "Product Title",
            "Daily Run Rate",
            "Sales Units",
            "Sellable On Hand Units",
            "Amazon In-Transit Units",
            "Own In-Transit Units",
            "Open PO Quantity",
            "Total Supply Cover DOH",
            "DOC Including Open PO",
            "Gap to Target Units",
            "Target DOH",
            "Cover Bucket",
            "Cover Alert",
            "Remarks",
        ]
    )
    ws.append(["A3", "SKU-A3", "No Sales", 0, 0, 100, 0, 0, 0, "No Sales", "No Sales", 0, 30, "No Sales", "No sales", ""])
    ws.append(["A2", "SKU-A2", "Risk Two", 2, 60, 20, 0, 0, 0, 10, 10, 50, 30, "High Risk", "Urgent", ""])
    ws.append(["A1", "SKU-A1", "Risk One", 5, 150, 10, 0, 0, 0, 2, 2, 100, 30, "Critical", "Immediate", ""])
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
