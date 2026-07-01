"""Read-only Streamlit dashboard for the latest Inventory Cover workbook."""

from __future__ import annotations

from dataclasses import dataclass
import hmac
import io
from pathlib import Path
import sys
import time
from typing import Any, Mapping

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import pandas as pd
import streamlit as st

from inventory_cover.config import GoogleDriveReportConfig
from inventory_cover.inventory_cover_schemas import COVER_BUCKETS
from inventory_cover.io.google_drive_report_store import (
    GoogleDriveFileMetadata,
    GoogleDriveReportStore,
    GoogleDriveReportStoreError,
)


APP_TITLE = "Inventory Cover Dashboard"
DEFAULT_REPORT_FILE_NAME = "Inventory_Cover_Report_latest.xlsx"
MAIN_SHEET_CANDIDATES = ("Inventory_Cover_Report", "Inventory_Cover_Master")
PRIORITY_ORDER = {
    "Critical": 0,
    "High Risk": 1,
    "Watch": 2,
    "Near Target": 3,
    "Healthy": 4,
    "No Sales": 5,
}
KPI_COLUMNS = (
    ("Total SKUs", None, "count"),
    ("Critical SKUs", "Critical", "bucket"),
    ("High Risk SKUs", "High Risk", "bucket"),
    ("Watch SKUs", "Watch", "bucket"),
    ("Near Target SKUs", "Near Target", "bucket"),
    ("Healthy SKUs", "Healthy", "bucket"),
    ("No Sales SKUs", "No Sales", "bucket"),
    ("Total Gap to Target Units", "Gap to Target Units", "sum"),
    ("Total Open PO Units", "Open PO Quantity", "sum"),
    ("Total Sellable On Hand Units", "Sellable On Hand Units", "sum"),
    ("Total Amazon In-Transit Units", "Amazon In-Transit Units", "sum"),
    ("Total Own In-Transit Units", "Own In-Transit Units", "sum"),
)
PRIORITY_COLUMNS = (
    "ASIN",
    "Model Number",
    "Model Number / SKU",
    "SKU",
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
    "Team Remarks",
    "Remarks",
)
SEARCH_COLUMNS = ("ASIN", "SKU", "Model Number", "Model Number / SKU", "Product Title")
RISK_BUCKETS = ("Critical", "High Risk", "Watch")


class DashboardConfigError(ValueError):
    """Raised for missing Streamlit dashboard configuration."""


class DashboardDataError(ValueError):
    """Raised for unreadable or structurally invalid workbook data."""


@dataclass(frozen=True)
class DashboardSettings:
    folder_id: str
    report_file_name: str
    service_account_info: dict[str, Any]


@dataclass(frozen=True)
class DashboardWorkbook:
    dataframe: pd.DataFrame
    sheet_name: str
    sheet_names: tuple[str, ...]


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    _inject_style()
    _require_password()

    st.title(APP_TITLE)
    st.caption(
        "The Excel workbook remains the official working report; this dashboard is a read-only visual summary."
    )

    if "refresh_nonce" not in st.session_state:
        st.session_state.refresh_nonce = 0

    try:
        settings = _read_settings_from_secrets(st.secrets)
    except DashboardConfigError as exc:
        st.error(str(exc))
        st.stop()

    header_left, header_right = st.columns([4, 1])
    with header_left:
        st.write(f"Report source filename: `{settings.report_file_name}`")
    with header_right:
        if st.button("Refresh", use_container_width=True):
            st.session_state.refresh_nonce = int(time.time())
            _download_report_from_drive.clear()

    try:
        report_bytes, metadata = _download_report_from_drive(
            settings.folder_id,
            settings.report_file_name,
            st.session_state.refresh_nonce,
        )
        workbook = load_inventory_report_from_bytes(report_bytes)
    except (DashboardConfigError, DashboardDataError, GoogleDriveReportStoreError) as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001 - dashboard should show a clean top-level failure.
        st.error(f"The dashboard could not load the latest report: {exc}")
        st.stop()

    _render_freshness(metadata, workbook)
    df = workbook.dataframe
    missing = _missing_expected_fields(df)
    if missing:
        st.warning("Some expected report fields were not found: " + ", ".join(missing))

    _render_kpis(df)
    _render_bucket_summary(df)
    _render_priority_table(df)
    _render_full_table(df)
    _render_charts(df)
    _render_no_sales(df)
    _render_data_quality(df)
    _render_download(report_bytes, metadata, settings.report_file_name)


@st.cache_data(ttl=300, show_spinner=False)
def _download_report_from_drive(
    folder_id: str,
    report_file_name: str,
    refresh_nonce: int,
) -> tuple[bytes, GoogleDriveFileMetadata]:
    del refresh_nonce
    settings = _read_settings_from_secrets(st.secrets)
    store = GoogleDriveReportStore.from_service_account_info(
        settings.service_account_info,
        scopes=GoogleDriveReportConfig().scopes,
    )
    return store.download_file_by_name(folder_id, report_file_name)


def load_inventory_report_from_bytes(workbook_bytes: bytes) -> DashboardWorkbook:
    """Read the official workbook bytes and pick the most likely report sheet."""

    if not workbook_bytes:
        raise DashboardDataError("The downloaded workbook was empty.")
    try:
        excel = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="openpyxl")
    except Exception as exc:
        raise DashboardDataError(f"The downloaded file is not a readable Excel workbook: {exc}") from exc

    sheet_names = tuple(str(name) for name in excel.sheet_names)
    candidates = [name for name in MAIN_SHEET_CANDIDATES if name in sheet_names]
    candidates.extend(name for name in sheet_names if name not in candidates)

    best_df: pd.DataFrame | None = None
    best_sheet = ""
    for sheet_name in candidates:
        try:
            frame = pd.read_excel(excel, sheet_name=sheet_name)
        except Exception:
            continue
        frame = _clean_dataframe(frame)
        if _looks_like_inventory_cover_sheet(frame):
            return DashboardWorkbook(dataframe=frame, sheet_name=sheet_name, sheet_names=sheet_names)
        if best_df is None and not frame.empty:
            best_df = frame
            best_sheet = sheet_name
    if best_df is not None:
        return DashboardWorkbook(dataframe=best_df, sheet_name=best_sheet, sheet_names=sheet_names)
    raise DashboardDataError("No readable data sheet was found in the downloaded workbook.")


def bucket_distribution(df: pd.DataFrame) -> pd.DataFrame:
    if "Cover Bucket" not in df.columns:
        return pd.DataFrame(columns=["Cover Bucket", "SKU Count"])
    counts = df["Cover Bucket"].fillna("Not available").astype(str).value_counts()
    order = {bucket: index for index, bucket in enumerate(COVER_BUCKETS)}
    summary = counts.rename_axis("Cover Bucket").reset_index(name="SKU Count")
    summary["_order"] = summary["Cover Bucket"].map(order).fillna(999)
    return summary.sort_values(["_order", "Cover Bucket"]).drop(columns="_order").reset_index(drop=True)


def priority_action_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.copy()
    if "Cover Bucket" in frame.columns:
        frame["_priority"] = frame["Cover Bucket"].map(PRIORITY_ORDER).fillna(999)
    else:
        frame["_priority"] = 999
    frame["_cover"] = _numeric_series(frame, "Total Supply Cover DOH")
    frame["_gap"] = _numeric_series(frame, "Gap to Target Units")
    frame["_drr"] = _numeric_series(frame, "Daily Run Rate")
    sorted_frame = frame.sort_values(
        by=["_priority", "_cover", "_gap", "_drr"],
        ascending=[True, True, False, False],
        na_position="last",
    )
    columns = _available_columns(sorted_frame, PRIORITY_COLUMNS)
    return sorted_frame[columns].reset_index(drop=True)


def filter_inventory_dataframe(
    df: pd.DataFrame,
    *,
    search_text: str = "",
    buckets: tuple[str, ...] = (),
    risk_only: bool = False,
) -> pd.DataFrame:
    frame = df.copy()
    if buckets and "Cover Bucket" in frame.columns:
        frame = frame[frame["Cover Bucket"].astype(str).isin(set(buckets))]
    if risk_only and "Cover Bucket" in frame.columns:
        frame = frame[frame["Cover Bucket"].astype(str).isin(RISK_BUCKETS)]
    search = search_text.strip().lower()
    if search:
        columns = _available_columns(frame, SEARCH_COLUMNS)
        if columns:
            haystack = frame[columns].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            frame = frame[haystack.str.contains(search, regex=False)]
    return frame.reset_index(drop=True)


def _read_settings_from_secrets(secrets: Mapping[str, Any]) -> DashboardSettings:
    folder_id = _secret_text(secrets, "GDRIVE_FOLDER_ID")
    report_file_name = _secret_text(secrets, "GDRIVE_REPORT_FILE_NAME") or DEFAULT_REPORT_FILE_NAME
    service_account_info = _secret_mapping(secrets, "gdrive_service_account")
    if not folder_id:
        raise DashboardConfigError("GDRIVE_FOLDER_ID is missing from Streamlit secrets.")
    if not service_account_info:
        raise DashboardConfigError("gdrive_service_account is missing from Streamlit secrets.")
    return DashboardSettings(
        folder_id=folder_id,
        report_file_name=report_file_name,
        service_account_info=service_account_info,
    )


def _require_password() -> None:
    password = _secret_text(st.secrets, "APP_PASSWORD")
    if not password:
        st.error("Dashboard password is not configured.")
        st.stop()
    if st.session_state.get("authenticated") is True:
        return
    with st.form("dashboard_login"):
        entered = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")
    if not submitted:
        st.stop()
    if hmac.compare_digest(entered, password):
        st.session_state.authenticated = True
        st.rerun()
    st.error("Invalid password.")
    st.stop()


def _render_freshness(metadata: GoogleDriveFileMetadata, workbook: DashboardWorkbook) -> None:
    cols = st.columns(4)
    cols[0].metric("Drive file", metadata.name or "Not available")
    cols[1].metric("Drive modified", metadata.modified_time or "Not available")
    cols[2].metric("Workbook sheet", workbook.sheet_name)
    cols[3].metric("Rows", f"{len(workbook.dataframe):,}")


def _render_kpis(df: pd.DataFrame) -> None:
    st.subheader("KPI Summary")
    metrics = _kpi_values(df)
    for index in range(0, len(metrics), 4):
        cols = st.columns(4)
        for col, (label, value) in zip(cols, metrics[index : index + 4]):
            col.metric(label, value)


def _render_bucket_summary(df: pd.DataFrame) -> None:
    st.subheader("Risk Bucket Summary")
    summary = bucket_distribution(df)
    if summary.empty:
        st.info("Cover Bucket is not available in this workbook.")
        return
    table_col, chart_col = st.columns([1, 2])
    table_col.dataframe(summary, hide_index=True, use_container_width=True)
    chart_col.bar_chart(summary.set_index("Cover Bucket")["SKU Count"])


def _render_priority_table(df: pd.DataFrame) -> None:
    st.subheader("Priority Action Table")
    priority = priority_action_dataframe(df)
    if priority.empty:
        st.info("No inventory cover rows are available.")
        return
    st.dataframe(priority.head(50), hide_index=True, use_container_width=True)


def _render_full_table(df: pd.DataFrame) -> None:
    st.subheader("Full Inventory Cover Table")
    search_col, bucket_col, risk_col = st.columns([2, 2, 1])
    search_text = search_col.text_input("Search", placeholder="ASIN, SKU, model, or title")
    bucket_options = tuple(bucket_distribution(df)["Cover Bucket"]) if "Cover Bucket" in df.columns else ()
    selected_buckets = bucket_col.multiselect("Cover Bucket", bucket_options, default=list(bucket_options))
    risk_only = risk_col.checkbox("Risk items only")
    filtered = filter_inventory_dataframe(
        df,
        search_text=search_text,
        buckets=tuple(selected_buckets),
        risk_only=risk_only,
    )
    st.dataframe(filtered, hide_index=True, use_container_width=True)
    st.download_button(
        "Download filtered CSV",
        data=filtered.to_csv(index=False).encode("utf-8"),
        file_name="inventory_cover_filtered.csv",
        mime="text/csv",
    )


def _render_charts(df: pd.DataFrame) -> None:
    st.subheader("Charts")
    chart_cols = st.columns(2)
    gap = _top_numeric(df, "Gap to Target Units", largest=True)
    if not gap.empty:
        chart_cols[0].bar_chart(gap)
    else:
        chart_cols[0].info("Gap to Target Units is not available.")

    low_cover = _lowest_cover(df)
    if not low_cover.empty:
        chart_cols[1].bar_chart(low_cover)
    else:
        chart_cols[1].info("Total Supply Cover DOH is not available for selling SKUs.")

    supply = _supply_summary(df)
    if not supply.empty:
        st.bar_chart(supply)


def _render_no_sales(df: pd.DataFrame) -> None:
    if "Cover Bucket" not in df.columns:
        return
    rows = df[df["Cover Bucket"].astype(str) == "No Sales"]
    if rows.empty:
        return
    columns = _available_columns(
        rows,
        (
            "ASIN",
            "Model Number / SKU",
            "SKU",
            "Product Title",
            "Sellable On Hand Units",
            "Amazon In-Transit Units",
            "Own In-Transit Units",
            "Open PO Quantity",
            "Cover Alert",
        ),
    )
    with st.expander(f"No Sales SKUs ({len(rows):,})"):
        st.write("Rows in this section have zero Daily Run Rate in the official report.")
        st.dataframe(rows[columns], hide_index=True, use_container_width=True)


def _render_data_quality(df: pd.DataFrame) -> None:
    st.subheader("Data Quality / Audit")
    quality_columns = _available_columns(df, ("Data Quality Flag", "Cover Alert"))
    if not quality_columns:
        st.info("No validation or audit fields are available in the selected workbook sheet.")
        return
    if "Data Quality Flag" in df.columns:
        flags = df["Data Quality Flag"].fillna("OK").astype(str)
        warning_rows = df[flags.str.upper() != "OK"]
        st.metric("Rows with data quality flags", f"{len(warning_rows):,}")
        if not warning_rows.empty:
            with st.expander("Validation details"):
                st.dataframe(warning_rows, hide_index=True, use_container_width=True)
    else:
        st.metric("Rows with cover alerts", f"{df['Cover Alert'].notna().sum():,}")


def _render_download(report_bytes: bytes, metadata: GoogleDriveFileMetadata, fallback_name: str) -> None:
    st.subheader("Download")
    st.download_button(
        "Download official Excel workbook",
        data=report_bytes,
        file_name=metadata.name or fallback_name,
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.dropna(how="all").copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    frame = frame.loc[:, [column for column in frame.columns if column and not column.startswith("Unnamed:")]]
    return frame


def _looks_like_inventory_cover_sheet(df: pd.DataFrame) -> bool:
    if df.empty:
        return False
    columns = set(df.columns)
    return "Cover Bucket" in columns or {"ASIN", "Total Supply Cover DOH"}.issubset(columns)


def _missing_expected_fields(df: pd.DataFrame) -> tuple[str, ...]:
    expected = (
        "Cover Bucket",
        "Total Supply Cover DOH",
        "Gap to Target Units",
        "Daily Run Rate",
        "Sellable On Hand Units",
        "Open PO Quantity",
    )
    return tuple(column for column in expected if column not in df.columns)


def _kpi_values(df: pd.DataFrame) -> list[tuple[str, str]]:
    values: list[tuple[str, str]] = []
    for label, column, mode in KPI_COLUMNS:
        if mode == "count":
            values.append((label, f"{len(df):,}"))
        elif mode == "bucket" and "Cover Bucket" in df.columns:
            count = int((df["Cover Bucket"].astype(str) == str(column)).sum())
            values.append((label, f"{count:,}"))
        elif mode == "sum" and column in df.columns:
            total = _numeric_series(df, str(column)).sum(skipna=True)
            values.append((label, _format_number(total)))
    return values


def _top_numeric(df: pd.DataFrame, column: str, *, largest: bool) -> pd.Series:
    if column not in df.columns:
        return pd.Series(dtype="float64")
    frame = df.copy()
    frame[column] = _numeric_series(frame, column)
    frame = frame.dropna(subset=[column])
    if frame.empty:
        return pd.Series(dtype="float64")
    label = _label_series(frame)
    ordered = frame.assign(_label=label).sort_values(column, ascending=not largest).head(10)
    return ordered.set_index("_label")[column]


def _lowest_cover(df: pd.DataFrame) -> pd.Series:
    if "Total Supply Cover DOH" not in df.columns:
        return pd.Series(dtype="float64")
    frame = df.copy()
    frame["_cover"] = _numeric_series(frame, "Total Supply Cover DOH")
    if "Daily Run Rate" in frame.columns:
        frame["_drr"] = _numeric_series(frame, "Daily Run Rate")
        frame = frame[frame["_drr"] > 0]
    frame = frame.dropna(subset=["_cover"]).sort_values("_cover").head(10)
    if frame.empty:
        return pd.Series(dtype="float64")
    return frame.set_index(_label_series(frame))["_cover"]


def _supply_summary(df: pd.DataFrame) -> pd.Series:
    columns = (
        "Sellable On Hand Units",
        "Amazon In-Transit Units",
        "Own In-Transit Units",
        "Open PO Quantity",
    )
    totals = {
        column: _numeric_series(df, column).sum(skipna=True)
        for column in columns
        if column in df.columns
    }
    return pd.Series(totals, dtype="float64")


def _numeric_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series([pd.NA] * len(df), index=df.index, dtype="Float64")
    return pd.to_numeric(df[column], errors="coerce")


def _label_series(df: pd.DataFrame) -> pd.Series:
    for column in ("ASIN", "Model Number / SKU", "SKU", "Product Title"):
        if column in df.columns:
            return df[column].fillna("").astype(str).replace("", "Unlabeled")
    return pd.Series([f"Row {index + 1}" for index in range(len(df))], index=df.index)


def _available_columns(df: pd.DataFrame, columns: tuple[str, ...]) -> list[str]:
    return [column for column in columns if column in df.columns]


def _format_number(value: float) -> str:
    if pd.isna(value):
        return "Not available"
    if abs(float(value) - round(float(value))) < 0.005:
        return f"{int(round(float(value))):,}"
    return f"{float(value):,.2f}"


def _secret_text(secrets: Mapping[str, Any], key: str) -> str:
    try:
        value = secrets.get(key, "")
    except Exception:
        return ""
    return "" if value is None else str(value).strip()


def _secret_mapping(secrets: Mapping[str, Any], key: str) -> dict[str, Any]:
    try:
        value = secrets.get(key, {})
    except Exception:
        return {}
    if value is None:
        return {}
    return dict(value)


def _inject_style() -> None:
    st.markdown(
        """
        <style>
        .block-container {padding-top: 2rem; padding-bottom: 3rem;}
        [data-testid="stMetric"] {
            background: #f8fafc;
            border: 1px solid #e5e7eb;
            padding: 0.75rem 0.85rem;
            border-radius: 6px;
        }
        h1, h2, h3 {letter-spacing: 0;}
        </style>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
