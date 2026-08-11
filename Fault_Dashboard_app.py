import sys
import os
from pathlib import Path
from typing import Tuple, List, Optional

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Charger Fault Dashboard", layout="wide")

DEFAULT_PATH = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("FAULT_SHEET_PATH", "")
DATE_COLS = ["Issue Date", "Resolution Date", "Restoration Date"]


def safe_dataframe(df_to_show: pd.DataFrame):
    """Safely display dataframe across different Streamlit versions."""
    try:
        st.dataframe(df_to_show, use_container_width=True, hide_index=True)
    except TypeError:
        try:
            st.dataframe(df_to_show, use_container_width=True)
        except Exception:
            st.dataframe(df_to_show)


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Find the first matching column name (case-insensitive, trimmed)."""
    cols_map = {str(c).strip().lower(): str(c) for c in df.columns}
    for cand in candidates:
        cand_clean = cand.strip().lower()
        if cand_clean in cols_map:
            return cols_map[cand_clean]
    return None


@st.cache_data(ttl=300)
def load_data(file) -> pd.DataFrame:
    """
    Load an Excel file into a DataFrame and normalize date fields and derived time columns.
    """
    df = pd.read_excel(file)

    # Clean column headers (strip whitespace and drop duplicate columns)
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.duplicated()]

    # Standardize common column names flexibly
    col_mappings = {
        "Issue Date": ["Issue Date", "Issue date", "Created Date", "Date", "Ticket Date"],
        "Resolution Date": ["Resolution Date", "Resolution date", "Closed Date"],
        "Restoration Date": ["Restoration Date", "Restoration date"],
        "Station ID": ["Station ID", "Station id", "StationID", "Station_ID"],
        "Station Name": ["Station Name", "Station name", "StationName", "Station_Name"],
        "Issue Type": ["Issue Type", "Issue type", "Fault Type", "Category"],
        "Issue Sub-Type": ["Issue Sub-Type", "Issue sub-type", "Sub Category", "Sub Type", "Fault Subtype"],
        "Zone": ["Zone", "Region", "State", "Circle"],
        "Severity": ["Severity", "Priority"],
        "Charger Make": ["Charger Make", "Charger Company", "OEM", "Make", "Vendor"],
        "Status": ["Status", "Ticket Status", "State"],
        "TAT Compliance": ["TAT Compliance", "TAT", "SLA Compliance"]
    }

    for std_name, candidates in col_mappings.items():
        if std_name not in df.columns:
            found = find_col(df, candidates)
            if found and found != std_name:
                df.rename(columns={found: std_name}, inplace=True)

    # Convert date columns to datetime safely
    for col in DATE_COLS:
        if col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
            except Exception:
                df[col] = pd.to_datetime(df[col], errors="coerce")

    # Derive useful time columns from Issue Date for filtering and reporting
    if "Issue Date" in df.columns and df["Issue Date"].notna().any():
        issue_dt = df["Issue Date"]
        df["Issue Day"] = issue_dt.dt.date
        df["Issue Month"] = issue_dt.dt.to_period("M").astype(str)
        df["Issue Month"] = df["Issue Month"].replace("NaT", np.nan)

        quarter_str = issue_dt.dt.year.astype(str) + "-Q" + issue_dt.dt.quarter.astype(str)
        df["Quarter"] = quarter_str.where(issue_dt.notna(), np.nan)

    return df


def get_repetitive_faults(df: pd.DataFrame) -> pd.DataFrame:
    required = ["Station ID", "Station Name", "Issue Sub-Type"]
    group_cols = [c for c in required if c in df.columns]

    if "Issue Sub-Type" not in df.columns or not ("Station ID" in df.columns or "Station Name" in df.columns):
        return pd.DataFrame(columns=required + ["Occurrences"])

    if df.empty:
        return pd.DataFrame(columns=group_cols + ["Occurrences"])

    counts = (
        df.groupby(group_cols)
        .size()
        .reset_index(name="Occurrences")
    )
    return counts[counts["Occurrences"] >= 2].sort_values("Occurrences", ascending=False)


def get_top_issue_breakdown(df: pd.DataFrame, top_types: int = 5, top_subtypes: int = 5) -> pd.DataFrame:
    if "Issue Type" not in df.columns or "Issue Sub-Type" not in df.columns or df.empty:
        return pd.DataFrame()

    type_counts = df["Issue Type"].dropna().value_counts().head(top_types)
    rows = []
    for issue_type, type_total in type_counts.items():
        sub_counts = (
            df[df["Issue Type"] == issue_type]["Issue Sub-Type"]
            .dropna()
            .value_counts()
            .head(top_subtypes)
        )
        for sub_type, sub_count in sub_counts.items():
            rows.append({
                "Issue Type": issue_type,
                "Issue Type Total": type_total,
                "Issue Sub-Type": sub_type,
                "Sub-Type Count": sub_count,
            })
    return pd.DataFrame(rows)


def get_top_stations(df: pd.DataFrame, n: int = 25) -> pd.DataFrame:
    group_cols = [c for c in ["Station ID", "Station Name"] if c in df.columns]
    if not group_cols or df.empty:
        return pd.DataFrame(columns=["Station Name", "Fault Count"])

    return (
        df.groupby(group_cols)
        .size()
        .reset_index(name="Fault Count")
        .sort_values("Fault Count", ascending=False)
        .head(n)
    )


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filters")
    filtered = df.copy()

    rendered_filters = set()
    filter_configs = [
        ("Quarter", "Quarter"),
        ("Zone", "Zone"),
        ("Severity", "Severity"),
        ("Charger Make", "Charger Make"),
        ("Issue Type", "Issue Type"),
        ("Status", "Status")
    ]

    for col, label in filter_configs:
        if col in filtered.columns and col not in rendered_filters:
            rendered_filters.add(col)
            raw_options = filtered[col].dropna().unique()
            # De-duplicate whitespace/case clean options
            unique_opts = list(dict.fromkeys([
                str(x).strip() for x in raw_options
                if pd.notna(x) and str(x).strip() != "" and str(x).strip().lower() not in ["nat", "nan"]
            ]))
            options = sorted(unique_opts, key=lambda s: s.lower())

            if not options:
                continue

            selected = st.sidebar.multiselect(
                label,
                options=options,
                default=options,
                key=f"filter_widget_{col}"
            )

            filtered = filtered[filtered[col].astype(str).str.strip().isin(selected)]

    # Issue Date range filter — includes latest reports up to 23:59:59 of end date
    if "Issue Date" in df.columns and df["Issue Date"].notna().any():
        valid_dates = df["Issue Date"].dropna()
        if not valid_dates.empty:
            min_d = valid_dates.min().date()
            max_d = valid_dates.max().date()

            date_range = st.sidebar.date_input(
                "Issue Date range",
                value=(min_d, max_d) if min_d != max_d else min_d,
                min_value=min_d,
                max_value=max_d,
                key="filter_widget_issue_date_range"
            )

            if isinstance(date_range, (list, tuple)):
                if len(date_range) == 2:
                    start = pd.Timestamp(date_range[0]).floor("D")
                    # Fix: ensure end date covers the full day till 23:59:59 so latest reports are not cut off
                    end = pd.Timestamp(date_range[1]).replace(hour=23, minute=59, second=59, microsecond=999999)
                    filtered = filtered[(filtered["Issue Date"] >= start) & (filtered["Issue Date"] <= end)]
                elif len(date_range) == 1:
                    start = pd.Timestamp(date_range[0]).floor("D")
                    end = pd.Timestamp(date_range[0]).replace(hour=23, minute=59, second=59, microsecond=999999)
                    filtered = filtered[(filtered["Issue Date"] >= start) & (filtered["Issue Date"] <= end)]

    return filtered


def render_kpis(df: pd.DataFrame, repetitive: pd.DataFrame):
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Faults", len(df))

    if "Status" in df.columns and len(df):
        open_count = df["Status"].astype(str).str.strip().str.lower().isin(["open", "pending", "in progress"]).sum()
        col2.metric("Open Faults", int(open_count))
    else:
        col2.metric("Open Faults", 0)

    col3.metric("Repetitive Fault Pairs", len(repetitive))

    if "TAT Compliance" in df.columns and len(df):
        valid_tat = df["TAT Compliance"].dropna()
        if len(valid_tat):
            compliant = valid_tat.astype(str).str.strip().str.lower().isin(["yes", "y", "true", "1", "compliant"]).mean()
            col4.metric("TAT Compliance", f"{compliant * 100:.1f}%")
        else:
            col4.metric("TAT Compliance", "N/A")
    else:
        col4.metric("TAT Compliance", "N/A")


def render_charts(df: pd.DataFrame, top_stations: pd.DataFrame):
    left, right = st.columns(2)

    if "Zone" in df.columns and not df.empty:
        zone_counts = df["Zone"].dropna().value_counts().reset_index()
        zone_counts.columns = ["Zone", "Faults"]
        if not zone_counts.empty:
            fig = px.bar(
                zone_counts,
                x="Zone",
                y="Faults",
                title="Region-wise Faults (Zone)",
                text_auto=True,
                color_discrete_sequence=["#1f77b4"]
            )
            fig.update_layout(xaxis_title="Zone", yaxis_title="Faults", margin=dict(l=20, r=20, t=40, b=20))
            left.plotly_chart(fig, use_container_width=True)

    if "Charger Make" in df.columns and not df.empty:
        make_counts = df["Charger Make"].dropna().value_counts().reset_index()
        make_counts.columns = ["Charger Make", "Faults"]
        if not make_counts.empty:
            fig = px.pie(
                make_counts,
                names="Charger Make",
                values="Faults",
                title="Faults by Charger Company",
                hole=0.3
            )
            fig.update_layout(margin=dict(l=20, r=20, t=40, b=20))
            right.plotly_chart(fig, use_container_width=True)

    if not top_stations.empty and "Station Name" in top_stations.columns:
        fig = px.bar(
            top_stations,
            x="Fault Count",
            y="Station Name",
            orientation="h",
            title="Top 25 Stations by Fault Count",
            text_auto=True,
            color_discrete_sequence=["#2ca02c"]
        )
        fig.update_layout(
            yaxis=dict(autorange="reversed", title="Station Name"),
            xaxis_title="Fault Count",
            height=max(400, len(top_stations) * 25),
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig, use_container_width=True)


def get_time_aggregates(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Return (daily, monthly, quarterly) aggregations of fault counts.
    Each DataFrame has columns [Time, Faults].
    """
    if "Issue Date" not in df.columns or df.empty or df["Issue Date"].isna().all():
        empty = pd.DataFrame(columns=["Time", "Faults"])
        return empty, empty, empty

    valid = df[df["Issue Date"].notna()].copy()

    daily = valid.groupby(valid["Issue Date"].dt.date).size().reset_index(name="Faults")
    daily.columns = ["Time", "Faults"]
    daily["Time"] = daily["Time"].astype(str)

    monthly = valid.groupby(valid["Issue Date"].dt.to_period("M").astype(str)).size().reset_index(name="Faults")
    monthly.columns = ["Time", "Faults"]

    quarterly = valid.groupby(valid["Issue Date"].dt.year.astype(str) + "-Q" + valid["Issue Date"].dt.quarter.astype(str)).size().reset_index(name="Faults")
    quarterly.columns = ["Time", "Faults"]

    return daily, monthly, quarterly


def render_time_reports(df: pd.DataFrame):
    daily, monthly, quarterly = get_time_aggregates(df)

    st.subheader("Date-wise (Daily) Faults")
    if daily.empty:
        st.info("No Issue Date data available for daily report.")
    else:
        fig = px.bar(
            daily,
            x="Time",
            y="Faults",
            title="Faults by Day",
            text_auto=True,
            color_discrete_sequence=["#1f77b4"]
        )
        fig.update_layout(xaxis_title="Date", yaxis_title="Faults", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)
        # Show latest daily reports first in table view
        safe_dataframe(daily.sort_values("Time", ascending=False))

    st.subheader("Month-wise Faults")
    if monthly.empty:
        st.info("No Issue Date data available for monthly report.")
    else:
        fig = px.bar(
            monthly,
            x="Time",
            y="Faults",
            title="Faults by Month",
            text_auto=True,
            color_discrete_sequence=["#ff7f0e"]
        )
        fig.update_layout(xaxis_title="Month (YYYY-MM)", yaxis_title="Faults", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)
        # Show latest monthly reports first in table view
        safe_dataframe(monthly.sort_values("Time", ascending=False))

    st.subheader("Quarter-wise Faults")
    if quarterly.empty:
        st.info("No Issue Date data available for quarterly report.")
    else:
        fig = px.bar(
            quarterly,
            x="Time",
            y="Faults",
            title="Faults by Quarter",
            text_auto=True,
            color_discrete_sequence=["#2ca02c"]
        )
        fig.update_layout(xaxis_title="Quarter (YYYY-Q#)", yaxis_title="Faults", margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)
        # Show latest quarterly reports first in table view
        safe_dataframe(quarterly.sort_values("Time", ascending=False))


def main():
    st.title("Charger Fault Report Dashboard")

    default_file = DEFAULT_PATH
    if not default_file and os.path.exists("Demo.xlsx"):
        default_file = "Demo.xlsx"

    st.sidebar.header("Data Source")
    file_path = st.sidebar.text_input("File path on this machine", value=default_file, key="input_file_path")
    uploaded = st.sidebar.file_uploader("...or upload a file", type=["xlsx"], key="input_file_uploader")

    if st.sidebar.button("Refresh data", key="btn_refresh"):
        try:
            st.cache_data.clear()
            st.sidebar.success("Cache cleared")
        except Exception:
            pass

    if uploaded is not None:
        source = uploaded
    elif file_path:
        path = Path(file_path)
        if not path.exists():
            st.error(f"No file found at: {file_path}")
            return
        source = path
    else:
        st.info("Enter a file path above, or upload a file, to get started.")
        return

    try:
        df = load_data(source)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return

    if df.empty:
        st.warning("Loaded file is empty.")
        return

    filtered = apply_filters(df)

    # Show active date range banner if Issue Date exists
    if "Issue Date" in filtered.columns and filtered["Issue Date"].notna().any():
        min_date_str = filtered["Issue Date"].min().strftime("%Y-%m-%d")
        max_date_str = filtered["Issue Date"].max().strftime("%Y-%m-%d")
        st.info(f"Showing **{len(filtered)}** records from **{min_date_str}** to **{max_date_str}** (Total in file: {len(df)})")
    else:
        st.info(f"Showing **{len(filtered)}** of **{len(df)}** total records")

    repetitive = get_repetitive_faults(filtered)
    top_stations = get_top_stations(filtered)

    render_kpis(filtered, repetitive)
    render_charts(filtered, top_stations)

    # Time-based reports
    render_time_reports(filtered)

    st.subheader("Top 5 Issue Types (with Top 5 Sub-Types each)")
    issue_breakdown = get_top_issue_breakdown(filtered)
    if issue_breakdown.empty:
        st.write("No 'Issue Type' / 'Issue Sub-Type' breakdown available.")
    else:
        for issue_type in issue_breakdown["Issue Type"].unique():
            group = issue_breakdown[issue_breakdown["Issue Type"] == issue_type]
            total = group["Issue Type Total"].iloc[0]
            with st.expander(f"{issue_type} — {total} faults"):
                fig = px.bar(
                    group,
                    x="Sub-Type Count",
                    y="Issue Sub-Type",
                    orientation="h",
                    title=f"Top Sub-Types for {issue_type}",
                    text_auto=True,
                    color_discrete_sequence=["#d62728"]
                )
                fig.update_layout(
                    yaxis=dict(autorange="reversed", title="Issue Sub-Type"),
                    xaxis_title="Sub-Type Count",
                    height=max(250, len(group) * 35),
                    margin=dict(l=20, r=20, t=40, b=20)
                )
                st.plotly_chart(fig, use_container_width=True)
                safe_dataframe(group[["Issue Sub-Type", "Sub-Type Count"]])

    st.subheader("Repetitive Faults (same station, same issue sub-type, 2+ times)")
    if repetitive.empty:
        st.info("No repetitive faults found for the selected filters.")
    else:
        safe_dataframe(repetitive)

    st.subheader("Top 25 Stations")
    if top_stations.empty:
        st.info("No station fault data available.")
    else:
        safe_dataframe(top_stations)

    # Latest Reports & Raw Data View
    st.subheader("Latest Fault Reports (Filtered Raw Data)")
    if "Issue Date" in filtered.columns:
        display_df = filtered.sort_values("Issue Date", ascending=False)
    else:
        display_df = filtered

    safe_dataframe(display_df)


if __name__ == "__main__":
    main()
