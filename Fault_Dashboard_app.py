import sys
import os
import io
import datetime
from pathlib import Path
from typing import Tuple, List, Optional, Union

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# Check optional PDF generation library
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Check optional PPT generation library
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    PYTHON_PPTX_AVAILABLE = True
except ImportError:
    PYTHON_PPTX_AVAILABLE = False

# Streamlit Page Configuration
st.set_page_config(
    page_title="Charger Fault Analytics Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inject Modern CSS Styling
st.markdown("""
<style>
    /* Metric Cards Styling */
    .metric-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 18px 20px;
        color: #ffffff;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        text-align: center;
    }
    .metric-label {
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #94a3b8;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        line-height: 1.2;
    }
    .val-blue { color: #38bdf8; }
    .val-red { color: #f87171; }
    .val-amber { color: #fbbf24; }
    .val-green { color: #4ade80; }

    /* Clean Tab Headers */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
    }
</style>
""", unsafe_allow_html=True)

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


def read_file_bytes(file_source: Union[str, Path, io.BytesIO, any]) -> Tuple[bytes, str]:
    """Extract raw bytes and filename safely from a path or Streamlit UploadedFile object."""
    if isinstance(file_source, (str, Path)):
        path = Path(file_source)
        filename = path.name
        with open(path, "rb") as f:
            return f.read(), filename
    elif hasattr(file_source, "name"):
        filename = getattr(file_source, "name", "data.xlsx")
        if hasattr(file_source, "getvalue"):
            b = file_source.getvalue()
        else:
            b = file_source.read()
        if hasattr(file_source, "seek"):
            file_source.seek(0)
        return b, filename
    elif isinstance(file_source, bytes):
        return file_source, "data.xlsx"
    else:
        raise ValueError("Unsupported file source format.")


@st.cache_data(ttl=300)
def parse_and_clean_data(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Load Excel or CSV raw bytes into a DataFrame and perform clean, robust data standardization.
    Automatically handles missing values, dates, and column mappings.
    """
    ext = filename.lower().split(".")[-1]

    if ext == "csv":
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="latin1", encoding_errors="replace")
    elif ext in ["xls", "xlsx", "xlsm", "xlsb"]:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl" if ext == "xlsx" else None)
        except Exception:
            df = pd.read_excel(io.BytesIO(file_bytes))
    else:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding_errors="replace")

    if df.empty:
        return df

    # Clean column headers
    df.columns = [str(c).strip() for c in df.columns]
    # Remove empty unnamed columns
    df = df.loc[:, ~df.columns.str.startswith("Unnamed:") | df.notna().any()]
    df = df.loc[:, ~df.columns.duplicated()]

    # Global string null cleaning
    null_strings = {"nan", "nat", "none", "null", "-", "n/a", "na", "<na>", "undefined", ""}
    for col in df.columns:
        if df[col].dtype == object or isinstance(df[col].dtype, pd.CategoricalDtype):
            df[col] = df[col].apply(
                lambda val: np.nan if pd.isna(val) or str(val).strip().lower() in null_strings else str(val).strip()
            )

    # Standardize column names flexibly
    col_mappings = {
        "Issue Date": ["Issue Date", "Issue date", "Created Date", "Date", "Ticket Date", "Fault Date", "Report Date"],
        "Resolution Date": ["Resolution Date", "Resolution date", "Closed Date", "Resolve Date"],
        "Restoration Date": ["Restoration Date", "Restoration date", "Restored Date"],
        "Station ID": ["Station ID", "Station id", "StationID", "Station_ID", "Site ID", "Location ID"],
        "Station Name": ["Station Name", "Station name", "StationName", "Station_Name", "Site Name", "Location"],
        "Issue Type": ["Issue Type", "Issue type", "Fault Type", "Category", "Main Category", "Issue Category"],
        "Issue Sub-Type": ["Issue Sub-Type", "Issue sub-type", "Sub Category", "Sub Type", "Fault Subtype", "Sub-Category"],
        "Zone": ["Zone", "Region", "State", "Circle", "Area", "Territory"],
        "Severity": ["Severity", "Priority", "Impact", "Urgency"],
        "Charger Make": ["Charger Make", "Charger Company", "OEM", "Make", "Vendor", "Brand", "Manufacturer"],
        "Status": ["Status", "Ticket Status", "State", "Current Status"],
        "TAT Compliance": ["TAT Compliance", "TAT", "SLA Compliance", "SLA Met", "TAT Met", "SLA Status"]
    }

    for std_name, candidates in col_mappings.items():
        if std_name not in df.columns:
            found = find_col(df, candidates)
            if found and found != std_name:
                df.rename(columns={found: std_name}, inplace=True)

    # Standardize datetime columns
    for col in DATE_COLS:
        if col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
            except Exception:
                df[col] = pd.to_datetime(df[col], errors="coerce")

    # Standardize Status formatting
    if "Status" in df.columns:
        df["Status"] = df["Status"].fillna("Unknown").astype(str).str.strip().str.title()

    # Standardize TAT Compliance values
    if "TAT Compliance" in df.columns:
        def clean_tat(v):
            if pd.isna(v):
                return "Unknown"
            sv = str(v).strip().lower()
            if sv in ["yes", "y", "1", "true", "compliant", "met", "within tat", "tat met"]:
                return "Compliant"
            elif sv in ["no", "n", "0", "false", "non-compliant", "breached", "missed", "tat breached"]:
                return "Non-Compliant"
            return str(v).strip().title()
        
        df["TAT Compliance Standardized"] = df["TAT Compliance"].apply(clean_tat)

    # Derive time columns from Issue Date for filtering and reporting
    if "Issue Date" in df.columns and df["Issue Date"].notna().any():
        issue_dt = df["Issue Date"]
        df["Issue Day"] = issue_dt.dt.date
        df["Issue Month"] = issue_dt.dt.to_period("M").astype(str).replace("NaT", np.nan)

        q_year = issue_dt.dt.year.astype(str)
        q_num = issue_dt.dt.quarter.astype(str)
        df["Quarter"] = (q_year + "-Q" + q_num).where(issue_dt.notna(), np.nan)

    return df


def load_data(file_source) -> pd.DataFrame:
    """Wrapper to safely extract bytes and run cached parser."""
    file_bytes, filename = read_file_bytes(file_source)
    return parse_and_clean_data(file_bytes, filename)


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
    st.sidebar.header("Filter Criteria")
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
            unique_opts = list(dict.fromkeys([
                str(x).strip() for x in raw_options
                if pd.notna(x) and str(x).strip() != "" and str(x).strip().lower() not in ["nat", "nan", "unknown"]
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

    # Issue Date range filter
    if "Issue Date" in df.columns and df["Issue Date"].notna().any():
        valid_dates = df["Issue Date"].dropna()
        if not valid_dates.empty:
            min_d = valid_dates.min().date()
            max_d = valid_dates.max().date()

            date_range = st.sidebar.date_input(
                "Issue Date Range",
                value=(min_d, max_d) if min_d != max_d else min_d,
                min_value=min_d,
                max_value=max_d,
                key="filter_widget_issue_date_range"
            )

            if isinstance(date_range, (list, tuple)):
                if len(date_range) == 2:
                    start = pd.Timestamp(date_range[0]).floor("D")
                    end = pd.Timestamp(date_range[1]).replace(hour=23, minute=59, second=59, microsecond=999999)
                    filtered = filtered[(filtered["Issue Date"] >= start) & (filtered["Issue Date"] <= end)]
                elif len(date_range) == 1:
                    start = pd.Timestamp(date_range[0]).floor("D")
                    end = pd.Timestamp(date_range[0]).replace(hour=23, minute=59, second=59, microsecond=999999)
                    filtered = filtered[(filtered["Issue Date"] >= start) & (filtered["Issue Date"] <= end)]

    return filtered


# PDF Report Generator
def generate_pdf_report(
    df: pd.DataFrame,
    filtered_df: pd.DataFrame,
    repetitive_df: pd.DataFrame,
    top_stations_df: pd.DataFrame,
    metadata_info: dict
) -> bytes:
    """Generate executive PDF report using ReportLab."""
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab package is missing. Install with `pip install reportlab`.")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontSize=20,
        leading=24,
        textColor=colors.HexColor('#0F172A'),
        fontName='Helvetica-Bold',
        spaceAfter=6
    )

    subtitle_style = ParagraphStyle(
        'SubTitle',
        parent=styles['Normal'],
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#475569'),
        spaceAfter=15
    )

    heading2_style = ParagraphStyle(
        'Heading2Custom',
        parent=styles['Heading2'],
        fontSize=12,
        leading=15,
        textColor=colors.HexColor('#1E293B'),
        fontName='Helvetica-Bold',
        spaceBefore=10,
        spaceAfter=6
    )

    body_style = ParagraphStyle(
        'BodyCustom',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#334155')
    )

    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=10,
        textColor=colors.white,
        fontName='Helvetica-Bold',
        alignment=1
    )

    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#1E293B')
    )

    elements = []

    # Title & Subtitle Banner
    elements.append(Paragraph("Charger Fault Analytics Report", title_style))
    date_str = metadata_info.get("date_range", "All Time")
    gen_time = metadata_info.get("generated_at", "")
    rec_count = len(filtered_df)
    total_rec = len(df)
    elements.append(Paragraph(f"<b>Data Date Range:</b> {date_str} &nbsp;|&nbsp; <b>Filtered Records:</b> {rec_count} of {total_rec} &nbsp;|&nbsp; <b>Generated:</b> {gen_time}", subtitle_style))
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#CBD5E1'), spaceAfter=12))

    # KPI Summary Cards
    total_faults = len(filtered_df)
    if "Status" in filtered_df.columns and len(filtered_df):
        open_faults = filtered_df["Status"].astype(str).str.strip().str.lower().isin(["open", "pending", "in progress"]).sum()
    else:
        open_faults = 0

    repetitive_count = len(repetitive_df)

    tat_pct = "N/A"
    tat_col = "TAT Compliance Standardized" if "TAT Compliance Standardized" in filtered_df.columns else "TAT Compliance"
    if tat_col in filtered_df.columns and len(filtered_df):
        valid_tat = filtered_df[tat_col].dropna()
        if len(valid_tat):
            compliant_ratio = valid_tat.astype(str).str.strip().str.lower().isin(["compliant", "yes", "y", "true", "1"]).mean()
            tat_pct = f"{compliant_ratio * 100:.1f}%"

    kpi_data = [
        [
            Paragraph("<b>Total Faults</b>", table_header_style),
            Paragraph("<b>Open Faults</b>", table_header_style),
            Paragraph("<b>Repetitive Fault Pairs</b>", table_header_style),
            Paragraph("<b>TAT Compliance</b>", table_header_style)
        ],
        [
            Paragraph(f"<font size=13 color='#0284C7'><b>{total_faults:,}</b></font>", ParagraphStyle('KPI1', alignment=1)),
            Paragraph(f"<font size=13 color='#DC2626'><b>{open_faults:,}</b></font>", ParagraphStyle('KPI2', alignment=1)),
            Paragraph(f"<font size=13 color='#D97706'><b>{repetitive_count:,}</b></font>", ParagraphStyle('KPI3', alignment=1)),
            Paragraph(f"<font size=13 color='#16A34A'><b>{tat_pct}</b></font>", ParagraphStyle('KPI4', alignment=1))
        ]
    ]

    kpi_table = Table(kpi_data, colWidths=[130, 130, 130, 130])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#F8FAFC')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(kpi_table)
    elements.append(Spacer(1, 10))

    # Regional & Charger Make Summary
    elements.append(Paragraph("Regional & Charger Vendor Breakdown", heading2_style))

    zone_summary = []
    if "Zone" in filtered_df.columns and not filtered_df.empty:
        z_counts = filtered_df["Zone"].dropna().value_counts().head(5)
        for z, c in z_counts.items():
            pct = (c / len(filtered_df)) * 100
            zone_summary.append(f"• <b>{z}:</b> {c} faults ({pct:.1f}%)")

    make_summary = []
    if "Charger Make" in filtered_df.columns and not filtered_df.empty:
        m_counts = filtered_df["Charger Make"].dropna().value_counts().head(5)
        for m, c in m_counts.items():
            pct = (c / len(filtered_df)) * 100
            make_summary.append(f"• <b>{m}:</b> {c} faults ({pct:.1f}%)")

    z_text = "<br/>".join(zone_summary) if zone_summary else "No Zone data"
    m_text = "<br/>".join(make_summary) if make_summary else "No Charger Make data"

    reg_data = [
        [Paragraph("<b>Top Zones</b>", table_header_style), Paragraph("<b>Top Charger Vendors</b>", table_header_style)],
        [Paragraph(z_text, body_style), Paragraph(m_text, body_style)]
    ]
    reg_table = Table(reg_data, colWidths=[260, 260])
    reg_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#334155')),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#FFFFFF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(reg_table)
    elements.append(Spacer(1, 10))

    # Repetitive Faults Table
    elements.append(Paragraph("Repetitive Fault Patterns (Station + Sub-Type >= 2)", heading2_style))
    if repetitive_df.empty:
        elements.append(Paragraph("<i>No repetitive fault patterns detected.</i>", body_style))
    else:
        rep_rows = [[
            Paragraph("<b>Station ID</b>", table_header_style),
            Paragraph("<b>Station Name</b>", table_header_style),
            Paragraph("<b>Issue Sub-Type</b>", table_header_style),
            Paragraph("<b>Occurrences</b>", table_header_style)
        ]]
        for _, row in repetitive_df.head(8).iterrows():
            st_id = str(row.get("Station ID", "N/A"))
            st_name = str(row.get("Station Name", "N/A"))
            sub_type = str(row.get("Issue Sub-Type", "N/A"))
            occ = str(row.get("Occurrences", "0"))
            rep_rows.append([
                Paragraph(st_id, table_cell_style),
                Paragraph(st_name, table_cell_style),
                Paragraph(sub_type, table_cell_style),
                Paragraph(f"<b>{occ}</b>", table_cell_style)
            ])
        rep_table = Table(rep_rows, colWidths=[100, 180, 160, 80])
        rep_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0F172A')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F8FAFC'), colors.HexColor('#FFFFFF')]),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
            ('PADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (3, 0), (3, -1), 'CENTER')
        ]))
        elements.append(rep_table)

    elements.append(Spacer(1, 10))

    # Top Stations Table
    elements.append(Paragraph("Top Affected Stations", heading2_style))
    if top_stations_df.empty:
        elements.append(Paragraph("<i>No station fault data available.</i>", body_style))
    else:
        st_rows = [[
            Paragraph("<b>#</b>", table_header_style),
            Paragraph("<b>Station ID</b>", table_header_style),
            Paragraph("<b>Station Name</b>", table_header_style),
            Paragraph("<b>Fault Count</b>", table_header_style)
        ]]
        for idx, row in enumerate(top_stations_df.head(8).iterrows(), 1):
            r_data = row[1]
            st_id = str(r_data.get("Station ID", "N/A")) if "Station ID" in r_data else "-"
            st_name = str(r_data.get("Station Name", "N/A"))
            cnt = str(r_data.get("Fault Count", "0"))
            st_rows.append([
                Paragraph(str(idx), table_cell_style),
                Paragraph(st_id, table_cell_style),
                Paragraph(st_name, table_cell_style),
                Paragraph(f"<b>{cnt}</b>", table_cell_style)
            ])
        top_st_table = Table(st_rows, colWidths=[30, 130, 260, 100])
        top_st_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F8FAFC'), colors.HexColor('#FFFFFF')]),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
            ('PADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (3, 0), (3, -1), 'CENTER')
        ]))
        elements.append(top_st_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


# PowerPoint Presentation Generator
def generate_ppt_report(
    df: pd.DataFrame,
    filtered_df: pd.DataFrame,
    repetitive_df: pd.DataFrame,
    top_stations_df: pd.DataFrame,
    metadata_info: dict
) -> bytes:
    """Generate high quality PowerPoint slide deck using python-pptx."""
    if not PYTHON_PPTX_AVAILABLE:
        raise RuntimeError("python-pptx package is missing. Install with `pip install python-pptx`.")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="CHARGER FAULT ANALYTICS"):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(15, 23, 42)
        shape.line.fill.background()

        txBox = slide.shapes.add_textbox(Inches(0.6), Inches(0.15), Inches(12), Inches(0.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = category_text.upper()
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(56, 189, 248)

        p2 = tf.add_paragraph()
        p2.text = title_text
        p2.font.size = Pt(20)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(255, 255, 255)

    # ----------------------------------------------------
    # SLIDE 1: Title Slide
    # ----------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    bg = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(15, 23, 42)

    tbox = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.5))
    tf = tbox.text_frame
    p1 = tf.paragraphs[0]
    p1.text = "Charger Fault Analytics Presentation"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(255, 255, 255)

    p2 = tf.add_paragraph()
    p2.text = "Executive Summary & Operational Performance Report"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(148, 163, 184)
    p2.space_before = Pt(10)

    p3 = tf.add_paragraph()
    date_str = metadata_info.get("date_range", "All Time")
    gen_time = metadata_info.get("generated_at", "")
    p3.text = f"Date Filter: {date_str}   |   Active Records: {len(filtered_df):,}   |   Generated: {gen_time}"
    p3.font.size = Pt(12)
    p3.font.color.rgb = RGBColor(56, 189, 248)
    p3.space_before = Pt(25)

    # ----------------------------------------------------
    # SLIDE 2: KPI Metrics Slide
    # ----------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Executive Key Performance Indicators")

    total_faults = len(filtered_df)
    if "Status" in filtered_df.columns and len(filtered_df):
        open_faults = int(filtered_df["Status"].astype(str).str.strip().str.lower().isin(["open", "pending", "in progress"]).sum())
    else:
        open_faults = 0
    repetitive_count = len(repetitive_df)

    tat_pct = "N/A"
    tat_col = "TAT Compliance Standardized" if "TAT Compliance Standardized" in filtered_df.columns else "TAT Compliance"
    if tat_col in filtered_df.columns and len(filtered_df):
        valid_tat = filtered_df[tat_col].dropna()
        if len(valid_tat):
            compliant_ratio = valid_tat.astype(str).str.strip().str.lower().isin(["compliant", "yes", "y", "true", "1"]).mean()
            tat_pct = f"{compliant_ratio * 100:.1f}%"

    metrics = [
        ("TOTAL FAULTS", f"{total_faults:,}", RGBColor(2, 132, 199)),
        ("OPEN FAULTS", f"{open_faults:,}", RGBColor(220, 38, 38)),
        ("REPETITIVE PAIRS", f"{repetitive_count:,}", RGBColor(217, 119, 6)),
        ("TAT COMPLIANCE", tat_pct, RGBColor(22, 163, 74))
    ]

    card_w = Inches(2.7)
    card_h = Inches(2.2)
    start_x = Inches(0.8)
    gap = Inches(0.35)
    start_y = Inches(1.6)

    for i, (label, val, color) in enumerate(metrics):
        cx = start_x + i * (card_w + gap)
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, start_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(248, 250, 252)
        card.line.color.rgb = RGBColor(226, 232, 240)

        tb = s2.shapes.add_textbox(cx, start_y + Inches(0.25), card_w, Inches(1.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        p1.text = label
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(100, 116, 139)

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.text = val
        p2.font.size = Pt(28)
        p2.font.bold = True
        p2.font.color.rgb = color
        p2.space_before = Pt(14)

    summary_box = s2.shapes.add_textbox(Inches(0.8), Inches(4.2), Inches(11.7), Inches(2.8))
    stf = summary_box.text_frame
    stf.word_wrap = True
    p = stf.paragraphs[0]
    p.text = "Operational Insights & Observations:"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor(15, 23, 42)

    bullets = [
        f"A total of {total_faults:,} fault tickets analyzed in the current view.",
        f"{open_faults:,} tickets currently remain Open or Pending resolution.",
        f"{repetitive_count:,} repetitive fault combinations (same station + issue subtype >= 2 occurrences) identified.",
        f"Overall Service Level / TAT compliance rate stands at {tat_pct}."
    ]

    for b in bullets:
        bp = stf.add_paragraph()
        bp.text = "• " + b
        bp.font.size = Pt(12)
        bp.font.color.rgb = RGBColor(51, 65, 85)
        bp.space_before = Pt(6)

    # ----------------------------------------------------
    # SLIDE 3: Regional & Vendor Breakdown
    # ----------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Regional & OEM Vendor Fault Distribution")

    z_counts = filtered_df["Zone"].dropna().value_counts().head(5) if "Zone" in filtered_df.columns else []
    z_rows = max(2, len(z_counts) + 1)

    table_shape1 = s3.shapes.add_table(z_rows, 3, Inches(0.8), Inches(1.6), Inches(5.6), Inches(4.8))
    t1 = table_shape1.table
    t1.columns[0].width = Inches(0.8)
    t1.columns[1].width = Inches(3.2)
    t1.columns[2].width = Inches(1.6)

    t1.cell(0, 0).text = "#"
    t1.cell(0, 1).text = "Zone / Region"
    t1.cell(0, 2).text = "Fault Count"

    for c in range(3):
        t1.cell(0, c).fill.solid()
        t1.cell(0, c).fill.fore_color.rgb = RGBColor(30, 41, 59)
        for p in t1.cell(0, c).text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

    if len(z_counts):
        for i, (z_name, z_val) in enumerate(z_counts.items(), 1):
            t1.cell(i, 0).text = str(i)
            t1.cell(i, 1).text = str(z_name)
            t1.cell(i, 2).text = f"{z_val:,}"
            for c in range(3):
                t1.cell(i, c).fill.solid()
                t1.cell(i, c).fill.fore_color.rgb = RGBColor(248, 250, 252) if i % 2 == 1 else RGBColor(255, 255, 255)

    m_counts = filtered_df["Charger Make"].dropna().value_counts().head(5) if "Charger Make" in filtered_df.columns else []
    m_rows = max(2, len(m_counts) + 1)

    table_shape2 = s3.shapes.add_table(m_rows, 3, Inches(6.9), Inches(1.6), Inches(5.6), Inches(4.8))
    t2 = table_shape2.table
    t2.columns[0].width = Inches(0.8)
    t2.columns[1].width = Inches(3.2)
    t2.columns[2].width = Inches(1.6)

    t2.cell(0, 0).text = "#"
    t2.cell(0, 1).text = "Charger Make / OEM"
    t2.cell(0, 2).text = "Fault Count"

    for c in range(3):
        t2.cell(0, c).fill.solid()
        t2.cell(0, c).fill.fore_color.rgb = RGBColor(30, 41, 59)
        for p in t2.cell(0, c).text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

    if len(m_counts):
        for i, (m_name, m_val) in enumerate(m_counts.items(), 1):
            t2.cell(i, 0).text = str(i)
            t2.cell(i, 1).text = str(m_name)
            t2.cell(i, 2).text = f"{m_val:,}"
            for c in range(3):
                t2.cell(i, c).fill.solid()
                t2.cell(i, c).fill.fore_color.rgb = RGBColor(248, 250, 252) if i % 2 == 1 else RGBColor(255, 255, 255)

    # ----------------------------------------------------
    # SLIDE 4: Repetitive Fault Pairs
    # ----------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Repetitive Fault Pairs & Top Stations")

    top_rep = repetitive_df.head(6) if not repetitive_df.empty else pd.DataFrame()
    r_rows = max(2, len(top_rep) + 1)

    ts_shape = s4.shapes.add_table(r_rows, 4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.0))
    st_table = ts_shape.table
    st_table.columns[0].width = Inches(2.2)
    st_table.columns[1].width = Inches(4.5)
    st_table.columns[2].width = Inches(3.5)
    st_table.columns[3].width = Inches(1.5)

    st_table.cell(0, 0).text = "Station ID"
    st_table.cell(0, 1).text = "Station Name"
    st_table.cell(0, 2).text = "Issue Sub-Type"
    st_table.cell(0, 3).text = "Occurrences"

    for c in range(4):
        st_table.cell(0, c).fill.solid()
        st_table.cell(0, c).fill.fore_color.rgb = RGBColor(15, 23, 42)
        for p in st_table.cell(0, c).text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

    if not top_rep.empty:
        for idx, (_, r_data) in enumerate(top_rep.iterrows(), 1):
            st_table.cell(idx, 0).text = str(r_data.get("Station ID", "N/A"))
            st_table.cell(idx, 1).text = str(r_data.get("Station Name", "N/A"))
            st_table.cell(idx, 2).text = str(r_data.get("Issue Sub-Type", "N/A"))
            st_table.cell(idx, 3).text = str(r_data.get("Occurrences", "0"))

            for c in range(4):
                st_table.cell(idx, c).fill.solid()
                st_table.cell(idx, c).fill.fore_color.rgb = RGBColor(248, 250, 252) if idx % 2 == 1 else RGBColor(255, 255, 255)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def render_kpis(df: pd.DataFrame, repetitive: pd.DataFrame):
    col1, col2, col3, col4 = st.columns(4)

    total_count = len(df)
    col1.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Total Faults</div>
        <div class="metric-value val-blue">{total_count:,}</div>
    </div>
    """, unsafe_allow_html=True)

    if "Status" in df.columns and len(df):
        open_count = df["Status"].astype(str).str.strip().str.lower().isin(["open", "pending", "in progress"]).sum()
    else:
        open_count = 0

    col2.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Open Faults</div>
        <div class="metric-value val-red">{int(open_count):,}</div>
    </div>
    """, unsafe_allow_html=True)

    col3.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Repetitive Fault Pairs</div>
        <div class="metric-value val-amber">{len(repetitive):,}</div>
    </div>
    """, unsafe_allow_html=True)

    tat_col = "TAT Compliance Standardized" if "TAT Compliance Standardized" in df.columns else "TAT Compliance"
    if tat_col in df.columns and len(df):
        valid_tat = df[tat_col].dropna()
        if len(valid_tat):
            compliant = valid_tat.astype(str).str.strip().str.lower().isin(["compliant", "yes", "y", "true", "1"]).mean()
            tat_str = f"{compliant * 100:.1f}%"
        else:
            tat_str = "N/A"
    else:
        tat_str = "N/A"

    col4.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">TAT Compliance</div>
        <div class="metric-value val-green">{tat_str}</div>
    </div>
    """, unsafe_allow_html=True)


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
                color_discrete_sequence=["#0284c7"]
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
                hole=0.35,
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig.update_layout(margin=dict(l=20, r=20, t=40, b=20))
            right.plotly_chart(fig, use_container_width=True)

    if not top_stations.empty and "Station Name" in top_stations.columns:
        fig = px.bar(
            top_stations,
            x="Fault Count",
            y="Station Name",
            orientation="h",
            title="Top Affected Stations by Fault Count",
            text_auto=True,
            color_discrete_sequence=["#16a34a"]
        )
        fig.update_layout(
            yaxis=dict(autorange="reversed", title="Station Name"),
            xaxis_title="Fault Count",
            height=max(400, len(top_stations) * 25),
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig, use_container_width=True)


def get_time_aggregates(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if "Issue Date" not in df.columns or df.empty or df["Issue Date"].isna().all():
        empty = pd.DataFrame(columns=["Time", "Faults"])
        return empty, empty, empty

    valid = df[df["Issue Date"].notna()].copy()

    daily = valid.groupby(valid["Issue Date"].dt.date).size().reset_index(name="Faults")
    daily.columns = ["Time", "Faults"]
    daily["Time"] = daily["Time"].astype(str)

    monthly = valid.groupby(valid["Issue Date"].dt.to_period("M").astype(str)).size().reset_index(name="Faults")
    monthly.columns = ["Time", "Faults"]

    quarterly = valid.groupby(
        valid["Issue Date"].dt.year.astype(str) + "-Q" + valid["Issue Date"].dt.quarter.astype(str)
    ).size().reset_index(name="Faults")
    quarterly.columns = ["Time", "Faults"]

    return daily, monthly, quarterly


def render_time_reports(df: pd.DataFrame):
    daily, monthly, quarterly = get_time_aggregates(df)

    t1, t2, t3 = st.tabs(["Daily Trends", "Monthly Trends", "Quarterly Trends"])

    with t1:
        if daily.empty:
            st.info("No Issue Date data available for daily report.")
        else:
            fig = px.bar(
                daily,
                x="Time",
                y="Faults",
                title="Faults by Day",
                text_auto=True,
                color_discrete_sequence=["#0284c7"]
            )
            fig.update_layout(xaxis_title="Date", yaxis_title="Faults", margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig, use_container_width=True)
            safe_dataframe(daily.sort_values("Time", ascending=False))

    with t2:
        if monthly.empty:
            st.info("No Issue Date data available for monthly report.")
        else:
            fig = px.bar(
                monthly,
                x="Time",
                y="Faults",
                title="Faults by Month",
                text_auto=True,
                color_discrete_sequence=["#ea580c"]
            )
            fig.update_layout(xaxis_title="Month (YYYY-MM)", yaxis_title="Faults", margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig, use_container_width=True)
            safe_dataframe(monthly.sort_values("Time", ascending=False))

    with t3:
        if quarterly.empty:
            st.info("No Issue Date data available for quarterly report.")
        else:
            fig = px.bar(
                quarterly,
                x="Time",
                y="Faults",
                title="Faults by Quarter",
                text_auto=True,
                color_discrete_sequence=["#16a34a"]
            )
            fig.update_layout(xaxis_title="Quarter (YYYY-Q#)", yaxis_title="Faults", margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig, use_container_width=True)
            safe_dataframe(quarterly.sort_values("Time", ascending=False))


def main():
    st.title("⚡ Charger Fault Analytics & Reporting Dashboard")

    default_file = DEFAULT_PATH
    if not default_file and os.path.exists("Demo.xlsx"):
        default_file = "Demo.xlsx"

    st.sidebar.header("Data Ingestion")
    file_path = st.sidebar.text_input("Local File Path", value=default_file, key="input_file_path")
    uploaded = st.sidebar.file_uploader(
        "Upload Sheet (xlsx, xls, csv)",
        type=["xlsx", "xls", "csv"],
        key="input_file_uploader"
    )

    if st.sidebar.button("Clear Cache & Refresh", key="btn_refresh"):
        try:
            st.cache_data.clear()
            st.sidebar.success("Cache refreshed!")
        except Exception:
            pass

    if uploaded is not None:
        source = uploaded
    elif file_path:
        path = Path(file_path)
        if not path.exists():
            st.error(f"File not found at specified path: `{file_path}`")
            return
        source = path
    else:
        st.info("💡 Upload an Excel (`.xlsx`, `.xls`) or CSV (`.csv`) sheet in the sidebar to view metrics.")
        return

    try:
        df = load_data(source)
    except Exception as e:
        st.error(f"❌ Error reading data file: {e}")
        st.info("Please verify the file is not corrupted and contains standard column headers.")
        return

    if df.empty:
        st.warning("⚠️ The loaded data sheet contains no valid records.")
        return

    filtered = apply_filters(df)

    # Sidebar Export / Download Section
    st.sidebar.markdown("---")
    st.sidebar.header("📥 Export Reports")

    date_str = "All Time"
    if "Issue Date" in filtered.columns and filtered["Issue Date"].notna().any():
        min_date_str = filtered["Issue Date"].min().strftime("%Y-%m-%d")
        max_date_str = filtered["Issue Date"].max().strftime("%Y-%m-%d")
        date_str = f"{min_date_str} to {max_date_str}"
        st.info(f"Showing **{len(filtered):,}** records from **{min_date_str}** to **{max_date_str}** (Total in file: {len(df):,})")
    else:
        st.info(f"Showing **{len(filtered):,}** of **{len(df):,}** total records")

    repetitive = get_repetitive_faults(filtered)
    top_stations = get_top_stations(filtered)

    meta_info = {
        "date_range": date_str,
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    }

    # PDF Download Button
    if REPORTLAB_AVAILABLE:
        try:
            pdf_bytes = generate_pdf_report(df, filtered, repetitive, top_stations, meta_info)
            st.sidebar.download_button(
                label="📄 Download PDF Report",
                data=pdf_bytes,
                file_name=f"Charger_Fault_Report_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as pdf_err:
            st.sidebar.warning(f"PDF generation error: {pdf_err}")
    else:
        st.sidebar.warning("`reportlab` missing for PDF download. Run `pip install reportlab`.")

    # PPT Download Button
    if PYTHON_PPTX_AVAILABLE:
        try:
            ppt_bytes = generate_ppt_report(df, filtered, repetitive, top_stations, meta_info)
            st.sidebar.download_button(
                label="📊 Download PPT Presentation",
                data=ppt_bytes,
                file_name=f"Charger_Fault_Presentation_{datetime.date.today()}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
        except Exception as ppt_err:
            st.sidebar.warning(f"PPT generation error: {ppt_err}")
    else:
        st.sidebar.warning("`python-pptx` missing for PPT download. Run `pip install python-pptx`.")

    # Cleaned Excel Download
    try:
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
            filtered.to_excel(writer, index=False, sheet_name="Cleaned Data")
            if not repetitive.empty:
                repetitive.to_excel(writer, index=False, sheet_name="Repetitive Faults")
            top_stations.to_excel(writer, index=False, sheet_name="Top Stations")
        buffer_excel.seek(0)
        st.sidebar.download_button(
            label="📗 Download Cleaned Data (Excel)",
            data=buffer_excel.getvalue(),
            file_name=f"Cleaned_Fault_Data_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception:
        pass

    # Dashboard Main Display Tabs
    t_overview, t_trends, t_issues, t_raw = st.tabs([
        "📊 KPI Overview & Charts",
        "📈 Time-based Reports",
        "⚠️ Issue Breakdown & Repetitive",
        "📋 Filtered Data Table"
    ])

    with t_overview:
        render_kpis(filtered, repetitive)
        st.write("")
        render_charts(filtered, top_stations)

    with t_trends:
        render_time_reports(filtered)

    with t_issues:
        st.subheader("Top 5 Issue Types (with Top Sub-Types)")
        issue_breakdown = get_top_issue_breakdown(filtered)
        if issue_breakdown.empty:
            st.write("No 'Issue Type' / 'Issue Sub-Type' breakdown available.")
        else:
            for issue_type in issue_breakdown["Issue Type"].unique():
                group = issue_breakdown[issue_breakdown["Issue Type"] == issue_type]
                total = group["Issue Type Total"].iloc[0]
                with st.expander(f"{issue_type} — {total} total faults"):
                    fig = px.bar(
                        group,
                        x="Sub-Type Count",
                        y="Issue Sub-Type",
                        orientation="h",
                        title=f"Top Sub-Types for {issue_type}",
                        text_auto=True,
                        color_discrete_sequence=["#dc2626"]
                    )
                    fig.update_layout(
                        yaxis=dict(autorange="reversed", title="Issue Sub-Type"),
                        xaxis_title="Sub-Type Count",
                        height=max(250, len(group) * 35),
                        margin=dict(l=20, r=20, t=40, b=20)
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    safe_dataframe(group[["Issue Sub-Type", "Sub-Type Count"]])

        st.subheader("Repetitive Faults (Same station & issue sub-type, 2+ occurrences)")
        if repetitive.empty:
            st.info("No repetitive faults found for current filter selection.")
        else:
            safe_dataframe(repetitive)

    with t_raw:
        st.subheader("Station Rankings (Top 25)")
        if top_stations.empty:
            st.info("No station fault data available.")
        else:
            safe_dataframe(top_stations)

        st.subheader("Filtered Fault Records")
        if "Issue Date" in filtered.columns:
            display_df = filtered.sort_values("Issue Date", ascending=False)
        else:
            display_df = filtered

        safe_dataframe(display_df)


if __name__ == "__main__":
    main()
