from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from streamlit_autorefresh import st_autorefresh

try:
    import msal
except Exception:
    msal = None

# -----------------------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Project Tracking Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# CONFIG
# -----------------------------------------------------------------------------
LOCAL_EXCEL_FILE = "Project_Tracking_v7.xlsx"
SHEET_NAME = "Project_Master"


def _secret(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, default)).strip()
    except Exception:
        return default


try:
    REFRESH_SECONDS = int(_secret("REFRESH_SECONDS", "60") or 60)
except Exception:
    REFRESH_SECONDS = 60

ENGINEERING_COLS = [
    "Equi. DesigN",
    "Technical Submittal",
    "Drawing",
    "ELS",
    "BOM",
]

DELIVERY_COLS = [
    "Out_Door",
    "Indoor",
    "CR Panels",
    "CR Ins. Materials]",
    "Doors",
    "Ref. Inst. Materials",
    "CCP",
    "Display CCP",
    "Floor Heater",
    "Cabinets",
    "Any Special",
]

BASE_COLUMNS = [
    "S_No",
    "Date_Time",
    "Job_Ref",
    "LPO_Ref",
    "Customer",
    "Project_Name",
    "Region",
    "Location",
    "Payment_Terms",
    *ENGINEERING_COLS,
    *DELIVERY_COLS,
    "Work_Status",
    "Remarks",
    "Material_Status",
    "Overall_Progress",
    "Priority",
    "Status",
    "Engineering_Pct",
    "Delivery_Pct",
    "Execution_Pct",
]

STATUS_ORDER = ["Completed", "In Progress", "On Hold", "Not Started", "Cancelled"]
MATERIAL_ORDER = ["Delivered", "Partially Delivered", "Ordered", "Not Ordered"]
PRIORITY_ORDER = ["High", "Medium", "Low"]

STATUS_COLORS = {
    "Completed": "#16a34a",
    "In Progress": "#0284c7",
    "On Hold": "#f97316",
    "Not Started": "#64748b",
    "Cancelled": "#dc2626",
}
PRIORITY_COLORS = {"High": "#dc2626", "Medium": "#d97706", "Low": "#16a34a"}
MATERIAL_COLORS = {
    "Delivered": "#16a34a",
    "Partially Delivered": "#ca8a04",
    "Ordered": "#2563eb",
    "Not Ordered": "#64748b",
}

# -----------------------------------------------------------------------------
# STYLE - LIGHT, READABLE, MANAGER-FRIENDLY
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    :root {
        --bg: #f5f7fb;
        --card: #ffffff;
        --text: #0f172a;
        --muted: #64748b;
        --border: #e2e8f0;
        --blue: #2563eb;
        --green: #16a34a;
        --orange: #f97316;
        --red: #dc2626;
        --shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
    }

    .stApp {
        background: var(--bg);
        color: var(--text);
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    h1, h2, h3, p, label, span, div {
        color: var(--text);
    }

    [data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid var(--border);
    }

    [data-testid="stSidebar"] * {
        color: var(--text) !important;
    }

    [data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"],
    [data-testid="stSidebar"] .stTextInput input {
        background: #f8fafc !important;
        border-color: #cbd5e1 !important;
        color: #0f172a !important;
    }

    .hero {
        background: linear-gradient(135deg, #ffffff 0%, #eef4ff 100%);
        border: 1px solid var(--border);
        border-radius: 26px;
        padding: 28px 32px;
        box-shadow: var(--shadow);
        margin-bottom: 18px;
    }

    .hero-title {
        font-size: 42px;
        font-weight: 850;
        line-height: 1.1;
        letter-spacing: -0.04em;
        margin: 0;
        color: #0f172a;
    }

    .hero-subtitle {
        margin-top: 10px;
        font-size: 16px;
        color: #475569;
        font-weight: 500;
    }

    .source-row {
        margin-top: 18px;
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: center;
    }

    .pill {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 8px 12px;
        border-radius: 999px;
        background: #f8fafc;
        border: 1px solid #dbeafe;
        color: #334155;
        font-size: 13px;
        font-weight: 650;
    }

    .pill.good { background: #ecfdf5; border-color: #bbf7d0; color: #166534; }
    .pill.blue { background: #eff6ff; border-color: #bfdbfe; color: #1d4ed8; }
    .pill.orange { background: #fff7ed; border-color: #fed7aa; color: #c2410c; }

    .kpi-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 18px 18px 16px;
        box-shadow: var(--shadow);
        min-height: 126px;
    }

    .kpi-label {
        color: #64748b;
        font-size: 12px;
        font-weight: 750;
        text-transform: uppercase;
        letter-spacing: .06em;
        margin-bottom: 10px;
    }

    .kpi-value {
        font-size: 34px;
        font-weight: 850;
        color: #0f172a;
        letter-spacing: -0.03em;
    }

    .kpi-note {
        margin-top: 8px;
        color: #64748b;
        font-size: 12px;
        font-weight: 500;
    }

    .section-header {
        margin: 26px 0 12px;
        font-size: 24px;
        font-weight: 850;
        letter-spacing: -0.03em;
        color: #0f172a;
    }

    .section-caption {
        color: #64748b;
        margin: -6px 0 14px;
        font-size: 14px;
    }

    .phase-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 20px;
        box-shadow: var(--shadow);
    }

    .phase-title {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 12px;
        font-weight: 850;
        font-size: 16px;
        color: #0f172a;
    }

    .phase-percent {
        font-size: 30px;
        letter-spacing: -0.04em;
    }

    .phase-sub {
        color: #64748b;
        font-size: 13px;
        margin-top: 12px;
        min-height: 34px;
    }

    .bar-bg {
        height: 12px;
        background: #e2e8f0;
        border-radius: 999px;
        overflow: hidden;
        margin-top: 14px;
    }

    .bar-fill {
        height: 12px;
        border-radius: 999px;
    }

    .chart-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 14px 14px 4px;
        box-shadow: var(--shadow);
        margin-bottom: 18px;
    }

    .chart-title {
        font-weight: 800;
        font-size: 16px;
        margin: 4px 6px 0;
        color: #0f172a;
    }

    .table-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 18px;
        box-shadow: var(--shadow);
        margin-top: 14px;
    }

    div[data-testid="stDataFrame"] {
        border: 1px solid var(--border);
        border-radius: 14px;
        overflow: hidden;
    }

    .stAlert {
        border-radius: 16px;
    }

    button[kind="secondary"], .stDownloadButton button {
        border-radius: 12px !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 1px solid var(--border);
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 12px 12px 0 0;
        padding: 10px 16px;
        background: #ffffff;
        border: 1px solid var(--border);
        border-bottom: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# DATA SOURCE
# -----------------------------------------------------------------------------

def make_onedrive_download_url(url: str) -> str:
    """Works for many OneDrive/SharePoint sharing links."""
    url = url.strip()
    if not url:
        return url
    if "download=1" in url:
        return url
    joiner = "&" if "?" in url else "?"
    return f"{url}{joiner}download=1"


@st.cache_data(ttl=REFRESH_SECONDS)
def load_excel_from_direct_url(url: str) -> bytes:
    direct_url = make_onedrive_download_url(url)
    response = requests.get(direct_url, timeout=60, allow_redirects=True)
    response.raise_for_status()
    content_type = response.headers.get("content-type", "").lower()
    if "text/html" in content_type and len(response.content) < 500_000:
        raise RuntimeError(
            "The OneDrive URL returned a web page, not the Excel file. "
            "Use an Anyone-with-link share URL, direct download link, or Microsoft Graph mode."
        )
    return response.content


@st.cache_data(ttl=REFRESH_SECONDS)
def get_graph_token() -> str:
    if msal is None:
        raise RuntimeError("msal is not installed. Add msal to requirements.txt")

    tenant_id = _secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    client_id = _secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    client_secret = _secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        raise RuntimeError("Missing GRAPH_TENANT_ID, GRAPH_CLIENT_ID, or GRAPH_CLIENT_SECRET secrets.")

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Could not get Microsoft Graph token: {result}")
    return result["access_token"]


@st.cache_data(ttl=REFRESH_SECONDS)
def load_excel_from_graph() -> bytes:
    user_id = _secret("GRAPH_USER_ID")
    file_path = _secret("ONEDRIVE_FILE_PATH")
    if not user_id or not file_path:
        raise RuntimeError("Missing GRAPH_USER_ID or ONEDRIVE_FILE_PATH secrets.")

    token = get_graph_token()
    if not file_path.startswith("/"):
        file_path = "/" + file_path

    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:{file_path}:/content"
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    response.raise_for_status()
    return response.content


def get_excel_bytes(uploaded_file) -> tuple[bytes, str]:
    if uploaded_file is not None:
        return uploaded_file.read(), uploaded_file.name

    excel_url = _secret("EXCEL_FILE_URL")
    if excel_url:
        return load_excel_from_direct_url(excel_url), "OneDrive live Excel"

    if _secret("GRAPH_USER_ID") and _secret("ONEDRIVE_FILE_PATH"):
        return load_excel_from_graph(), "Microsoft Graph OneDrive"

    local_path = Path(LOCAL_EXCEL_FILE)
    if local_path.exists():
        return local_path.read_bytes(), LOCAL_EXCEL_FILE

    st.error(
        "No Excel source found. Upload the workbook, set EXCEL_FILE_URL, "
        "or configure Microsoft Graph secrets."
    )
    st.stop()

# -----------------------------------------------------------------------------
# PARSING
# -----------------------------------------------------------------------------

def _clean_text(value, default: str = "") -> str:
    if pd.isna(value):
        return default
    text = str(value).strip()
    return text if text else default


def _clean_status(value: str) -> str:
    value = _clean_text(value, "Not Started")
    mapping = {
        "completed": "Completed",
        "complete": "Completed",
        "in progress": "In Progress",
        "on progress": "In Progress",
        "progress": "In Progress",
        "on hold": "On Hold",
        "hold": "On Hold",
        "cancelled": "Cancelled",
        "canceled": "Cancelled",
        "not started": "Not Started",
        "notstart": "Not Started",
    }
    cleaned = mapping.get(value.lower(), value)
    return cleaned if cleaned in STATUS_ORDER else "Not Started"


def _clean_priority(value: str) -> str:
    value = _clean_text(value, "Unspecified")
    text = value.lower().strip()
    if text in {"high", "h"}:
        return "High"
    if text in {"medium", "med", "m"}:
        return "Medium"
    if text in {"low", "l"}:
        return "Low"
    return "Unspecified"


def _to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


@st.cache_data(ttl=REFRESH_SECONDS)
def parse_excel(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME, header=None, engine="openpyxl")

    # The workbook has title/group headers in row 0 and real column headers in row 1.
    # We read by header name where possible so the dashboard does not break if a
    # column is inserted in the Excel file. Fixed positions are used as a fallback.
    header = raw.iloc[1].fillna("").astype(str).str.strip().tolist()
    header_norm = [h.lower().replace(" ", "").replace("_", "") for h in header]

    def find_col(names: list[str], fallback: int) -> int:
        wanted = [n.lower().replace(" ", "").replace("_", "") for n in names]
        for w in wanted:
            if w in header_norm:
                return header_norm.index(w)
        return fallback

    idx = {
        "S_No": find_col(["S_No", "S No"], 0),
        "Date_Time": find_col(["Date_Time", "Date Time"], 1),
        "Job_Ref": find_col(["Job_Ref", "Job Ref"], 2),
        "LPO_Ref": find_col(["LPO_Ref", "LPO Ref"], 3),
        "Customer": find_col(["Customer"], 4),
        "Project_Name": find_col(["Project_Name", "Project Name"], 5),
        "Region": find_col(["Region"], 6),
        "Location": find_col(["Location"], 7),
        "Payment_Terms": find_col(["Payment Terms", "Payment_Terms"], 8),
        "Work_Status": 25,
        "Remarks": 26,
        "Material_Status": 27,
        "Overall_Progress": find_col(["Overall Progress %", "Overall Progress", "Progress"], 28),
        "Priority": find_col(["Priority"], 29),
        "Status": find_col(["Status"], 30),
        "Engineering_Pct": find_col(["Engineering %", "Engineering%"], 31),
        "Delivery_Pct": find_col(["Delivery%", "Delivery %"], 32),
        "Execution_Pct": find_col(["Execution %", "Execution%"], 33),
    }

    rows = raw.iloc[2:].copy()
    data = pd.DataFrame()
    for col, i in idx.items():
        data[col] = rows.iloc[:, i] if i < rows.shape[1] else None

    for name, pos in zip(ENGINEERING_COLS, range(9, 14)):
        data[name] = rows.iloc[:, pos] if pos < rows.shape[1] else ""
    for name, pos in zip(DELIVERY_COLS, range(14, 25)):
        data[name] = rows.iloc[:, pos] if pos < rows.shape[1] else ""

    data = data.dropna(how="all")
    data = data[~(data["S_No"].isna() & data["Project_Name"].isna())]

    for col in ["Job_Ref", "LPO_Ref", "Customer", "Project_Name", "Region", "Location"]:
        default = "Unknown" if col in ["Customer", "Region"] else ""
        data[col] = data[col].apply(lambda x: _clean_text(x, default))

    data["Material_Status"] = data["Material_Status"].apply(lambda x: _clean_text(x, "Not Ordered"))
    # Keep material chart readable by grouping unexpected values.
    data.loc[~data["Material_Status"].isin(MATERIAL_ORDER), "Material_Status"] = "Not Ordered"
    data["Status"] = data["Status"].apply(_clean_status)
    data["Priority"] = data["Priority"].apply(_clean_priority)
    data["Work_Status"] = data["Work_Status"].apply(lambda x: _clean_text(x, ""))

    for col in ["Overall_Progress", "Engineering_Pct", "Delivery_Pct", "Execution_Pct"]:
        data[col] = _to_number(data[col]).clip(lower=0, upper=100)

    for group_name, cols in [("Eng", ENGINEERING_COLS), ("Del", DELIVERY_COLS)]:
        status_block = data[cols].fillna("").astype(str).apply(lambda s: s.str.strip().str.upper())
        data[f"{group_name}_Done"] = (status_block == "DONE").sum(axis=1)
        data[f"{group_name}_Partial"] = (status_block == "PART.DONE").sum(axis=1)
        data[f"{group_name}_NA"] = (status_block == "N/A").sum(axis=1)

    data["S_No"] = data["S_No"].fillna("").astype(str).str.replace(".0", "", regex=False)
    return data.reset_index(drop=True)

# -----------------------------------------------------------------------------
# FILTERS
# -----------------------------------------------------------------------------

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar:
        st.markdown("### Filters")
        st.caption("Use these filters to focus the whole dashboard.")
        search = st.text_input("Search", placeholder="Project, customer, location, job ref...")

        quick = st.radio(
            "Quick views",
            ["All", "Completed", "In Progress", "On Hold", "Not Started"],
            horizontal=False,
        )

        st.divider()
        status = st.multiselect("Status", sorted(df["Status"].dropna().unique()), default=[])
        region = st.multiselect("Region", sorted(df["Region"].dropna().unique()), default=[])
        customer = st.multiselect("Customer", sorted(df["Customer"].dropna().unique()), default=[])
        material = st.multiselect("Material", sorted(df["Material_Status"].dropna().unique()), default=[])
        priority = st.multiselect("Priority", sorted(df["Priority"].dropna().unique()), default=[])

        st.divider()
        st.caption(f"Data refresh: every {REFRESH_SECONDS}s")

    out = df.copy()
    if quick != "All":
        out = out[out["Status"] == quick]
    if search:
        text = search.lower().strip()
        mask = (
            out["Project_Name"].str.lower().str.contains(text, na=False)
            | out["Customer"].str.lower().str.contains(text, na=False)
            | out["Location"].str.lower().str.contains(text, na=False)
            | out["Job_Ref"].astype(str).str.lower().str.contains(text, na=False)
            | out["LPO_Ref"].astype(str).str.lower().str.contains(text, na=False)
        )
        out = out[mask]
    if status:
        out = out[out["Status"].isin(status)]
    if region:
        out = out[out["Region"].isin(region)]
    if customer:
        out = out[out["Customer"].isin(customer)]
    if material:
        out = out[out["Material_Status"].isin(material)]
    if priority:
        out = out[out["Priority"].isin(priority)]
    return out

# -----------------------------------------------------------------------------
# UI HELPERS
# -----------------------------------------------------------------------------

def pct(value: float) -> str:
    return f"{value:.1f}%"


def html_escape(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def count_by_order(df: pd.DataFrame, col: str, order: Iterable[str]) -> pd.DataFrame:
    counts = df[col].value_counts().reindex(order, fill_value=0).reset_index()
    counts.columns = [col, "Count"]
    extra = df[~df[col].isin(order)][col].value_counts().reset_index()
    extra.columns = [col, "Count"]
    return pd.concat([counts, extra], ignore_index=True)


def render_kpi(label: str, value: str, note: str = "", accent: str = "#2563eb") -> None:
    st.markdown(
        f"""
        <div class="kpi-card" style="border-top: 4px solid {accent};">
            <div class="kpi-label">{html_escape(label)}</div>
            <div class="kpi-value">{html_escape(value)}</div>
            <div class="kpi-note">{html_escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_phase_card(title: str, value: float, subtitle: str, color: str) -> None:
    st.markdown(
        f"""
        <div class="phase-card">
            <div class="phase-title">
                <span>{html_escape(title)}</span>
                <span class="phase-percent" style="color:{color};">{value:.1f}%</span>
            </div>
            <div class="bar-bg"><div class="bar-fill" style="width:{max(0, min(value, 100)):.1f}%; background:{color};"></div></div>
            <div class="phase-sub">{html_escape(subtitle)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def add_chart_card_title(title: str) -> None:
    st.markdown(f'<div class="chart-card"><div class="chart-title">{html_escape(title)}</div>', unsafe_allow_html=True)


def close_chart_card() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


def style_fig(fig: go.Figure, height: int = 360) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=24, r=24, t=30, b=24),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#0f172a", family="Inter, Segoe UI, sans-serif", size=13),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#e2e8f0", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="#e2e8f0", zeroline=False)
    return fig


def readable_count_chart(
    df: pd.DataFrame,
    label_col: str,
    count_col: str = "Count",
    colors: dict[str, str] | None = None,
    height: int = 300,
):
    """Simple horizontal count chart with count + percentage labels."""
    chart_df = df.copy()
    chart_df = chart_df[chart_df[count_col] > 0].copy()
    if chart_df.empty:
        st.info("No data for this view.")
        return
    total = chart_df[count_col].sum()
    chart_df["Percent"] = chart_df[count_col] / total * 100
    chart_df["Text"] = chart_df.apply(lambda r: f"{int(r[count_col])}  ({r['Percent']:.0f}%)", axis=1)
    chart_df = chart_df.sort_values(count_col, ascending=True)

    color_values = [colors.get(str(v), "#64748b") if colors else "#2563eb" for v in chart_df[label_col]]
    fig = go.Figure(
        go.Bar(
            x=chart_df[count_col],
            y=chart_df[label_col],
            orientation="h",
            text=chart_df["Text"],
            textposition="outside",
            marker_color=color_values,
            hovertemplate="%{y}: %{x}<extra></extra>",
        )
    )
    max_count = max(float(chart_df[count_col].max()), 1)
    fig = style_fig(fig, height=height)
    fig.update_layout(showlegend=False, margin=dict(l=20, r=70, t=10, b=20))
    fig.update_xaxes(range=[0, max_count * 1.22], title="Projects", dtick=1 if max_count <= 10 else None)
    fig.update_yaxes(title="")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def progress_bucket_chart(bucket_df: pd.DataFrame, height: int = 300):
    colors = {
        "0-25%": "#dc2626",
        "26-50%": "#f97316",
        "51-75%": "#f59e0b",
        "76-99%": "#2563eb",
        "100%": "#16a34a",
    }
    readable_count_chart(bucket_df, "Progress Bucket", "Count", colors, height)


def bar_chart(df: pd.DataFrame, x: str, y: str, orientation: str = "v", color: str = "#2563eb", height: int = 360):
    fig = px.bar(df, x=x, y=y, orientation=orientation, text_auto=True)
    fig.update_traces(marker_color=color, textfont_color="#0f172a")
    fig = style_fig(fig, height=height)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
st_autorefresh(interval=REFRESH_SECONDS * 1000, key="refresh")

# Dashboard reads directly from the configured Excel source.
# Manual upload section removed for manager-facing view.
uploaded = None

try:
    file_bytes, source_name = get_excel_bytes(uploaded)
    df = parse_excel(file_bytes)
except Exception as exc:
    st.error(f"Could not load dashboard data: {exc}")
    st.stop()

filtered = apply_filters(df)

# HERO
avg_progress = filtered["Overall_Progress"].mean() if len(filtered) else 0
last_refresh_text = pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")
source_badge = "good" if "OneDrive" in source_name or "Graph" in source_name else "blue"

st.markdown(
    f"""
    <div class="hero">
        <div class="hero-title">Project Tracking Dashboard</div>
        <div class="hero-subtitle">Cold Rooms, Cabinets & Refrigeration Projects — clean live view from Excel</div>
        <div class="source-row">
            <span class="pill {source_badge}">● Current source: {html_escape(source_name)}</span>
            <span class="pill blue">↻ Refresh: {REFRESH_SECONDS}s</span>
            <span class="pill orange">Showing: {len(filtered)} / {len(df)} projects</span>
            <span class="pill">Updated: {last_refresh_text}</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# KPIs
kpi_cols = st.columns(6)
with kpi_cols[0]:
    render_kpi("Total Projects", str(len(filtered)), "Projects in current view", "#2563eb")
with kpi_cols[1]:
    render_kpi("Completed", str(int((filtered["Status"] == "Completed").sum())), "Ready / closed", "#16a34a")
with kpi_cols[2]:
    render_kpi("In Progress", str(int((filtered["Status"] == "In Progress").sum())), "Currently active", "#0284c7")
with kpi_cols[3]:
    render_kpi("On Hold", str(int((filtered["Status"] == "On Hold").sum())), "Needs attention", "#f97316")
with kpi_cols[4]:
    render_kpi("Not Started", str(int((filtered["Status"] == "Not Started").sum())), "Yet to begin", "#64748b")
with kpi_cols[5]:
    render_kpi("Avg Progress", pct(avg_progress), "Across current view", "#7c3aed")

# PHASE CARDS
st.markdown('<div class="section-header">Phase Progress Overview</div>', unsafe_allow_html=True)
st.markdown('<div class="section-caption">Average completion by Engineering, Delivery and Execution phases.</div>', unsafe_allow_html=True)
phase_cols = st.columns(3)
with phase_cols[0]:
    render_phase_card(
        "Engineering",
        filtered["Engineering_Pct"].mean() if len(filtered) else 0,
        "Design / Submittal / Drawing / ELS / BOM",
        "#2563eb",
    )
with phase_cols[1]:
    render_phase_card(
        "Delivery",
        filtered["Delivery_Pct"].mean() if len(filtered) else 0,
        f"Material delivery across {len(DELIVERY_COLS)} tracked items",
        "#7c3aed",
    )
with phase_cols[2]:
    render_phase_card(
        "Execution",
        filtered["Execution_Pct"].mean() if len(filtered) else 0,
        "Site installation, commissioning and handover progress",
        "#f97316",
    )

# TABS FOR READABILITY
tab_overview, tab_progress, tab_details = st.tabs(["Overview", "Progress Analysis", "Project Details"])

with tab_overview:
    st.markdown('<div class="section-header">Portfolio Overview</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        add_chart_card_title("Project Status")
        st.caption("How many projects are completed, active, on hold, or not started.")
        readable_count_chart(count_by_order(filtered, "Status", STATUS_ORDER), "Status", "Count", STATUS_COLORS, height=300)
        close_chart_card()
    with c2:
        add_chart_card_title("Material Status")
        st.caption("Material readiness by project count.")
        readable_count_chart(count_by_order(filtered, "Material_Status", MATERIAL_ORDER), "Material_Status", "Count", MATERIAL_COLORS, height=300)
        close_chart_card()
    with c3:
        add_chart_card_title("Priority")
        st.caption("Only High, Medium and Low are shown; unknown values are grouped.")
        priority_counts = count_by_order(filtered, "Priority", ["High", "Medium", "Low", "Unspecified"])
        priority_colors = {**PRIORITY_COLORS, "Unspecified": "#94a3b8"}
        readable_count_chart(priority_counts, "Priority", "Count", priority_colors, height=300)
        close_chart_card()

    c4, c5 = st.columns(2)
    with c4:
        add_chart_card_title("Region Breakdown")
        region_df = filtered["Region"].value_counts().reset_index()
        region_df.columns = ["Region", "Count"]
        readable_count_chart(region_df, "Region", "Count", height=360)
        close_chart_card()
    with c5:
        add_chart_card_title("Top Customers")
        cust_df = filtered["Customer"].value_counts().head(10).reset_index()
        cust_df.columns = ["Customer", "Count"]
        bar_chart(cust_df.sort_values("Count"), "Count", "Customer", orientation="h", color="#0f766e", height=380)
        close_chart_card()

with tab_progress:
    st.markdown('<div class="section-header">Progress Analysis</div>', unsafe_allow_html=True)
    c6, c7 = st.columns([7, 5])
    with c6:
        add_chart_card_title("Delivery Items Status")
        delivery_done = []
        delivery_partial = []
        for col in DELIVERY_COLS:
            vals = filtered[col].fillna("").astype(str).str.strip().str.upper()
            delivery_done.append(int((vals == "DONE").sum()))
            delivery_partial.append(int((vals == "PART.DONE").sum()))
        delivery_df = pd.DataFrame({"Item": DELIVERY_COLS, "Done": delivery_done, "Partial": delivery_partial})
        fig = go.Figure()
        fig.add_trace(go.Bar(x=delivery_df["Item"], y=delivery_df["Done"], name="Done", marker_color="#16a34a"))
        fig.add_trace(go.Bar(x=delivery_df["Item"], y=delivery_df["Partial"], name="Partial", marker_color="#f59e0b"))
        fig.update_layout(barmode="stack")
        fig = style_fig(fig, height=420)
        fig.update_layout(margin=dict(l=20, r=20, t=20, b=110))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        close_chart_card()

    with c7:
        add_chart_card_title("Overall Progress Buckets")
        bins = [-0.1, 25, 50, 75, 99.999, 100]
        labels = ["0-25%", "26-50%", "51-75%", "76-99%", "100%"]
        bucket = pd.cut(filtered["Overall_Progress"], bins=bins, labels=labels)
        bucket_df = bucket.value_counts().reindex(labels, fill_value=0).reset_index()
        bucket_df.columns = ["Progress Bucket", "Count"]
        bucket_colors = {
            "0-25%": "#dc2626",
            "26-50%": "#f97316",
            "51-75%": "#f59e0b",
            "76-99%": "#2563eb",
            "100%": "#16a34a",
        }
        progress_bucket_chart(bucket_df, height=320)
        close_chart_card()

    add_chart_card_title("Phase Progress by Project")
    phase_df = filtered.sort_values("Overall_Progress", ascending=False).head(25)
    fig = go.Figure()
    fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Engineering_Pct"], name="Engineering", orientation="h", marker_color="#2563eb"))
    fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Delivery_Pct"], name="Delivery", orientation="h", marker_color="#7c3aed"))
    fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Execution_Pct"], name="Execution", orientation="h", marker_color="#f97316"))
    fig.update_layout(barmode="group", xaxis_title="Completion %", yaxis_title="", yaxis={"autorange": "reversed"})
    fig = style_fig(fig, height=720)
    fig.update_layout(margin=dict(l=20, r=20, t=30, b=20), xaxis=dict(range=[0, 100]))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    close_chart_card()

with tab_details:
    st.markdown('<div class="section-header">Project Details</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-caption">Sortable project list. Use the sidebar filters to reduce this table.</div>', unsafe_allow_html=True)

    show_cols = [
        "S_No",
        "Customer",
        "Project_Name",
        "Region",
        "Location",
        "Status",
        "Engineering_Pct",
        "Delivery_Pct",
        "Execution_Pct",
        "Overall_Progress",
        "Material_Status",
        "Priority",
        "Work_Status",
    ]
    table_df = filtered[show_cols].sort_values("Overall_Progress", ascending=False).copy()
    table_df = table_df.rename(
        columns={
            "S_No": "#",
            "Project_Name": "Project",
            "Engineering_Pct": "Eng %",
            "Delivery_Pct": "Delivery %",
            "Execution_Pct": "Exec %",
            "Overall_Progress": "Overall %",
            "Material_Status": "Material",
            "Work_Status": "Work Status",
        }
    )

    st.dataframe(
        table_df,
        use_container_width=True,
        hide_index=True,
        height=650,
        column_config={
            "Eng %": st.column_config.ProgressColumn("Eng %", min_value=0, max_value=100, format="%.0f%%"),
            "Delivery %": st.column_config.ProgressColumn("Delivery %", min_value=0, max_value=100, format="%.0f%%"),
            "Exec %": st.column_config.ProgressColumn("Exec %", min_value=0, max_value=100, format="%.0f%%"),
            "Overall %": st.column_config.ProgressColumn("Overall %", min_value=0, max_value=100, format="%.0f%%"),
        },
    )

    csv = table_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered project list as CSV",
        data=csv,
        file_name="filtered_project_dashboard.csv",
        mime="text/csv",
    )
