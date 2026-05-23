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
    page_icon=":bar_chart:",
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
    "Completed": "#10b981",
    "In Progress": "#3b82f6",
    "On Hold": "#f59e0b",
    "Not Started": "#6b7280",
    "Cancelled": "#ef4444",
}
PRIORITY_COLORS = {"High": "#ef4444", "Medium": "#f59e0b", "Low": "#10b981"}
MATERIAL_COLORS = {
    "Delivered": "#10b981",
    "Partially Delivered": "#f59e0b",
    "Ordered": "#3b82f6",
    "Not Ordered": "#6b7280",
}

# -----------------------------------------------------------------------------
# STYLE - PREMIUM DARK INDUSTRIAL DASHBOARD
# -----------------------------------------------------------------------------
st.markdown(
    """
    <link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&family=Outfit:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
    :root {
        --bg:        #0a0c10;
        --surface:   #111318;
        --surface2:  #181c24;
        --border:    rgba(255,255,255,0.07);
        --border2:   rgba(255,255,255,0.12);
        --text:      #f0f2f7;
        --muted:     #6b7280;
        --dim:       #9ca3af;

        --blue:      #3b82f6;
        --blue-glow: rgba(59,130,246,0.18);
        --green:     #10b981;
        --green-glow:rgba(16,185,129,0.18);
        --amber:     #f59e0b;
        --amber-glow:rgba(245,158,11,0.18);
        --red:       #ef4444;
        --red-glow:  rgba(239,68,68,0.18);
        --purple:    #a78bfa;
        --purple-glow:rgba(167,139,250,0.18);
        --teal:      #2dd4bf;
        --teal-glow: rgba(45,212,191,0.18);

        --font-display: 'Syne', sans-serif;
        --font-body:    'Outfit', sans-serif;
        --font-mono:    'JetBrains Mono', monospace;

        --radius-sm: 10px;
        --radius:    16px;
        --radius-lg: 22px;
        --radius-xl: 28px;

        --shadow-card: 0 0 0 1px rgba(255,255,255,0.05), 0 4px 24px rgba(0,0,0,0.4);
        --shadow-glow: 0 0 40px rgba(59,130,246,0.12);
    }

    /* --- GLOBAL --- */
    .stApp {
        background: var(--bg);
        color: var(--text);
        font-family: var(--font-body);
    }

    .block-container {
        padding-top: 1.6rem !important;
        padding-bottom: 4rem !important;
        max-width: 1560px !important;
    }

    h1,h2,h3,h4,h5,h6 { font-family: var(--font-display); color: var(--text); }
    p, label, span, li { font-family: var(--font-body); color: var(--text); }

    /* --- SIDEBAR --- */
    [data-testid="stSidebar"] {
        background: var(--surface) !important;
        border-right: 1px solid var(--border2) !important;
        padding-top: 1rem;
    }

    [data-testid="stSidebar"] * {
        color: var(--text) !important;
        font-family: var(--font-body) !important;
    }

    [data-testid="stSidebar"] h3 {
        font-family: var(--font-display) !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        letter-spacing: 0.02em !important;
        color: var(--text) !important;
        padding-bottom: 4px;
    }

    [data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] > div,
    [data-testid="stSidebar"] .stTextInput input,
    [data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div {
        background: var(--surface2) !important;
        border-color: var(--border2) !important;
        border-radius: var(--radius-sm) !important;
        color: var(--text) !important;
    }

    [data-testid="stSidebar"] .stRadio label {
        background: var(--surface2);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 6px 12px;
        margin: 2px 0;
        transition: all .2s;
        display: block;
    }

    [data-testid="stSidebar"] .stRadio label:hover {
        border-color: var(--blue);
        background: var(--blue-glow);
    }

    /* --- HERO --- */
    .hero {
        position: relative;
        overflow: hidden;
        background: var(--surface);
        border: 1px solid var(--border2);
        border-radius: var(--radius-xl);
        padding: 36px 40px;
        box-shadow: var(--shadow-card);
        margin-bottom: 24px;
    }

    .hero::before {
        content: '';
        position: absolute;
        top: -60px; right: -60px;
        width: 320px; height: 320px;
        background: radial-gradient(circle, rgba(59,130,246,0.14) 0%, transparent 70%);
        pointer-events: none;
    }

    .hero::after {
        content: '';
        position: absolute;
        bottom: -80px; left: 30%;
        width: 400px; height: 200px;
        background: radial-gradient(ellipse, rgba(167,139,250,0.07) 0%, transparent 70%);
        pointer-events: none;
    }

    .hero-eyebrow {
        font-family: var(--font-mono);
        font-size: 11px;
        font-weight: 600;
        letter-spacing: .2em;
        text-transform: uppercase;
        color: var(--blue);
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .hero-eyebrow::before {
        content: '';
        display: inline-block;
        width: 24px; height: 2px;
        background: var(--blue);
        border-radius: 2px;
    }

    .hero-title {
        font-family: var(--font-display);
        font-size: 48px;
        font-weight: 800;
        line-height: 1.05;
        letter-spacing: -0.03em;
        margin: 0 0 8px;
        color: var(--text);
    }

    .hero-title span {
        background: linear-gradient(135deg, #60a5fa, #a78bfa);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    .hero-subtitle {
        font-size: 15px;
        color: var(--muted);
        font-weight: 400;
        margin-top: 4px;
        letter-spacing: 0.01em;
    }

    .source-row {
        margin-top: 22px;
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
    }

    .pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 14px;
        border-radius: 999px;
        background: var(--surface2);
        border: 1px solid var(--border2);
        color: var(--dim);
        font-size: 12px;
        font-weight: 500;
        font-family: var(--font-mono);
        letter-spacing: .02em;
    }

    .pill.good  { background: var(--green-glow);  border-color: rgba(16,185,129,.3);  color: #34d399; }
    .pill.blue  { background: var(--blue-glow);   border-color: rgba(59,130,246,.3);  color: #60a5fa; }
    .pill.orange{ background: var(--amber-glow);  border-color: rgba(245,158,11,.3);  color: #fbbf24; }

    /* --- KPI CARDS --- */
    .kpi-card {
        position: relative;
        overflow: hidden;
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius-lg);
        padding: 22px 20px 18px;
        box-shadow: var(--shadow-card);
        min-height: 130px;
        transition: transform .2s, box-shadow .2s, border-color .2s;
    }

    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: var(--shadow-card), 0 12px 40px rgba(0,0,0,0.5);
        border-color: var(--border2);
    }

    .kpi-card::after {
        content: '';
        position: absolute;
        bottom: 0; left: 0; right: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.06), transparent);
    }

    .kpi-accent {
        position: absolute;
        top: 0; left: 0;
        width: 4px; height: 100%;
        border-radius: var(--radius-lg) 0 0 var(--radius-lg);
    }

    .kpi-label {
        font-family: var(--font-mono);
        color: var(--muted);
        font-size: 10px;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: .12em;
        margin-bottom: 10px;
        padding-left: 12px;
    }

    .kpi-value {
        font-family: var(--font-display);
        font-size: 38px;
        font-weight: 800;
        color: var(--text);
        letter-spacing: -0.03em;
        line-height: 1;
        padding-left: 12px;
    }

    .kpi-note {
        margin-top: 8px;
        color: var(--muted);
        font-size: 12px;
        font-weight: 400;
        padding-left: 12px;
    }

    /* --- SECTION HEADERS --- */
    .section-header {
        font-family: var(--font-display);
        margin: 32px 0 6px;
        font-size: 22px;
        font-weight: 700;
        letter-spacing: -0.02em;
        color: var(--text);
        display: flex;
        align-items: center;
        gap: 10px;
    }

    .section-header::before {
        content: '';
        display: inline-block;
        width: 4px; height: 22px;
        background: linear-gradient(180deg, var(--blue), var(--purple));
        border-radius: 2px;
        flex-shrink: 0;
    }

    .section-caption {
        color: var(--muted);
        margin: 0 0 18px;
        font-size: 13.5px;
        font-weight: 400;
        padding-left: 14px;
    }

    /* --- PHASE CARDS --- */
    .phase-card {
        position: relative;
        overflow: hidden;
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius-lg);
        padding: 22px 22px 20px;
        box-shadow: var(--shadow-card);
        transition: transform .2s, border-color .2s;
    }

    .phase-card:hover {
        transform: translateY(-2px);
        border-color: var(--border2);
    }

    .phase-card::before {
        content: '';
        position: absolute;
        top: -40px; right: -40px;
        width: 120px; height: 120px;
        border-radius: 50%;
        opacity: 0.08;
        pointer-events: none;
    }

    .phase-title {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 12px;
    }

    .phase-name {
        font-family: var(--font-display);
        font-weight: 700;
        font-size: 15px;
        color: var(--text);
        letter-spacing: .01em;
    }

    .phase-percent {
        font-family: var(--font-display);
        font-size: 36px;
        font-weight: 800;
        letter-spacing: -0.04em;
        line-height: 1;
    }

    .phase-sub {
        color: var(--muted);
        font-size: 12.5px;
        margin-top: 14px;
        min-height: 32px;
        line-height: 1.5;
    }

    .bar-bg {
        height: 6px;
        background: rgba(255,255,255,0.07);
        border-radius: 999px;
        overflow: hidden;
        margin-top: 16px;
    }

    .bar-fill {
        height: 6px;
        border-radius: 999px;
        position: relative;
    }

    .bar-fill::after {
        content: '';
        position: absolute;
        right: 0; top: 0;
        width: 6px; height: 6px;
        border-radius: 50%;
        background: inherit;
        filter: brightness(1.5);
    }

    /* --- CHART CARDS --- */
    .chart-card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius-lg);
        padding: 20px 18px 8px;
        box-shadow: var(--shadow-card);
        margin-bottom: 20px;
    }

    .chart-title {
        font-family: var(--font-display);
        font-weight: 700;
        font-size: 15px;
        margin: 0 0 4px 2px;
        color: var(--text);
        letter-spacing: .01em;
    }

    /* --- DATAFRAME --- */
    div[data-testid="stDataFrame"] {
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
        overflow: hidden !important;
        background: var(--surface);
    }

    div[data-testid="stDataFrame"] table {
        background: var(--surface) !important;
    }

    /* --- TABS --- */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: var(--surface) !important;
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 4px;
        width: fit-content;
        margin-bottom: 20px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: var(--radius-sm) !important;
        padding: 10px 20px !important;
        background: transparent !important;
        border: none !important;
        font-family: var(--font-body) !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        color: var(--muted) !important;
        transition: all .2s !important;
    }

    .stTabs [aria-selected="true"] {
        background: var(--surface2) !important;
        color: var(--text) !important;
        border: 1px solid var(--border2) !important;
    }

    .stTabs [data-baseweb="tab-highlight"] {
        display: none !important;
    }

    .stTabs [data-baseweb="tab-border"] {
        display: none !important;
    }

    /* --- ALERTS --- */
    .stAlert {
        background: var(--surface2) !important;
        border: 1px solid var(--border2) !important;
        border-radius: var(--radius) !important;
        color: var(--text) !important;
    }

    /* --- DOWNLOAD BUTTON --- */
    .stDownloadButton button {
        background: var(--surface2) !important;
        border: 1px solid var(--border2) !important;
        border-radius: var(--radius-sm) !important;
        color: var(--text) !important;
        font-family: var(--font-body) !important;
        font-weight: 500 !important;
        transition: all .2s !important;
    }

    .stDownloadButton button:hover {
        border-color: var(--blue) !important;
        background: var(--blue-glow) !important;
        color: #60a5fa !important;
    }

    /* --- MULTISELECT TAGS --- */
    [data-baseweb="tag"] {
        background: var(--blue-glow) !important;
        border: 1px solid rgba(59,130,246,.3) !important;
        border-radius: 6px !important;
    }

    [data-baseweb="tag"] span {
        color: #93c5fd !important;
    }

    /* --- SCROLLBAR --- */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: var(--bg); }
    ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.2); }

    /* --- DIVIDER --- */
    hr { border-color: var(--border) !important; }
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
    return mapping.get(value.lower(), value)


def _to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


@st.cache_data(ttl=REFRESH_SECONDS)
def parse_excel(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    data = raw.iloc[2:, : len(BASE_COLUMNS)].copy()
    data.columns = BASE_COLUMNS

    data = data.dropna(how="all")
    data = data[~(data["S_No"].isna() & data["Project_Name"].isna())]

    for col in ["Job_Ref", "LPO_Ref", "Customer", "Project_Name", "Region", "Location", "Priority"]:
        default = "Unknown" if col in ["Customer", "Region"] else ""
        data[col] = data[col].apply(lambda x: _clean_text(x, default))

    data["Material_Status"] = data["Material_Status"].apply(lambda x: _clean_text(x, "Not Ordered"))
    data["Status"] = data["Status"].apply(_clean_status)
    data["Work_Status"] = data["Work_Status"].apply(lambda x: _clean_text(x, ""))

    for col in ["Overall_Progress", "Engineering_Pct", "Delivery_Pct", "Execution_Pct"]:
        data[col] = _to_number(data[col]).clip(lower=0, upper=100)

    for group_name, cols in [("Eng", ENGINEERING_COLS), ("Del", DELIVERY_COLS)]:
        status_block = data[cols].fillna("").astype(str).apply(lambda s: s.str.strip().str.upper())
        data[f"{group_name}_Done"] = (status_block == "DONE").sum(axis=1)
        data[f"{group_name}_Partial"] = (status_block == "PART.DONE").sum(axis=1)
        data[f"{group_name}_NA"] = (status_block == "N/A").sum(axis=1)

    data["S_No"] = data["S_No"].fillna("").astype(str).str.replace(".0", "", regex=False)
    data["Priority"] = data["Priority"].replace("", "Medium")
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


def render_kpi(label: str, value: str, note: str = "", accent: str = "#3b82f6") -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-accent" style="background: linear-gradient(180deg, {accent}, {accent}88);"></div>
            <div class="kpi-label">{html_escape(label)}</div>
            <div class="kpi-value" style="color: {accent};">{html_escape(value)}</div>
            <div class="kpi-note">{html_escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_phase_card(title: str, value: float, subtitle: str, color: str) -> None:
    st.markdown(
        f"""
        <div class="phase-card" style="border-top: 3px solid {color};">
            <div class="phase-title">
                <span class="phase-name">{html_escape(title)}</span>
                <span class="phase-percent" style="color:{color};">{value:.1f}%</span>
            </div>
            <div class="bar-bg"><div class="bar-fill" style="width:{max(0, min(value, 100)):.1f}%; background: linear-gradient(90deg, {color}88, {color});"></div></div>
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
        font=dict(color="#9ca3af", family="Outfit, sans-serif", size=12),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
            bgcolor="rgba(0,0,0,0)", font=dict(color="#9ca3af"),
        ),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(255,255,255,0.05)", zeroline=False, color="#6b7280")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(255,255,255,0.05)", zeroline=False, color="#6b7280")
    return fig


def pie_chart(df: pd.DataFrame, names: str, values: str, colors: dict[str, str] | None = None):
    color_sequence = None
    if colors:
        color_sequence = [colors.get(str(name), "#94a3b8") for name in df[names]]
    fig = px.pie(df, names=names, values=values, hole=0.62, color_discrete_sequence=color_sequence)
    fig.update_traces(
        textposition="outside",
        textinfo="percent+label",
        marker=dict(line=dict(color="#ffffff", width=2)),
    )
    fig = style_fig(fig, height=340)
    fig.update_layout(showlegend=False, margin=dict(l=20, r=20, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def bar_chart(df: pd.DataFrame, x: str, y: str, orientation: str = "v", color: str = "#3b82f6", height: int = 360):
    fig = px.bar(df, x=x, y=y, orientation=orientation, text_auto=True)
    fig.update_traces(marker_color=color, textfont_color="#9ca3af", marker_line_width=0)
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
        <div class="hero-eyebrow">Live Operations</div>
        <div class="hero-title">Project Tracking <span>Dashboard</span></div>
        <div class="hero-subtitle">Cold Rooms &middot; Cabinets &middot; Refrigeration Projects — live from Excel</div>
        <div class="source-row">
            <span class="pill {source_badge}">&#9679; {html_escape(source_name)}</span>
            <span class="pill blue">&#8635; {REFRESH_SECONDS}s refresh</span>
            <span class="pill orange">&#9672; {len(filtered)} / {len(df)} projects</span>
            <span class="pill">&#8857; {last_refresh_text}</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# KPIs
kpi_cols = st.columns(6)
with kpi_cols[0]:
    render_kpi("Total Projects", str(len(filtered)), "Projects in current view", "#3b82f6")
with kpi_cols[1]:
    render_kpi("Completed", str(int((filtered["Status"] == "Completed").sum())), "Ready / closed", "#10b981")
with kpi_cols[2]:
    render_kpi("In Progress", str(int((filtered["Status"] == "In Progress").sum())), "Currently active", "#2dd4bf")
with kpi_cols[3]:
    render_kpi("On Hold", str(int((filtered["Status"] == "On Hold").sum())), "Needs attention", "#f59e0b")
with kpi_cols[4]:
    render_kpi("Not Started", str(int((filtered["Status"] == "Not Started").sum())), "Yet to begin", "#6b7280")
with kpi_cols[5]:
    render_kpi("Avg Progress", pct(avg_progress), "Across current view", "#a78bfa")

# PHASE CARDS
st.markdown('<div class="section-header">Phase Progress Overview</div>', unsafe_allow_html=True)
st.markdown('<div class="section-caption">Average completion by Engineering, Delivery and Execution phases.</div>', unsafe_allow_html=True)
phase_cols = st.columns(3)
with phase_cols[0]:
    render_phase_card(
        "Engineering",
        filtered["Engineering_Pct"].mean() if len(filtered) else 0,
        "Design / Submittal / Drawing / ELS / BOM",
        "#3b82f6",
    )
with phase_cols[1]:
    render_phase_card(
        "Delivery",
        filtered["Delivery_Pct"].mean() if len(filtered) else 0,
        f"Material delivery across {len(DELIVERY_COLS)} tracked items",
        "#a78bfa",
    )
with phase_cols[2]:
    render_phase_card(
        "Execution",
        filtered["Execution_Pct"].mean() if len(filtered) else 0,
        "Site installation, commissioning and handover progress",
        "#f59e0b",
    )

# TABS FOR READABILITY
tab_overview, tab_progress, tab_details = st.tabs(["Overview", "Progress Analysis", "Project Details"])

with tab_overview:
    st.markdown('<div class="section-header">Portfolio Overview</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        add_chart_card_title("Project Status")
        pie_chart(count_by_order(filtered, "Status", STATUS_ORDER), "Status", "Count", STATUS_COLORS)
        close_chart_card()
    with c2:
        add_chart_card_title("Material Status")
        pie_chart(count_by_order(filtered, "Material_Status", MATERIAL_ORDER), "Material_Status", "Count", MATERIAL_COLORS)
        close_chart_card()
    with c3:
        add_chart_card_title("Priority")
        pie_chart(count_by_order(filtered, "Priority", PRIORITY_ORDER), "Priority", "Count", PRIORITY_COLORS)
        close_chart_card()

    c4, c5 = st.columns(2)
    with c4:
        add_chart_card_title("Region Breakdown")
        region_df = filtered["Region"].value_counts().reset_index()
        region_df.columns = ["Region", "Count"]
        bar_chart(region_df, "Region", "Count", color="#3b82f6", height=380)
        close_chart_card()
    with c5:
        add_chart_card_title("Top Customers")
        cust_df = filtered["Customer"].value_counts().head(10).reset_index()
        cust_df.columns = ["Customer", "Count"]
        bar_chart(cust_df.sort_values("Count"), "Count", "Customer", orientation="h", color="#2dd4bf", height=380)
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
        fig.add_trace(go.Bar(x=delivery_df["Item"], y=delivery_df["Done"], name="Done", marker_color="#10b981", marker_line_width=0))
        fig.add_trace(go.Bar(x=delivery_df["Item"], y=delivery_df["Partial"], name="Partial", marker_color="#f59e0b", marker_line_width=0))
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
            "0-25%":  "#ef4444",
            "26-50%": "#f59e0b",
            "51-75%": "#eab308",
            "76-99%": "#3b82f6",
            "100%":   "#10b981",
        }
        pie_chart(bucket_df, "Progress Bucket", "Count", bucket_colors)
        close_chart_card()

    add_chart_card_title("Phase Progress by Project")
    phase_df = filtered.sort_values("Overall_Progress", ascending=False).head(25)
    fig = go.Figure()
    fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Engineering_Pct"], name="Engineering", orientation="h", marker_color="#3b82f6", marker_line_width=0))
    fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Delivery_Pct"], name="Delivery", orientation="h", marker_color="#a78bfa", marker_line_width=0))
    fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Execution_Pct"], name="Execution", orientation="h", marker_color="#f59e0b", marker_line_width=0))
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
