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

# ---------------------------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Project Tracking Dashboard",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
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

STATUS_ORDER   = ["Completed", "In Progress", "On Hold", "Not Started", "Cancelled"]
MATERIAL_ORDER = ["Delivered", "Partially Delivered", "Ordered", "Not Ordered"]
PRIORITY_ORDER = ["High", "Medium", "Low"]

STATUS_COLORS = {
    "Completed":   "#16a34a",
    "In Progress": "#2563eb",
    "On Hold":     "#d97706",
    "Not Started": "#94a3b8",
    "Cancelled":   "#dc2626",
}
PRIORITY_COLORS = {
    "High":   "#dc2626",
    "Medium": "#d97706",
    "Low":    "#16a34a",
}
MATERIAL_COLORS = {
    "Delivered":           "#16a34a",
    "Partially Delivered": "#d97706",
    "Ordered":             "#2563eb",
    "Not Ordered":         "#94a3b8",
}

# ---------------------------------------------------------------------------
# STYLES
# ---------------------------------------------------------------------------
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
}

.stApp {
    background: #f1f5f9 !important;
}

section[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e2e8f0 !important;
}

section[data-testid="stSidebar"] * {
    color: #1e293b !important;
}

.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 4rem !important;
    max-width: 1560px !important;
}

/* ---- HERO ---- */
.hero-wrap {
    background: linear-gradient(135deg, #1e40af 0%, #1d4ed8 50%, #2563eb 100%);
    border-radius: 18px;
    padding: 30px 36px;
    margin-bottom: 24px;
}
.hero-title {
    font-size: 32px;
    font-weight: 800;
    color: #ffffff;
    margin: 0 0 4px;
    letter-spacing: -0.02em;
}
.hero-sub {
    font-size: 14px;
    color: rgba(255,255,255,0.75);
    margin-bottom: 18px;
}
.pill-row { display: flex; flex-wrap: wrap; gap: 8px; }
.pill {
    font-size: 12px;
    font-weight: 500;
    padding: 4px 12px;
    border-radius: 999px;
    background: rgba(255,255,255,0.15);
    color: #ffffff;
    border: 1px solid rgba(255,255,255,0.25);
}
.pill-green { background: rgba(22,163,74,0.25);  border-color: rgba(22,163,74,0.5);  color: #bbf7d0; }
.pill-blue  { background: rgba(255,255,255,0.20); border-color: rgba(255,255,255,0.4); color: #ffffff; }
.pill-amber { background: rgba(251,191,36,0.25);  border-color: rgba(251,191,36,0.5);  color: #fef9c3; }

/* ---- KPI CARD ---- */
.kpi-wrap {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 18px 20px 16px;
    min-height: 120px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.kpi-wrap::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: var(--ac, #2563eb);
    border-radius: 14px 14px 0 0;
}
.kpi-lbl {
    font-size: 11px;
    font-weight: 600;
    letter-spacing: .08em;
    text-transform: uppercase;
    color: #64748b;
    margin-bottom: 8px;
}
.kpi-val {
    font-size: 36px;
    font-weight: 800;
    line-height: 1;
    color: #0f172a;
    letter-spacing: -0.03em;
}
.kpi-note {
    font-size: 12px;
    color: #94a3b8;
    margin-top: 6px;
}

/* ---- SECTION HEADER ---- */
.sec-hdr {
    font-size: 18px;
    font-weight: 700;
    color: #0f172a;
    margin: 24px 0 4px;
    padding-left: 10px;
    border-left: 3px solid #2563eb;
}
.sec-cap {
    font-size: 13px;
    color: #64748b;
    margin: 0 0 14px 10px;
}

/* ---- PHASE CARD ---- */
.phase-wrap {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-top: 4px solid var(--ac, #2563eb);
    border-radius: 14px;
    padding: 18px 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.phase-name {
    font-weight: 700;
    font-size: 14px;
    color: #374151;
}
.phase-pct {
    font-size: 32px;
    font-weight: 800;
    color: var(--ac, #2563eb);
    letter-spacing: -0.03em;
    line-height: 1;
}
.phase-bar-bg {
    height: 8px;
    background: #f1f5f9;
    border-radius: 999px;
    margin: 12px 0 8px;
    overflow: hidden;
}
.phase-bar-fill {
    height: 8px;
    border-radius: 999px;
    background: var(--ac, #2563eb);
}
.phase-sub { font-size: 12px; color: #94a3b8; }

/* ---- CHART CARD ---- */
.chart-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 18px 16px 8px;
    margin-bottom: 18px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.chart-hdr {
    font-size: 14px;
    font-weight: 700;
    color: #0f172a;
    margin-bottom: 4px;
}

/* ---- TABS ---- */
.stTabs [data-baseweb="tab-list"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 12px !important;
    padding: 4px !important;
    gap: 2px !important;
    width: fit-content !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05) !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border: none !important;
    border-radius: 8px !important;
    color: #64748b !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    padding: 8px 20px !important;
}
.stTabs [aria-selected="true"] {
    background: #2563eb !important;
    color: #ffffff !important;
    border: none !important;
}
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"] { display: none !important; }

/* ---- DATAFRAME ---- */
[data-testid="stDataFrame"] {
    border: 1px solid #e2e8f0 !important;
    border-radius: 12px !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05) !important;
}

/* ---- DOWNLOAD BUTTON ---- */
.stDownloadButton > button {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    color: #374151 !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 500 !important;
}
.stDownloadButton > button:hover {
    border-color: #2563eb !important;
    color: #2563eb !important;
    background: #eff6ff !important;
}

hr { border-color: #e2e8f0 !important; }
</style>
"""

st.html(_CSS)

# ---------------------------------------------------------------------------
# DATA SOURCE
# ---------------------------------------------------------------------------

def make_onedrive_download_url(url: str) -> str:
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
            "Use an Anyone-with-link share URL or Microsoft Graph mode."
        )
    return response.content


@st.cache_data(ttl=REFRESH_SECONDS)
def get_graph_token() -> str:
    if msal is None:
        raise RuntimeError("msal is not installed. Add msal to requirements.txt")
    tenant_id     = _secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    client_id     = _secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    client_secret = _secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        raise RuntimeError("Missing GRAPH_TENANT_ID, GRAPH_CLIENT_ID, or GRAPH_CLIENT_SECRET.")
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
    user_id   = _secret("GRAPH_USER_ID")
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
    st.error("No Excel source found. Set EXCEL_FILE_URL or configure Microsoft Graph secrets.")
    st.stop()

# ---------------------------------------------------------------------------
# PARSING
# ---------------------------------------------------------------------------

def _clean_text(value, default: str = "") -> str:
    if pd.isna(value):
        return default
    text = str(value).strip()
    return text if text else default


def _clean_status(value: str) -> str:
    value = _clean_text(value, "Not Started")
    mapping = {
        "completed":   "Completed",
        "complete":    "Completed",
        "in progress": "In Progress",
        "on progress": "In Progress",
        "progress":    "In Progress",
        "on hold":     "On Hold",
        "hold":        "On Hold",
        "cancelled":   "Cancelled",
        "canceled":    "Cancelled",
        "not started": "Not Started",
        "notstart":    "Not Started",
    }
    return mapping.get(value.lower(), value)


def _to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


@st.cache_data(ttl=REFRESH_SECONDS)
def parse_excel(file_bytes: bytes) -> pd.DataFrame:
    raw  = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    data = raw.iloc[2:, : len(BASE_COLUMNS)].copy()
    data.columns = BASE_COLUMNS
    data = data.dropna(how="all")
    data = data[~(data["S_No"].isna() & data["Project_Name"].isna())]
    for col in ["Job_Ref", "LPO_Ref", "Customer", "Project_Name", "Region", "Location", "Priority"]:
        default = "Unknown" if col in ["Customer", "Region"] else ""
        data[col] = data[col].apply(lambda x: _clean_text(x, default))
    data["Material_Status"] = data["Material_Status"].apply(lambda x: _clean_text(x, "Not Ordered"))
    data["Status"]           = data["Status"].apply(_clean_status)
    data["Work_Status"]      = data["Work_Status"].apply(lambda x: _clean_text(x, ""))
    for col in ["Overall_Progress", "Engineering_Pct", "Delivery_Pct", "Execution_Pct"]:
        data[col] = _to_number(data[col]).clip(lower=0, upper=100)
    for group_name, cols in [("Eng", ENGINEERING_COLS), ("Del", DELIVERY_COLS)]:
        block = data[cols].fillna("").astype(str).apply(lambda s: s.str.strip().str.upper())
        data[f"{group_name}_Done"]    = (block == "DONE").sum(axis=1)
        data[f"{group_name}_Partial"] = (block == "PART.DONE").sum(axis=1)
        data[f"{group_name}_NA"]      = (block == "N/A").sum(axis=1)
    data["S_No"]     = data["S_No"].fillna("").astype(str).str.replace(".0", "", regex=False)
    data["Priority"] = data["Priority"].replace("", "Medium")
    return data.reset_index(drop=True)

# ---------------------------------------------------------------------------
# FILTERS
# ---------------------------------------------------------------------------

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar:
        st.markdown("### Filters")
        st.caption("Focus the whole dashboard.")
        search = st.text_input("Search", placeholder="Project, customer, location...")
        quick  = st.radio(
            "Quick view",
            ["All", "Completed", "In Progress", "On Hold", "Not Started"],
            horizontal=False,
        )
        st.divider()
        status   = st.multiselect("Status",   sorted(df["Status"].dropna().unique()),         default=[])
        region   = st.multiselect("Region",   sorted(df["Region"].dropna().unique()),          default=[])
        customer = st.multiselect("Customer", sorted(df["Customer"].dropna().unique()),        default=[])
        material = st.multiselect("Material", sorted(df["Material_Status"].dropna().unique()), default=[])
        priority = st.multiselect("Priority", sorted(df["Priority"].dropna().unique()),        default=[])
        st.divider()
        st.caption(f"Auto-refresh every {REFRESH_SECONDS}s")

    out = df.copy()
    if quick != "All":
        out = out[out["Status"] == quick]
    if search:
        t = search.lower().strip()
        mask = (
            out["Project_Name"].str.lower().str.contains(t, na=False)
            | out["Customer"].str.lower().str.contains(t, na=False)
            | out["Location"].str.lower().str.contains(t, na=False)
            | out["Job_Ref"].astype(str).str.lower().str.contains(t, na=False)
            | out["LPO_Ref"].astype(str).str.lower().str.contains(t, na=False)
        )
        out = out[mask]
    if status:   out = out[out["Status"].isin(status)]
    if region:   out = out[out["Region"].isin(region)]
    if customer: out = out[out["Customer"].isin(customer)]
    if material: out = out[out["Material_Status"].isin(material)]
    if priority: out = out[out["Priority"].isin(priority)]
    return out

# ---------------------------------------------------------------------------
# UI HELPERS
# ---------------------------------------------------------------------------

def pct(value: float) -> str:
    return f"{value:.1f}%"


def h(text: str) -> str:
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
    extra  = df[~df[col].isin(order)][col].value_counts().reset_index()
    extra.columns = [col, "Count"]
    return pd.concat([counts, extra], ignore_index=True)


def render_kpi(label: str, value: str, note: str, accent: str) -> None:
    st.html(f"""
    <div class="kpi-wrap" style="--ac:{accent}">
        <div class="kpi-lbl">{h(label)}</div>
        <div class="kpi-val">{h(value)}</div>
        <div class="kpi-note">{h(note)}</div>
    </div>
    """)


def render_phase_card(title: str, value: float, subtitle: str, color: str) -> None:
    w = max(0.0, min(value, 100.0))
    st.html(f"""
    <div class="phase-wrap" style="--ac:{color}">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px">
            <div class="phase-name">{h(title)}</div>
            <div class="phase-pct">{value:.1f}%</div>
        </div>
        <div class="phase-bar-bg"><div class="phase-bar-fill" style="width:{w:.1f}%"></div></div>
        <div class="phase-sub">{h(subtitle)}</div>
    </div>
    """)


def chart_card_open(title: str) -> None:
    st.html(f'<div class="chart-card"><div class="chart-hdr">{h(title)}</div></div>')


def style_fig(fig: go.Figure, height: int = 360) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=28, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#374151", family="Inter, sans-serif", size=12),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
            bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"),
        ),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#f1f5f9", zeroline=False, color="#6b7280", tickfont=dict(color="#6b7280"))
    fig.update_yaxes(showgrid=True, gridcolor="#f1f5f9", zeroline=False, color="#6b7280", tickfont=dict(color="#6b7280"))
    return fig


def donut_chart(df: pd.DataFrame, names: str, values: str, colors: dict | None = None) -> None:
    """Clean donut chart: inside labels for big slices, legend below for all."""
    # Filter out zero-count slices so labels don't crowd
    df = df[df[values] > 0].copy()
    seq = [colors.get(str(n), "#94a3b8") for n in df[names]] if colors else None

    fig = px.pie(
        df, names=names, values=values,
        hole=0.55,
        color_discrete_sequence=seq,
    )
    fig.update_traces(
        textposition="inside",
        textinfo="percent",
        textfont=dict(size=13, color="#ffffff"),
        insidetextorientation="radial",
        marker=dict(line=dict(color="#ffffff", width=2)),
        hovertemplate="<b>%{label}</b><br>Count: %{value}<br>%{percent}<extra></extra>",
    )
    total = int(df[values].sum())
    fig.update_layout(
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.08,
            xanchor="center",
            x=0.5,
            font=dict(size=12, color="#374151"),
            bgcolor="rgba(0,0,0,0)",
            itemwidth=80,
        ),
        annotations=[dict(
            text=f"<b>{total}</b><br><span style='font-size:11px;color:#6b7280'>Total</span>",
            x=0.5, y=0.5,
            font=dict(size=16, color="#0f172a"),
            showarrow=False,
        )],
        margin=dict(l=10, r=10, t=10, b=60),
        height=340,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#374151", family="Inter, sans-serif", size=12),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def bar_chart(df: pd.DataFrame, x: str, y: str, orientation: str = "v",
              color: str = "#2563eb", height: int = 360) -> None:
    fig = px.bar(df, x=x, y=y, orientation=orientation, text_auto=True)
    fig.update_traces(
        marker_color=color,
        textfont=dict(color="#374151", size=11),
        marker_line_width=0,
    )
    fig = style_fig(fig, height=height)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def colored_bar_chart(df: pd.DataFrame, x: str, y: str, color_col: str,
                      color_map: dict, height: int = 380) -> None:
    """Horizontal bar chart where each bar is colored by its category value."""
    fig = px.bar(
        df, x=x, y=y, orientation="h",
        color=color_col, color_discrete_map=color_map,
        text=x,
    )
    fig.update_traces(
        textfont=dict(color="#374151", size=11),
        marker_line_width=0,
        textposition="outside",
    )
    fig = style_fig(fig, height=height)
    fig.update_layout(showlegend=True)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
st_autorefresh(interval=REFRESH_SECONDS * 1000, key="refresh")

uploaded = None

try:
    file_bytes, source_name = get_excel_bytes(uploaded)
    df = parse_excel(file_bytes)
except Exception as exc:
    st.error(f"Could not load dashboard data: {exc}")
    st.stop()

filtered = apply_filters(df)

avg_progress      = filtered["Overall_Progress"].mean() if len(filtered) else 0
last_refresh_text = pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")
badge_cls         = "pill-green" if ("OneDrive" in source_name or "Graph" in source_name) else "pill-blue"

# HERO
st.html(f"""
<div class="hero-wrap">
    <div class="hero-title">Project Tracking Dashboard</div>
    <div class="hero-sub">Cold Rooms &middot; Cabinets &middot; Refrigeration Projects &mdash; live from Excel</div>
    <div class="pill-row">
        <span class="pill {badge_cls}">{h(source_name)}</span>
        <span class="pill pill-blue">Refresh every {REFRESH_SECONDS}s</span>
        <span class="pill pill-amber">Showing {len(filtered)} of {len(df)} projects</span>
        <span class="pill">Updated {last_refresh_text}</span>
    </div>
</div>
""")

# KPIs
k = st.columns(6)
with k[0]: render_kpi("Total Projects", str(len(filtered)),                                         "In current view",   "#2563eb")
with k[1]: render_kpi("Completed",      str(int((filtered["Status"]=="Completed").sum())),          "Ready / closed",    "#16a34a")
with k[2]: render_kpi("In Progress",    str(int((filtered["Status"]=="In Progress").sum())),        "Currently active",  "#0284c7")
with k[3]: render_kpi("On Hold",        str(int((filtered["Status"]=="On Hold").sum())),            "Needs attention",   "#d97706")
with k[4]: render_kpi("Not Started",    str(int((filtered["Status"]=="Not Started").sum())),        "Yet to begin",      "#64748b")
with k[5]: render_kpi("Avg Progress",   pct(avg_progress),                                         "Across view",       "#7c3aed")

# PHASE CARDS
st.html('<div class="sec-hdr">Phase Progress Overview</div><div class="sec-cap">Average completion across Engineering, Delivery and Execution phases.</div>')
p = st.columns(3)
with p[0]: render_phase_card("Engineering", filtered["Engineering_Pct"].mean() if len(filtered) else 0, "Design / Submittal / Drawing / ELS / BOM",            "#2563eb")
with p[1]: render_phase_card("Delivery",    filtered["Delivery_Pct"].mean()    if len(filtered) else 0, f"Material delivery across {len(DELIVERY_COLS)} items", "#7c3aed")
with p[2]: render_phase_card("Execution",   filtered["Execution_Pct"].mean()   if len(filtered) else 0, "Installation, commissioning and handover",              "#d97706")

# TABS
tab_ov, tab_pr, tab_dt = st.tabs(["  Overview  ", "  Progress Analysis  ", "  Project Details  "])

with tab_ov:
    st.html('<div class="sec-hdr">Portfolio Overview</div><div class="sec-cap">Status, material, and priority breakdown across all visible projects.</div>')

    # Row 1: three donuts -- but keep slices readable
    c1, c2, c3 = st.columns(3)
    with c1:
        chart_card_open("Project Status")
        donut_chart(count_by_order(filtered, "Status", STATUS_ORDER), "Status", "Count", STATUS_COLORS)
    with c2:
        chart_card_open("Material Status")
        donut_chart(count_by_order(filtered, "Material_Status", MATERIAL_ORDER), "Material_Status", "Count", MATERIAL_COLORS)
    with c3:
        chart_card_open("Priority Distribution")
        donut_chart(count_by_order(filtered, "Priority", PRIORITY_ORDER), "Priority", "Count", PRIORITY_COLORS)

    # Row 2: Region bar + Top customers
    c4, c5 = st.columns(2)
    with c4:
        chart_card_open("Projects by Region")
        rdf = filtered["Region"].value_counts().reset_index()
        rdf.columns = ["Region", "Count"]
        rdf = rdf.sort_values("Count", ascending=False)
        bar_chart(rdf, "Region", "Count", color="#2563eb", height=360)
    with c5:
        chart_card_open("Top 10 Customers by Projects")
        cdf = filtered["Customer"].value_counts().head(10).reset_index()
        cdf.columns = ["Customer", "Count"]
        cdf = cdf.sort_values("Count", ascending=True)
        bar_chart(cdf, "Count", "Customer", orientation="h", color="#0284c7", height=360)

with tab_pr:
    st.html('<div class="sec-hdr">Progress Analysis</div>')

    c6, c7 = st.columns([3, 2])
    with c6:
        chart_card_open("Delivery Items Status")
        done_vals    = [int((filtered[col].fillna("").astype(str).str.strip().str.upper() == "DONE").sum())      for col in DELIVERY_COLS]
        partial_vals = [int((filtered[col].fillna("").astype(str).str.strip().str.upper() == "PART.DONE").sum()) for col in DELIVERY_COLS]
        not_done_vals = [len(filtered) - d - p for d, p in zip(done_vals, partial_vals)]
        ddf = pd.DataFrame({
            "Item":     DELIVERY_COLS,
            "Done":     done_vals,
            "Partial":  partial_vals,
            "Not Done": not_done_vals,
        })
        fig = go.Figure()
        fig.add_trace(go.Bar(x=ddf["Item"], y=ddf["Done"],     name="Done",     marker_color="#16a34a", marker_line_width=0))
        fig.add_trace(go.Bar(x=ddf["Item"], y=ddf["Partial"],  name="Partial",  marker_color="#d97706", marker_line_width=0))
        fig.add_trace(go.Bar(x=ddf["Item"], y=ddf["Not Done"], name="Not Done", marker_color="#e2e8f0", marker_line_width=0))
        fig.update_layout(barmode="stack")
        fig = style_fig(fig, height=380)
        fig.update_layout(margin=dict(l=16, r=16, t=20, b=110))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with c7:
        chart_card_open("Overall Progress Distribution")
        bins   = [-0.1, 25, 50, 75, 99.999, 100]
        labels = ["0-25%", "26-50%", "51-75%", "76-99%", "100%"]
        bkt    = pd.cut(filtered["Overall_Progress"], bins=bins, labels=labels)
        bkt_df = bkt.value_counts().reindex(labels, fill_value=0).reset_index()
        bkt_df.columns = ["Progress Bucket", "Count"]
        # Use a bar chart here -- much cleaner than a pie for ordered buckets
        clrs   = ["#ef4444", "#f97316", "#eab308", "#2563eb", "#16a34a"]
        fig    = go.Figure()
        for i, (_, row) in enumerate(bkt_df.iterrows()):
            fig.add_trace(go.Bar(
                x=[row["Progress Bucket"]],
                y=[row["Count"]],
                name=row["Progress Bucket"],
                marker_color=clrs[i],
                marker_line_width=0,
                text=[str(int(row["Count"]))],
                textposition="outside",
                showlegend=False,
            ))
        fig = style_fig(fig, height=380)
        fig.update_layout(
            margin=dict(l=16, r=16, t=20, b=20),
            xaxis_title="Progress Range",
            yaxis_title="No. of Projects",
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # Phase progress by project
    chart_card_open("Phase Progress by Project (Top 25 by Overall Progress)")
    ph_df = filtered.sort_values("Overall_Progress", ascending=False).head(25)
    fig   = go.Figure()
    fig.add_trace(go.Bar(y=ph_df["Project_Name"], x=ph_df["Engineering_Pct"], name="Engineering", orientation="h", marker_color="#2563eb", marker_line_width=0))
    fig.add_trace(go.Bar(y=ph_df["Project_Name"], x=ph_df["Delivery_Pct"],    name="Delivery",    orientation="h", marker_color="#7c3aed", marker_line_width=0))
    fig.add_trace(go.Bar(y=ph_df["Project_Name"], x=ph_df["Execution_Pct"],   name="Execution",   orientation="h", marker_color="#d97706", marker_line_width=0))
    fig.update_layout(
        barmode="group",
        xaxis_title="Completion %",
        yaxis={"autorange": "reversed"},
        xaxis=dict(range=[0, 100]),
    )
    fig = style_fig(fig, height=max(400, len(ph_df) * 28))
    fig.update_layout(margin=dict(l=16, r=16, t=24, b=16))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

with tab_dt:
    st.html('<div class="sec-hdr">Project Details</div><div class="sec-cap">Full sortable list. Use sidebar filters to narrow down.</div>')
    show_cols = [
        "S_No", "Customer", "Project_Name", "Region", "Location", "Status",
        "Engineering_Pct", "Delivery_Pct", "Execution_Pct", "Overall_Progress",
        "Material_Status", "Priority", "Work_Status",
    ]
    tbl = filtered[show_cols].sort_values("Overall_Progress", ascending=False).copy()
    tbl = tbl.rename(columns={
        "S_No": "#",
        "Project_Name": "Project",
        "Engineering_Pct": "Eng %",
        "Delivery_Pct": "Delivery %",
        "Execution_Pct": "Exec %",
        "Overall_Progress": "Overall %",
        "Material_Status": "Material",
        "Work_Status": "Work Status",
    })
    st.dataframe(
        tbl,
        use_container_width=True,
        hide_index=True,
        height=650,
        column_config={
            "Eng %":      st.column_config.ProgressColumn("Eng %",      min_value=0, max_value=100, format="%.0f%%"),
            "Delivery %": st.column_config.ProgressColumn("Delivery %", min_value=0, max_value=100, format="%.0f%%"),
            "Exec %":     st.column_config.ProgressColumn("Exec %",     min_value=0, max_value=100, format="%.0f%%"),
            "Overall %":  st.column_config.ProgressColumn("Overall %",  min_value=0, max_value=100, format="%.0f%%"),
        },
    )
    st.download_button(
        "Download filtered project list as CSV",
        data=tbl.to_csv(index=False).encode("utf-8"),
        file_name="filtered_projects.csv",
        mime="text/csv",
    )
