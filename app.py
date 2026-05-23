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

st.set_page_config(
    page_title="Project Tracking Dashboard",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)

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

ENGINEERING_COLS = ["Equi. DesigN", "Technical Submittal", "Drawing", "ELS", "BOM"]
DELIVERY_COLS = [
    "Out_Door", "Indoor", "CR Panels", "CR Ins. Materials]", "Doors",
    "Ref. Inst. Materials", "CCP", "Display CCP", "Floor Heater", "Cabinets", "Any Special",
]
BASE_COLUMNS = [
    "S_No", "Date_Time", "Job_Ref", "LPO_Ref", "Customer", "Project_Name",
    "Region", "Location", "Payment_Terms",
    *ENGINEERING_COLS, *DELIVERY_COLS,
    "Work_Status", "Remarks", "Material_Status", "Overall_Progress",
    "Priority", "Status", "Engineering_Pct", "Delivery_Pct", "Execution_Pct",
]

STATUS_ORDER   = ["Completed", "In Progress", "On Hold", "Not Started", "Cancelled"]
MATERIAL_ORDER = ["Delivered", "Partially Delivered", "Ordered", "Not Ordered"]
PRIORITY_ORDER = ["High", "Medium", "Low"]

# Exact colors from the HTML reference
STATUS_COLORS   = {"Completed": "#22c55e", "In Progress": "#14b8a6", "On Hold": "#f97316", "Not Started": "#64748b", "Cancelled": "#ef4444"}
PRIORITY_COLORS = {"High": "#ef4444", "Medium": "#eab308", "Low": "#22c55e"}
MATERIAL_COLORS = {"Delivered": "#22c55e", "Partially Delivered": "#eab308", "Ordered": "#f97316", "Not Ordered": "#ef4444"}

# ---------------------------------------------------------------------------
# CSS -- dark theme matching HTML reference exactly
# ---------------------------------------------------------------------------
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Segoe UI', 'Inter', Tahoma, Arial, sans-serif !important; }

.stApp { background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%) !important; min-height: 100vh; }

section[data-testid="stSidebar"] {
    background: #1e293b !important;
    border-right: 1px solid #334155 !important;
}
section[data-testid="stSidebar"] * { color: #f1f5f9 !important; }
section[data-testid="stSidebar"] .stTextInput input,
section[data-testid="stSidebar"] .stSelectbox > div > div,
section[data-testid="stSidebar"] .stMultiSelect > div > div {
    background: #334155 !important; border-color: #475569 !important; color: #f1f5f9 !important;
}

.block-container { padding-top: 1.2rem !important; padding-bottom: 4rem !important; max-width: 1560px !important; }

/* HEADER */
.dash-header {
    display: flex; justify-content: space-between; align-items: flex-start;
    margin-bottom: 16px; padding-bottom: 14px;
    border-bottom: 1px solid #334155;
    flex-wrap: wrap; gap: 12px;
}
.dash-header h1 { font-size: 24px; font-weight: 700; color: #f1f5f9; margin: 0; }
.dash-header .sub { color: #94a3b8; font-size: 13px; margin-top: 4px; }
.dash-header .right { text-align: right; color: #94a3b8; font-size: 12px; }

/* QUICK BAR */
.quick-bar {
    display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; align-items: center;
}
.quick-label { color: #94a3b8; font-size: 12px; text-transform: uppercase; letter-spacing: .5px; }

/* KPI */
.kpi-grid { display: grid; grid-template-columns: repeat(8, 1fr); gap: 12px; margin-bottom: 18px; }
.kpi {
    background: #1e293b; border-radius: 12px; padding: 16px;
    border-left: 4px solid var(--ac, #3b82f6);
    cursor: pointer; transition: transform .15s;
}
.kpi:hover { transform: translateY(-2px); }
.kpi .lbl { color: #94a3b8; font-size: 11px; text-transform: uppercase; letter-spacing: .5px; }
.kpi .val { font-size: 28px; font-weight: 700; margin-top: 5px; color: #f1f5f9; }
.kpi .delta { font-size: 11px; color: #94a3b8; margin-top: 3px; }

/* SECTION TITLE */
.sec-title { font-size: 12px; color: #94a3b8; margin: 6px 0 12px; text-transform: uppercase; letter-spacing: .5px; font-weight: 600; }

/* PHASE CARDS */
.phase-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 18px; }
.phase-card {
    background: #1e293b; border-radius: 12px; padding: 18px;
    border-top: 4px solid var(--ac, #3b82f6);
}
.phase-card .ph-title { font-size: 14px; font-weight: 600; display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; color: #f1f5f9; }
.phase-card .ph-pct { font-size: 26px; font-weight: 700; }
.phase-card .ph-bar-bg { height: 10px; background: #334155; border-radius: 5px; overflow: hidden; margin-bottom: 10px; }
.phase-card .ph-bar-fill { height: 100%; border-radius: 5px; }
.phase-card .ph-stats { display: flex; justify-content: space-between; font-size: 11px; color: #94a3b8; }
.phase-card .ph-stats strong { color: #f1f5f9; font-size: 13px; }

/* CARD */
.card {
    background: #1e293b; border-radius: 12px; padding: 16px; margin-bottom: 16px;
}
.card h3 {
    font-size: 14px; font-weight: 600; margin-bottom: 12px;
    display: flex; justify-content: space-between; align-items: center; color: #f1f5f9;
}
.card h3 .tag { font-size: 11px; color: #94a3b8; font-weight: 400; }

/* BADGES */
.badge {
    display: inline-block; padding: 2px 8px; border-radius: 10px;
    font-size: 11px; font-weight: 600;
}
.badge-delivered   { background: rgba(34,197,94,.2);  color: #86efac; }
.badge-partial     { background: rgba(234,179,8,.2);  color: #fde047; }
.badge-ordered     { background: rgba(249,115,22,.2); color: #fdba74; }
.badge-notordered  { background: rgba(239,68,68,.2);  color: #fca5a5; }
.badge-high        { background: rgba(239,68,68,.2);  color: #fca5a5; }
.badge-medium      { background: rgba(234,179,8,.2);  color: #fde047; }
.badge-low         { background: rgba(34,197,94,.2);  color: #86efac; }
.badge-completed   { background: rgba(34,197,94,.25); color: #86efac; }
.badge-progress    { background: rgba(20,184,166,.25);color: #5eead4; }
.badge-hold        { background: rgba(249,115,22,.25);color: #fdba74; }
.badge-cancelled   { background: rgba(239,68,68,.25); color: #fca5a5; }
.badge-notstarted  { background: rgba(100,116,139,.3); color: #cbd5e1; }

/* MINI BAR */
.bar-cell { display: flex; align-items: center; gap: 6px; min-width: 90px; }
.bar-bg-mini { flex: 1; height: 6px; background: #334155; border-radius: 3px; overflow: hidden; min-width: 40px; }
.bar-fill-mini { height: 100%; border-radius: 3px; }
.bar-eng  { background: linear-gradient(90deg,#60a5fa,#3b82f6); }
.bar-del  { background: linear-gradient(90deg,#c084fc,#a855f7); }
.bar-exec { background: linear-gradient(90deg,#fb923c,#f97316); }
.bar-all  { background: linear-gradient(90deg,#ef4444,#eab308,#22c55e); }
.bar-pct  { font-size: 11px; font-weight: 600; min-width: 32px; text-align: right; color: #cbd5e1; }

/* TABLE */
.tbl-wrap { overflow: auto; max-height: 540px; border-radius: 8px; }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
thead { position: sticky; top: 0; background: #334155; z-index: 5; }
th { text-align: left; padding: 10px 8px; font-weight: 600; border-bottom: 2px solid #475569; white-space: nowrap; color: #f1f5f9; }
td { padding: 8px 8px; border-bottom: 1px solid #334155; color: #cbd5e1; }
tr:hover td { background: rgba(59,130,246,.08); }

/* TABS */
.stTabs [data-baseweb="tab-list"] {
    background: #1e293b !important; border: 1px solid #334155 !important;
    border-radius: 10px !important; padding: 3px !important; gap: 2px !important;
    width: fit-content !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important; border: none !important; border-radius: 7px !important;
    color: #94a3b8 !important; font-size: 13px !important; font-weight: 500 !important;
    padding: 7px 18px !important;
}
.stTabs [aria-selected="true"] {
    background: #3b82f6 !important; color: #ffffff !important;
}
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"] { display: none !important; }

/* SIDEBAR RADIO */
section[data-testid="stSidebar"] .stRadio label {
    background: #334155; border: 1px solid #475569; border-radius: 8px;
    padding: 6px 12px; margin: 2px 0; display: block; transition: all .15s; color: #f1f5f9 !important;
}
section[data-testid="stSidebar"] .stRadio label:hover { border-color: #3b82f6; background: rgba(59,130,246,.15); }

/* DOWNLOAD BUTTON */
.stDownloadButton > button {
    background: #334155 !important; border: 1px solid #475569 !important;
    border-radius: 8px !important; color: #f1f5f9 !important; font-size: 13px !important;
}
.stDownloadButton > button:hover { background: #3b82f6 !important; border-color: #3b82f6 !important; }

hr { border-color: #334155 !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #0f172a; }
::-webkit-scrollbar-thumb { background: #334155; border-radius: 3px; }
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
    r = requests.get(direct_url, timeout=60, allow_redirects=True)
    r.raise_for_status()
    if "text/html" in r.headers.get("content-type", "").lower() and len(r.content) < 500_000:
        raise RuntimeError("OneDrive URL returned HTML, not Excel. Check your share link.")
    return r.content


@st.cache_data(ttl=REFRESH_SECONDS)
def get_graph_token() -> str:
    if msal is None:
        raise RuntimeError("msal not installed")
    tenant_id     = _secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    client_id     = _secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    client_secret = _secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        raise RuntimeError("Missing GRAPH secrets")
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Graph token error: {result}")
    return result["access_token"]


@st.cache_data(ttl=REFRESH_SECONDS)
def load_excel_from_graph() -> bytes:
    user_id = _secret("GRAPH_USER_ID")
    file_path = _secret("ONEDRIVE_FILE_PATH")
    if not user_id or not file_path:
        raise RuntimeError("Missing GRAPH_USER_ID or ONEDRIVE_FILE_PATH")
    token = get_graph_token()
    if not file_path.startswith("/"):
        file_path = "/" + file_path
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:{file_path}:/content"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.content


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

def _ct(value, default=""):
    if pd.isna(value):
        return default
    t = str(value).strip()
    return t if t else default


def _cs(value):
    v = _ct(value, "Not Started")
    m = {
        "completed": "Completed", "complete": "Completed",
        "in progress": "In Progress", "on progress": "In Progress", "progress": "In Progress",
        "on hold": "On Hold", "hold": "On Hold",
        "cancelled": "Cancelled", "canceled": "Cancelled",
        "not started": "Not Started", "notstart": "Not Started",
    }
    return m.get(v.lower(), v)


def _num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


@st.cache_data(ttl=REFRESH_SECONDS)
def parse_excel(file_bytes: bytes) -> pd.DataFrame:
    raw  = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    data = raw.iloc[2:, : len(BASE_COLUMNS)].copy()
    data.columns = BASE_COLUMNS
    data = data.dropna(how="all")
    data = data[~(data["S_No"].isna() & data["Project_Name"].isna())]
    for col in ["Job_Ref", "LPO_Ref", "Customer", "Project_Name", "Region", "Location", "Priority"]:
        default = "Unknown" if col in ["Customer", "Region"] else ""
        data[col] = data[col].apply(lambda x: _ct(x, default))
    data["Material_Status"] = data["Material_Status"].apply(lambda x: _ct(x, "Not Ordered"))
    data["Status"]          = data["Status"].apply(_cs)
    data["Work_Status"]     = data["Work_Status"].apply(lambda x: _ct(x, ""))
    for col in ["Overall_Progress", "Engineering_Pct", "Delivery_Pct", "Execution_Pct"]:
        data[col] = _num(data[col]).clip(0, 100)
    for gn, cols in [("Eng", ENGINEERING_COLS), ("Del", DELIVERY_COLS)]:
        blk = data[cols].fillna("").astype(str).apply(lambda s: s.str.strip().str.upper())
        data[f"{gn}_Done"]    = (blk == "DONE").sum(axis=1)
        data[f"{gn}_Partial"] = (blk == "PART.DONE").sum(axis=1)
        data[f"{gn}_NA"]      = (blk == "N/A").sum(axis=1)
    data["S_No"]     = data["S_No"].fillna("").astype(str).str.replace(".0", "", regex=False)
    data["Priority"] = data["Priority"].replace("", "Medium")
    return data.reset_index(drop=True)

# ---------------------------------------------------------------------------
# FILTERS
# ---------------------------------------------------------------------------

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar:
        st.markdown("### Filters")
        search = st.text_input("Search", placeholder="Project, customer, location, job ref...")
        quick  = st.radio("Quick view", ["All", "Completed", "In Progress", "On Hold", "Not Started"])
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
        t = search.lower()
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
# HELPERS
# ---------------------------------------------------------------------------

def pct(v): return f"{v:.1f}%"

def h(t):
    return str(t).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def count_by_order(df, col, order):
    counts = df[col].value_counts().reindex(order, fill_value=0).reset_index()
    counts.columns = [col, "Count"]
    extra = df[~df[col].isin(order)][col].value_counts().reset_index()
    extra.columns = [col, "Count"]
    return pd.concat([counts, extra], ignore_index=True)

def style_fig(fig, height=280):
    fig.update_layout(
        height=height,
        margin=dict(l=16, r=16, t=24, b=16),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#cbd5e1", family="Segoe UI, Inter, sans-serif", size=11),
        legend=dict(
            font=dict(color="#94a3b8", size=11),
            bgcolor="rgba(0,0,0,0)",
        ),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#334155", zeroline=False, color="#64748b", tickfont=dict(color="#94a3b8"))
    fig.update_yaxes(showgrid=True, gridcolor="#334155", zeroline=False, color="#64748b", tickfont=dict(color="#94a3b8"))
    return fig

def donut(df, names, values, colors):
    df = df[df[values] > 0].copy()
    seq = [colors.get(str(n), "#64748b") for n in df[names]]
    fig = px.pie(df, names=names, values=values, hole=0.60, color_discrete_sequence=seq)
    fig.update_traces(
        textposition="inside", textinfo="percent",
        textfont=dict(size=12, color="#ffffff"),
        insidetextorientation="radial",
        marker=dict(line=dict(color="#1e293b", width=2)),
        hovertemplate="<b>%{label}</b><br>%{value} projects (%{percent})<extra></extra>",
    )
    total = int(df[values].sum())
    fig.update_layout(
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5,
                    font=dict(size=11, color="#94a3b8"), bgcolor="rgba(0,0,0,0)"),
        annotations=[dict(text=f"<b>{total}</b>", x=0.5, y=0.5,
                          font=dict(size=18, color="#f1f5f9"), showarrow=False)],
        margin=dict(l=8, r=8, t=8, b=50),
        height=260,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#94a3b8", family="Segoe UI, Inter, sans-serif", size=11),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def hbar(df, x, y, colors_list, height=280):
    fig = px.bar(df, x=x, y=y, orientation="h", text=x)
    if isinstance(colors_list, list):
        fig.update_traces(marker_color=colors_list, marker_line_width=0,
                          textfont=dict(color="#94a3b8", size=11), textposition="outside")
    else:
        fig.update_traces(marker_color=colors_list, marker_line_width=0,
                          textfont=dict(color="#94a3b8", size=11), textposition="outside")
    fig = style_fig(fig, height=height)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def vbar(df, x, y, color="#3b82f6", height=260):
    fig = px.bar(df, x=x, y=y, text=y)
    fig.update_traces(marker_color=color, marker_line_width=0, borderRadius=6,
                      textfont=dict(color="#94a3b8", size=11), textposition="outside")
    fig = style_fig(fig, height=height)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def mini_bar(val, cls):
    p = max(0, min(100, float(val or 0)))
    return (f'<div class="bar-cell">'
            f'<div class="bar-bg-mini"><div class="bar-fill-mini {cls}" style="width:{p:.0f}%"></div></div>'
            f'<div class="bar-pct">{p:.0f}%</div></div>')

def badge(text, css_class):
    return f'<span class="badge {css_class}">{h(text)}</span>'

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
st_autorefresh(interval=REFRESH_SECONDS * 1000, key="refresh")

try:
    file_bytes, source_name = get_excel_bytes(None)
    df = parse_excel(file_bytes)
except Exception as exc:
    st.error(f"Could not load data: {exc}")
    st.stop()

filtered = apply_filters(df)

total = len(filtered)
avg_prog  = filtered["Overall_Progress"].mean() if total else 0
n_comp    = int((filtered["Status"] == "Completed").sum())
n_prog    = int((filtered["Status"] == "In Progress").sum())
n_hold    = int((filtered["Status"] == "On Hold").sum())
n_ns      = int((filtered["Status"] == "Not Started").sum())
n_cancel  = int((filtered["Status"] == "Cancelled").sum())
n_mat_del = int((filtered["Material_Status"] == "Delivered").sum())
eng_avg   = filtered["Engineering_Pct"].mean() if total else 0
del_avg   = filtered["Delivery_Pct"].mean() if total else 0
exec_avg  = filtered["Execution_Pct"].mean() if total else 0
eng_done  = int((filtered["Engineering_Pct"] >= 100).sum())
del_done  = int((filtered["Delivery_Pct"] >= 100).sum())
exec_done = int((filtered["Execution_Pct"] >= 100).sum())
now_str   = pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")

# HEADER
st.html(f"""
<div class="dash-header">
  <div>
    <h1>Project Tracking Dashboard</h1>
    <div class="sub">Cold Rooms, Cabinets &amp; Refrigeration Projects - Live view from Excel</div>
  </div>
  <div class="right">
    <div>Data as of: {h(now_str)}</div>
    <div>Source: {h(source_name)} | Total: {total} projects</div>
  </div>
</div>
""")

# KPIs
pct_comp = f"{round(n_comp/total*100)}%" if total else ""
pct_prog = f"{round(n_prog/total*100)}%" if total else ""
pct_mat  = f"{round(n_mat_del/total*100)}%" if total else ""

st.html(f"""
<div class="kpi-grid">
  <div class="kpi" style="--ac:#3b82f6"><div class="lbl">Total Projects</div><div class="val">{total}</div><div class="delta">Showing {total} of {len(df)}</div></div>
  <div class="kpi" style="--ac:#22c55e"><div class="lbl">Completed</div><div class="val">{n_comp}</div><div class="delta">{pct_comp}</div></div>
  <div class="kpi" style="--ac:#14b8a6"><div class="lbl">In Progress</div><div class="val">{n_prog}</div><div class="delta">{pct_prog}</div></div>
  <div class="kpi" style="--ac:#f97316"><div class="lbl">On Hold</div><div class="val">{n_hold}</div><div class="delta"></div></div>
  <div class="kpi" style="--ac:#64748b"><div class="lbl">Not Started</div><div class="val">{n_ns}</div><div class="delta"></div></div>
  <div class="kpi" style="--ac:#ef4444"><div class="lbl">Cancelled</div><div class="val">{n_cancel}</div><div class="delta"></div></div>
  <div class="kpi" style="--ac:#ec4899"><div class="lbl">Overall Avg %</div><div class="val">{avg_prog:.1f}%</div><div class="delta"></div></div>
  <div class="kpi" style="--ac:#eab308"><div class="lbl">Mat. Delivered</div><div class="val">{n_mat_del}</div><div class="delta">{pct_mat}</div></div>
</div>
""")

# PHASE CARDS
st.html('<div class="sec-title">Phase Progress Overview</div>')
st.html(f"""
<div class="phase-grid">
  <div class="phase-card" style="--ac:#3b82f6">
    <div class="ph-title">Engineering <span class="ph-pct" style="color:#60a5fa">{eng_avg:.1f}%</span></div>
    <div class="ph-bar-bg"><div class="ph-bar-fill" style="width:{eng_avg:.1f}%;background:linear-gradient(90deg,#3b82f6,#60a5fa)"></div></div>
    <div class="ph-stats"><span>Design / Submittal / Drawing / ELS / BOM</span><span><strong>{eng_done}</strong>/{total} done</span></div>
  </div>
  <div class="phase-card" style="--ac:#a855f7">
    <div class="ph-title">Delivery <span class="ph-pct" style="color:#c084fc">{del_avg:.1f}%</span></div>
    <div class="ph-bar-bg"><div class="ph-bar-fill" style="width:{del_avg:.1f}%;background:linear-gradient(90deg,#a855f7,#c084fc)"></div></div>
    <div class="ph-stats"><span>Material delivery - {len(DELIVERY_COLS)} items</span><span><strong>{del_done}</strong>/{total} done</span></div>
  </div>
  <div class="phase-card" style="--ac:#f97316">
    <div class="ph-title">Execution <span class="ph-pct" style="color:#fb923c">{exec_avg:.1f}%</span></div>
    <div class="ph-bar-bg"><div class="ph-bar-fill" style="width:{exec_avg:.1f}%;background:linear-gradient(90deg,#f97316,#fb923c)"></div></div>
    <div class="ph-stats"><span>On-site installation, commissioning</span><span><strong>{exec_done}</strong>/{total} done</span></div>
  </div>
</div>
""")

# TABS
tab_ov, tab_pr, tab_dt = st.tabs(["  Overview  ", "  Progress Analysis  ", "  Project Details  "])

with tab_ov:
    # Row 1: 3 donuts
    c1, c2, c3 = st.columns(3)
    with c1:
        st.html('<div class="card"><h3>Project Status</h3>')
        donut(count_by_order(filtered, "Status", STATUS_ORDER), "Status", "Count", STATUS_COLORS)
        st.html('</div>')
    with c2:
        st.html('<div class="card"><h3>Material Status</h3>')
        donut(count_by_order(filtered, "Material_Status", MATERIAL_ORDER), "Material_Status", "Count", MATERIAL_COLORS)
        st.html('</div>')
    with c3:
        st.html('<div class="card"><h3>Priority</h3>')
        donut(count_by_order(filtered, "Priority", PRIORITY_ORDER), "Priority", "Count", PRIORITY_COLORS)
        st.html('</div>')

    # Row 2: Region + Customers
    c4, c5 = st.columns(2)
    with c4:
        st.html('<div class="card"><h3>Region Breakdown</h3>')
        rdf = filtered["Region"].value_counts().reset_index()
        rdf.columns = ["Region", "Count"]
        rdf = rdf.sort_values("Count", ascending=False)
        palette = ["#3b82f6","#a855f7","#14b8a6","#ec4899","#f97316","#64748b","#eab308","#22c55e"]
        clrs = [palette[i % len(palette)] for i in range(len(rdf))]
        vbar(rdf, "Region", "Count",
             color=[palette[i % len(palette)] for i in range(len(rdf))],
             height=280)
        st.html('</div>')
    with c5:
        st.html('<div class="card"><h3>Top Customers</h3>')
        cdf = filtered["Customer"].value_counts().head(10).reset_index()
        cdf.columns = ["Customer", "Count"]
        cdf = cdf.sort_values("Count", ascending=True)
        hbar(cdf, "Count", "Customer", "#3b82f6", height=300)
        st.html('</div>')

with tab_pr:
    # Delivery items (horizontal stacked)
    st.html('<div class="card"><h3>Delivery Items Status <span class="tag">DONE across all projects</span></h3>')
    done_v    = [int((filtered[c].fillna("").astype(str).str.strip().str.upper() == "DONE").sum())      for c in DELIVERY_COLS]
    part_v    = [int((filtered[c].fillna("").astype(str).str.strip().str.upper() == "PART.DONE").sum()) for c in DELIVERY_COLS]
    ndone_v   = [int((filtered[c].fillna("").astype(str).str.strip().str.upper() == "N.DONE").sum())    for c in DELIVERY_COLS]
    na_v      = [int((filtered[c].fillna("").astype(str).str.strip().str.upper() == "N/A").sum())       for c in DELIVERY_COLS]
    ddf = pd.DataFrame({"Item": DELIVERY_COLS, "Done": done_v, "Partial": part_v, "Not Done": ndone_v, "N/A": na_v})
    fig = go.Figure()
    fig.add_trace(go.Bar(y=ddf["Item"], x=ddf["Done"],     name="Done",     orientation="h", marker_color="#22c55e", marker_line_width=0))
    fig.add_trace(go.Bar(y=ddf["Item"], x=ddf["Partial"],  name="Partial",  orientation="h", marker_color="#eab308", marker_line_width=0))
    fig.add_trace(go.Bar(y=ddf["Item"], x=ddf["Not Done"], name="Not Done", orientation="h", marker_color="#ef4444", marker_line_width=0))
    fig.add_trace(go.Bar(y=ddf["Item"], x=ddf["N/A"],      name="N/A",      orientation="h", marker_color="#64748b", marker_line_width=0))
    fig.update_layout(barmode="stack", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                      font=dict(size=11, color="#94a3b8"), bgcolor="rgba(0,0,0,0)"))
    fig = style_fig(fig, height=360)
    fig.update_layout(xaxis_title="Number of Projects", margin=dict(l=16,r=16,t=36,b=16))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    st.html('</div>')

    # Progress buckets
    c6, c7 = st.columns([3, 2])
    with c6:
        st.html('<div class="card"><h3>Phase Progress by Project <span class="tag">Engineering / Delivery / Execution</span></h3>')
        ph_df = filtered.sort_values("Overall_Progress", ascending=False).head(25)
        labels_pp = [f"#{r.S_No} {(r.Project_Name or r.Customer)[:28]}" for _, r in ph_df.iterrows()]
        fig = go.Figure()
        fig.add_trace(go.Bar(y=labels_pp, x=ph_df["Engineering_Pct"].values, name="Engineering %", orientation="h", marker_color="#3b82f6", marker_line_width=0, borderRadius=3))
        fig.add_trace(go.Bar(y=labels_pp, x=ph_df["Delivery_Pct"].values,    name="Delivery %",    orientation="h", marker_color="#a855f7", marker_line_width=0, borderRadius=3))
        fig.add_trace(go.Bar(y=labels_pp, x=ph_df["Execution_Pct"].values,   name="Execution %",   orientation="h", marker_color="#f97316", marker_line_width=0, borderRadius=3))
        fig.update_layout(
            barmode="group",
            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1,
                        font=dict(size=11, color="#94a3b8"), bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(range=[0, 100], title="Percent Complete"),
            yaxis=dict(autorange="reversed"),
        )
        fig = style_fig(fig, height=max(420, len(ph_df) * 26))
        fig.update_layout(margin=dict(l=16, r=16, t=40, b=16))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        st.html('</div>')
    with c7:
        st.html('<div class="card"><h3>Overall Progress Buckets</h3>')
        bins   = [-0.1, 25, 50, 75, 100]
        labels = ["0-25%", "26-50%", "51-75%", "76-100%"]
        bkt    = pd.cut(filtered["Overall_Progress"], bins=bins, labels=labels)
        bkt_df = bkt.value_counts().reindex(labels, fill_value=0).reset_index()
        bkt_df.columns = ["Bucket", "Count"]
        bkt_clr = ["#ef4444", "#f97316", "#eab308", "#22c55e"]
        fig = go.Figure()
        for i, row in bkt_df.iterrows():
            fig.add_trace(go.Bar(
                x=[row["Bucket"]], y=[row["Count"]],
                name=row["Bucket"], marker_color=bkt_clr[i],
                marker_line_width=0, text=[str(int(row["Count"]))],
                textposition="outside", showlegend=False,
                textfont=dict(color="#94a3b8"),
            ))
        fig = style_fig(fig, height=300)
        fig.update_layout(xaxis_title="Progress Range", yaxis_title="Projects",
                          margin=dict(l=16, r=16, t=24, b=16))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        st.html('</div>')

with tab_dt:
    st.html('<div class="card"><h3>Project Details <span class="tag" id="rowCount"></span></h3>')

    # Filters row
    fc1, fc2, fc3, fc4, fc5, fc6 = st.columns([2.5, 1.5, 1.5, 1.5, 1.5, 1.5])
    with fc1: srch = st.text_input("Search", placeholder="Project, customer, location...", label_visibility="collapsed", key="tbl_srch")
    with fc2: f_st = st.selectbox("Status",   ["All Statuses"]   + sorted(df["Status"].dropna().unique().tolist()),         label_visibility="collapsed", key="tbl_st")
    with fc3: f_rg = st.selectbox("Region",   ["All Regions"]    + sorted(df["Region"].dropna().unique().tolist()),          label_visibility="collapsed", key="tbl_rg")
    with fc4: f_cu = st.selectbox("Customer", ["All Customers"]  + sorted(df["Customer"].dropna().unique().tolist()),        label_visibility="collapsed", key="tbl_cu")
    with fc5: f_mt = st.selectbox("Material", ["All Material"]   + sorted(df["Material_Status"].dropna().unique().tolist()), label_visibility="collapsed", key="tbl_mt")
    with fc6: f_pr = st.selectbox("Priority", ["All Priorities"] + sorted(df["Priority"].dropna().unique().tolist()),        label_visibility="collapsed", key="tbl_pr")

    tbl = filtered.copy()
    if srch:
        t = srch.lower()
        mask = (
            tbl["Project_Name"].str.lower().str.contains(t, na=False)
            | tbl["Customer"].str.lower().str.contains(t, na=False)
            | tbl["Location"].str.lower().str.contains(t, na=False)
        )
        tbl = tbl[mask]
    if f_st != "All Statuses":   tbl = tbl[tbl["Status"] == f_st]
    if f_rg != "All Regions":    tbl = tbl[tbl["Region"] == f_rg]
    if f_cu != "All Customers":  tbl = tbl[tbl["Customer"] == f_cu]
    if f_mt != "All Material":   tbl = tbl[tbl["Material_Status"] == f_mt]
    if f_pr != "All Priorities": tbl = tbl[tbl["Priority"] == f_pr]
    tbl = tbl.sort_values("Overall_Progress", ascending=False)

    status_cls  = {"Completed": "badge-completed", "In Progress": "badge-progress", "On Hold": "badge-hold", "Cancelled": "badge-cancelled", "Not Started": "badge-notstarted"}
    material_cls= {"Delivered": "badge-delivered", "Partially Delivered": "badge-partial", "Ordered": "badge-ordered", "Not Ordered": "badge-notordered"}
    priority_cls= {"High": "badge-high", "Medium": "badge-medium", "Low": "badge-low"}

    rows_html = ""
    for _, r in tbl.iterrows():
        rows_html += (
            f"<tr>"
            f"<td>{h(r.S_No)}</td>"
            f"<td>{h(r.Customer)}</td>"
            f"<td>{h(r.Project_Name)}</td>"
            f"<td>{h(r.Region)}</td>"
            f"<td>{badge(r.Status, status_cls.get(r.Status,''))}</td>"
            f"<td>{mini_bar(r.Engineering_Pct, 'bar-eng')}</td>"
            f"<td>{mini_bar(r.Delivery_Pct, 'bar-del')}</td>"
            f"<td>{mini_bar(r.Execution_Pct, 'bar-exec')}</td>"
            f"<td>{mini_bar(r.Overall_Progress, 'bar-all')}</td>"
            f"<td>{badge(r.Material_Status, material_cls.get(r.Material_Status,''))}</td>"
            f"<td>{badge(r.Priority, priority_cls.get(r.Priority,''))}</td>"
            f"</tr>"
        )

    st.html(f"""
    <div style="font-size:12px;color:#94a3b8;margin-bottom:8px">{len(tbl)} rows</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr>
          <th>#</th><th>Customer</th><th>Project</th><th>Region</th><th>Status</th>
          <th>Eng %</th><th>Delivery %</th><th>Exec %</th><th>Overall</th>
          <th>Material</th><th>Priority</th>
        </tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    """)
    st.html('</div>')

    csv = tbl[["S_No","Customer","Project_Name","Region","Location","Status",
               "Engineering_Pct","Delivery_Pct","Execution_Pct","Overall_Progress",
               "Material_Status","Priority","Work_Status"]].to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered list as CSV", data=csv, file_name="projects.csv", mime="text/csv")
