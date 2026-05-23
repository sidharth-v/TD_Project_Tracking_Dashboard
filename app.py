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
except Exception:  # msal is only needed for Microsoft Graph mode
    msal = None

st.set_page_config(page_title="Project Tracking Dashboard", layout="wide")

# ---------------- CONFIG ----------------
LOCAL_EXCEL_FILE = "Project_Tracking_v7.xlsx"
SHEET_NAME = "Project_Master"
REFRESH_SECONDS = int(st.secrets.get("REFRESH_SECONDS", 60)) if hasattr(st, "secrets") else 60

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

# ---------------- STYLE ----------------
st.markdown(
    """
    <style>
    .stApp { background: linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%); color:#f1f5f9; }
    [data-testid="stMetric"] { background:#1e293b; padding:16px; border-radius:14px; border-left:4px solid #3b82f6; }
    [data-testid="stMetricLabel"] { color:#94a3b8; }
    .phase-card { background:#1e293b; border-radius:14px; padding:18px; margin-bottom:8px; }
    .phase-title { display:flex; justify-content:space-between; align-items:center; font-weight:700; }
    .phase-sub { color:#94a3b8; font-size:12px; margin-top:8px; }
    .bar-bg { height:10px; background:#334155; border-radius:8px; overflow:hidden; margin-top:12px; }
    .bar-fill { height:10px; border-radius:8px; }
    div[data-testid="stDataFrame"] { background:#1e293b; border-radius:12px; }
    section[data-testid="stSidebar"] { background:#111827; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------- DATA SOURCE ----------------
def _secret(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, default)).strip()
    except Exception:
        return default


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
            "The OneDrive URL returned an HTML page, not the Excel file. "
            "Use a direct download/share link or Microsoft Graph mode."
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
    """
    Business OneDrive / SharePoint app-only mode.
    Required secrets:
      GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET,
      GRAPH_USER_ID, ONEDRIVE_FILE_PATH
    The Azure app needs Microsoft Graph Application permission Files.Read.All.
    """
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
    # 1) Manual upload always wins, for testing.
    if uploaded_file is not None:
        return uploaded_file.read(), uploaded_file.name

    # 2) Public/share OneDrive direct-download URL.
    excel_url = _secret("EXCEL_FILE_URL")
    if excel_url:
        return load_excel_from_direct_url(excel_url), "OneDrive direct URL"

    # 3) Private business OneDrive/SharePoint through Microsoft Graph.
    if _secret("GRAPH_USER_ID") and _secret("ONEDRIVE_FILE_PATH"):
        return load_excel_from_graph(), "Microsoft Graph OneDrive"

    # 4) Local file for development.
    local_path = Path(LOCAL_EXCEL_FILE)
    if local_path.exists():
        return local_path.read_bytes(), LOCAL_EXCEL_FILE

    st.error(
        "No Excel source found. Upload the workbook, set EXCEL_FILE_URL, "
        "or configure Microsoft Graph secrets."
    )
    st.stop()

# ---------------- PARSING ----------------
def _clean_text(value, default="") -> str:
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
    return data.reset_index(drop=True)

# ---------------- FILTERS ----------------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar:
        st.header("Filters")
        search = st.text_input("Search project, customer, location")

        quick = st.radio(
            "Quick Views",
            ["All", "Completed", "In Progress", "On Hold", "Not Started"],
            horizontal=False,
        )

        status = st.multiselect("Status", sorted(df["Status"].dropna().unique()), default=[])
        region = st.multiselect("Region", sorted(df["Region"].dropna().unique()), default=[])
        customer = st.multiselect("Customer", sorted(df["Customer"].dropna().unique()), default=[])
        material = st.multiselect("Material", sorted(df["Material_Status"].dropna().unique()), default=[])
        priority = st.multiselect("Priority", sorted(df["Priority"].dropna().unique()), default=[])

    out = df.copy()
    if quick != "All":
        out = out[out["Status"] == quick]
    if search:
        text = search.lower()
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

# ---------------- UI HELPERS ----------------
def count_by_order(df: pd.DataFrame, col: str, order: Iterable[str]) -> pd.DataFrame:
    counts = df[col].value_counts().reindex(order, fill_value=0).reset_index()
    counts.columns = [col, "Count"]
    extra = df[~df[col].isin(order)][col].value_counts().reset_index()
    extra.columns = [col, "Count"]
    return pd.concat([counts, extra], ignore_index=True)


def render_phase_card(title: str, value: float, subtitle: str, color: str):
    st.markdown(
        f"""
        <div class="phase-card">
            <div class="phase-title"><span>{title}</span><span style="color:{color};font-size:26px;">{value:.1f}%</span></div>
            <div class="bar-bg"><div class="bar-fill" style="width:{value:.1f}%; background:{color};"></div></div>
            <div class="phase-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def pie_chart(df: pd.DataFrame, names: str, values: str, title: str):
    fig = px.pie(df, names=names, values=values, hole=0.55, title=title)
    fig.update_layout(
        height=330,
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#f1f5f9",
        legend_font_color="#f1f5f9",
    )
    st.plotly_chart(fig, use_container_width=True)


def bar_chart(df: pd.DataFrame, x: str, y: str, title: str, orientation="v"):
    fig = px.bar(df, x=x, y=y, title=title, orientation=orientation)
    fig.update_layout(
        height=360,
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#f1f5f9",
    )
    st.plotly_chart(fig, use_container_width=True)

# ---------------- MAIN ----------------
st_autorefresh(interval=REFRESH_SECONDS * 1000, key="refresh")

st.title("Project Tracking Dashboard")
st.caption("Cold Rooms, Cabinets & Refrigeration Projects — live view from Excel")

uploaded = st.file_uploader("Optional: upload Excel file for testing", type=["xlsx", "xlsm", "xls"])
file_bytes, source_name = get_excel_bytes(uploaded)
df = parse_excel(file_bytes)
filtered = apply_filters(df)

st.caption(
    f"Current data: **{source_name}** · Auto-refresh: **{REFRESH_SECONDS}s** · "
    f"Showing **{len(filtered)} / {len(df)}** projects"
)

# KPI row
kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
kpi1.metric("Total Projects", len(filtered))
kpi2.metric("Completed", int((filtered["Status"] == "Completed").sum()))
kpi3.metric("In Progress", int((filtered["Status"] == "In Progress").sum()))
kpi4.metric("On Hold", int((filtered["Status"] == "On Hold").sum()))
kpi5.metric("Not Started", int((filtered["Status"] == "Not Started").sum()))
kpi6.metric("Avg Progress", f"{filtered['Overall_Progress'].mean() if len(filtered) else 0:.1f}%")

st.subheader("Phase Progress Overview")
c1, c2, c3 = st.columns(3)
with c1:
    render_phase_card(
        "Engineering",
        filtered["Engineering_Pct"].mean() if len(filtered) else 0,
        "Design / Submittal / Drawing / ELS / BOM",
        "#60a5fa",
    )
with c2:
    render_phase_card(
        "Delivery",
        filtered["Delivery_Pct"].mean() if len(filtered) else 0,
        f"Material delivery — {len(DELIVERY_COLS)} items",
        "#c084fc",
    )
with c3:
    render_phase_card(
        "Execution",
        filtered["Execution_Pct"].mean() if len(filtered) else 0,
        "On-site installation and commissioning",
        "#fb923c",
    )

# Charts
r1c1, r1c2, r1c3 = st.columns(3)
with r1c1:
    pie_chart(count_by_order(filtered, "Status", STATUS_ORDER), "Status", "Count", "Project Status")
with r1c2:
    pie_chart(count_by_order(filtered, "Material_Status", MATERIAL_ORDER), "Material_Status", "Count", "Material Status")
with r1c3:
    pie_chart(count_by_order(filtered, "Priority", PRIORITY_ORDER), "Priority", "Count", "Priority")

r2c1, r2c2 = st.columns(2)
with r2c1:
    region_df = filtered["Region"].value_counts().reset_index()
    region_df.columns = ["Region", "Count"]
    bar_chart(region_df, "Region", "Count", "Region Breakdown")
with r2c2:
    cust_df = filtered["Customer"].value_counts().head(10).reset_index()
    cust_df.columns = ["Customer", "Count"]
    bar_chart(cust_df.sort_values("Count"), "Count", "Customer", "Top Customers", orientation="h")

r3c1, r3c2 = st.columns([7, 5])
with r3c1:
    delivery_done = []
    delivery_partial = []
    for col in DELIVERY_COLS:
        vals = filtered[col].fillna("").astype(str).str.strip().str.upper()
        delivery_done.append(int((vals == "DONE").sum()))
        delivery_partial.append(int((vals == "PART.DONE").sum()))
    delivery_df = pd.DataFrame({"Item": DELIVERY_COLS, "Done": delivery_done, "Partial": delivery_partial})
    fig = go.Figure()
    fig.add_trace(go.Bar(x=delivery_df["Item"], y=delivery_df["Done"], name="Done"))
    fig.add_trace(go.Bar(x=delivery_df["Item"], y=delivery_df["Partial"], name="Partial"))
    fig.update_layout(
        title="Delivery Items Status",
        barmode="stack",
        height=380,
        margin=dict(l=20, r=20, t=50, b=90),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#f1f5f9",
    )
    st.plotly_chart(fig, use_container_width=True)

with r3c2:
    bins = [-0.1, 25, 50, 75, 99.999, 100]
    labels = ["0-25%", "26-50%", "51-75%", "76-99%", "100%"]
    bucket = pd.cut(filtered["Overall_Progress"], bins=bins, labels=labels)
    bucket_df = bucket.value_counts().reindex(labels, fill_value=0).reset_index()
    bucket_df.columns = ["Progress Bucket", "Count"]
    pie_chart(bucket_df, "Progress Bucket", "Count", "Overall Progress Buckets")

st.subheader("Phase Progress by Project")
phase_df = filtered.sort_values("Overall_Progress", ascending=False).head(25)
fig = go.Figure()
fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Engineering_Pct"], name="Engineering", orientation="h"))
fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Delivery_Pct"], name="Delivery", orientation="h"))
fig.add_trace(go.Bar(y=phase_df["Project_Name"], x=phase_df["Execution_Pct"], name="Execution", orientation="h"))
fig.update_layout(
    barmode="group",
    height=650,
    xaxis_title="%",
    yaxis_title="",
    margin=dict(l=20, r=20, t=30, b=20),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font_color="#f1f5f9",
    yaxis={"autorange": "reversed"},
)
st.plotly_chart(fig, use_container_width=True)

st.subheader("Project Details")
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
st.dataframe(
    filtered[show_cols].sort_values("Overall_Progress", ascending=False),
    use_container_width=True,
    hide_index=True,
)
