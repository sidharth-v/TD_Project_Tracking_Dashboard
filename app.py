import base64
import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from streamlit_autorefresh import st_autorefresh


# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(
    page_title="Project Tracking Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# -----------------------------
# Force same light UI everywhere
# -----------------------------
st.markdown(
    """
<style>
:root {
    color-scheme: light !important;
}

html, body, .stApp, [data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #f4f7fb 0%, #eef2ff 100%) !important;
    color: #111827 !important;
}

[data-testid="stHeader"] {
    background: #ffffff !important;
}

[data-testid="stToolbar"] {
    background: #ffffff !important;
}

[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e5e7eb !important;
}

[data-testid="stSidebar"] * {
    color: #111827 !important;
}

.block-container {
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
    max-width: 1500px !important;
}

h1, h2, h3, h4, h5, h6, p, label, span, div {
    color: #111827;
}

.main-hero {
    background: #ffffff;
    border: 1px solid #dbe3ef;
    border-radius: 18px;
    padding: 26px 28px;
    margin-bottom: 18px;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.08);
}

.main-hero h1 {
    margin: 0;
    font-size: 34px;
    line-height: 1.15;
    font-weight: 800;
    color: #0f172a;
}

.main-hero .sub {
    color: #475569;
    margin-top: 10px;
    font-size: 16px;
}

.pill-row {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 22px;
}

.pill {
    padding: 9px 14px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 700;
    border: 1px solid #dbe3ef;
    background: #f8fafc;
    color: #334155;
}

.pill.green {
    background: #ecfdf5;
    color: #166534;
    border-color: #bbf7d0;
}

.pill.blue {
    background: #eff6ff;
    color: #1d4ed8;
    border-color: #bfdbfe;
}

.pill.orange {
    background: #fff7ed;
    color: #c2410c;
    border-color: #fed7aa;
}

.section-title {
    font-size: 15px;
    color: #475569;
    margin: 22px 0 12px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 800;
}

.kpi-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 20px;
    border: 1px solid #dbe3ef;
    box-shadow: 0 8px 22px rgba(15, 23, 42, 0.07);
    min-height: 135px;
    border-top: 5px solid #2563eb;
}

.kpi-card.green { border-top-color: #16a34a; }
.kpi-card.teal { border-top-color: #0891b2; }
.kpi-card.orange { border-top-color: #f97316; }
.kpi-card.gray { border-top-color: #64748b; }
.kpi-card.purple { border-top-color: #7c3aed; }
.kpi-card.red { border-top-color: #dc2626; }

.kpi-label {
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #64748b;
    font-weight: 800;
}

.kpi-value {
    font-size: 34px;
    margin-top: 16px;
    font-weight: 850;
    color: #0f172a;
}

.kpi-note {
    font-size: 13px;
    margin-top: 8px;
    color: #64748b;
}

.phase-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 20px;
    border: 1px solid #dbe3ef;
    box-shadow: 0 8px 22px rgba(15, 23, 42, 0.07);
    min-height: 160px;
    border-top: 5px solid #2563eb;
}

.phase-card.del { border-top-color: #7c3aed; }
.phase-card.exec { border-top-color: #f97316; }

.phase-top {
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.phase-title {
    font-size: 16px;
    font-weight: 800;
    color: #111827;
}

.phase-pct {
    font-size: 32px;
    font-weight: 850;
    color: #111827;
}

.phase-bar-bg {
    width: 100%;
    height: 12px;
    background: #e5e7eb;
    border-radius: 999px;
    overflow: hidden;
    margin-top: 18px;
}

.phase-bar-fill {
    height: 100%;
    border-radius: 999px;
    background: linear-gradient(90deg, #2563eb, #60a5fa);
}

.phase-card.del .phase-bar-fill {
    background: linear-gradient(90deg, #7c3aed, #c084fc);
}

.phase-card.exec .phase-bar-fill {
    background: linear-gradient(90deg, #f97316, #fb923c);
}

.phase-stats {
    margin-top: 14px;
    display: flex;
    justify-content: space-between;
    color: #64748b;
    font-size: 13px;
}

.chart-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 18px;
    border: 1px solid #dbe3ef;
    box-shadow: 0 8px 22px rgba(15, 23, 42, 0.07);
    margin-bottom: 18px;
}

.chart-title {
    font-size: 18px;
    font-weight: 850;
    color: #111827;
    margin-bottom: 8px;
}

.chart-tag {
    font-size: 12px;
    color: #64748b;
    font-weight: 600;
}

.badge {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 800;
}

.badge-completed {
    background: #dcfce7;
    color: #166534;
}

.badge-progress {
    background: #cffafe;
    color: #155e75;
}

.badge-hold {
    background: #ffedd5;
    color: #9a3412;
}

.badge-notstarted {
    background: #e2e8f0;
    color: #334155;
}

.badge-cancelled {
    background: #fee2e2;
    color: #991b1b;
}

.badge-high {
    background: #fee2e2;
    color: #991b1b;
}

.badge-medium {
    background: #fef9c3;
    color: #854d0e;
}

.badge-low {
    background: #dcfce7;
    color: #166534;
}

.badge-delivered {
    background: #dcfce7;
    color: #166534;
}

.badge-partial {
    background: #fef9c3;
    color: #854d0e;
}

.badge-ordered {
    background: #ffedd5;
    color: #9a3412;
}

.badge-notordered {
    background: #fee2e2;
    color: #991b1b;
}

[data-testid="stDataFrame"] {
    background: #ffffff !important;
    border-radius: 16px !important;
}

div[data-testid="stMultiSelect"] div {
    color: #111827 !important;
}

div[data-baseweb="select"] > div {
    background: #ffffff !important;
    color: #111827 !important;
}

input {
    background: #ffffff !important;
    color: #111827 !important;
}

hr {
    border-color: #dbe3ef;
}
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------
# Auto refresh
# -----------------------------
REFRESH_SECONDS = int(st.secrets.get("REFRESH_SECONDS", 60))
st_autorefresh(interval=REFRESH_SECONDS * 1000, key="dashboard_refresh")


# -----------------------------
# Column mapping
# Based on your Project_Master structure
# -----------------------------
BASE_COLUMNS = [
    "S_No",
    "Unknown_1",
    "Job_Ref",
    "LPO_Ref",
    "Customer",
    "Project_Name",
    "Region",
    "Location",
    "Unknown_8",
    "Equi_Design",
    "Technical_Submittal",
    "Drawing",
    "ELS",
    "BOM",
    "Out_Door",
    "Indoor",
    "CR_Panels",
    "CR_Ins_Materials",
    "Doors",
    "Ref_Inst_Materials",
    "CCP",
    "Display_CCP",
    "Floor_Heater",
    "Cabinets",
    "Any_Special",
    "Work_Status",
    "Unknown_26",
    "Material_Status",
    "Overall_Progress",
    "Priority",
    "Status",
    "Engineering_Pct",
    "Delivery_Pct",
    "Execution_Pct",
]

ENGINEERING_ITEMS = [
    "Equi_Design",
    "Technical_Submittal",
    "Drawing",
    "ELS",
    "BOM",
]

DELIVERY_ITEMS = [
    "Out_Door",
    "Indoor",
    "CR_Panels",
    "CR_Ins_Materials",
    "Doors",
    "Ref_Inst_Materials",
    "CCP",
    "Display_CCP",
    "Floor_Heater",
    "Cabinets",
    "Any_Special",
]


# -----------------------------
# Helpers
# -----------------------------
def clean_text(value, default=""):
    if pd.isna(value):
        return default
    value = str(value).strip()
    if value.lower() in ["nan", "none"]:
        return default
    return value


def clean_status(value):
    value = clean_text(value, "Not Started").strip().title()

    replacements = {
        "Inprogress": "In Progress",
        "In-Progress": "In Progress",
        "Onhold": "On Hold",
        "On-Hold": "On Hold",
        "Notstarted": "Not Started",
        "Not-Started": "Not Started",
    }

    value = replacements.get(value, value)

    allowed = ["Completed", "In Progress", "On Hold", "Not Started", "Cancelled"]
    if value not in allowed:
        return "Not Started"

    return value


def clean_material(value):
    value = clean_text(value, "Not Ordered").strip().title()

    replacements = {
        "Partial": "Partially Delivered",
        "Partial Delivered": "Partially Delivered",
        "Partially Delivered": "Partially Delivered",
        "Notordered": "Not Ordered",
        "Not Ordered": "Not Ordered",
    }

    value = replacements.get(value, value)

    allowed = ["Delivered", "Partially Delivered", "Ordered", "Not Ordered"]
    if value not in allowed:
        return "Not Ordered"

    return value


def clean_priority(value):
    value = clean_text(value, "Medium").strip().title()

    if value in ["High", "Medium", "Low"]:
        return value

    return "Medium"


def status_badge(status):
    status = clean_status(status)
    cls = {
        "Completed": "badge-completed",
        "In Progress": "badge-progress",
        "On Hold": "badge-hold",
        "Not Started": "badge-notstarted",
        "Cancelled": "badge-cancelled",
    }.get(status, "badge-notstarted")
    return f'<span class="badge {cls}">{status}</span>'


def priority_badge(priority):
    priority = clean_priority(priority)
    cls = {
        "High": "badge-high",
        "Medium": "badge-medium",
        "Low": "badge-low",
    }.get(priority, "badge-medium")
    return f'<span class="badge {cls}">{priority}</span>'


def material_badge(material):
    material = clean_material(material)
    cls = {
        "Delivered": "badge-delivered",
        "Partially Delivered": "badge-partial",
        "Ordered": "badge-ordered",
        "Not Ordered": "badge-notordered",
    }.get(material, "badge-notordered")
    return f'<span class="badge {cls}">{material}</span>'


def pct(value):
    try:
        return round(float(value), 1)
    except Exception:
        return 0.0


def convert_onedrive_link(url):
    """
    Works for many personal OneDrive share links.
    If direct download already works, it returns the same URL.
    """
    if not url:
        return url

    url = url.strip()

    if "download=1" in url:
        return url

    if "1drv.ms" in url:
        encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
        return f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"

    if "onedrive.live.com" in url and "download=1" not in url:
        joiner = "&" if "?" in url else "?"
        return f"{url}{joiner}download=1"

    return url


@st.cache_data(ttl=REFRESH_SECONDS)
def download_excel_from_url(url):
    final_url = convert_onedrive_link(url)
    response = requests.get(final_url, timeout=45)
    response.raise_for_status()
    return response.content


@st.cache_data(ttl=REFRESH_SECONDS)
def parse_excel(file_bytes):
    raw = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="Project_Master",
        header=None,
        engine="openpyxl",
    )

    raw = raw.iloc[:, : len(BASE_COLUMNS)].copy()
    raw.columns = BASE_COLUMNS

    df = raw.iloc[2:].copy()
    df = df.dropna(how="all")
    df = df[~(df["S_No"].isna() & df["Project_Name"].isna())].copy()

    text_cols = [
        "Job_Ref",
        "LPO_Ref",
        "Customer",
        "Project_Name",
        "Region",
        "Location",
        "Work_Status",
    ]

    for col in text_cols:
        df[col] = df[col].apply(lambda x: clean_text(x, "Unknown"))

    df["Customer"] = df["Customer"].replace("", "Unassigned")
    df["Project_Name"] = df["Project_Name"].replace("", "Unnamed Project")
    df["Region"] = df["Region"].replace("", "Unknown")
    df["Location"] = df["Location"].replace("", "Unknown")

    df["Status"] = df["Status"].apply(clean_status)
    df["Material_Status"] = df["Material_Status"].apply(clean_material)
    df["Priority"] = df["Priority"].apply(clean_priority)

    for col in ["Overall_Progress", "Engineering_Pct", "Delivery_Pct", "Execution_Pct"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).clip(0, 100)

    for col in ENGINEERING_ITEMS + DELIVERY_ITEMS:
        df[col] = df[col].apply(lambda x: clean_text(x, ""))

    df["Eng_Done"] = df[ENGINEERING_ITEMS].apply(lambda r: sum(v == "DONE" for v in r), axis=1)
    df["Eng_Part"] = df[ENGINEERING_ITEMS].apply(lambda r: sum(v == "PART.DONE" for v in r), axis=1)
    df["Eng_NA"] = df[ENGINEERING_ITEMS].apply(lambda r: sum(v == "N/A" for v in r), axis=1)

    df["Del_Done"] = df[DELIVERY_ITEMS].apply(lambda r: sum(v == "DONE" for v in r), axis=1)
    df["Del_Part"] = df[DELIVERY_ITEMS].apply(lambda r: sum(v == "PART.DONE" for v in r), axis=1)
    df["Del_NA"] = df[DELIVERY_ITEMS].apply(lambda r: sum(v == "N/A" for v in r), axis=1)

    return df.reset_index(drop=True)


def load_data():
    excel_url = st.secrets.get("EXCEL_FILE_URL", "")

    if not excel_url:
        st.error("No Excel source found. Add EXCEL_FILE_URL in Streamlit secrets.")
        st.stop()

    try:
        file_bytes = download_excel_from_url(excel_url)
        df = parse_excel(file_bytes)
        return df, "OneDrive live Excel"
    except Exception as e:
        st.error(f"Could not load dashboard data: {e}")
        st.stop()


# -----------------------------
# Chart helpers
# -----------------------------
CHART_COLORS = {
    "Completed": "#16a34a",
    "In Progress": "#0891b2",
    "On Hold": "#f97316",
    "Not Started": "#64748b",
    "Cancelled": "#dc2626",
    "Delivered": "#16a34a",
    "Partially Delivered": "#eab308",
    "Ordered": "#f97316",
    "Not Ordered": "#dc2626",
    "High": "#dc2626",
    "Medium": "#eab308",
    "Low": "#16a34a",
}


def chart_layout(fig, height=320, left=10, right=70):
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=left, r=right, t=10, b=10),
        showlegend=False,
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font=dict(color="#111827", size=13),
        xaxis=dict(
            title=None,
            showgrid=True,
            gridcolor="#e5e7eb",
            zeroline=False,
        ),
        yaxis=dict(
            title=None,
            autorange="reversed",
        ),
    )
    return fig


def make_count_bar(df, column, order=None, height=310):
    counts = (
        df[column]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace("", "Unknown")
        .value_counts()
        .reset_index()
    )

    counts.columns = [column, "Count"]

    if order:
        counts[column] = pd.Categorical(counts[column], categories=order, ordered=True)
        counts = counts.sort_values(column)
        counts = counts[counts["Count"].notna()]

    total = counts["Count"].sum()
    counts["Percent"] = (counts["Count"] / total * 100).round(1)
    counts["Label"] = counts["Count"].astype(str) + " (" + counts["Percent"].astype(str) + "%)"

    fig = px.bar(
        counts,
        x="Count",
        y=column,
        orientation="h",
        text="Label",
        color=column,
        color_discrete_map=CHART_COLORS,
    )

    fig.update_traces(
        textposition="outside",
        cliponaxis=False,
        marker_line_width=0,
    )

    return chart_layout(fig, height=height, right=95)


def make_top_bar(df, column, top_n=10, height=340, color="#2563eb"):
    counts = (
        df[column]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace("", "Unknown")
        .value_counts()
        .head(top_n)
        .reset_index()
    )

    counts.columns = [column, "Count"]

    fig = px.bar(
        counts,
        x="Count",
        y=column,
        orientation="h",
        text="Count",
    )

    fig.update_traces(
        textposition="outside",
        cliponaxis=False,
        marker_color=color,
    )

    return chart_layout(fig, height=height, right=60)


def make_delivery_items_chart(df):
    rows = []

    for col in DELIVERY_ITEMS:
        label = col.replace("_", " ")
        done = (df[col] == "DONE").sum()
        part = (df[col] == "PART.DONE").sum()
        not_done = (df[col] == "N.DONE").sum()

        rows.append(
            {
                "Item": label,
                "Done": done,
                "Part Done": part,
                "Not Done": not_done,
            }
        )

    chart_df = pd.DataFrame(rows)

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=chart_df["Item"],
            x=chart_df["Done"],
            orientation="h",
            name="Done",
            marker_color="#16a34a",
            text=chart_df["Done"],
            textposition="outside",
            cliponaxis=False,
        )
    )
    fig.add_trace(
        go.Bar(
            y=chart_df["Item"],
            x=chart_df["Part Done"],
            orientation="h",
            name="Part Done",
            marker_color="#eab308",
            text=chart_df["Part Done"],
            textposition="outside",
            cliponaxis=False,
        )
    )
    fig.add_trace(
        go.Bar(
            y=chart_df["Item"],
            x=chart_df["Not Done"],
            orientation="h",
            name="Not Done",
            marker_color="#dc2626",
            text=chart_df["Not Done"],
            textposition="outside",
            cliponaxis=False,
        )
    )

    fig.update_layout(
        barmode="group",
        template="plotly_white",
        height=410,
        margin=dict(l=10, r=70, t=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font=dict(color="#111827", size=13),
        xaxis=dict(title=None, showgrid=True, gridcolor="#e5e7eb", zeroline=False),
        yaxis=dict(title=None, autorange="reversed"),
    )

    return fig


def make_progress_bucket_chart(df):
    bins = [0, 25, 50, 75, 90, 100]
    labels = ["0-25%", "26-50%", "51-75%", "76-90%", "91-100%"]

    bucket = pd.cut(
        df["Overall_Progress"],
        bins=bins,
        labels=labels,
        include_lowest=True,
    )

    counts = bucket.value_counts().reindex(labels).fillna(0).astype(int).reset_index()
    counts.columns = ["Progress Bucket", "Count"]

    fig = px.bar(
        counts,
        x="Count",
        y="Progress Bucket",
        orientation="h",
        text="Count",
    )

    fig.update_traces(
        marker_color="#7c3aed",
        textposition="outside",
        cliponaxis=False,
    )

    return chart_layout(fig, height=320, right=60)


def make_phase_by_project_chart(df):
    temp = df.sort_values("Overall_Progress", ascending=False).head(25).copy()

    project_labels = temp["Project_Name"].astype(str).str.slice(0, 42)

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            y=project_labels,
            x=temp["Engineering_Pct"],
            name="Engineering",
            orientation="h",
            marker_color="#2563eb",
        )
    )

    fig.add_trace(
        go.Bar(
            y=project_labels,
            x=temp["Delivery_Pct"],
            name="Delivery",
            orientation="h",
            marker_color="#7c3aed",
        )
    )

    fig.add_trace(
        go.Bar(
            y=project_labels,
            x=temp["Execution_Pct"],
            name="Execution",
            orientation="h",
            marker_color="#f97316",
        )
    )

    fig.update_layout(
        barmode="group",
        template="plotly_white",
        height=620,
        margin=dict(l=10, r=40, t=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font=dict(color="#111827", size=12),
        xaxis=dict(
            title=None,
            range=[0, 105],
            ticksuffix="%",
            showgrid=True,
            gridcolor="#e5e7eb",
            zeroline=False,
        ),
        yaxis=dict(title=None, autorange="reversed"),
    )

    return fig


def render_chart_card(title, fig, tag=None):
    tag_html = f'<span class="chart-tag">{tag}</span>' if tag else ""
    st.markdown(
        f"""
<div class="chart-card">
    <div class="chart-title">{title} {tag_html}</div>
</div>
""",
        unsafe_allow_html=True,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_kpi(label, value, note, color_class=""):
    st.markdown(
        f"""
<div class="kpi-card {color_class}">
    <div class="kpi-label">{label}</div>
    <div class="kpi-value">{value}</div>
    <div class="kpi-note">{note}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def render_phase(title, avg_pct, done_label, color_class=""):
    avg_pct = pct(avg_pct)
    st.markdown(
        f"""
<div class="phase-card {color_class}">
    <div class="phase-top">
        <div class="phase-title">{title}</div>
        <div class="phase-pct">{avg_pct:.1f}%</div>
    </div>
    <div class="phase-bar-bg">
        <div class="phase-bar-fill" style="width:{avg_pct}%;"></div>
    </div>
    <div class="phase-stats">
        <span>Average completion</span>
        <strong>{done_label}</strong>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )


# -----------------------------
# Load data
# -----------------------------
df, source_name = load_data()


# -----------------------------
# Sidebar filters
# -----------------------------
st.sidebar.title("Filters")

search = st.sidebar.text_input("Search project, customer, location")

quick_view = st.sidebar.radio(
    "Quick Views",
    ["All", "Completed", "In Progress", "On Hold", "Not Started"],
)

statuses = sorted(df["Status"].dropna().unique().tolist())
regions = sorted(df["Region"].dropna().unique().tolist())
customers = sorted(df["Customer"].dropna().unique().tolist())
materials = sorted(df["Material_Status"].dropna().unique().tolist())
priorities = ["High", "Medium", "Low"]

selected_statuses = st.sidebar.multiselect("Status", statuses)
selected_regions = st.sidebar.multiselect("Region", regions)
selected_customers = st.sidebar.multiselect("Customer", customers)
selected_materials = st.sidebar.multiselect("Material", materials)
selected_priorities = st.sidebar.multiselect("Priority", priorities)

filtered_df = df.copy()

if quick_view != "All":
    filtered_df = filtered_df[filtered_df["Status"] == quick_view]

if search:
    pattern = re.escape(search.strip())
    mask = (
        filtered_df["Project_Name"].astype(str).str.contains(pattern, case=False, na=False)
        | filtered_df["Customer"].astype(str).str.contains(pattern, case=False, na=False)
        | filtered_df["Location"].astype(str).str.contains(pattern, case=False, na=False)
        | filtered_df["Region"].astype(str).str.contains(pattern, case=False, na=False)
        | filtered_df["Job_Ref"].astype(str).str.contains(pattern, case=False, na=False)
    )
    filtered_df = filtered_df[mask]

if selected_statuses:
    filtered_df = filtered_df[filtered_df["Status"].isin(selected_statuses)]

if selected_regions:
    filtered_df = filtered_df[filtered_df["Region"].isin(selected_regions)]

if selected_customers:
    filtered_df = filtered_df[filtered_df["Customer"].isin(selected_customers)]

if selected_materials:
    filtered_df = filtered_df[filtered_df["Material_Status"].isin(selected_materials)]

if selected_priorities:
    filtered_df = filtered_df[filtered_df["Priority"].isin(selected_priorities)]


# -----------------------------
# Header
# -----------------------------
updated = datetime.now().strftime("%d %b %Y, %I:%M %p")

st.markdown(
    f"""
<div class="main-hero">
    <h1>Project Tracking Dashboard</h1>
    <div class="sub">Cold Rooms, Cabinets & Refrigeration Projects — Live view from Excel</div>
    <div class="pill-row">
        <div class="pill green">● Current source: {source_name}</div>
        <div class="pill blue">↻ Refresh: {REFRESH_SECONDS}s</div>
        <div class="pill orange">Showing: {len(filtered_df)} / {len(df)} projects</div>
        <div class="pill">Updated: {updated}</div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# -----------------------------
# KPI cards
# -----------------------------
total_projects = len(filtered_df)
completed = int((filtered_df["Status"] == "Completed").sum())
in_progress = int((filtered_df["Status"] == "In Progress").sum())
on_hold = int((filtered_df["Status"] == "On Hold").sum())
not_started = int((filtered_df["Status"] == "Not Started").sum())
avg_progress = filtered_df["Overall_Progress"].mean() if len(filtered_df) else 0

k1, k2, k3, k4, k5, k6 = st.columns(6)

with k1:
    render_kpi("Total Projects", total_projects, "Projects in current view", "")
with k2:
    render_kpi("Completed", completed, "Ready / closed", "green")
with k3:
    render_kpi("In Progress", in_progress, "Currently active", "teal")
with k4:
    render_kpi("On Hold", on_hold, "Needs attention", "orange")
with k5:
    render_kpi("Not Started", not_started, "Yet to begin", "gray")
with k6:
    render_kpi("Avg Progress", f"{avg_progress:.1f}%", "Across current view", "purple")


# -----------------------------
# Phase cards
# -----------------------------
st.markdown('<div class="section-title">Phase Progress Overview</div>', unsafe_allow_html=True)

p1, p2, p3 = st.columns(3)

with p1:
    eng_avg = filtered_df["Engineering_Pct"].mean() if len(filtered_df) else 0
    eng_done = int(filtered_df["Eng_Done"].sum()) if len(filtered_df) else 0
    render_phase("Engineering", eng_avg, f"{eng_done} items done", "")

with p2:
    del_avg = filtered_df["Delivery_Pct"].mean() if len(filtered_df) else 0
    del_done = int(filtered_df["Del_Done"].sum()) if len(filtered_df) else 0
    render_phase("Delivery", del_avg, f"{del_done} items done", "del")

with p3:
    exec_avg = filtered_df["Execution_Pct"].mean() if len(filtered_df) else 0
    render_phase("Execution", exec_avg, "Site/work progress", "exec")


# -----------------------------
# Charts row 1
# Same sections as given dashboard, readable chart type
# -----------------------------
c1, c2, c3 = st.columns(3)

with c1:
    fig = make_count_bar(
        filtered_df,
        "Status",
        order=["Completed", "In Progress", "On Hold", "Not Started", "Cancelled"],
    )
    render_chart_card("Project Status", fig)

with c2:
    fig = make_count_bar(
        filtered_df,
        "Material_Status",
        order=["Delivered", "Partially Delivered", "Ordered", "Not Ordered"],
    )
    render_chart_card("Material Status", fig)

with c3:
    fig = make_count_bar(
        filtered_df,
        "Priority",
        order=["High", "Medium", "Low"],
    )
    render_chart_card("Priority", fig)


# -----------------------------
# Charts row 2
# -----------------------------
c1, c2 = st.columns(2)

with c1:
    fig = make_top_bar(filtered_df, "Region", top_n=10, color="#0891b2")
    render_chart_card("Region Breakdown", fig)

with c2:
    fig = make_top_bar(filtered_df, "Customer", top_n=10, color="#2563eb")
    render_chart_card("Top Customers", fig)


# -----------------------------
# Charts row 3
# -----------------------------
c1, c2 = st.columns([7, 5])

with c1:
    fig = make_delivery_items_chart(filtered_df)
    render_chart_card("Delivery Items Status", fig, "DONE / PART DONE / NOT DONE across projects")

with c2:
    fig = make_progress_bucket_chart(filtered_df)
    render_chart_card("Overall Progress Buckets", fig)


# -----------------------------
# Phase by project
# -----------------------------
fig = make_phase_by_project_chart(filtered_df)
render_chart_card(
    "Phase Progress by Project",
    fig,
    "Engineering / Delivery / Execution — sorted by Overall %",
)


# -----------------------------
# Project details table
# -----------------------------
st.markdown('<div class="section-title">Project Details</div>', unsafe_allow_html=True)

table_df = filtered_df[
    [
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
].copy()

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
    height=620,
)
