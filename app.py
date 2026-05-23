from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
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
    initial_sidebar_state="collapsed",
)

LOCAL_EXCEL_FILE = "Project_Tracking_v7.xlsx"
SHEET_NAME = "Project_Master"

def _secret(name, default=""):
    try:
        return str(st.secrets.get(name, default)).strip()
    except Exception:
        return default

try:
    REFRESH_SECONDS = int(_secret("REFRESH_SECONDS", "60") or 60)
except Exception:
    REFRESH_SECONDS = 60

ENGINEERING_COLS = ["Equi. DesigN","Technical Submittal","Drawing","ELS","BOM"]
DELIVERY_COLS    = ["Out_Door","Indoor","CR Panels","CR Ins. Materials]","Doors",
                    "Ref. Inst. Materials","CCP","Display CCP","Floor Heater","Cabinets","Any Special"]
BASE_COLUMNS = [
    "S_No","Date_Time","Job_Ref","LPO_Ref","Customer","Project_Name",
    "Region","Location","Payment_Terms",
    *ENGINEERING_COLS, *DELIVERY_COLS,
    "Work_Status","Remarks","Material_Status","Overall_Progress",
    "Priority","Status","Engineering_Pct","Delivery_Pct","Execution_Pct",
]

STATUS_ORDER   = ["Completed","In Progress","On Hold","Not Started","Cancelled"]
MATERIAL_ORDER = ["Delivered","Partially Delivered","Ordered","Not Ordered"]
PRIORITY_ORDER = ["High","Medium","Low"]

STATUS_COLORS   = {"Completed":"#22c55e","In Progress":"#14b8a6","On Hold":"#f97316","Not Started":"#64748b","Cancelled":"#ef4444"}
PRIORITY_COLORS = {"High":"#ef4444","Medium":"#eab308","Low":"#22c55e"}
MATERIAL_COLORS = {"Delivered":"#22c55e","Partially Delivered":"#eab308","Ordered":"#f97316","Not Ordered":"#ef4444"}

# -- CSS -----------------------------------------------------------------------
st.html("""
<style>
html,body,[class*="css"]{font-family:'Segoe UI',Tahoma,Arial,sans-serif!important;}
.stApp{background:linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%)!important;}
.block-container{padding:1rem 2rem 4rem!important;max-width:1600px!important;}
section[data-testid="stSidebar"]{display:none!important;}
[data-testid="collapsedControl"]{display:none!important;}

/* header */
.hdr{display:flex;justify-content:space-between;align-items:flex-start;
  margin-bottom:16px;padding-bottom:14px;border-bottom:1px solid #334155;flex-wrap:wrap;gap:12px;}
.hdr h1{font-size:24px;font-weight:700;color:#f1f5f9;margin:0;}
.hdr .sub{color:#94a3b8;font-size:13px;margin-top:4px;}
.hdr .right{text-align:right;color:#94a3b8;font-size:12px;}

/* quick bar */
.qbar{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;align-items:center;}
.qlabel{color:#94a3b8;font-size:12px;text-transform:uppercase;letter-spacing:.5px;margin-right:4px;}
.qbtn{display:inline-block;padding:7px 14px;border-radius:8px;font-size:13px;font-weight:500;
  cursor:pointer;border:none;color:#fff;text-decoration:none;}

/* kpi grid */
.kpi-grid{display:grid;grid-template-columns:repeat(8,1fr);gap:12px;margin-bottom:18px;}
.kpi{background:#1e293b;border-radius:12px;padding:16px 14px;border-left:4px solid var(--ac,#3b82f6);}
.kpi .lbl{color:#94a3b8;font-size:11px;text-transform:uppercase;letter-spacing:.5px;}
.kpi .val{font-size:28px;font-weight:700;margin-top:5px;color:#f1f5f9;}
.kpi .dlt{font-size:11px;color:#94a3b8;margin-top:3px;}

/* section title */
.stitle{font-size:12px;color:#94a3b8;margin:6px 0 10px;text-transform:uppercase;letter-spacing:.5px;font-weight:600;}

/* phase cards */
.pgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:18px;}
.pc{background:#1e293b;border-radius:12px;padding:18px;border-top:4px solid var(--ac,#3b82f6);}
.pc .pt{font-size:14px;font-weight:600;display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;color:#f1f5f9;}
.pc .pp{font-size:26px;font-weight:700;}
.pc .pb{height:10px;background:#334155;border-radius:5px;overflow:hidden;margin-bottom:10px;}
.pc .pf{height:100%;border-radius:5px;}
.pc .ps{display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;}
.pc .ps strong{color:#f1f5f9;font-size:13px;}

/* card */
.card{background:#1e293b;border-radius:12px;padding:16px;margin-bottom:16px;}
.card h3{font-size:14px;font-weight:600;margin-bottom:10px;display:flex;
  justify-content:space-between;align-items:center;color:#f1f5f9;}
.card h3 .tag{font-size:11px;color:#94a3b8;font-weight:400;}

/* badges */
.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;}
.s-completed{background:rgba(34,197,94,.25);color:#86efac;}
.s-progress{background:rgba(20,184,166,.25);color:#5eead4;}
.s-hold{background:rgba(249,115,22,.25);color:#fdba74;}
.s-cancelled{background:rgba(239,68,68,.25);color:#fca5a5;}
.s-notstarted{background:rgba(100,116,139,.3);color:#cbd5e1;}
.m-delivered{background:rgba(34,197,94,.2);color:#86efac;}
.m-partial{background:rgba(234,179,8,.2);color:#fde047;}
.m-ordered{background:rgba(249,115,22,.2);color:#fdba74;}
.m-notordered{background:rgba(239,68,68,.2);color:#fca5a5;}
.p-high{background:rgba(239,68,68,.2);color:#fca5a5;}
.p-medium{background:rgba(234,179,8,.2);color:#fde047;}
.p-low{background:rgba(34,197,94,.2);color:#86efac;}

/* mini progress bar */
.bc{display:flex;align-items:center;gap:6px;min-width:90px;}
.bb{flex:1;height:6px;background:#334155;border-radius:3px;overflow:hidden;min-width:45px;}
.bf{height:100%;border-radius:3px;}
.eng{background:linear-gradient(90deg,#60a5fa,#3b82f6);}
.del{background:linear-gradient(90deg,#c084fc,#a855f7);}
.exc{background:linear-gradient(90deg,#fb923c,#f97316);}
.all{background:linear-gradient(90deg,#ef4444,#eab308,#22c55e);}
.bt{font-size:11px;font-weight:600;min-width:32px;text-align:right;color:#cbd5e1;}

/* table */
.twrap{overflow:auto;max-height:560px;border-radius:8px;}
table{width:100%;border-collapse:collapse;font-size:12px;}
thead{position:sticky;top:0;background:#334155;z-index:5;}
th{text-align:left;padding:10px 8px;font-weight:600;border-bottom:2px solid #475569;
   white-space:nowrap;color:#f1f5f9;}
td{padding:9px 8px;border-bottom:1px solid #334155;color:#cbd5e1;}
tr:hover td{background:rgba(59,130,246,.08);}

/* filters row */
.frow{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;align-items:center;}
.frow input,.frow select{
  background:#334155;color:#f1f5f9;border:1px solid #475569;
  padding:7px 10px;border-radius:8px;font-size:13px;outline:none;
  font-family:'Segoe UI',Tahoma,Arial,sans-serif;}
.frow input{min-width:220px;}

/* download btn */
.stDownloadButton>button{
  background:#334155!important;border:1px solid #475569!important;
  border-radius:8px!important;color:#f1f5f9!important;font-size:13px!important;}
.stDownloadButton>button:hover{background:#3b82f6!important;border-color:#3b82f6!important;}

/* streamlit element cleanup */
[data-testid="stDecoration"]{display:none!important;}
.stTabs [data-baseweb="tab-list"]{
  background:#1e293b!important;border:1px solid #334155!important;
  border-radius:10px!important;padding:3px!important;gap:2px!important;width:fit-content!important;}
.stTabs [data-baseweb="tab"]{
  background:transparent!important;border:none!important;border-radius:7px!important;
  color:#94a3b8!important;font-size:13px!important;font-weight:500!important;padding:7px 18px!important;}
.stTabs [aria-selected="true"]{background:#3b82f6!important;color:#fff!important;}
.stTabs [data-baseweb="tab-highlight"],.stTabs [data-baseweb="tab-border"]{display:none!important;}

hr{border-color:#334155!important;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:#0f172a;}
::-webkit-scrollbar-thumb{background:#334155;border-radius:3px;}
</style>
""")

# -- DATA SOURCE ---------------------------------------------------------------
def make_onedrive_url(url):
    url = url.strip()
    if not url or "download=1" in url:
        return url
    return url + ("&" if "?" in url else "?") + "download=1"

@st.cache_data(ttl=REFRESH_SECONDS)
def load_url(url):
    r = requests.get(make_onedrive_url(url), timeout=60, allow_redirects=True)
    r.raise_for_status()
    if "text/html" in r.headers.get("content-type","").lower() and len(r.content)<500_000:
        raise RuntimeError("URL returned HTML not Excel. Check share link.")
    return r.content

@st.cache_data(ttl=REFRESH_SECONDS)
def load_graph():
    if msal is None: raise RuntimeError("msal not installed")
    tid = _secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    cid = _secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    cs  = _secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")
    if not all([tid,cid,cs]): raise RuntimeError("Missing Graph secrets")
    app = msal.ConfidentialClientApplication(cid,
        authority=f"https://login.microsoftonline.com/{tid}", client_credential=cs)
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res: raise RuntimeError(f"Token error: {res}")
    uid = _secret("GRAPH_USER_ID")
    fp  = _secret("ONEDRIVE_FILE_PATH")
    if not fp.startswith("/"): fp = "/" + fp
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{uid}/drive/root:{fp}:/content",
        headers={"Authorization": f"Bearer {res['access_token']}"}, timeout=60)
    r.raise_for_status()
    return r.content

def get_bytes():
    eu = _secret("EXCEL_FILE_URL")
    if eu: return load_url(eu), "OneDrive"
    if _secret("GRAPH_USER_ID"): return load_graph(), "Graph OneDrive"
    lp = Path(LOCAL_EXCEL_FILE)
    if lp.exists(): return lp.read_bytes(), LOCAL_EXCEL_FILE
    st.error("No data source. Set EXCEL_FILE_URL secret."); st.stop()

# -- PARSE ---------------------------------------------------------------------
def ct(v, d=""):
    if pd.isna(v): return d
    t = str(v).strip(); return t if t else d

def cs(v):
    v = ct(v, "Not Started")
    m = {"completed":"Completed","complete":"Completed","in progress":"In Progress",
         "on progress":"In Progress","progress":"In Progress","on hold":"On Hold",
         "hold":"On Hold","cancelled":"Cancelled","canceled":"Cancelled",
         "not started":"Not Started","notstart":"Not Started"}
    return m.get(v.lower(), v)

@st.cache_data(ttl=REFRESH_SECONDS)
def parse(file_bytes):
    raw  = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME,
                         header=None, engine="openpyxl")
    data = raw.iloc[2:, :len(BASE_COLUMNS)].copy()
    data.columns = BASE_COLUMNS
    data = data.dropna(how="all")
    data = data[~(data["S_No"].isna() & data["Project_Name"].isna())]
    for col in ["Job_Ref","LPO_Ref","Customer","Project_Name","Region","Location","Priority"]:
        data[col] = data[col].apply(lambda x: ct(x,"Unknown" if col in ["Customer","Region"] else ""))
    data["Material_Status"] = data["Material_Status"].apply(lambda x: ct(x,"Not Ordered"))
    data["Status"]          = data["Status"].apply(cs)
    data["Work_Status"]     = data["Work_Status"].apply(lambda x: ct(x,""))
    for col in ["Overall_Progress","Engineering_Pct","Delivery_Pct","Execution_Pct"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0).clip(0,100)
    for gn, cols in [("Eng",ENGINEERING_COLS),("Del",DELIVERY_COLS)]:
        blk = data[cols].fillna("").astype(str).apply(lambda s:s.str.strip().str.upper())
        data[f"{gn}_Done"]    = (blk=="DONE").sum(axis=1)
        data[f"{gn}_Partial"] = (blk=="PART.DONE").sum(axis=1)
        data[f"{gn}_NA"]      = (blk=="N/A").sum(axis=1)
    data["S_No"]     = data["S_No"].fillna("").astype(str).str.replace(".0","",regex=False)
    data["Priority"] = data["Priority"].replace("","Medium")
    return data.reset_index(drop=True)

# -- HELPERS -------------------------------------------------------------------
def h(t):
    return str(t).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def pct(v): return f"{float(v):.1f}%"

def count_order(df, col, order):
    c = df[col].value_counts().reindex(order, fill_value=0).reset_index()
    c.columns = [col,"Count"]
    return c

def sf(fig, height=280):
    fig.update_layout(
        height=height, margin=dict(l=12,r=12,t=28,b=12),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#cbd5e1", family="Segoe UI,Arial,sans-serif", size=11),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8",size=11)),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#334155", zeroline=False,
                     tickfont=dict(color="#94a3b8"))
    fig.update_yaxes(showgrid=True, gridcolor="#334155", zeroline=False,
                     tickfont=dict(color="#94a3b8"))
    return fig

def donut_chart(df, names, values, colors, height=270):
    df = df[df[values]>0].copy()
    seq = [colors.get(str(n),"#64748b") for n in df[names]]
    fig = px.pie(df, names=names, values=values, hole=0.60,
                 color_discrete_sequence=seq)
    fig.update_traces(
        textposition="inside", textinfo="percent",
        textfont=dict(size=12,color="#fff"),
        insidetextorientation="radial",
        marker=dict(line=dict(color="#1e293b",width=2)),
        hovertemplate="<b>%{label}</b><br>%{value} (%{percent})<extra></extra>",
    )
    total = int(df[values].sum())
    fig.update_layout(
        showlegend=True,
        legend=dict(orientation="h",yanchor="top",y=-0.05,xanchor="center",x=0.5,
                    font=dict(size=11,color="#94a3b8"),bgcolor="rgba(0,0,0,0)"),
        annotations=[dict(text=f"<b>{total}</b>",x=0.5,y=0.5,
                          font=dict(size=18,color="#f1f5f9"),showarrow=False)],
        height=height, margin=dict(l=6,r=6,t=6,b=50),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#94a3b8",family="Segoe UI,Arial,sans-serif",size=11),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

def bar_v(df, x, y, colors, height=260):
    fig = go.Figure()
    for i,(lbl,cnt) in enumerate(zip(df[x],df[y])):
        c = colors[i] if isinstance(colors,list) else colors
        fig.add_trace(go.Bar(x=[lbl],y=[cnt],marker_color=c,marker_line_width=0,
                             showlegend=False,text=[cnt],textposition="outside",
                             textfont=dict(color="#94a3b8",size=11)))
    fig = sf(fig, height)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

def bar_h(df, x, y, color="#3b82f6", height=280):
    fig = px.bar(df, x=x, y=y, orientation="h", text=x)
    fig.update_traces(marker_color=color, marker_line_width=0,
                      textfont=dict(color="#94a3b8",size=11), textposition="outside")
    fig = sf(fig, height)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

def mbar(val, cls):
    p = max(0,min(100,float(val or 0)))
    return (f'<div class="bc"><div class="bb"><div class="bf {cls}" style="width:{p:.0f}%"></div></div>'
            f'<div class="bt">{p:.0f}%</div></div>')

# -- LOAD ----------------------------------------------------------------------
st_autorefresh(interval=REFRESH_SECONDS*1000, key="ar")
try:
    fb, src = get_bytes()
    df_all  = parse(fb)
except Exception as e:
    st.error(f"Cannot load data: {e}"); st.stop()

# -- SESSION STATE for quick filter --------------------------------------------
if "qf" not in st.session_state:
    st.session_state.qf = "All"

# -- HEADER --------------------------------------------------------------------
now = pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")
st.html(f"""
<div class="hdr">
  <div>
    <h1>Project Tracking Dashboard</h1>
    <div class="sub">Cold Rooms, Cabinets &amp; Refrigeration Projects - Live view from Excel</div>
  </div>
  <div class="right">
    <div>Data as of: {h(now)}</div>
    <div>Source: {h(src)} &nbsp;|&nbsp; Total: {len(df_all)} projects</div>
  </div>
</div>
""")

# -- QUICK FILTER BUTTONS ------------------------------------------------------
qcols = st.columns([1,1,1,1,1,1,8])
labels_btns = [
    ("All",       "#475569"),
    ("Completed", "#22c55e"),
    ("In Progress","#3b82f6"),
    ("On Hold",   "#f97316"),
    ("Not Started","#64748b"),
    ("Reset",     "#475569"),
]
for i,(lbl,color) in enumerate(labels_btns):
    with qcols[i]:
        if st.button(lbl, key=f"qb_{lbl}",
                     use_container_width=True,
                     type="secondary"):
            st.session_state.qf = "All" if lbl=="Reset" else lbl

# -- SIDEBAR-STYLE FILTERS (shown as expander) ---------------------------------
with st.expander("Filters", expanded=False):
    fc = st.columns(6)
    with fc[0]: srch = st.text_input("Search","",placeholder="Project, customer, location...")
    with fc[1]: f_st = st.selectbox("Status",  ["All"]+sorted(df_all["Status"].dropna().unique().tolist()),  key="f_st")
    with fc[2]: f_rg = st.selectbox("Region",  ["All"]+sorted(df_all["Region"].dropna().unique().tolist()),  key="f_rg")
    with fc[3]: f_cu = st.selectbox("Customer",["All"]+sorted(df_all["Customer"].dropna().unique().tolist()),key="f_cu")
    with fc[4]: f_mt = st.selectbox("Material",["All"]+sorted(df_all["Material_Status"].dropna().unique().tolist()),key="f_mt")
    with fc[5]: f_pr = st.selectbox("Priority",["All"]+sorted(df_all["Priority"].dropna().unique().tolist()),key="f_pr")

# -- APPLY FILTERS -------------------------------------------------------------
df = df_all.copy()
if st.session_state.qf not in ("All","Reset"):
    df = df[df["Status"]==st.session_state.qf]
if srch:
    t = srch.lower()
    df = df[df["Project_Name"].str.lower().str.contains(t,na=False)
          | df["Customer"].str.lower().str.contains(t,na=False)
          | df["Location"].str.lower().str.contains(t,na=False)
          | df["Job_Ref"].astype(str).str.lower().str.contains(t,na=False)]
if f_st!="All":   df=df[df["Status"]==f_st]
if f_rg!="All":   df=df[df["Region"]==f_rg]
if f_cu!="All":   df=df[df["Customer"]==f_cu]
if f_mt!="All":   df=df[df["Material_Status"]==f_mt]
if f_pr!="All":   df=df[df["Priority"]==f_pr]

N = len(df)
n_comp   = int((df["Status"]=="Completed").sum())
n_prog   = int((df["Status"]=="In Progress").sum())
n_hold   = int((df["Status"]=="On Hold").sum())
n_ns     = int((df["Status"]=="Not Started").sum())
n_can    = int((df["Status"]=="Cancelled").sum())
n_matd   = int((df["Material_Status"]=="Delivered").sum())
avg_p    = df["Overall_Progress"].mean() if N else 0
eng_avg  = df["Engineering_Pct"].mean() if N else 0
del_avg  = df["Delivery_Pct"].mean() if N else 0
exc_avg  = df["Execution_Pct"].mean() if N else 0
eng_done = int((df["Engineering_Pct"]>=100).sum())
del_done = int((df["Delivery_Pct"]>=100).sum())
exc_done = int((df["Execution_Pct"]>=100).sum())

def pp(n): return f"{round(n/N*100)}%" if N else ""

# -- KPI ROW -------------------------------------------------------------------
st.html(f"""
<div class="kpi-grid">
  <div class="kpi" style="--ac:#3b82f6"><div class="lbl">Total Projects</div>
    <div class="val">{N}</div><div class="dlt">Showing {N} of {len(df_all)}</div></div>
  <div class="kpi" style="--ac:#22c55e"><div class="lbl">Completed</div>
    <div class="val">{n_comp}</div><div class="dlt">{pp(n_comp)}</div></div>
  <div class="kpi" style="--ac:#14b8a6"><div class="lbl">In Progress</div>
    <div class="val">{n_prog}</div><div class="dlt">{pp(n_prog)}</div></div>
  <div class="kpi" style="--ac:#f97316"><div class="lbl">On Hold</div>
    <div class="val">{n_hold}</div><div class="dlt">{pp(n_hold)}</div></div>
  <div class="kpi" style="--ac:#64748b"><div class="lbl">Not Started</div>
    <div class="val">{n_ns}</div><div class="dlt">{pp(n_ns)}</div></div>
  <div class="kpi" style="--ac:#ef4444"><div class="lbl">Cancelled</div>
    <div class="val">{n_can}</div><div class="dlt">{pp(n_can)}</div></div>
  <div class="kpi" style="--ac:#ec4899"><div class="lbl">Overall Avg %</div>
    <div class="val">{avg_p:.1f}%</div><div class="dlt">&nbsp;</div></div>
  <div class="kpi" style="--ac:#eab308"><div class="lbl">Mat. Delivered</div>
    <div class="val">{n_matd}</div><div class="dlt">{pp(n_matd)}</div></div>
</div>
""")

# -- PHASE CARDS ---------------------------------------------------------------
st.html('<div class="stitle">Phase Progress Overview</div>')
st.html(f"""
<div class="pgrid">
  <div class="pc" style="--ac:#3b82f6">
    <div class="pt">Engineering
      <span class="pp" style="color:#60a5fa;font-size:26px;font-weight:700">{eng_avg:.1f}%</span></div>
    <div class="pb"><div class="pf" style="width:{eng_avg:.1f}%;background:linear-gradient(90deg,#3b82f6,#60a5fa)"></div></div>
    <div class="ps"><span>Design / Submittal / Drawing / ELS / BOM</span>
      <span><strong>{eng_done}</strong>/{N} done</span></div></div>
  <div class="pc" style="--ac:#a855f7">
    <div class="pt">Delivery
      <span class="pp" style="color:#c084fc;font-size:26px;font-weight:700">{del_avg:.1f}%</span></div>
    <div class="pb"><div class="pf" style="width:{del_avg:.1f}%;background:linear-gradient(90deg,#a855f7,#c084fc)"></div></div>
    <div class="ps"><span>Material delivery - {len(DELIVERY_COLS)} items</span>
      <span><strong>{del_done}</strong>/{N} done</span></div></div>
  <div class="pc" style="--ac:#f97316">
    <div class="pt">Execution
      <span class="pp" style="color:#fb923c;font-size:26px;font-weight:700">{exc_avg:.1f}%</span></div>
    <div class="pb"><div class="pf" style="width:{exc_avg:.1f}%;background:linear-gradient(90deg,#f97316,#fb923c)"></div></div>
    <div class="ps"><span>On-site installation, commissioning</span>
      <span><strong>{exc_done}</strong>/{N} done</span></div></div>
</div>
""")

# -- ROW 1: 3 DONUTS -----------------------------------------------------------
c1,c2,c3 = st.columns(3)
with c1:
    st.html('<div class="card"><h3>Project Status</h3>')
    donut_chart(count_order(df,"Status",STATUS_ORDER),"Status","Count",STATUS_COLORS)
    st.html('</div>')
with c2:
    st.html('<div class="card"><h3>Material Status</h3>')
    donut_chart(count_order(df,"Material_Status",MATERIAL_ORDER),"Material_Status","Count",MATERIAL_COLORS)
    st.html('</div>')
with c3:
    st.html('<div class="card"><h3>Priority</h3>')
    donut_chart(count_order(df,"Priority",PRIORITY_ORDER),"Priority","Count",PRIORITY_COLORS)
    st.html('</div>')

# -- ROW 2: REGION + CUSTOMERS -------------------------------------------------
c4,c5 = st.columns(2)
with c4:
    st.html('<div class="card"><h3>Region Breakdown</h3>')
    rdf = df["Region"].value_counts().reset_index()
    rdf.columns = ["Region","Count"]
    rdf = rdf.sort_values("Count",ascending=False)
    pal = ["#3b82f6","#a855f7","#14b8a6","#ec4899","#f97316","#64748b","#eab308","#22c55e"]
    bar_v(rdf,"Region","Count",[pal[i%len(pal)] for i in range(len(rdf))], height=270)
    st.html('</div>')
with c5:
    st.html('<div class="card"><h3>Top Customers</h3>')
    cdf = df["Customer"].value_counts().head(10).reset_index()
    cdf.columns = ["Customer","Count"]
    cdf = cdf.sort_values("Count",ascending=True)
    bar_h(cdf,"Count","Customer","#3b82f6",height=290)
    st.html('</div>')

# -- ROW 3: DELIVERY ITEMS + PROGRESS BUCKETS ---------------------------------
c6,c7 = st.columns([7,5])
with c6:
    st.html('<div class="card"><h3>Delivery Items Status <span class="tag">DONE across all projects</span></h3>')
    done_v  = [int((df[c].fillna("").astype(str).str.strip().str.upper()=="DONE").sum())      for c in DELIVERY_COLS]
    part_v  = [int((df[c].fillna("").astype(str).str.strip().str.upper()=="PART.DONE").sum()) for c in DELIVERY_COLS]
    nd_v    = [int((df[c].fillna("").astype(str).str.strip().str.upper()=="N.DONE").sum())    for c in DELIVERY_COLS]
    na_v    = [int((df[c].fillna("").astype(str).str.strip().str.upper()=="N/A").sum())       for c in DELIVERY_COLS]
    ddf = pd.DataFrame({"Item":DELIVERY_COLS,"Done":done_v,"Partial":part_v,"Not Done":nd_v,"N/A":na_v})
    fig = go.Figure()
    fig.add_trace(go.Bar(y=ddf["Item"],x=ddf["Done"],    name="Done",    orientation="h",marker_color="#22c55e",marker_line_width=0))
    fig.add_trace(go.Bar(y=ddf["Item"],x=ddf["Partial"], name="Partial", orientation="h",marker_color="#eab308",marker_line_width=0))
    fig.add_trace(go.Bar(y=ddf["Item"],x=ddf["Not Done"],name="Not Done",orientation="h",marker_color="#ef4444",marker_line_width=0))
    fig.add_trace(go.Bar(y=ddf["Item"],x=ddf["N/A"],     name="N/A",     orientation="h",marker_color="#64748b",marker_line_width=0))
    fig.update_layout(barmode="stack",
        legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1,
                    font=dict(size=11,color="#94a3b8"),bgcolor="rgba(0,0,0,0)"))
    fig = sf(fig,360)
    fig.update_layout(xaxis_title="Number of Projects",margin=dict(l=12,r=12,t=36,b=12))
    st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
    st.html('</div>')
with c7:
    st.html('<div class="card"><h3>Overall Progress Buckets</h3>')
    bins   = [-0.1,25,50,75,100]
    lbls   = ["0-25%","26-50%","51-75%","76-100%"]
    bkt    = pd.cut(df["Overall_Progress"],bins=bins,labels=lbls)
    bdf    = bkt.value_counts().reindex(lbls,fill_value=0).reset_index()
    bdf.columns = ["Bucket","Count"]
    bclr   = ["#ef4444","#f97316","#eab308","#22c55e"]
    bar_v(bdf,"Bucket","Count",bclr,height=360)
    st.html('</div>')

# -- ROW 4: PHASE PROGRESS BY PROJECT -----------------------------------------
st.html('<div class="card"><h3>Phase Progress by Project <span class="tag">Engineering / Delivery / Execution - sorted by Overall %</span></h3>')
ph = df.sort_values("Overall_Progress",ascending=False).head(25)
ylbls = [f"#{r.S_No} {(r.Project_Name or r.Customer)[:30]}" for _,r in ph.iterrows()]
fig = go.Figure()
fig.add_trace(go.Bar(y=ylbls,x=ph["Engineering_Pct"].values,name="Engineering %",
    orientation="h",marker_color="#3b82f6",marker_line_width=0))
fig.add_trace(go.Bar(y=ylbls,x=ph["Delivery_Pct"].values,name="Delivery %",
    orientation="h",marker_color="#a855f7",marker_line_width=0))
fig.add_trace(go.Bar(y=ylbls,x=ph["Execution_Pct"].values,name="Execution %",
    orientation="h",marker_color="#f97316",marker_line_width=0))
fig.update_layout(
    barmode="group",
    legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                font=dict(size=11,color="#94a3b8"),bgcolor="rgba(0,0,0,0)"),
    xaxis=dict(range=[0,100],title="Percent Complete",tickfont=dict(color="#94a3b8")),
    yaxis=dict(autorange="reversed",tickfont=dict(color="#94a3b8")),
)
fig = sf(fig, max(420, len(ph)*26))
fig.update_layout(margin=dict(l=12,r=12,t=40,b=12))
st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
st.html('</div>')

# -- PROJECT DETAILS TABLE -----------------------------------------------------
st.html('<div class="card"><h3>Project Details</h3>')

fc2 = st.columns([2.5,1.4,1.4,1.4,1.4,1.4,1])
with fc2[0]: ts = st.text_input("ts","",placeholder="Search project, customer, location...",label_visibility="collapsed",key="ts")
with fc2[1]: ts_st = st.selectbox("ss",["All Statuses"]+sorted(df_all["Status"].dropna().unique().tolist()),label_visibility="collapsed",key="ss")
with fc2[2]: ts_rg = st.selectbox("sr",["All Regions"]+sorted(df_all["Region"].dropna().unique().tolist()),label_visibility="collapsed",key="sr")
with fc2[3]: ts_cu = st.selectbox("sc",["All Customers"]+sorted(df_all["Customer"].dropna().unique().tolist()),label_visibility="collapsed",key="sc")
with fc2[4]: ts_mt = st.selectbox("sm",["All Material"]+sorted(df_all["Material_Status"].dropna().unique().tolist()),label_visibility="collapsed",key="sm")
with fc2[5]: ts_pr = st.selectbox("sp",["All Priorities"]+sorted(df_all["Priority"].dropna().unique().tolist()),label_visibility="collapsed",key="sp")
with fc2[6]: 
    st.write("")
    reset = st.button("Reset", key="treset")

tbl = df.copy()
if ts:
    t2 = ts.lower()
    tbl = tbl[tbl["Project_Name"].str.lower().str.contains(t2,na=False)
            | tbl["Customer"].str.lower().str.contains(t2,na=False)
            | tbl["Location"].str.lower().str.contains(t2,na=False)]
if ts_st!="All Statuses":   tbl=tbl[tbl["Status"]==ts_st]
if ts_rg!="All Regions":    tbl=tbl[tbl["Region"]==ts_rg]
if ts_cu!="All Customers":  tbl=tbl[tbl["Customer"]==ts_cu]
if ts_mt!="All Material":   tbl=tbl[tbl["Material_Status"]==ts_mt]
if ts_pr!="All Priorities": tbl=tbl[tbl["Priority"]==ts_pr]
tbl = tbl.sort_values("Overall_Progress",ascending=False)

sc  = {"Completed":"s-completed","In Progress":"s-progress","On Hold":"s-hold","Cancelled":"s-cancelled","Not Started":"s-notstarted"}
mc  = {"Delivered":"m-delivered","Partially Delivered":"m-partial","Ordered":"m-ordered","Not Ordered":"m-notordered"}
pc2 = {"High":"p-high","Medium":"p-medium","Low":"p-low"}

def build_row(r):
    ss = sc.get(r.Status, "")
    ms = mc.get(r.Material_Status, "")
    ps = pc2.get(r.Priority, "")
    return (
        "<tr>"
        + f"<td>{h(r.S_No)}</td>"
        + f"<td>{h(r.Customer)}</td>"
        + f"<td>{h(r.Project_Name)}</td>"
        + f"<td>{h(r.Region)}</td>"
        + f"<td><span class='badge {ss}'>{h(r.Status)}</span></td>"
        + f"<td>{mbar(r.Engineering_Pct, 'eng')}</td>"
        + f"<td>{mbar(r.Delivery_Pct, 'del')}</td>"
        + f"<td>{mbar(r.Execution_Pct, 'exc')}</td>"
        + f"<td>{mbar(r.Overall_Progress, 'all')}</td>"
        + f"<td><span class='badge {ms}'>{h(r.Material_Status)}</span></td>"
        + f"<td><span class='badge {ps}'>{h(r.Priority)}</span></td>"
        + "</tr>"
    )
rows_html = "".join(build_row(r) for _, r in tbl.iterrows())

st.html(f"""
<div style="font-size:12px;color:#94a3b8;margin-bottom:8px">{len(tbl)} rows</div>
<div class="twrap">
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
st.download_button("Download CSV", data=csv, file_name="projects.csv", mime="text/csv")
