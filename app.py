from __future__ import annotations
from io import BytesIO
from pathlib import Path
import json, requests, pandas as pd, streamlit as st
import streamlit.components.v1 as components
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

# hide streamlit chrome
st.markdown("""
<style>
#MainMenu,footer,header,[data-testid="stToolbar"],[data-testid="stDecoration"]{display:none!important;}
.block-container{padding:0!important;margin:0!important;max-width:100%!important;}
section[data-testid="stSidebar"]{display:none!important;}
[data-testid="collapsedControl"]{display:none!important;}
</style>""", unsafe_allow_html=True)

LOCAL_EXCEL_FILE = "Project_Tracking_v7.xlsx"
SHEET_NAME       = "Project_Master"

def _secret(name, default=""):
    try:    return str(st.secrets.get(name, default)).strip()
    except: return default

try:    REFRESH_SECONDS = int(_secret("REFRESH_SECONDS","60") or 60)
except: REFRESH_SECONDS = 60

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

# ---------- data source ----------
def make_dl_url(url):
    url = url.strip()
    if not url or "download=1" in url: return url
    return url + ("&" if "?" in url else "?") + "download=1"

@st.cache_data(ttl=REFRESH_SECONDS)
def load_url(url):
    r = requests.get(make_dl_url(url), timeout=60, allow_redirects=True)
    r.raise_for_status()
    ct = r.headers.get("content-type","").lower()
    if "text/html" in ct and len(r.content) < 500_000:
        raise RuntimeError("URL returned HTML not Excel. Check your OneDrive share link.")
    return r.content

@st.cache_data(ttl=REFRESH_SECONDS)
def load_graph():
    if msal is None: raise RuntimeError("msal not installed")
    tid = _secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    cid = _secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    cs  = _secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")
    if not all([tid,cid,cs]): raise RuntimeError("Missing Graph secrets")
    app = msal.ConfidentialClientApplication(
        cid, authority=f"https://login.microsoftonline.com/{tid}", client_credential=cs)
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res: raise RuntimeError(f"Token error: {res}")
    uid = _secret("GRAPH_USER_ID"); fp = _secret("ONEDRIVE_FILE_PATH")
    if not fp.startswith("/"): fp = "/" + fp
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{uid}/drive/root:{fp}:/content",
        headers={"Authorization": f"Bearer {res['access_token']}"}, timeout=60)
    r.raise_for_status(); return r.content

def get_bytes():
    eu = _secret("EXCEL_FILE_URL")
    if eu: return load_url(eu), "OneDrive"
    if _secret("GRAPH_USER_ID"): return load_graph(), "Graph OneDrive"
    lp = Path(LOCAL_EXCEL_FILE)
    if lp.exists(): return lp.read_bytes(), LOCAL_EXCEL_FILE
    st.error("No data source. Set EXCEL_FILE_URL in Streamlit secrets."); st.stop()

# ---------- parse ----------
def ct(v, d=""):
    if pd.isna(v): return d
    t = str(v).strip(); return t if t else d

def cs(v):
    v = ct(v,"Not Started")
    m = {"completed":"Completed","complete":"Completed",
         "in progress":"In Progress","on progress":"In Progress","progress":"In Progress",
         "on hold":"On Hold","hold":"On Hold",
         "cancelled":"Cancelled","canceled":"Cancelled",
         "not started":"Not Started","notstart":"Not Started"}
    return m.get(v.lower(), v)

@st.cache_data(ttl=REFRESH_SECONDS)
def parse(file_bytes):
    raw  = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    data = raw.iloc[2:, :len(BASE_COLUMNS)].copy()
    data.columns = BASE_COLUMNS
    data = data.dropna(how="all")
    data = data[~(data["S_No"].isna() & data["Project_Name"].isna())]
    for col in ["Job_Ref","LPO_Ref","Customer","Project_Name","Region","Location","Priority"]:
        data[col] = data[col].apply(lambda x: ct(x, "Unknown" if col in ["Customer","Region"] else ""))
    data["Material_Status"] = data["Material_Status"].apply(lambda x: ct(x,"Not Ordered"))
    data["Status"]          = data["Status"].apply(cs)
    data["Work_Status"]     = data["Work_Status"].apply(lambda x: ct(x,""))
    for col in ["Overall_Progress","Engineering_Pct","Delivery_Pct","Execution_Pct"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0).clip(0,100)
    for gn, cols in [("Eng",ENGINEERING_COLS),("Del",DELIVERY_COLS)]:
        blk = data[cols].fillna("").astype(str).apply(lambda s: s.str.strip().str.upper())
        data[f"{gn}_Done"]    = (blk=="DONE").sum(axis=1)
        data[f"{gn}_Partial"] = (blk=="PART.DONE").sum(axis=1)
        data[f"{gn}_NA"]      = (blk=="N/A").sum(axis=1)
        data[f"{gn}_vals"]    = blk.values.tolist()
    data["S_No"]     = data["S_No"].fillna("").astype(str).str.replace(".0","",regex=False)
    data["Priority"] = data["Priority"].replace("","Medium")
    return data.reset_index(drop=True)

def df_to_json(df):
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "s_no":            str(r.S_No),
            "job_ref":         str(r.Job_Ref),
            "lpo_ref":         str(r.LPO_Ref),
            "customer":        str(r.Customer),
            "project_name":    str(r.Project_Name),
            "region":          str(r.Region),
            "location":        str(r.Location),
            "work_status":     str(r.Work_Status),
            "material_status": str(r.Material_Status),
            "progress":        round(float(r.Overall_Progress), 1),
            "priority":        str(r.Priority),
            "status":          str(r.Status),
            "eng_pct":         round(float(r.Engineering_Pct), 1),
            "del_pct":         round(float(r.Delivery_Pct), 1),
            "exec_pct":        round(float(r.Execution_Pct), 1),
            "eng_done":        int(r.Eng_Done),
            "del_done":        int(r.Del_Done),
            "del_vals":        [str(v) for v in r.Del_vals],
        })
    return json.dumps(rows, ensure_ascii=True)

# ---------- load ----------
st_autorefresh(interval=REFRESH_SECONDS*1000, key="ar")

try:
    fb, src = get_bytes()
    df      = parse(fb)
except Exception as e:
    st.error(f"Cannot load data: {e}"); st.stop()

now         = pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")
data_json   = df_to_json(df)
del_items_json = json.dumps(DELIVERY_COLS, ensure_ascii=True)

# ---------- HTML dashboard ----------
HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;
  background:linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%);
  color:#f1f5f9;min-height:100vh;padding:20px;}
.header{display:flex;justify-content:space-between;align-items:flex-start;
  margin-bottom:18px;padding-bottom:14px;border-bottom:1px solid #334155;gap:20px;flex-wrap:wrap;}
.header h1{font-size:24px;font-weight:700;color:#f1f5f9;}
.header .sub{color:#94a3b8;font-size:13px;margin-top:4px;}
.header .right{text-align:right;color:#94a3b8;font-size:12px;line-height:1.8;}
.qbar{display:flex;gap:8px;margin-bottom:18px;flex-wrap:wrap;align-items:center;}
.ql{color:#94a3b8;font-size:12px;text-transform:uppercase;letter-spacing:.5px;margin-right:4px;}
.btn{padding:8px 14px;border-radius:8px;border:none;font-size:13px;font-weight:500;
  cursor:pointer;color:#fff;transition:all .2s;}
.btn:hover{filter:brightness(1.15);}
.btn-green{background:#22c55e;}.btn-blue{background:#3b82f6;}
.btn-orange{background:#f97316;}.btn-gray{background:#475569;}
.kpis{display:grid;grid-template-columns:repeat(8,1fr);gap:12px;margin-bottom:20px;}
@media(max-width:1200px){.kpis{grid-template-columns:repeat(4,1fr);}}
.kpi{background:#1e293b;border-radius:12px;padding:16px 14px;border-left:4px solid var(--ac,#3b82f6);
  cursor:pointer;transition:transform .15s;}
.kpi:hover{transform:translateY(-2px);}
.kpi .lbl{color:#94a3b8;font-size:11px;text-transform:uppercase;letter-spacing:.5px;}
.kpi .val{font-size:30px;font-weight:700;margin-top:6px;color:#f1f5f9;}
.kpi .dlt{font-size:11px;color:#94a3b8;margin-top:3px;}
.stitle{font-size:12px;color:#94a3b8;margin:6px 0 12px;text-transform:uppercase;letter-spacing:.5px;font-weight:600;}
.pgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px;}
@media(max-width:900px){.pgrid{grid-template-columns:1fr;}}
.pc{background:#1e293b;border-radius:12px;padding:18px;border-top:4px solid var(--ac,#3b82f6);}
.pc .pt{font-size:14px;font-weight:600;display:flex;justify-content:space-between;
  align-items:center;margin-bottom:12px;color:#f1f5f9;}
.pc .pp{font-size:26px;font-weight:700;}
.pc .pb{height:10px;background:#334155;border-radius:5px;overflow:hidden;margin-bottom:10px;}
.pc .pf{height:100%;border-radius:5px;}
.pc .ps{display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;}
.pc .ps strong{color:#f1f5f9;font-size:13px;}
.row{display:grid;gap:16px;margin-bottom:20px;}
.row-3{grid-template-columns:repeat(3,1fr);}
.row-2{grid-template-columns:1fr 1fr;}
.row-7-5{grid-template-columns:7fr 5fr;}
@media(max-width:900px){.row-3,.row-2,.row-7-5{grid-template-columns:1fr;}}
.card{background:#1e293b;border-radius:12px;padding:18px;}
.card h3{font-size:14px;font-weight:600;margin-bottom:12px;
  display:flex;justify-content:space-between;align-items:center;color:#f1f5f9;}
.card h3 .tag{font-size:11px;color:#94a3b8;font-weight:400;}
.chart-wrap{position:relative;height:280px;}
.chart-wrap.tall{height:380px;}
.chart-wrap.xtall{height:460px;}
.frow{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;align-items:center;}
.frow input,.frow select{
  background:#334155;color:#f1f5f9;border:1px solid #475569;
  padding:7px 10px;border-radius:8px;font-size:12px;outline:none;
  font-family:'Segoe UI',Tahoma,Arial,sans-serif;}
.frow input{min-width:200px;}
.twrap{overflow:auto;max-height:540px;border-radius:8px;}
table{width:100%;border-collapse:collapse;font-size:12px;}
thead{position:sticky;top:0;background:#334155;z-index:5;}
th{text-align:left;padding:10px 8px;font-weight:600;
  border-bottom:2px solid #475569;white-space:nowrap;color:#f1f5f9;}
td{padding:8px 8px;border-bottom:1px solid #334155;color:#cbd5e1;}
tr:hover td{background:rgba(59,130,246,.08);}
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
.bc{display:flex;align-items:center;gap:6px;min-width:90px;}
.bb{flex:1;height:6px;background:#334155;border-radius:3px;overflow:hidden;min-width:45px;}
.bf{height:100%;border-radius:3px;}
.beng{background:linear-gradient(90deg,#60a5fa,#3b82f6);}
.bdel{background:linear-gradient(90deg,#c084fc,#a855f7);}
.bexc{background:linear-gradient(90deg,#fb923c,#f97316);}
.ball{background:linear-gradient(90deg,#ef4444,#eab308,#22c55e);}
.bt{font-size:11px;font-weight:600;min-width:30px;text-align:right;color:#cbd5e1;}
.rc{font-size:12px;color:#94a3b8;margin-bottom:8px;}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>Project Tracking Dashboard</h1>
    <div class="sub">Cold Rooms, Cabinets &amp; Refrigeration Projects &mdash; Live from Excel</div>
  </div>
  <div class="right">
    <div>Updated: __NOW__</div>
    <div>Source: __SRC__ &nbsp;|&nbsp; Total: <span id="totalCount">0</span> projects</div>
  </div>
</div>

<div class="qbar">
  <span class="ql">Quick Views:</span>
  <button class="btn btn-green" onclick="qf('Completed')">Completed</button>
  <button class="btn btn-blue"  onclick="qf('In Progress')">In Progress</button>
  <button class="btn btn-orange" onclick="qf('On Hold')">On Hold</button>
  <button class="btn btn-gray"  onclick="qf('Not Started')">Not Started</button>
  <button class="btn btn-gray"  onclick="qf('Cancelled')">Cancelled</button>
  <button class="btn btn-gray"  onclick="resetQ()">Show All</button>
</div>

<div class="kpis" id="kpiRow"></div>

<div class="stitle">Phase Progress Overview</div>
<div class="pgrid" id="phaseCards"></div>

<div class="row row-3">
  <div class="card"><h3>Project Status</h3><div class="chart-wrap"><canvas id="cStatus"></canvas></div></div>
  <div class="card"><h3>Material Status</h3><div class="chart-wrap"><canvas id="cMaterial"></canvas></div></div>
  <div class="card"><h3>Priority Distribution</h3><div class="chart-wrap"><canvas id="cPriority"></canvas></div></div>
</div>

<div class="row row-2">
  <div class="card"><h3>Projects by Region</h3><div class="chart-wrap"><canvas id="cRegion"></canvas></div></div>
  <div class="card"><h3>Top 10 Customers</h3><div class="chart-wrap"><canvas id="cCustomer"></canvas></div></div>
</div>

<div class="row row-7-5">
  <div class="card">
    <h3>Delivery Items Status <span class="tag">Done / Partial / Not Done / N/A per item</span></h3>
    <div class="chart-wrap tall"><canvas id="cDelivery"></canvas></div>
  </div>
  <div class="card">
    <h3>Overall Progress Buckets</h3>
    <div class="chart-wrap tall"><canvas id="cBuckets"></canvas></div>
  </div>
</div>

<div class="card" style="margin-bottom:20px">
  <h3>Phase Progress by Project <span class="tag">Engineering / Delivery / Execution &mdash; top 25 by overall %</span></h3>
  <div class="chart-wrap xtall"><canvas id="cPhase"></canvas></div>
</div>

<div class="card">
  <h3>Project Details <span class="tag" id="rowCount"></span></h3>
  <div class="frow">
    <input  type="text"   id="srch"  placeholder="Search project, customer, location, job ref...">
    <select id="fStatus"> <option value="">All Statuses</option></select>
    <select id="fRegion">  <option value="">All Regions</option></select>
    <select id="fCustomer"><option value="">All Customers</option></select>
    <select id="fMaterial"><option value="">All Material</option></select>
    <select id="fPriority"><option value="">All Priorities</option></select>
    <button class="btn btn-gray" onclick="resetFilters()">Reset</button>
  </div>
  <div class="twrap">
    <table>
      <thead><tr>
        <th>#</th><th>Customer</th><th>Project</th><th>Region</th><th>Status</th>
        <th>Eng %</th><th>Delivery %</th><th>Exec %</th><th>Overall %</th>
        <th>Material</th><th>Priority</th>
      </tr></thead>
      <tbody id="tBody"></tbody>
    </table>
  </div>
</div>

<script>
Chart.defaults.color='#94a3b8';
Chart.defaults.borderColor='#334155';
Chart.defaults.font.family="'Segoe UI',Tahoma,Arial,sans-serif";

const ALL_DATA = __DATA_JSON__;
const DEL_ITEMS = __DEL_ITEMS_JSON__;
let DATA = ALL_DATA.slice();
let activeQ = '';
const charts = {};

function esc(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function avg(arr){ return arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : 0; }
function uniq(key){ return [...new Set(ALL_DATA.map(r=>r[key]).filter(Boolean))].sort(); }

function destroyAll(){ Object.values(charts).forEach(c=>c.destroy()); Object.keys(charts).forEach(k=>delete charts[k]); }

// ---- KPIs ----
function buildKPIs(rows){
  const N=rows.length;
  const by=s=>rows.filter(r=>r.status===s).length;
  const comp=by('Completed'),prog=by('In Progress'),hold=by('On Hold'),ns=by('Not Started'),can=by('Cancelled');
  const matd=rows.filter(r=>r.material_status==='Delivered').length;
  const avgP=avg(rows.map(r=>+r.progress||0));
  const pp=n=>N?Math.round(n/N*100)+'%':'';
  document.getElementById('totalCount').textContent=N;
  document.getElementById('kpiRow').innerHTML=[
    ['Total Projects', N, 'In current view', '#3b82f6',''],
    ['Completed',      comp, pp(comp),       '#22c55e', "qf('Completed')"],
    ['In Progress',    prog, pp(prog),        '#14b8a6', "qf('In Progress')"],
    ['On Hold',        hold, pp(hold),        '#f97316', "qf('On Hold')"],
    ['Not Started',    ns,   pp(ns),          '#64748b', "qf('Not Started')"],
    ['Cancelled',      can,  pp(can),         '#ef4444', "qf('Cancelled')"],
    ['Overall Avg %',  avgP.toFixed(1)+'%','Across projects','#ec4899',''],
    ['Mat. Delivered', matd, pp(matd),        '#eab308', ''],
  ].map(([lbl,val,dlt,ac,fn])=>
    `<div class="kpi" style="--ac:${ac}" onclick="${fn}">
      <div class="lbl">${lbl}</div>
      <div class="val">${val}</div>
      <div class="dlt">${dlt}</div>
    </div>`
  ).join('');
}

// ---- Phase Cards ----
function buildPhase(rows){
  const N=rows.length||1;
  const e=avg(rows.map(r=>+r.eng_pct||0));
  const d=avg(rows.map(r=>+r.del_pct||0));
  const x=avg(rows.map(r=>+r.exec_pct||0));
  const ed=rows.filter(r=>r.eng_pct>=100).length;
  const dd=rows.filter(r=>r.del_pct>=100).length;
  const xd=rows.filter(r=>r.exec_pct>=100).length;
  document.getElementById('phaseCards').innerHTML=`
    <div class="pc" style="--ac:#3b82f6">
      <div class="pt">Engineering<span class="pp" style="color:#60a5fa">${e.toFixed(1)}%</span></div>
      <div class="pb"><div class="pf" style="width:${e.toFixed(1)}%;background:linear-gradient(90deg,#3b82f6,#60a5fa)"></div></div>
      <div class="ps"><span>Design / Submittal / Drawing / ELS / BOM</span><span><strong>${ed}</strong>/${N} done</span></div>
    </div>
    <div class="pc" style="--ac:#a855f7">
      <div class="pt">Delivery<span class="pp" style="color:#c084fc">${d.toFixed(1)}%</span></div>
      <div class="pb"><div class="pf" style="width:${d.toFixed(1)}%;background:linear-gradient(90deg,#a855f7,#c084fc)"></div></div>
      <div class="ps"><span>Material delivery &mdash; ${DEL_ITEMS.length} items tracked</span><span><strong>${dd}</strong>/${N} done</span></div>
    </div>
    <div class="pc" style="--ac:#f97316">
      <div class="pt">Execution<span class="pp" style="color:#fb923c">${x.toFixed(1)}%</span></div>
      <div class="pb"><div class="pf" style="width:${x.toFixed(1)}%;background:linear-gradient(90deg,#f97316,#fb923c)"></div></div>
      <div class="ps"><span>On-site installation &amp; commissioning</span><span><strong>${xd}</strong>/${N} done</span></div>
    </div>`;
}

// ---- Charts ----
const COLORS=['#3b82f6','#a855f7','#14b8a6','#ec4899','#f97316','#64748b','#eab308','#22c55e','#ef4444','#0ea5e9'];
const STATUS_CLR={'Completed':'#22c55e','In Progress':'#14b8a6','On Hold':'#f97316','Not Started':'#64748b','Cancelled':'#ef4444'};
const MAT_CLR={'Delivered':'#22c55e','Partially Delivered':'#eab308','Ordered':'#f97316','Not Ordered':'#ef4444'};
const PRI_CLR={'High':'#ef4444','Medium':'#eab308','Low':'#22c55e'};

function donut(id, labels, data, colors){
  return new Chart(document.getElementById(id),{
    type:'doughnut',
    data:{labels,datasets:[{data,backgroundColor:colors,borderWidth:2,borderColor:'#1e293b'}]},
    options:{
      cutout:'60%',
      plugins:{
        legend:{position:'bottom',labels:{font:{size:11},padding:10,color:'#94a3b8',boxWidth:12}},
        tooltip:{callbacks:{label:ctx=>' '+ctx.label+': '+ctx.raw+' ('+Math.round(ctx.raw/ctx.dataset.data.reduce((a,b)=>a+b,0)*100)+'%)'}}
      }
    }
  });
}

function buildCharts(rows){
  destroyAll();

  // status donut
  const stO=['Completed','In Progress','On Hold','Not Started','Cancelled'];
  const stC=stO.map(s=>rows.filter(r=>r.status===s).length);
  charts.st=donut('cStatus',stO,stC,stO.map(s=>STATUS_CLR[s]));

  // material donut
  const mtO=['Delivered','Partially Delivered','Ordered','Not Ordered'];
  const mtC=mtO.map(s=>rows.filter(r=>r.material_status===s).length);
  charts.mt=donut('cMaterial',mtO,mtC,mtO.map(s=>MAT_CLR[s]));

  // priority donut
  const prO=['High','Medium','Low'];
  const prC=prO.map(s=>rows.filter(r=>r.priority===s).length);
  charts.pr=donut('cPriority',prO,prC,prO.map(s=>PRI_CLR[s]));

  // region bar
  const rgM={};rows.forEach(r=>{rgM[r.region]=(rgM[r.region]||0)+1;});
  const rgE=Object.entries(rgM).sort((a,b)=>b[1]-a[1]);
  charts.rg=new Chart(document.getElementById('cRegion'),{
    type:'bar',
    data:{labels:rgE.map(e=>e[0]),datasets:[{
      data:rgE.map(e=>e[1]),
      backgroundColor:rgE.map((_,i)=>COLORS[i%COLORS.length]),
      borderRadius:6,borderWidth:0
    }]},
    options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{stepSize:1}},x:{ticks:{color:'#94a3b8'}}},
      animation:{duration:600}}
  });

  // customers bar
  const csM={};rows.forEach(r=>{csM[r.customer]=(csM[r.customer]||0)+1;});
  const top10=Object.entries(csM).sort((a,b)=>b[1]-a[1]).slice(0,10);
  charts.cs=new Chart(document.getElementById('cCustomer'),{
    type:'bar',
    data:{labels:top10.map(e=>e[0]),datasets:[{
      data:top10.map(e=>e[1]),
      backgroundColor:'#3b82f6',borderRadius:6,borderWidth:0
    }]},
    options:{indexAxis:'y',plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true,ticks:{stepSize:1,color:'#94a3b8'}},y:{ticks:{color:'#94a3b8'}}},
      animation:{duration:600}}
  });

  // delivery items stacked horizontal
  const nI=DEL_ITEMS.length;
  const done=new Array(nI).fill(0),part=new Array(nI).fill(0),ndone=new Array(nI).fill(0),naArr=new Array(nI).fill(0);
  rows.forEach(r=>{
    (r.del_vals||[]).forEach((v,i)=>{
      if(i>=nI)return;
      const u=String(v||'').trim().toUpperCase();
      if(u==='DONE')done[i]++;
      else if(u==='PART.DONE')part[i]++;
      else if(u==='N.DONE')ndone[i]++;
      else if(u==='N/A')naArr[i]++;
    });
  });
  charts.di=new Chart(document.getElementById('cDelivery'),{
    type:'bar',
    data:{labels:DEL_ITEMS,datasets:[
      {label:'Done',    data:done, backgroundColor:'#22c55e',stack:'s',borderWidth:0},
      {label:'Partial', data:part, backgroundColor:'#eab308',stack:'s',borderWidth:0},
      {label:'Not Done',data:ndone,backgroundColor:'#ef4444',stack:'s',borderWidth:0},
      {label:'N/A',     data:naArr,backgroundColor:'#475569',stack:'s',borderWidth:0},
    ]},
    options:{
      indexAxis:'y',
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}}},
      scales:{
        x:{beginAtZero:true,stacked:true,ticks:{color:'#94a3b8'}},
        y:{stacked:true,ticks:{color:'#94a3b8'}}
      },animation:{duration:600}
    }
  });

  // progress buckets
  const bk={'0-25%':0,'26-50%':0,'51-75%':0,'76-100%':0};
  rows.forEach(r=>{
    const p=+r.progress||0;
    if(p<=25)bk['0-25%']++;else if(p<=50)bk['26-50%']++;else if(p<=75)bk['51-75%']++;else bk['76-100%']++;
  });
  charts.bk=new Chart(document.getElementById('cBuckets'),{
    type:'bar',
    data:{labels:Object.keys(bk),datasets:[{
      data:Object.values(bk),
      backgroundColor:['#ef4444','#f97316','#eab308','#22c55e'],
      borderRadius:8,borderWidth:0
    }]},
    options:{
      plugins:{legend:{display:false},
        datalabels:{display:false}},
      scales:{y:{beginAtZero:true,ticks:{color:'#94a3b8'}},x:{ticks:{color:'#94a3b8'}}},
      animation:{duration:600}
    }
  });

  // phase by project grouped bar
  const sorted=[...rows].sort((a,b)=>(+b.progress||0)-(+a.progress||0)).slice(0,25);
  const ylbls=sorted.map(r=>'#'+r.s_no+' '+(r.project_name||r.customer).substring(0,30));
  charts.pp=new Chart(document.getElementById('cPhase'),{
    type:'bar',
    data:{labels:ylbls,datasets:[
      {label:'Engineering %',data:sorted.map(r=>+r.eng_pct||0), backgroundColor:'#3b82f6',borderRadius:3,borderWidth:0},
      {label:'Delivery %',   data:sorted.map(r=>+r.del_pct||0), backgroundColor:'#a855f7',borderRadius:3,borderWidth:0},
      {label:'Execution %',  data:sorted.map(r=>+r.exec_pct||0),backgroundColor:'#f97316',borderRadius:3,borderWidth:0},
    ]},
    options:{
      indexAxis:'y',
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}},
        tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+ctx.raw+'%'}}},
      scales:{
        x:{beginAtZero:true,max:100,title:{display:true,text:'Percent Complete',color:'#94a3b8'},ticks:{color:'#94a3b8'}},
        y:{ticks:{color:'#94a3b8'}}
      },animation:{duration:600}
    }
  });
}

// ---- Table ----
function mbar(val,cls){
  const p=Math.max(0,Math.min(100,+val||0));
  return `<div class="bc"><div class="bb"><div class="bf ${cls}" style="width:${p}%"></div></div><div class="bt">${p}%</div></div>`;
}
const SC={'Completed':'s-completed','In Progress':'s-progress','On Hold':'s-hold','Cancelled':'s-cancelled','Not Started':'s-notstarted'};
const MC={'Delivered':'m-delivered','Partially Delivered':'m-partial','Ordered':'m-ordered','Not Ordered':'m-notordered'};
const PC={'High':'p-high','Medium':'p-medium','Low':'p-low'};

function buildTable(rows){
  document.getElementById('rowCount').textContent=rows.length+' rows';
  document.getElementById('tBody').innerHTML=rows.map(r=>`
    <tr>
      <td>${esc(r.s_no)}</td>
      <td>${esc(r.customer)}</td>
      <td>${esc(r.project_name)}</td>
      <td>${esc(r.region)}</td>
      <td><span class="badge ${SC[r.status]||''}">${esc(r.status)}</span></td>
      <td>${mbar(r.eng_pct,'beng')}</td>
      <td>${mbar(r.del_pct,'bdel')}</td>
      <td>${mbar(r.exec_pct,'bexc')}</td>
      <td>${mbar(r.progress,'ball')}</td>
      <td><span class="badge ${MC[r.material_status]||''}">${esc(r.material_status)}</span></td>
      <td><span class="badge ${PC[r.priority]||''}">${esc(r.priority)}</span></td>
    </tr>`).join('');
}

// ---- Filters ----
function populateFilters(){
  const add=(id,opts)=>{
    const sel=document.getElementById(id);
    const cur=sel.value;
    while(sel.options.length>1)sel.remove(1);
    opts.forEach(v=>{const o=document.createElement('option');o.value=o.textContent=v;sel.appendChild(o);});
    if(cur)sel.value=cur;
  };
  add('fStatus',['Not Started','In Progress','Completed','On Hold','Cancelled']);
  add('fRegion', uniq('region'));
  add('fCustomer',uniq('customer'));
  add('fMaterial',['Delivered','Partially Delivered','Ordered','Not Ordered']);
  add('fPriority',['High','Medium','Low']);
}

function getFiltered(){
  const q=(document.getElementById('srch').value||'').toLowerCase();
  const fs=document.getElementById('fStatus').value;
  const fr=document.getElementById('fRegion').value;
  const fc=document.getElementById('fCustomer').value;
  const fm=document.getElementById('fMaterial').value;
  const fp=document.getElementById('fPriority').value;
  return DATA.filter(r=>{
    if(fs&&r.status!==fs)return false;
    if(fr&&r.region!==fr)return false;
    if(fc&&r.customer!==fc)return false;
    if(fm&&r.material_status!==fm)return false;
    if(fp&&r.priority!==fp)return false;
    if(q){
      const hay=(r.project_name+' '+r.customer+' '+r.location+' '+r.job_ref+' '+r.lpo_ref).toLowerCase();
      if(!hay.includes(q))return false;
    }
    return true;
  });
}

function refresh(){
  const rows=getFiltered();
  buildKPIs(rows);buildPhase(rows);buildCharts(rows);buildTable(rows);
}

function qf(status){
  activeQ=status;
  DATA=ALL_DATA.filter(r=>r.status===status);
  resetFilters(false);
}
function resetQ(){
  activeQ='';DATA=ALL_DATA.slice();resetFilters(false);
}
function resetFilters(full=true){
  if(full){activeQ='';DATA=ALL_DATA.slice();}
  ['srch','fStatus','fRegion','fCustomer','fMaterial','fPriority'].forEach(id=>{
    const el=document.getElementById(id);
    if(el.tagName==='INPUT')el.value='';else el.value='';
  });
  refresh();
}

['srch','fStatus','fRegion','fCustomer','fMaterial','fPriority'].forEach(id=>{
  document.getElementById(id).addEventListener('input',refresh);
  document.getElementById(id).addEventListener('change',refresh);
});

populateFilters();
refresh();
</script>
</body>
</html>"""

# inject Python data into the HTML
HTML = HTML.replace("__DATA_JSON__", data_json)
HTML = HTML.replace("__DEL_ITEMS_JSON__", del_items_json)
HTML = HTML.replace("__NOW__", now)
HTML = HTML.replace("__SRC__", src)

components.html(HTML, height=4800, scrolling=True)
