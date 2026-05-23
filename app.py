from __future__ import annotations
from io import BytesIO
from pathlib import Path
import json, requests, pandas as pd
import streamlit as st
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

st.markdown("""
<style>
#MainMenu,footer,header,[data-testid="stToolbar"],
[data-testid="stDecoration"],[data-testid="collapsedControl"]{display:none!important;}
.block-container{padding:0!important;margin:0!important;max-width:100%!important;}
section[data-testid="stSidebar"]{display:none!important;}
</style>""", unsafe_allow_html=True)

# -- CONFIG --------------------------------------------------------------------
LOCAL_FILE   = "Project_Tracking_v7.xlsx"
SHEET_NAME   = "Project_Master"

def _secret(k, d=""):
    try:    return str(st.secrets.get(k, d)).strip()
    except: return d

try:    REFRESH = int(_secret("REFRESH_SECONDS","60") or 60)
except: REFRESH = 60

# Exact column positions confirmed from Project_Tracking_v7.xlsx
C_SNO     = 0
C_JOB     = 2
C_LPO     = 3
C_CUST    = 4
C_PROJ    = 5
C_REGION  = 6
C_LOC     = 7
C_ENG     = list(range(9,14))    # cols 9-13: Equi.Design, Tech Sub, Drawing, ELS, BOM
C_DEL     = list(range(14,25))   # cols 14-24: 11 delivery items
C_WSTAT   = 25
C_MATSTAT = 27
C_PROG    = 28
C_PRI     = 29
C_STATUS  = 30
C_ENGPCT  = 31
C_DELPCT  = 32
C_EXCPCT  = 33

DEL_NAMES = ["Out_Door","Indoor","CR Panels","CR Ins. Materials","Doors",
             "Ref. Inst. Materials","CCP","Display CCP","Floor Heater","Cabinets","Any Special"]

STATUS_MAP = {
    "completed":"Completed","complete":"Completed",
    "in progress":"In Progress","on progress":"In Progress","progress":"In Progress",
    "on hold":"On Hold","hold":"On Hold",
    "cancelled":"Cancelled","canceled":"Cancelled",
    "not started":"Not Started","notstart":"Not Started",
}

# -- DATA SOURCE ---------------------------------------------------------------
def _dl_url(url):
    url = url.strip()
    if not url or "download=1" in url: return url
    return url + ("&" if "?" in url else "?") + "download=1"

@st.cache_data(ttl=REFRESH)
def load_onedrive(url):
    r = requests.get(_dl_url(url), timeout=60, allow_redirects=True)
    r.raise_for_status()
    if "text/html" in r.headers.get("content-type","").lower() and len(r.content)<500_000:
        raise RuntimeError("OneDrive returned an HTML page, not the Excel file. Check your sharing link.")
    return r.content

@st.cache_data(ttl=REFRESH)
def load_graph():
    if msal is None: raise RuntimeError("msal not installed")
    tid = _secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    cid = _secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    cs  = _secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")
    app = msal.ConfidentialClientApplication(
        cid, authority=f"https://login.microsoftonline.com/{tid}", client_credential=cs)
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res: raise RuntimeError(str(res))
    uid = _secret("GRAPH_USER_ID")
    fp  = _secret("ONEDRIVE_FILE_PATH")
    if not fp.startswith("/"): fp = "/" + fp
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{uid}/drive/root:{fp}:/content",
        headers={"Authorization": f"Bearer {res['access_token']}"}, timeout=60)
    r.raise_for_status()
    return r.content

def get_file():
    eu = _secret("EXCEL_FILE_URL")
    if eu: return load_onedrive(eu), "OneDrive"
    if _secret("GRAPH_USER_ID"): return load_graph(), "Graph OneDrive"
    lp = Path(LOCAL_FILE)
    if lp.exists(): return lp.read_bytes(), LOCAL_FILE
    st.error("No data source found. Add EXCEL_FILE_URL to Streamlit secrets.")
    st.stop()

# -- PARSE ---------------------------------------------------------------------
def _v(cell, default=""):
    if cell is None or (isinstance(cell, float) and pd.isna(cell)): return default
    s = str(cell).strip()
    return s if s and s.lower() != "nan" else default

def _n(cell):
    try: return round(float(cell), 1)
    except: return 0.0

def _status(raw):
    s = _v(raw, "Not Started")
    return STATUS_MAP.get(s.lower(), s)

@st.cache_data(ttl=REFRESH)
def parse(file_bytes: bytes):
    raw  = pd.read_excel(BytesIO(file_bytes), sheet_name=SHEET_NAME,
                         header=None, engine="openpyxl")
    rows = []
    for idx in range(2, len(raw)):   # row 0=group header, row 1=col names, row 2+ = data
        r = raw.iloc[idx]
        sno  = _v(r[C_SNO])
        cust = _v(r[C_CUST], "Unknown")
        proj = _v(r[C_PROJ])
        if not sno and not cust and not proj:
            continue

        eng_vals = [_v(r[c]) for c in C_ENG]
        del_vals = [_v(r[c]) for c in C_DEL]

        def cnt(lst, val): return sum(1 for x in lst if x.upper()==val)

        rows.append({
            "s_no":            sno,
            "job_ref":         _v(r[C_JOB], "N/A"),
            "lpo_ref":         _v(r[C_LPO], "N/A"),
            "customer":        cust,
            "project_name":    proj,
            "region":          _v(r[C_REGION], "Unknown"),
            "location":        _v(r[C_LOC]),
            "work_status":     _v(r[C_WSTAT]),
            "material_status": _v(r[C_MATSTAT], "Not Ordered"),
            "progress":        _n(r[C_PROG]),
            "priority":        _v(r[C_PRI], "Medium") or "Medium",
            "status":          _status(r[C_STATUS]),
            "eng_pct":         _n(r[C_ENGPCT]),
            "del_pct":         _n(r[C_DELPCT]),
            "exec_pct":        _n(r[C_EXCPCT]),
            "eng_done":        cnt(eng_vals,"DONE"),
            "del_done":        cnt(del_vals,"DONE"),
            "del_vals":        del_vals,
        })
    return rows

# -- LOAD ----------------------------------------------------------------------
st_autorefresh(interval=REFRESH*1000, key="ar")

try:
    fb, src = get_file()
    projects = parse(fb)
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

now = pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")
data_js      = json.dumps(projects, ensure_ascii=True)
del_names_js = json.dumps(DEL_NAMES, ensure_ascii=True)

# -- HTML DASHBOARD ------------------------------------------------------------
HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;
  background:linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%);
  color:#f1f5f9;min-height:100vh;padding:20px 24px;}

/* HEADER */
.hdr{display:flex;justify-content:space-between;align-items:flex-start;
  margin-bottom:18px;padding-bottom:14px;border-bottom:1px solid #334155;gap:16px;flex-wrap:wrap;}
.hdr h1{font-size:24px;font-weight:700;color:#f1f5f9;}
.hdr .sub{color:#94a3b8;font-size:13px;margin-top:3px;}
.hdr .meta{text-align:right;color:#94a3b8;font-size:12px;line-height:2;}

/* QUICK BAR */
.qbar{display:flex;gap:8px;margin-bottom:18px;flex-wrap:wrap;align-items:center;}
.ql{color:#94a3b8;font-size:12px;text-transform:uppercase;letter-spacing:.5px;margin-right:2px;}
.btn{padding:8px 14px;border-radius:8px;border:none;font-size:13px;font-weight:500;
     cursor:pointer;color:#fff;transition:opacity .15s;}
.btn:hover{opacity:.85;}
.btn-green{background:#22c55e;} .btn-blue{background:#3b82f6;}
.btn-teal{background:#14b8a6;} .btn-orange{background:#f97316;}
.btn-gray{background:#475569;} .btn-red{background:#ef4444;}
.btn.active{outline:3px solid #fff;outline-offset:2px;}

/* KPIs */
.kpis{display:grid;grid-template-columns:repeat(8,1fr);gap:12px;margin-bottom:20px;}
@media(max-width:1100px){.kpis{grid-template-columns:repeat(4,1fr);}}
@media(max-width:600px){.kpis{grid-template-columns:repeat(2,1fr);}}
.kpi{background:#1e293b;border-radius:12px;padding:16px 14px 14px;
     border-left:4px solid var(--ac,#3b82f6);cursor:pointer;transition:transform .15s;}
.kpi:hover{transform:translateY(-2px);}
.kpi .lbl{color:#94a3b8;font-size:11px;text-transform:uppercase;letter-spacing:.5px;}
.kpi .val{font-size:30px;font-weight:700;margin-top:6px;color:#f1f5f9;}
.kpi .dlt{font-size:11px;color:#64748b;margin-top:3px;}

/* SECTION TITLE */
.stl{font-size:12px;color:#94a3b8;margin:4px 0 12px;
     text-transform:uppercase;letter-spacing:.5px;font-weight:600;}

/* PHASE CARDS */
.pgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px;}
@media(max-width:768px){.pgrid{grid-template-columns:1fr;}}
.pc{background:#1e293b;border-radius:12px;padding:18px;
    border-top:4px solid var(--ac,#3b82f6);}
.pc .pt{font-size:14px;font-weight:600;display:flex;justify-content:space-between;
        align-items:center;margin-bottom:12px;color:#f1f5f9;}
.pc .pp{font-weight:700;font-size:26px;}
.pc .pb{height:10px;background:#334155;border-radius:5px;overflow:hidden;margin-bottom:10px;}
.pc .pf{height:100%;border-radius:5px;}
.pc .ps{display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;}
.pc .ps strong{color:#f1f5f9;font-size:13px;}

/* ROWS */
.row3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:18px;}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:18px;}
.row75{display:grid;grid-template-columns:7fr 5fr;gap:16px;margin-bottom:18px;}
@media(max-width:900px){.row3,.row2,.row75{grid-template-columns:1fr;}}

/* CARD */
.card{background:#1e293b;border-radius:12px;padding:18px;margin-bottom:0;}
.card h3{font-size:14px;font-weight:600;margin-bottom:12px;color:#f1f5f9;
         display:flex;justify-content:space-between;align-items:center;}
.card h3 .tag{font-size:11px;color:#64748b;font-weight:400;}
.cw{position:relative;height:280px;}
.cw.tall{height:380px;}
.cw.xtall{height:460px;}

/* TABLE FILTERS */
.frow{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px;align-items:center;}
.frow input,.frow select{
  background:#334155;color:#f1f5f9;border:1px solid #475569;
  padding:7px 10px;border-radius:8px;font-size:12px;outline:none;
  font-family:'Segoe UI',Tahoma,Arial,sans-serif;}
.frow input{min-width:210px;}
.frow select{min-width:130px;}

/* TABLE */
.tw{overflow:auto;max-height:540px;border-radius:8px;}
table{width:100%;border-collapse:collapse;font-size:12px;}
thead{position:sticky;top:0;background:#334155;z-index:5;}
th{text-align:left;padding:10px 8px;font-weight:600;
   border-bottom:2px solid #475569;white-space:nowrap;color:#f1f5f9;}
td{padding:8px 8px;border-bottom:1px solid #263248;color:#cbd5e1;
   white-space:nowrap;}
tr:hover td{background:rgba(59,130,246,.07);}

/* BADGES */
.badge{display:inline-block;padding:3px 10px;border-radius:12px;
       font-size:11px;font-weight:600;white-space:nowrap;}
.sc{background:rgba(34,197,94,.2);color:#86efac;}
.sp{background:rgba(20,184,166,.2);color:#5eead4;}
.sh{background:rgba(249,115,22,.2);color:#fdba74;}
.sx{background:rgba(239,68,68,.2);color:#fca5a5;}
.sn{background:rgba(100,116,139,.25);color:#cbd5e1;}
.md{background:rgba(34,197,94,.2);color:#86efac;}
.mp{background:rgba(234,179,8,.2);color:#fde047;}
.mo{background:rgba(249,115,22,.2);color:#fdba74;}
.mn{background:rgba(239,68,68,.2);color:#fca5a5;}
.ph{background:rgba(239,68,68,.2);color:#fca5a5;}
.pm{background:rgba(234,179,8,.2);color:#fde047;}
.pl{background:rgba(34,197,94,.2);color:#86efac;}

/* MINI BARS */
.bc{display:flex;align-items:center;gap:6px;}
.bb{flex:1;height:6px;background:#334155;border-radius:3px;
    overflow:hidden;min-width:50px;}
.bf{height:100%;border-radius:3px;}
.beng{background:linear-gradient(90deg,#60a5fa,#3b82f6);}
.bdel{background:linear-gradient(90deg,#c084fc,#a855f7);}
.bexc{background:linear-gradient(90deg,#fb923c,#f97316);}
.ball{background:linear-gradient(90deg,#ef4444,#eab308,#22c55e);}
.bpct{font-size:11px;font-weight:600;min-width:32px;text-align:right;color:#94a3b8;}

/* ROW COUNT */
.rc{font-size:12px;color:#64748b;margin-bottom:8px;}
</style>
</head>
<body>

<!-- HEADER -->
<div class="hdr">
  <div>
    <h1>Project Tracking Dashboard</h1>
    <div class="sub">Cold Rooms, Cabinets &amp; Refrigeration Projects &mdash; Live from Excel</div>
  </div>
  <div class="meta">
    <div>Updated: TIMESTAMP</div>
    <div>Source: SOURCENAME &nbsp;|&nbsp; <span id="totCount">0</span> projects loaded</div>
  </div>
</div>

<!-- QUICK FILTERS -->
<div class="qbar">
  <span class="ql">Quick:</span>
  <button class="btn btn-green"  id="qAll"      onclick="quickFilter('')">All Projects</button>
  <button class="btn btn-green"  id="qComp"     onclick="quickFilter('Completed')">Completed</button>
  <button class="btn btn-blue"   id="qProg"     onclick="quickFilter('In Progress')">In Progress</button>
  <button class="btn btn-orange" id="qHold"     onclick="quickFilter('On Hold')">On Hold</button>
  <button class="btn btn-gray"   id="qNS"       onclick="quickFilter('Not Started')">Not Started</button>
  <button class="btn btn-red"    id="qCan"      onclick="quickFilter('Cancelled')">Cancelled</button>
</div>

<!-- KPIs -->
<div class="kpis" id="kpiRow"></div>

<!-- PHASE CARDS -->
<div class="stl">Phase Progress Overview</div>
<div class="pgrid" id="phaseCards"></div>

<!-- ROW 1: 3 donuts -->
<div class="row3">
  <div class="card"><h3>Project Status</h3>
    <div class="cw"><canvas id="cStatus"></canvas></div></div>
  <div class="card"><h3>Material Status</h3>
    <div class="cw"><canvas id="cMaterial"></canvas></div></div>
  <div class="card"><h3>Priority Distribution</h3>
    <div class="cw"><canvas id="cPriority"></canvas></div></div>
</div>

<!-- ROW 2: region + customers -->
<div class="row2">
  <div class="card"><h3>Projects by Region</h3>
    <div class="cw"><canvas id="cRegion"></canvas></div></div>
  <div class="card"><h3>Top 10 Customers</h3>
    <div class="cw"><canvas id="cCustomer"></canvas></div></div>
</div>

<!-- ROW 3: delivery + buckets -->
<div class="row75">
  <div class="card">
    <h3>Delivery Items Status <span class="tag">Done / Partial / Not Done / N/A per item</span></h3>
    <div class="cw tall"><canvas id="cDelivery"></canvas></div></div>
  <div class="card">
    <h3>Overall Progress Buckets</h3>
    <div class="cw tall"><canvas id="cBuckets"></canvas></div></div>
</div>

<!-- ROW 4: phase by project -->
<div class="card" style="margin-bottom:18px;">
  <h3>Phase Progress by Project
    <span class="tag">Engineering / Delivery / Execution &mdash; top 25 by overall %</span></h3>
  <div class="cw xtall"><canvas id="cPhase"></canvas></div>
</div>

<!-- TABLE -->
<div class="card">
  <h3>Project Details <span class="tag" id="rowCount"></span></h3>
  <div class="frow">
    <input  type="text"   id="srch"   placeholder="Search project, customer, location, job ref...">
    <select id="fSt"><option value="">All Statuses</option></select>
    <select id="fRg"><option value="">All Regions</option></select>
    <select id="fCu"><option value="">All Customers</option></select>
    <select id="fMt"><option value="">All Material</option></select>
    <select id="fPr"><option value="">All Priorities</option></select>
    <button class="btn btn-gray" onclick="clearFilters()">Clear</button>
  </div>
  <div class="tw">
    <table>
      <thead><tr>
        <th>#</th><th>Customer</th><th>Project</th><th>Region</th>
        <th>Status</th><th>Eng %</th><th>Del %</th><th>Exec %</th>
        <th>Overall %</th><th>Material</th><th>Priority</th><th>Work Status</th>
      </tr></thead>
      <tbody id="tBody"></tbody>
    </table>
  </div>
</div>

<script>
Chart.defaults.color = '#94a3b8';
Chart.defaults.borderColor = '#334155';
Chart.defaults.font.family = "'Segoe UI', Tahoma, Arial, sans-serif";

const ALL = DATA_PLACEHOLDER;
const DEL_ITEMS = DEL_PLACEHOLDER;
const charts = {};
let activeStatus = '';

function e(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function avg(a){ return a.length ? a.reduce((s,x)=>s+x,0)/a.length : 0; }
function uniq(k){ return [...new Set(ALL.map(r=>r[k]).filter(Boolean))].sort(); }
function pct(n,t){ return t ? Math.round(n/t*100)+'%' : ''; }
function destroy(k){ if(charts[k]){ charts[k].destroy(); delete charts[k]; } }
function destroyAll(){ Object.keys(charts).forEach(destroy); }

// --- KPIs -------------------------------------------------------------------
function buildKPIs(rows){
  const N=rows.length;
  const by=s=>rows.filter(r=>r.status===s).length;
  const C=by('Completed'), P=by('In Progress'), H=by('On Hold'),
        NS=by('Not Started'), X=by('Cancelled');
  const MD=rows.filter(r=>r.material_status==='Delivered').length;
  const avgP=avg(rows.map(r=>r.progress));
  document.getElementById('totCount').textContent=N;
  const kpis=[
    {l:'Total Projects', v:N,             d:'Showing '+N+' of '+ALL.length, ac:'#3b82f6', qk:''},
    {l:'Completed',      v:C,             d:pct(C,N),  ac:'#22c55e', qk:'Completed'},
    {l:'In Progress',    v:P,             d:pct(P,N),  ac:'#14b8a6', qk:'In Progress'},
    {l:'On Hold',        v:H,             d:pct(H,N),  ac:'#f97316', qk:'On Hold'},
    {l:'Not Started',    v:NS,            d:pct(NS,N), ac:'#64748b', qk:'Not Started'},
    {l:'Cancelled',      v:X,             d:pct(X,N),  ac:'#ef4444', qk:'Cancelled'},
    {l:'Overall Avg %',  v:avgP.toFixed(1)+'%', d:'Across filtered', ac:'#ec4899', qk:''},
    {l:'Mat. Delivered', v:MD,            d:pct(MD,N), ac:'#eab308', qk:''},
  ];
  document.getElementById('kpiRow').innerHTML = kpis.map(k=>
    `<div class="kpi" style="--ac:${k.ac}" onclick="${k.qk?"quickFilter('"+k.qk+"')":""}">
      <div class="lbl">${k.l}</div>
      <div class="val">${k.v}</div>
      <div class="dlt">${k.d}</div>
    </div>`
  ).join('');
}

// --- PHASE CARDS ------------------------------------------------------------
function buildPhase(rows){
  const N=rows.length||1;
  const e_=avg(rows.map(r=>r.eng_pct)), d_=avg(rows.map(r=>r.del_pct)), x_=avg(rows.map(r=>r.exec_pct));
  const ed=rows.filter(r=>r.eng_pct>=100).length;
  const dd=rows.filter(r=>r.del_pct>=100).length;
  const xd=rows.filter(r=>r.exec_pct>=100).length;
  document.getElementById('phaseCards').innerHTML=`
    <div class="pc" style="--ac:#3b82f6">
      <div class="pt">Engineering <span class="pp" style="color:#60a5fa">${e_.toFixed(1)}%</span></div>
      <div class="pb"><div class="pf" style="width:${e_.toFixed(1)}%;background:linear-gradient(90deg,#3b82f6,#60a5fa)"></div></div>
      <div class="ps"><span>Design / Submittal / Drawing / ELS / BOM</span><span><strong>${ed}</strong>&nbsp;/ ${N} done</span></div>
    </div>
    <div class="pc" style="--ac:#a855f7">
      <div class="pt">Delivery <span class="pp" style="color:#c084fc">${d_.toFixed(1)}%</span></div>
      <div class="pb"><div class="pf" style="width:${d_.toFixed(1)}%;background:linear-gradient(90deg,#a855f7,#c084fc)"></div></div>
      <div class="ps"><span>Material delivery &mdash; ${DEL_ITEMS.length} items tracked</span><span><strong>${dd}</strong>&nbsp;/ ${N} done</span></div>
    </div>
    <div class="pc" style="--ac:#f97316">
      <div class="pt">Execution <span class="pp" style="color:#fb923c">${x_.toFixed(1)}%</span></div>
      <div class="pb"><div class="pf" style="width:${x_.toFixed(1)}%;background:linear-gradient(90deg,#f97316,#fb923c)"></div></div>
      <div class="ps"><span>On-site installation &amp; commissioning</span><span><strong>${xd}</strong>&nbsp;/ ${N} done</span></div>
    </div>`;
}

// --- CHARTS -----------------------------------------------------------------
const PAL=['#3b82f6','#a855f7','#14b8a6','#ec4899','#f97316','#64748b','#eab308','#22c55e','#ef4444','#0ea5e9'];
const SC={'Completed':'#22c55e','In Progress':'#14b8a6','On Hold':'#f97316','Not Started':'#64748b','Cancelled':'#ef4444'};
const MC={'Delivered':'#22c55e','Partially Delivered':'#eab308','Ordered':'#f97316','Not Ordered':'#ef4444'};
const PC={'High':'#ef4444','Medium':'#eab308','Low':'#22c55e'};

function donut(id,labels,data,colors){
  destroy(id);
  charts[id]=new Chart(document.getElementById(id),{
    type:'doughnut',
    data:{labels,datasets:[{data,backgroundColor:colors,borderWidth:2,borderColor:'#1e293b'}]},
    options:{cutout:'60%',
      plugins:{
        legend:{position:'bottom',labels:{padding:10,font:{size:11},color:'#94a3b8',boxWidth:12}},
        tooltip:{callbacks:{label:ctx=>{
          const t=ctx.dataset.data.reduce((a,b)=>a+b,0);
          return ' '+ctx.label+': '+ctx.raw+' ('+Math.round(ctx.raw/t*100)+'%)';
        }}}
      }
    }
  });
}

function buildCharts(rows){
  // Status donut
  const stO=['Completed','In Progress','On Hold','Not Started','Cancelled'];
  donut('cStatus',stO,stO.map(s=>rows.filter(r=>r.status===s).length),stO.map(s=>SC[s]));

  // Material donut
  const mtO=['Delivered','Partially Delivered','Ordered','Not Ordered'];
  donut('cMaterial',mtO,mtO.map(s=>rows.filter(r=>r.material_status===s).length),mtO.map(s=>MC[s]));

  // Priority donut
  const prO=['High','Medium','Low'];
  donut('cPriority',prO,prO.map(s=>rows.filter(r=>r.priority===s).length),prO.map(s=>PC[s]));

  // Region bar
  destroy('cRegion');
  const rgM={};rows.forEach(r=>{rgM[r.region]=(rgM[r.region]||0)+1;});
  const rgE=Object.entries(rgM).sort((a,b)=>b[1]-a[1]);
  charts['cRegion']=new Chart(document.getElementById('cRegion'),{
    type:'bar',
    data:{labels:rgE.map(x=>x[0]),datasets:[{
      data:rgE.map(x=>x[1]),
      backgroundColor:rgE.map((_,i)=>PAL[i%PAL.length]),
      borderRadius:5,borderWidth:0
    }]},
    options:{plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true,ticks:{stepSize:1,color:'#94a3b8'},grid:{color:'#1e293b'}},
              x:{ticks:{color:'#94a3b8'},grid:{color:'#1e293b'}}}}
  });

  // Customers horizontal bar
  destroy('cCustomer');
  const csM={};rows.forEach(r=>{csM[r.customer]=(csM[r.customer]||0)+1;});
  const top=Object.entries(csM).sort((a,b)=>b[1]-a[1]).slice(0,10).reverse();
  charts['cCustomer']=new Chart(document.getElementById('cCustomer'),{
    type:'bar',
    data:{labels:top.map(x=>x[0]),datasets:[{
      data:top.map(x=>x[1]),backgroundColor:'#3b82f6',borderRadius:4,borderWidth:0
    }]},
    options:{indexAxis:'y',plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true,ticks:{stepSize:1,color:'#94a3b8'},grid:{color:'#1e293b'}},
              y:{ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}
  });

  // Delivery items stacked horizontal
  destroy('cDelivery');
  const nI=DEL_ITEMS.length;
  const done=new Array(nI).fill(0),part=new Array(nI).fill(0),
        ndone=new Array(nI).fill(0),naArr=new Array(nI).fill(0);
  rows.forEach(r=>{
    (r.del_vals||[]).forEach((v,i)=>{
      if(i>=nI)return;
      const u=(v||'').toUpperCase();
      if(u==='DONE')done[i]++;
      else if(u==='PART.DONE')part[i]++;
      else if(u==='N.DONE')ndone[i]++;
      else if(u==='N/A')naArr[i]++;
    });
  });
  charts['cDelivery']=new Chart(document.getElementById('cDelivery'),{
    type:'bar',
    data:{labels:DEL_ITEMS,datasets:[
      {label:'Done',     data:done, backgroundColor:'#22c55e',stack:'s',borderWidth:0},
      {label:'Partial',  data:part, backgroundColor:'#eab308',stack:'s',borderWidth:0},
      {label:'Not Done', data:ndone,backgroundColor:'#ef4444',stack:'s',borderWidth:0},
      {label:'N/A',      data:naArr,backgroundColor:'#475569',stack:'s',borderWidth:0},
    ]},
    options:{indexAxis:'y',
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}}},
      scales:{
        x:{beginAtZero:true,stacked:true,ticks:{color:'#94a3b8'},grid:{color:'#1e293b'}},
        y:{stacked:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}}
      }}
  });

  // Progress buckets
  destroy('cBuckets');
  const bk={'0-25%':0,'26-50%':0,'51-75%':0,'76-99%':0,'100%':0};
  rows.forEach(r=>{
    const p=r.progress;
    if(p<=25)bk['0-25%']++;
    else if(p<=50)bk['26-50%']++;
    else if(p<=75)bk['51-75%']++;
    else if(p<100)bk['76-99%']++;
    else bk['100%']++;
  });
  charts['cBuckets']=new Chart(document.getElementById('cBuckets'),{
    type:'bar',
    data:{labels:Object.keys(bk),datasets:[{
      data:Object.values(bk),
      backgroundColor:['#ef4444','#f97316','#eab308','#3b82f6','#22c55e'],
      borderRadius:6,borderWidth:0
    }]},
    options:{plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true,ticks:{color:'#94a3b8'},grid:{color:'#1e293b'}},
              x:{ticks:{color:'#94a3b8'},grid:{color:'#1e293b'}}}}
  });

  // Phase by project
  destroy('cPhase');
  const sorted=[...rows].sort((a,b)=>b.progress-a.progress).slice(0,25);
  const ylbls=sorted.map(r=>'#'+r.s_no+' '+(r.project_name||r.customer).substring(0,32));
  charts['cPhase']=new Chart(document.getElementById('cPhase'),{
    type:'bar',
    data:{labels:ylbls,datasets:[
      {label:'Engineering %',data:sorted.map(r=>r.eng_pct), backgroundColor:'#3b82f6',borderRadius:3,borderWidth:0},
      {label:'Delivery %',   data:sorted.map(r=>r.del_pct), backgroundColor:'#a855f7',borderRadius:3,borderWidth:0},
      {label:'Execution %',  data:sorted.map(r=>r.exec_pct),backgroundColor:'#f97316',borderRadius:3,borderWidth:0},
    ]},
    options:{indexAxis:'y',
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}},
        tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+ctx.raw+'%'}}},
      scales:{
        x:{beginAtZero:true,max:100,title:{display:true,text:'Percent Complete',color:'#94a3b8'},
           ticks:{color:'#94a3b8'},grid:{color:'#1e293b'}},
        y:{ticks:{color:'#94a3b8',font:{size:11}},grid:{color:'#263248'}}
      }}
  });
}

// --- TABLE -------------------------------------------------------------------
function mbar(v,cls){
  const p=Math.max(0,Math.min(100,+v||0));
  return `<div class="bc"><div class="bb"><div class="bf ${cls}" style="width:${p}%"></div></div><div class="bpct">${p}%</div></div>`;
}
const SB={'Completed':'sc','In Progress':'sp','On Hold':'sh','Cancelled':'sx','Not Started':'sn'};
const MB={'Delivered':'md','Partially Delivered':'mp','Ordered':'mo','Not Ordered':'mn'};
const PB={'High':'ph','Medium':'pm','Low':'pl'};

function buildTable(rows){
  document.getElementById('rowCount').textContent=rows.length+' rows';
  document.getElementById('tBody').innerHTML=rows.map(r=>`
    <tr>
      <td>${e(r.s_no)}</td>
      <td>${e(r.customer)}</td>
      <td style="max-width:260px;white-space:normal;word-break:break-word">${e(r.project_name)}</td>
      <td>${e(r.region)}</td>
      <td><span class="badge ${SB[r.status]||'sn'}">${e(r.status)}</span></td>
      <td>${mbar(r.eng_pct,'beng')}</td>
      <td>${mbar(r.del_pct,'bdel')}</td>
      <td>${mbar(r.exec_pct,'bexc')}</td>
      <td>${mbar(r.progress,'ball')}</td>
      <td><span class="badge ${MB[r.material_status]||'mn'}">${e(r.material_status)}</span></td>
      <td><span class="badge ${PB[r.priority]||'pm'}">${e(r.priority)}</span></td>
      <td style="max-width:200px;white-space:normal;word-break:break-word;color:#64748b;font-size:11px">${e(r.work_status)}</td>
    </tr>`).join('');
}

// --- FILTERS ----------------------------------------------------------------
function populateDropdowns(){
  function fill(id,opts){
    const s=document.getElementById(id);
    const cur=s.value;
    while(s.options.length>1)s.remove(1);
    opts.forEach(v=>{const o=new Option(v,v);s.add(o);});
    if(cur)s.value=cur;
  }
  fill('fSt',['Not Started','In Progress','Completed','On Hold','Cancelled']);
  fill('fRg',uniq('region'));
  fill('fCu',uniq('customer'));
  fill('fMt',['Delivered','Partially Delivered','Ordered','Not Ordered']);
  fill('fPr',['High','Medium','Low']);
}

function getFiltered(){
  const base = activeStatus ? ALL.filter(r=>r.status===activeStatus) : ALL;
  const q=(document.getElementById('srch').value||'').toLowerCase().trim();
  const fs=document.getElementById('fSt').value;
  const fr=document.getElementById('fRg').value;
  const fc=document.getElementById('fCu').value;
  const fm=document.getElementById('fMt').value;
  const fp=document.getElementById('fPr').value;
  return base.filter(r=>{
    if(fs&&r.status!==fs)return false;
    if(fr&&r.region!==fr)return false;
    if(fc&&r.customer!==fc)return false;
    if(fm&&r.material_status!==fm)return false;
    if(fp&&r.priority!==fp)return false;
    if(q){
      const hay=[r.project_name,r.customer,r.location,r.job_ref,r.lpo_ref,r.work_status]
                .join(' ').toLowerCase();
      if(!hay.includes(q))return false;
    }
    return true;
  });
}

function redraw(){
  const rows=getFiltered();
  buildKPIs(rows);
  buildPhase(rows);
  buildCharts(rows);
  buildTable(rows);
}

function quickFilter(status){
  activeStatus=status;
  // highlight active button
  ['qAll','qComp','qProg','qHold','qNS','qCan'].forEach(id=>{
    document.getElementById(id).classList.remove('active');
  });
  const map={'':'qAll','Completed':'qComp','In Progress':'qProg',
             'On Hold':'qHold','Not Started':'qNS','Cancelled':'qCan'};
  if(map[status]) document.getElementById(map[status]).classList.add('active');
  redraw();
}

function clearFilters(){
  activeStatus='';
  ['srch','fSt','fRg','fCu','fMt','fPr'].forEach(id=>{
    const el=document.getElementById(id);
    if(el.tagName==='INPUT')el.value=''; else el.value='';
  });
  ['qAll','qComp','qProg','qHold','qNS','qCan'].forEach(id=>
    document.getElementById(id).classList.remove('active'));
  redraw();
}

['srch','fSt','fRg','fCu','fMt','fPr'].forEach(id=>{
  document.getElementById(id).addEventListener('input',redraw);
  document.getElementById(id).addEventListener('change',redraw);
});

populateDropdowns();
quickFilter('');   // show all, highlight All button
</script>
</body>
</html>"""

HTML = HTML.replace("DATA_PLACEHOLDER",   data_js)
HTML = HTML.replace("DEL_PLACEHOLDER",    del_names_js)
HTML = HTML.replace("TIMESTAMP",          now)
HTML = HTML.replace("SOURCENAME",         src)

components.html(HTML, height=5200, scrolling=True)
