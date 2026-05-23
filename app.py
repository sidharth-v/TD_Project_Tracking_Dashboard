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

st.set_page_config(page_title="Project Tracking Dashboard",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""<style>
#MainMenu,footer,header,[data-testid="stToolbar"],
[data-testid="stDecoration"],[data-testid="collapsedControl"]{display:none!important;}
.block-container{padding:0!important;margin:0!important;max-width:100%!important;}
section[data-testid="stSidebar"]{display:none!important;}
</style>""", unsafe_allow_html=True)

LOCAL_FILE  = "Project_Tracking_v7.xlsx"
SHEET_NAME  = "Project_Master"

def _secret(k, d=""):
    try:    return str(st.secrets.get(k, d)).strip()
    except: return d

try:    REFRESH = int(_secret("REFRESH_SECONDS","60") or 60)
except: REFRESH = 60

# Exact column indices from Project_Tracking_v7.xlsx
C_SNO=0; C_JOB=2; C_LPO=3; C_CUST=4; C_PROJ=5
C_REGION=6; C_LOC=7; C_WSTAT=25; C_MATSTAT=27
C_PROG=28; C_PRI=29; C_STATUS=30
C_ENGPCT=31; C_DELPCT=32; C_EXCPCT=33
C_ENG=list(range(9,14))
C_DEL=list(range(14,25))

DEL_NAMES=["Out_Door","Indoor","CR Panels","CR Ins. Materials","Doors",
           "Ref. Inst. Materials","CCP","Display CCP","Floor Heater","Cabinets","Any Special"]

def _dl_url(url):
    url=url.strip()
    if not url or "download=1" in url: return url
    return url+("&" if "?" in url else "?")+"download=1"

@st.cache_data(ttl=REFRESH)
def load_onedrive(url):
    r=requests.get(_dl_url(url),timeout=60,allow_redirects=True); r.raise_for_status()
    if "text/html" in r.headers.get("content-type","").lower() and len(r.content)<500_000:
        raise RuntimeError("OneDrive link returned HTML. Check your share link.")
    return r.content

@st.cache_data(ttl=REFRESH)
def load_graph():
    if msal is None: raise RuntimeError("msal not installed")
    tid=_secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    cid=_secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    cs=_secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")
    app=msal.ConfidentialClientApplication(cid,
        authority=f"https://login.microsoftonline.com/{tid}",client_credential=cs)
    res=app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res: raise RuntimeError(str(res))
    uid=_secret("GRAPH_USER_ID"); fp=_secret("ONEDRIVE_FILE_PATH")
    if not fp.startswith("/"): fp="/"+fp
    r=requests.get(f"https://graph.microsoft.com/v1.0/users/{uid}/drive/root:{fp}:/content",
        headers={"Authorization":f"Bearer {res['access_token']}"},timeout=60)
    r.raise_for_status(); return r.content

def get_file():
    eu=_secret("EXCEL_FILE_URL")
    if eu: return load_onedrive(eu),"OneDrive"
    if _secret("GRAPH_USER_ID"): return load_graph(),"Graph OneDrive"
    lp=Path(LOCAL_FILE)
    if lp.exists(): return lp.read_bytes(),LOCAL_FILE
    st.error("No data source. Set EXCEL_FILE_URL in Streamlit secrets."); st.stop()

def _v(cell,d=""):
    if cell is None or (isinstance(cell,float) and pd.isna(cell)): return d
    s=str(cell).strip()
    return s if s and s.lower()!="nan" else d

def _n(cell):
    try: return round(float(cell),1)
    except: return 0.0

@st.cache_data(ttl=REFRESH)
def parse(file_bytes:bytes):
    raw=pd.read_excel(BytesIO(file_bytes),sheet_name=SHEET_NAME,header=None,engine="openpyxl")
    rows=[]
    for idx in range(2,len(raw)):
        r=raw.iloc[idx]
        sno=r[C_SNO]; cust=r[C_CUST]; proj=r[C_PROJ]
        if (pd.isna(sno) or str(sno).strip()=="") and (pd.isna(cust) or str(cust).strip()==""):
            continue
        try:   sno_str=str(int(float(sno)))
        except: sno_str=_v(sno)
        del_vals=[_v(r[c]) for c in C_DEL]
        rows.append({
            "s_no":         sno_str,
            "job_ref":      _v(r[C_JOB],"N/A"),
            "lpo_ref":      _v(r[C_LPO],"N/A"),
            "customer":     _v(r[C_CUST],"Unknown"),
            "project_name": _v(r[C_PROJ]),
            "region":       _v(r[C_REGION],"Unknown"),
            "location":     _v(r[C_LOC]),
            "work_status":  _v(r[C_WSTAT]),
            "material_status": _v(r[C_MATSTAT],"Not Ordered"),
            "progress":     _n(r[C_PROG]),
            "priority":     _v(r[C_PRI],"Medium") or "Medium",
            "status":       _v(r[C_STATUS],"Not Started"),
            "eng_pct":      _n(r[C_ENGPCT]),
            "del_pct":      _n(r[C_DELPCT]),
            "exec_pct":     _n(r[C_EXCPCT]),
            "del_vals":     del_vals,
        })
    return rows

st_autorefresh(interval=REFRESH*1000,key="ar")

try:
    fb,src=get_file()
    projects=parse(fb)
except Exception as ex:
    st.error(f"Failed to load: {ex}"); st.stop()

now=pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")

# Build complete HTML with data baked in as JS variables
# Use JSON then assign to window globals to avoid any injection issues
data_js     = json.dumps(projects,   ensure_ascii=True)
del_js      = json.dumps(DEL_NAMES,  ensure_ascii=True)
now_js      = json.dumps(now,        ensure_ascii=True)
src_js      = json.dumps(src,        ensure_ascii=True)

# Estimated height: header~120 + kpis~120 + phase~160 + 3donuts~380
# + 2bars~360 + delivery+buckets~480 + phase_chart~520 + table~700 = ~2840 + padding
DASH_HEIGHT = 3200

HTML = (
"""<!DOCTYPE html><html><head><meta charset="UTF-8">"""
"""<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>"""
"""<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;
  background:linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%);
  color:#f1f5f9;min-height:100vh;padding:18px 22px 40px;}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;
  margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #334155;flex-wrap:wrap;gap:10px;}
.hdr h1{font-size:22px;font-weight:700;color:#f1f5f9;}
.hdr .sub{color:#94a3b8;font-size:12px;margin-top:3px;}
.hdr .meta{text-align:right;color:#94a3b8;font-size:11px;line-height:2;}
.qbar{display:flex;gap:7px;margin-bottom:16px;flex-wrap:wrap;align-items:center;}
.ql{color:#94a3b8;font-size:11px;text-transform:uppercase;letter-spacing:.5px;margin-right:2px;}
.btn{padding:7px 13px;border-radius:7px;border:none;font-size:12px;font-weight:500;
     cursor:pointer;color:#fff;transition:opacity .15s,outline .1s;}
.btn:hover{opacity:.82;}
.btn.active{outline:2px solid #ffffff;outline-offset:2px;}
.btn-all{background:#475569;} .btn-comp{background:#22c55e;}
.btn-prog{background:#3b82f6;} .btn-hold{background:#f97316;}
.btn-ns{background:#64748b;} .btn-can{background:#ef4444;}
.kpis{display:grid;grid-template-columns:repeat(8,1fr);gap:10px;margin-bottom:18px;}
.kpi{background:#1e293b;border-radius:10px;padding:14px 12px 12px;
     border-left:4px solid var(--ac,#3b82f6);cursor:pointer;transition:transform .12s;}
.kpi:hover{transform:translateY(-2px);}
.kpi .lbl{color:#94a3b8;font-size:10px;text-transform:uppercase;letter-spacing:.5px;}
.kpi .val{font-size:28px;font-weight:700;margin-top:4px;color:#f1f5f9;}
.kpi .dlt{font-size:10px;color:#64748b;margin-top:2px;}
.stl{font-size:11px;color:#94a3b8;margin:2px 0 10px;
     text-transform:uppercase;letter-spacing:.5px;font-weight:600;}
.pgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px;}
.pc{background:#1e293b;border-radius:10px;padding:16px;border-top:4px solid var(--ac,#3b82f6);}
.pc .pt{font-size:13px;font-weight:600;display:flex;justify-content:space-between;
        align-items:center;margin-bottom:10px;color:#f1f5f9;}
.pc .pp{font-weight:700;font-size:24px;}
.pc .pb{height:8px;background:#334155;border-radius:4px;overflow:hidden;margin-bottom:8px;}
.pc .pf{height:100%;border-radius:4px;}
.pc .ps{display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;}
.pc .ps strong{color:#f1f5f9;font-size:12px;}
.row3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px;}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px;}
.row75{display:grid;grid-template-columns:7fr 5fr;gap:14px;margin-bottom:16px;}
.card{background:#1e293b;border-radius:10px;padding:16px;}
.card h3{font-size:13px;font-weight:600;margin-bottom:10px;color:#f1f5f9;
         display:flex;justify-content:space-between;align-items:center;}
.card h3 .tag{font-size:10px;color:#64748b;font-weight:400;}
.cw{position:relative;height:260px;}
.cw.tall{height:350px;}
.cw.xtall{height:440px;}
.frow{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:8px;align-items:center;}
.frow input,.frow select{
  background:#334155;color:#f1f5f9;border:1px solid #475569;
  padding:6px 9px;border-radius:7px;font-size:12px;outline:none;
  font-family:'Segoe UI',Tahoma,Arial,sans-serif;}
.frow input{min-width:190px;}
.tw{overflow:auto;max-height:520px;border-radius:7px;}
table{width:100%;border-collapse:collapse;font-size:12px;}
thead{position:sticky;top:0;background:#334155;z-index:5;}
th{text-align:left;padding:9px 7px;font-weight:600;
   border-bottom:2px solid #475569;white-space:nowrap;color:#f1f5f9;}
td{padding:7px 7px;border-bottom:1px solid #263248;color:#cbd5e1;}
tr:hover td{background:rgba(59,130,246,.07);}
.badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:11px;font-weight:600;}
.sc{background:rgba(34,197,94,.2);color:#86efac;}
.sp{background:rgba(20,184,166,.2);color:#5eead4;}
.sh{background:rgba(249,115,22,.2);color:#fdba74;}
.sx{background:rgba(239,68,68,.2);color:#fca5a5;}
.sn{background:rgba(100,116,139,.2);color:#cbd5e1;}
.md{background:rgba(34,197,94,.2);color:#86efac;}
.mp{background:rgba(234,179,8,.2);color:#fde047;}
.mo{background:rgba(249,115,22,.2);color:#fdba74;}
.mn{background:rgba(239,68,68,.2);color:#fca5a5;}
.ph{background:rgba(239,68,68,.2);color:#fca5a5;}
.pm{background:rgba(234,179,8,.2);color:#fde047;}
.pl{background:rgba(34,197,94,.2);color:#86efac;}
.bc{display:flex;align-items:center;gap:5px;}
.bb{flex:1;height:5px;background:#334155;border-radius:3px;overflow:hidden;min-width:44px;}
.bf{height:100%;border-radius:3px;}
.beng{background:linear-gradient(90deg,#60a5fa,#3b82f6);}
.bdel{background:linear-gradient(90deg,#c084fc,#a855f7);}
.bexc{background:linear-gradient(90deg,#fb923c,#f97316);}
.ball{background:linear-gradient(90deg,#ef4444,#eab308,#22c55e);}
.bpct{font-size:11px;font-weight:600;min-width:30px;text-align:right;color:#94a3b8;}
.rc{font-size:11px;color:#64748b;margin-bottom:7px;}
</style></head><body>
<div class="hdr">
  <div><h1>Project Tracking Dashboard</h1>
    <div class="sub">Cold Rooms, Cabinets &amp; Refrigeration Projects &mdash; Live from Excel</div></div>
  <div class="meta">
    <div>Updated: <span id="nowTxt"></span></div>
    <div>Source: <span id="srcTxt"></span> &nbsp;|&nbsp; <span id="totTxt">0</span> projects</div>
  </div>
</div>
<div class="qbar">
  <span class="ql">Quick View:</span>
  <button class="btn btn-all  active" id="qAll"  onclick="qf('')">All Projects</button>
  <button class="btn btn-comp"        id="qComp" onclick="qf('Completed')">Completed</button>
  <button class="btn btn-prog"        id="qProg" onclick="qf('In Progress')">In Progress</button>
  <button class="btn btn-hold"        id="qHold" onclick="qf('On Hold')">On Hold</button>
  <button class="btn btn-ns"          id="qNS"   onclick="qf('Not Started')">Not Started</button>
  <button class="btn btn-can"         id="qCan"  onclick="qf('Cancelled')">Cancelled</button>
</div>
<div class="kpis" id="kpiRow"></div>
<div class="stl">Phase Progress Overview</div>
<div class="pgrid" id="phaseRow"></div>
<div class="row3">
  <div class="card"><h3>Project Status</h3><div class="cw"><canvas id="cSt"></canvas></div></div>
  <div class="card"><h3>Material Status</h3><div class="cw"><canvas id="cMt"></canvas></div></div>
  <div class="card"><h3>Priority Distribution</h3><div class="cw"><canvas id="cPr"></canvas></div></div>
</div>
<div class="row2">
  <div class="card"><h3>Projects by Region</h3><div class="cw"><canvas id="cRg"></canvas></div></div>
  <div class="card"><h3>Top 10 Customers</h3><div class="cw"><canvas id="cCu"></canvas></div></div>
</div>
<div class="row75">
  <div class="card">
    <h3>Delivery Items Status <span class="tag">Done / Partial / Not Done / N/A per item</span></h3>
    <div class="cw tall"><canvas id="cDi"></canvas></div></div>
  <div class="card"><h3>Overall Progress Buckets</h3>
    <div class="cw tall"><canvas id="cBk"></canvas></div></div>
</div>
<div class="card" style="margin-bottom:16px">
  <h3>Phase Progress by Project <span class="tag">Eng / Delivery / Execution &mdash; top 25 by overall %</span></h3>
  <div class="cw xtall"><canvas id="cPh"></canvas></div>
</div>
<div class="card">
  <h3>Project Details <span class="tag" id="rcTxt"></span></h3>
  <div class="frow">
    <input id="srch" type="text" placeholder="Search project, customer, location, job ref...">
    <select id="fSt"><option value="">All Statuses</option></select>
    <select id="fRg"><option value="">All Regions</option></select>
    <select id="fCu"><option value="">All Customers</option></select>
    <select id="fMt"><option value="">All Material</option></select>
    <select id="fPr"><option value="">All Priorities</option></select>
    <button class="btn btn-ns" onclick="clearF()">Clear</button>
  </div>
  <div class="tw"><table>
    <thead><tr><th>#</th><th>Customer</th><th>Project</th><th>Region</th>
    <th>Status</th><th>Eng%</th><th>Del%</th><th>Exec%</th><th>Overall%</th>
    <th>Material</th><th>Priority</th><th>Work Status</th></tr></thead>
    <tbody id="tBody"></tbody>
  </table></div>
</div>
<script>
Chart.defaults.color='#94a3b8';
Chart.defaults.borderColor='#334155';
Chart.defaults.font.family="'Segoe UI',Tahoma,Arial,sans-serif";

// Data injected by Python
var ALL="""
+ data_js
+ """;
var DEL="""
+ del_js
+ """;
document.getElementById('nowTxt').textContent="""
+ now_js
+ """;
document.getElementById('srcTxt').textContent="""
+ src_js
+ """;

var charts={}, activeQ='';

function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function avg(a){return a.length?a.reduce((s,x)=>s+x,0)/a.length:0;}
function uniq(k){return[...new Set(ALL.map(r=>r[k]).filter(Boolean))].sort();}
function pp(n,t){return t?Math.round(n/t*100)+'%':'';}
function dc(k){if(charts[k]){charts[k].destroy();delete charts[k];}}
function dcAll(){Object.keys(charts).forEach(dc);}

function buildKPIs(rows){
  var N=rows.length;
  var C=rows.filter(r=>r.status==='Completed').length;
  var P=rows.filter(r=>r.status==='In Progress').length;
  var H=rows.filter(r=>r.status==='On Hold').length;
  var NS=rows.filter(r=>r.status==='Not Started').length;
  var X=rows.filter(r=>r.status==='Cancelled').length;
  var MD=rows.filter(r=>r.material_status==='Delivered').length;
  var ap=avg(rows.map(r=>r.progress)).toFixed(1);
  document.getElementById('totTxt').textContent=N;
  var defs=[
    {l:'Total Projects',v:N,         d:'Showing '+N+' of '+ALL.length, ac:'#3b82f6',q:''},
    {l:'Completed',     v:C,         d:pp(C,N),  ac:'#22c55e', q:'Completed'},
    {l:'In Progress',   v:P,         d:pp(P,N),  ac:'#14b8a6', q:'In Progress'},
    {l:'On Hold',       v:H,         d:pp(H,N),  ac:'#f97316', q:'On Hold'},
    {l:'Not Started',   v:NS,        d:pp(NS,N), ac:'#64748b', q:'Not Started'},
    {l:'Cancelled',     v:X,         d:pp(X,N),  ac:'#ef4444', q:'Cancelled'},
    {l:'Overall Avg',   v:ap+'%',    d:'Across filtered', ac:'#ec4899', q:''},
    {l:'Mat. Delivered',v:MD,        d:pp(MD,N), ac:'#eab308', q:''},
  ];
  document.getElementById('kpiRow').innerHTML=defs.map(k=>
    '<div class="kpi" style="--ac:'+k.ac+'" onclick="'+(k.q?"qf('"+k.q+"')":'')+'">'
    +'<div class="lbl">'+k.l+'</div>'
    +'<div class="val">'+k.v+'</div>'
    +'<div class="dlt">'+k.d+'</div></div>'
  ).join('');
}

function buildPhase(rows){
  var N=rows.length||1;
  var e=avg(rows.map(r=>r.eng_pct)),
      d=avg(rows.map(r=>r.del_pct)),
      x=avg(rows.map(r=>r.exec_pct));
  var ed=rows.filter(r=>r.eng_pct>=100).length;
  var dd=rows.filter(r=>r.del_pct>=100).length;
  var xd=rows.filter(r=>r.exec_pct>=100).length;
  document.getElementById('phaseRow').innerHTML=
    '<div class="pc" style="--ac:#3b82f6">'
    +'<div class="pt">Engineering <span class="pp" style="color:#60a5fa">'+e.toFixed(1)+'%</span></div>'
    +'<div class="pb"><div class="pf" style="width:'+e.toFixed(1)+'%;background:linear-gradient(90deg,#3b82f6,#60a5fa)"></div></div>'
    +'<div class="ps"><span>Design / Submittal / Drawing / ELS / BOM</span><span><strong>'+ed+'</strong> / '+N+' done</span></div></div>'
    +'<div class="pc" style="--ac:#a855f7">'
    +'<div class="pt">Delivery <span class="pp" style="color:#c084fc">'+d.toFixed(1)+'%</span></div>'
    +'<div class="pb"><div class="pf" style="width:'+d.toFixed(1)+'%;background:linear-gradient(90deg,#a855f7,#c084fc)"></div></div>'
    +'<div class="ps"><span>Material delivery &mdash; '+DEL.length+' items tracked</span><span><strong>'+dd+'</strong> / '+N+' done</span></div></div>'
    +'<div class="pc" style="--ac:#f97316">'
    +'<div class="pt">Execution <span class="pp" style="color:#fb923c">'+x.toFixed(1)+'%</span></div>'
    +'<div class="pb"><div class="pf" style="width:'+x.toFixed(1)+'%;background:linear-gradient(90deg,#f97316,#fb923c)"></div></div>'
    +'<div class="ps"><span>On-site installation &amp; commissioning</span><span><strong>'+xd+'</strong> / '+N+' done</span></div></div>';
}

var PAL=['#3b82f6','#a855f7','#14b8a6','#ec4899','#f97316','#64748b','#eab308','#22c55e','#ef4444','#0ea5e9'];
var SC={'Completed':'#22c55e','In Progress':'#14b8a6','On Hold':'#f97316','Not Started':'#64748b','Cancelled':'#ef4444'};
var MC={'Delivered':'#22c55e','Partially Delivered':'#eab308','Ordered':'#f97316','Not Ordered':'#ef4444'};
var PC={'High':'#ef4444','Medium':'#eab308','Low':'#22c55e'};

function donut(id,labels,data,colors){
  dc(id);
  var total=data.reduce(function(a,b){return a+b;},0);
  charts[id]=new Chart(document.getElementById(id),{
    type:'doughnut',
    data:{labels:labels,datasets:[{data:data,backgroundColor:colors,borderWidth:2,borderColor:'#1e293b'}]},
    options:{cutout:'60%',animation:{duration:500},
      plugins:{
        legend:{position:'bottom',labels:{padding:8,font:{size:11},color:'#94a3b8',boxWidth:12}},
        tooltip:{callbacks:{label:function(ctx){
          var t=ctx.dataset.data.reduce(function(a,b){return a+b;},0);
          return ' '+ctx.label+': '+ctx.raw+' ('+Math.round(ctx.raw/t*100)+'%)';
        }}}
      }
    }
  });
}

function buildCharts(rows){
  var stO=['Completed','In Progress','On Hold','Not Started','Cancelled'];
  donut('cSt',stO,stO.map(function(s){return rows.filter(function(r){return r.status===s;}).length;}),stO.map(function(s){return SC[s];}));

  var mtO=['Delivered','Partially Delivered','Ordered','Not Ordered'];
  donut('cMt',mtO,mtO.map(function(s){return rows.filter(function(r){return r.material_status===s;}).length;}),mtO.map(function(s){return MC[s];}));

  var prO=['High','Medium','Low'];
  donut('cPr',prO,prO.map(function(s){return rows.filter(function(r){return r.priority===s;}).length;}),prO.map(function(s){return PC[s];}));

  // Region
  dc('cRg');
  var rgM={};rows.forEach(function(r){rgM[r.region]=(rgM[r.region]||0)+1;});
  var rgE=Object.entries(rgM).sort(function(a,b){return b[1]-a[1];});
  charts['cRg']=new Chart(document.getElementById('cRg'),{type:'bar',
    data:{labels:rgE.map(function(x){return x[0];}),datasets:[{
      data:rgE.map(function(x){return x[1];}),
      backgroundColor:rgE.map(function(_,i){return PAL[i%PAL.length];}),
      borderRadius:5,borderWidth:0}]},
    options:{animation:{duration:500},plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},
              x:{ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});

  // Customers
  dc('cCu');
  var csM={};rows.forEach(function(r){csM[r.customer]=(csM[r.customer]||0)+1;});
  var top=Object.entries(csM).sort(function(a,b){return b[1]-a[1];}).slice(0,10).reverse();
  charts['cCu']=new Chart(document.getElementById('cCu'),{type:'bar',
    data:{labels:top.map(function(x){return x[0];}),datasets:[{
      data:top.map(function(x){return x[1];}),backgroundColor:'#3b82f6',borderRadius:4,borderWidth:0}]},
    options:{indexAxis:'y',animation:{duration:500},plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},
              y:{ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});

  // Delivery items stacked horizontal
  dc('cDi');
  var nI=DEL.length;
  var done=new Array(nI).fill(0),part=new Array(nI).fill(0),
      ndone=new Array(nI).fill(0),naA=new Array(nI).fill(0);
  rows.forEach(function(r){
    (r.del_vals||[]).forEach(function(v,i){
      if(i>=nI)return;
      var u=(v||'').toUpperCase();
      if(u==='DONE')done[i]++;
      else if(u==='PART.DONE')part[i]++;
      else if(u==='N.DONE')ndone[i]++;
      else if(u==='N/A')naA[i]++;
    });
  });
  charts['cDi']=new Chart(document.getElementById('cDi'),{type:'bar',
    data:{labels:DEL,datasets:[
      {label:'Done',    data:done, backgroundColor:'#22c55e',stack:'s',borderWidth:0},
      {label:'Partial', data:part, backgroundColor:'#eab308',stack:'s',borderWidth:0},
      {label:'Not Done',data:ndone,backgroundColor:'#ef4444',stack:'s',borderWidth:0},
      {label:'N/A',     data:naA,  backgroundColor:'#475569',stack:'s',borderWidth:0}]},
    options:{indexAxis:'y',animation:{duration:500},
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}}},
      scales:{x:{beginAtZero:true,stacked:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},
              y:{stacked:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});

  // Progress buckets
  dc('cBk');
  var bk={'0-25%':0,'26-50%':0,'51-75%':0,'76-99%':0,'100%':0};
  rows.forEach(function(r){
    var p=r.progress;
    if(p<=25)bk['0-25%']++;
    else if(p<=50)bk['26-50%']++;
    else if(p<=75)bk['51-75%']++;
    else if(p<100)bk['76-99%']++;
    else bk['100%']++;
  });
  charts['cBk']=new Chart(document.getElementById('cBk'),{type:'bar',
    data:{labels:Object.keys(bk),datasets:[{
      data:Object.values(bk),
      backgroundColor:['#ef4444','#f97316','#eab308','#3b82f6','#22c55e'],
      borderRadius:6,borderWidth:0}]},
    options:{animation:{duration:500},plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},
              x:{ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});

  // Phase by project
  dc('cPh');
  var sorted=rows.slice().sort(function(a,b){return b.progress-a.progress;}).slice(0,25);
  var ylbls=sorted.map(function(r){return '#'+r.s_no+' '+(r.project_name||r.customer).substring(0,30);});
  charts['cPh']=new Chart(document.getElementById('cPh'),{type:'bar',
    data:{labels:ylbls,datasets:[
      {label:'Engineering %',data:sorted.map(function(r){return r.eng_pct;}),backgroundColor:'#3b82f6',borderRadius:3,borderWidth:0},
      {label:'Delivery %',   data:sorted.map(function(r){return r.del_pct;}),backgroundColor:'#a855f7',borderRadius:3,borderWidth:0},
      {label:'Execution %',  data:sorted.map(function(r){return r.exec_pct;}),backgroundColor:'#f97316',borderRadius:3,borderWidth:0}]},
    options:{indexAxis:'y',animation:{duration:500},
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}},
        tooltip:{callbacks:{label:function(ctx){return ctx.dataset.label+': '+ctx.raw+'%';}}}},
      scales:{x:{beginAtZero:true,max:100,title:{display:true,text:'Percent Complete',color:'#94a3b8'},
                 ticks:{color:'#94a3b8'},grid:{color:'#263248'}},
              y:{ticks:{color:'#94a3b8',font:{size:11}},grid:{color:'#263248'}}}}});
}

function mbar(v,cls){
  var p=Math.max(0,Math.min(100,+v||0));
  return '<div class="bc"><div class="bb"><div class="bf '+cls+'" style="width:'+p+'%"></div></div><div class="bpct">'+p+'%</div></div>';
}
var SB={'Completed':'sc','In Progress':'sp','On Hold':'sh','Cancelled':'sx','Not Started':'sn'};
var MB={'Delivered':'md','Partially Delivered':'mp','Ordered':'mo','Not Ordered':'mn'};
var PB={'High':'ph','Medium':'pm','Low':'pl'};

function buildTable(rows){
  document.getElementById('rcTxt').textContent=rows.length+' rows';
  document.getElementById('tBody').innerHTML=rows.map(function(r){
    return '<tr>'
      +'<td>'+esc(r.s_no)+'</td>'
      +'<td>'+esc(r.customer)+'</td>'
      +'<td style="max-width:240px;white-space:normal;word-break:break-word">'+esc(r.project_name)+'</td>'
      +'<td>'+esc(r.region)+'</td>'
      +'<td><span class="badge '+(SB[r.status]||'sn')+'">'+esc(r.status)+'</span></td>'
      +'<td>'+mbar(r.eng_pct,'beng')+'</td>'
      +'<td>'+mbar(r.del_pct,'bdel')+'</td>'
      +'<td>'+mbar(r.exec_pct,'bexc')+'</td>'
      +'<td>'+mbar(r.progress,'ball')+'</td>'
      +'<td><span class="badge '+(MB[r.material_status]||'mn')+'">'+esc(r.material_status)+'</span></td>'
      +'<td><span class="badge '+(PB[r.priority]||'pm')+'">'+esc(r.priority)+'</span></td>'
      +'<td style="max-width:180px;white-space:normal;word-break:break-word;color:#64748b;font-size:11px">'+esc(r.work_status)+'</td>'
      +'</tr>';
  }).join('');
}

function fillDropdowns(){
  function fill(id,opts){
    var s=document.getElementById(id),cur=s.value;
    while(s.options.length>1)s.remove(1);
    opts.forEach(function(v){s.add(new Option(v,v));});
    if(cur)s.value=cur;
  }
  fill('fSt',['Not Started','In Progress','Completed','On Hold','Cancelled']);
  fill('fRg',uniq('region'));
  fill('fCu',uniq('customer'));
  fill('fMt',['Delivered','Partially Delivered','Ordered','Not Ordered']);
  fill('fPr',['High','Medium','Low']);
}

function filtered(){
  var base=activeQ?ALL.filter(function(r){return r.status===activeQ;}):ALL;
  var q=(document.getElementById('srch').value||'').toLowerCase().trim();
  var fs=document.getElementById('fSt').value;
  var fr=document.getElementById('fRg').value;
  var fc=document.getElementById('fCu').value;
  var fm=document.getElementById('fMt').value;
  var fp=document.getElementById('fPr').value;
  return base.filter(function(r){
    if(fs&&r.status!==fs)return false;
    if(fr&&r.region!==fr)return false;
    if(fc&&r.customer!==fc)return false;
    if(fm&&r.material_status!==fm)return false;
    if(fp&&r.priority!==fp)return false;
    if(q){
      var h=[r.project_name,r.customer,r.location,r.job_ref,r.lpo_ref,r.work_status].join(' ').toLowerCase();
      if(h.indexOf(q)<0)return false;
    }
    return true;
  });
}

function redraw(){
  var rows=filtered();
  buildKPIs(rows);buildPhase(rows);buildCharts(rows);buildTable(rows);
}

function qf(status){
  activeQ=status;
  ['qAll','qComp','qProg','qHold','qNS','qCan'].forEach(function(id){
    document.getElementById(id).classList.remove('active');
  });
  var map={'':'qAll','Completed':'qComp','In Progress':'qProg','On Hold':'qHold','Not Started':'qNS','Cancelled':'qCan'};
  if(map[status])document.getElementById(map[status]).classList.add('active');
  redraw();
}

function clearF(){
  activeQ='';
  document.getElementById('srch').value='';
  ['fSt','fRg','fCu','fMt','fPr'].forEach(function(id){document.getElementById(id).value='';});
  qf('');
}

['srch','fSt','fRg','fCu','fMt','fPr'].forEach(function(id){
  document.getElementById(id).addEventListener('input',redraw);
  document.getElementById(id).addEventListener('change',redraw);
});

fillDropdowns();
qf('');
</script></body></html>"""
)

components.html(HTML, height=DASH_HEIGHT, scrolling=True)
