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

LOCAL_FILE = "Project_Tracking_v7.xlsx"
SHEET_NAME = "Project_Master"
DEL_NAMES  = ["Out_Door","Indoor","CR Panels","CR Ins. Materials","Doors",
               "Ref. Inst. Materials","CCP","Display CCP","Floor Heater","Cabinets","Any Special"]

def _secret(k, d=""):
    try:    return str(st.secrets.get(k, d)).strip()
    except: return d

try:    REFRESH = int(_secret("REFRESH_SECONDS","60") or 60)
except: REFRESH = 60

def _dl_url(url):
    url = url.strip()
    if not url or "download=1" in url: return url
    return url + ("&" if "?" in url else "?") + "download=1"

@st.cache_data(ttl=REFRESH)
def load_onedrive(url):
    r = requests.get(_dl_url(url), timeout=60, allow_redirects=True)
    r.raise_for_status()
    if "text/html" in r.headers.get("content-type","").lower() and len(r.content)<500_000:
        raise RuntimeError("OneDrive returned HTML not Excel. Check share link.")
    return r.content

@st.cache_data(ttl=REFRESH)
def load_graph():
    if msal is None: raise RuntimeError("msal not installed")
    tid = _secret("GRAPH_TENANT_ID") or _secret("TENANT_ID")
    cid = _secret("GRAPH_CLIENT_ID") or _secret("CLIENT_ID")
    cs  = _secret("GRAPH_CLIENT_SECRET") or _secret("CLIENT_SECRET")
    app = msal.ConfidentialClientApplication(
        cid, authority="https://login.microsoftonline.com/"+tid, client_credential=cs)
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res: raise RuntimeError(str(res))
    uid = _secret("GRAPH_USER_ID")
    fp  = _secret("ONEDRIVE_FILE_PATH")
    if not fp.startswith("/"): fp = "/" + fp
    r = requests.get("https://graph.microsoft.com/v1.0/users/"+uid+"/drive/root:"+fp+":/content",
        headers={"Authorization": "Bearer "+res["access_token"]}, timeout=60)
    r.raise_for_status()
    return r.content

def get_file():
    eu = _secret("EXCEL_FILE_URL")
    if eu: return load_onedrive(eu), "OneDrive"
    if _secret("GRAPH_USER_ID"): return load_graph(), "Graph OneDrive"
    lp = Path(LOCAL_FILE)
    if lp.exists(): return lp.read_bytes(), LOCAL_FILE
    st.error("No data source. Set EXCEL_FILE_URL in Streamlit secrets.")
    st.stop()

def _v(cell, d=""):
    if cell is None or (isinstance(cell, float) and pd.isna(cell)): return d
    s = str(cell).strip()
    return s if s and s.lower() != "nan" else d

@st.cache_data(ttl=REFRESH)
def parse(fb):
    raw = pd.read_excel(BytesIO(fb), sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    rows = []
    for i in range(2, len(raw)):
        r = raw.iloc[i]
        if (pd.isna(r[0]) or str(r[0]).strip()=="") and (pd.isna(r[4]) or str(r[4]).strip()==""): continue
        try:    sno = str(int(float(r[0])))
        except: sno = _v(r[0])
        rows.append({
            "s_no":            sno,
            "job_ref":         _v(r[2], "N/A"),
            "lpo_ref":         _v(r[3], "N/A"),
            "customer":        _v(r[4], "Unknown"),
            "project_name":    _v(r[5]),
            "region":          _v(r[6], "Unknown"),
            "location":        _v(r[7]),
            "work_status":     _v(r[25]),
            "material_status": _v(r[27], "Not Ordered"),
            "progress":        round(float(r[28]), 1) if not pd.isna(r[28]) else 0.0,
            "priority":        _v(r[29], "Medium") or "Medium",
            "status":          _v(r[30], "Not Started"),
            "eng_pct":         round(float(r[31]), 1) if not pd.isna(r[31]) else 0.0,
            "del_pct":         round(float(r[32]), 1) if not pd.isna(r[32]) else 0.0,
            "exec_pct":        round(float(r[33]), 1) if not pd.isna(r[33]) else 0.0,
            "del_vals":        [_v(r[c]) for c in range(14, 25)],
        })
    return rows

st_autorefresh(interval=REFRESH*1000, key="ar")

try:
    fb, src = get_file()
    projects = parse(fb)
except Exception as ex:
    st.error("Failed to load: " + str(ex))
    st.stop()

now = pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")

# Build the JS data block as a plain Python string - NO triple quotes
data_block = "var ALL=" + json.dumps(projects, ensure_ascii=True) + ";"
data_block += "var DEL=" + json.dumps(DEL_NAMES, ensure_ascii=True) + ";"
data_block += "var NOW=" + json.dumps(now, ensure_ascii=True) + ";"
data_block += "var SRC=" + json.dumps(src, ensure_ascii=True) + ";"

# Read the HTML template from a separate file
css = """
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;background:linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%);color:#f1f5f9;min-height:100vh;padding:18px 22px 40px;}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #334155;flex-wrap:wrap;gap:10px;}
.hdr h1{font-size:22px;font-weight:700;color:#f1f5f9;}
.hdr .sub{color:#94a3b8;font-size:12px;margin-top:3px;}
.hdr .meta{text-align:right;color:#94a3b8;font-size:11px;line-height:2;}
.qbar{display:flex;gap:7px;margin-bottom:16px;flex-wrap:wrap;align-items:center;}
.ql{color:#94a3b8;font-size:11px;text-transform:uppercase;letter-spacing:.5px;margin-right:2px;}
.btn{padding:7px 13px;border-radius:7px;border:none;font-size:12px;font-weight:500;cursor:pointer;color:#fff;transition:opacity .15s;}
.btn:hover{opacity:.82;}
.btn.active{outline:2px solid #fff;outline-offset:2px;}
.kpis{display:grid;grid-template-columns:repeat(8,1fr);gap:10px;margin-bottom:18px;}
.kpi{background:#1e293b;border-radius:10px;padding:14px 12px 12px;border-left:4px solid var(--ac,#3b82f6);cursor:pointer;transition:transform .12s;}
.kpi:hover{transform:translateY(-2px);}
.kpi .lbl{color:#94a3b8;font-size:10px;text-transform:uppercase;letter-spacing:.5px;}
.kpi .val{font-size:28px;font-weight:700;margin-top:4px;color:#f1f5f9;}
.kpi .dlt{font-size:10px;color:#64748b;margin-top:2px;}
.stl{font-size:11px;color:#94a3b8;margin:2px 0 10px;text-transform:uppercase;letter-spacing:.5px;font-weight:600;}
.pgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px;}
.pc{background:#1e293b;border-radius:10px;padding:16px;border-top:4px solid var(--ac,#3b82f6);}
.pc .pt{font-size:13px;font-weight:600;display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;color:#f1f5f9;}
.pc .pp{font-weight:700;font-size:24px;}
.pc .pb{height:8px;background:#334155;border-radius:4px;overflow:hidden;margin-bottom:8px;}
.pc .pf{height:100%;border-radius:4px;}
.pc .ps{display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;}
.pc .ps strong{color:#f1f5f9;font-size:12px;}
.row3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px;}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px;}
.row75{display:grid;grid-template-columns:7fr 5fr;gap:14px;margin-bottom:16px;}
.card{background:#1e293b;border-radius:10px;padding:16px;}
.card h3{font-size:13px;font-weight:600;margin-bottom:10px;color:#f1f5f9;display:flex;justify-content:space-between;align-items:center;}
.card h3 .tag{font-size:10px;color:#64748b;font-weight:400;}
.cw{position:relative;height:260px;}
.cw.tall{height:350px;}
.cw.xtall{height:440px;}
.frow{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:8px;align-items:center;}
.frow input,.frow select{background:#334155;color:#f1f5f9;border:1px solid #475569;padding:6px 9px;border-radius:7px;font-size:12px;outline:none;font-family:'Segoe UI',Tahoma,Arial,sans-serif;}
.frow input{min-width:190px;}
.tw{overflow:auto;max-height:520px;border-radius:7px;}
table{width:100%;border-collapse:collapse;font-size:12px;}
thead{position:sticky;top:0;background:#334155;z-index:5;}
th{text-align:left;padding:9px 7px;font-weight:600;border-bottom:2px solid #475569;white-space:nowrap;color:#f1f5f9;}
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
"""

body = """
<div class="hdr">
  <div><h1>Project Tracking Dashboard</h1>
    <div class="sub">Cold Rooms, Cabinets &amp; Refrigeration Projects &mdash; Live from Excel</div></div>
  <div class="meta"><div>Updated: <span id="xNow"></span></div>
    <div>Source: <span id="xSrc"></span> &nbsp;|&nbsp; <span id="xTot">0</span> projects</div></div>
</div>
<div class="qbar">
  <span class="ql">Quick:</span>
  <button class="btn active" id="qAll"  style="background:#475569" onclick="qf('')">All</button>
  <button class="btn"        id="qCo"   style="background:#22c55e" onclick="qf('Completed')">Completed</button>
  <button class="btn"        id="qPr"   style="background:#3b82f6" onclick="qf('In Progress')">In Progress</button>
  <button class="btn"        id="qHo"   style="background:#f97316" onclick="qf('On Hold')">On Hold</button>
  <button class="btn"        id="qNS"   style="background:#64748b" onclick="qf('Not Started')">Not Started</button>
  <button class="btn"        id="qCa"   style="background:#ef4444" onclick="qf('Cancelled')">Cancelled</button>
</div>
<div class="kpis" id="kR"></div>
<div class="stl">Phase Progress Overview</div>
<div class="pgrid" id="pR"></div>
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
  <div class="card"><h3>Delivery Items Status <span class="tag">Done / Partial / Not Done / N/A</span></h3>
    <div class="cw tall"><canvas id="cDi"></canvas></div></div>
  <div class="card"><h3>Progress Buckets</h3>
    <div class="cw tall"><canvas id="cBk"></canvas></div></div>
</div>
<div class="card" style="margin-bottom:16px">
  <h3>Phase Progress by Project <span class="tag">Top 25 by overall % &mdash; Eng / Del / Exec</span></h3>
  <div class="cw xtall"><canvas id="cPh"></canvas></div>
</div>
<div class="card">
  <h3>Project Details <span class="tag" id="xRc"></span></h3>
  <div class="frow">
    <input id="xSr" type="text" placeholder="Search project, customer, location, job ref...">
    <select id="fSt"><option value="">All Statuses</option></select>
    <select id="fRg"><option value="">All Regions</option></select>
    <select id="fCu"><option value="">All Customers</option></select>
    <select id="fMt"><option value="">All Material</option></select>
    <select id="fPr"><option value="">All Priorities</option></select>
    <button class="btn" style="background:#475569" onclick="clearF()">Clear</button>
  </div>
  <div class="tw"><table>
    <thead><tr><th>#</th><th>Customer</th><th>Project</th><th>Region</th>
    <th>Status</th><th>Eng%</th><th>Del%</th><th>Exec%</th><th>Overall%</th>
    <th>Material</th><th>Priority</th><th>Work Status</th></tr></thead>
    <tbody id="tB"></tbody>
  </table></div>
</div>
"""

js = """
Chart.defaults.color='#94a3b8';
Chart.defaults.borderColor='#334155';
Chart.defaults.font.family="'Segoe UI',Tahoma,Arial,sans-serif";
document.getElementById('xNow').textContent=NOW;
document.getElementById('xSrc').textContent=SRC;
var charts={},activeQ='';
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function avg(a){return a.length?a.reduce(function(s,x){return s+x;},0)/a.length:0;}
function uniq(k){return[...new Set(ALL.map(function(r){return r[k];}).filter(Boolean))].sort();}
function pp(n,t){return t?Math.round(n/t*100)+'%':'';}
function dc(k){if(charts[k]){charts[k].destroy();delete charts[k];}}
function buildKPIs(rows){
  var N=rows.length;
  var C=rows.filter(function(r){return r.status==='Completed';}).length;
  var P=rows.filter(function(r){return r.status==='In Progress';}).length;
  var H=rows.filter(function(r){return r.status==='On Hold';}).length;
  var NS=rows.filter(function(r){return r.status==='Not Started';}).length;
  var X=rows.filter(function(r){return r.status==='Cancelled';}).length;
  var MD=rows.filter(function(r){return r.material_status==='Delivered';}).length;
  var ap=avg(rows.map(function(r){return r.progress;})).toFixed(1);
  document.getElementById('xTot').textContent=N;
  var d=[
    {l:'Total Projects',v:N,d:'Showing '+N+' of '+ALL.length,ac:'#3b82f6',q:''},
    {l:'Completed',v:C,d:pp(C,N),ac:'#22c55e',q:'Completed'},
    {l:'In Progress',v:P,d:pp(P,N),ac:'#14b8a6',q:'In Progress'},
    {l:'On Hold',v:H,d:pp(H,N),ac:'#f97316',q:'On Hold'},
    {l:'Not Started',v:NS,d:pp(NS,N),ac:'#64748b',q:'Not Started'},
    {l:'Cancelled',v:X,d:pp(X,N),ac:'#ef4444',q:'Cancelled'},
    {l:'Overall Avg',v:ap+'%',d:'Across filtered',ac:'#ec4899',q:''},
    {l:'Mat. Delivered',v:MD,d:pp(MD,N),ac:'#eab308',q:''},
  ];
  document.getElementById('kR').innerHTML=d.map(function(k){
    return '<div class="kpi" style="--ac:'+k.ac+'" onclick="'+(k.q?"qf('"+k.q+"')":'')+'">'
      +'<div class="lbl">'+k.l+'</div><div class="val">'+k.v+'</div><div class="dlt">'+k.d+'</div></div>';
  }).join('');
}
function buildPhase(rows){
  var N=rows.length||1;
  var e=avg(rows.map(function(r){return r.eng_pct;}));
  var d=avg(rows.map(function(r){return r.del_pct;}));
  var x=avg(rows.map(function(r){return r.exec_pct;}));
  var ed=rows.filter(function(r){return r.eng_pct>=100;}).length;
  var dd=rows.filter(function(r){return r.del_pct>=100;}).length;
  var xd=rows.filter(function(r){return r.exec_pct>=100;}).length;
  document.getElementById('pR').innerHTML=
    '<div class="pc" style="--ac:#3b82f6"><div class="pt">Engineering <span class="pp" style="color:#60a5fa">'+e.toFixed(1)+'%</span></div>'
    +'<div class="pb"><div class="pf" style="width:'+e.toFixed(1)+'%;background:linear-gradient(90deg,#3b82f6,#60a5fa)"></div></div>'
    +'<div class="ps"><span>Design / Submittal / Drawing / ELS / BOM</span><span><strong>'+ed+'</strong> / '+N+' done</span></div></div>'
    +'<div class="pc" style="--ac:#a855f7"><div class="pt">Delivery <span class="pp" style="color:#c084fc">'+d.toFixed(1)+'%</span></div>'
    +'<div class="pb"><div class="pf" style="width:'+d.toFixed(1)+'%;background:linear-gradient(90deg,#a855f7,#c084fc)"></div></div>'
    +'<div class="ps"><span>Material delivery &mdash; '+DEL.length+' items</span><span><strong>'+dd+'</strong> / '+N+' done</span></div></div>'
    +'<div class="pc" style="--ac:#f97316"><div class="pt">Execution <span class="pp" style="color:#fb923c">'+x.toFixed(1)+'%</span></div>'
    +'<div class="pb"><div class="pf" style="width:'+x.toFixed(1)+'%;background:linear-gradient(90deg,#f97316,#fb923c)"></div></div>'
    +'<div class="ps"><span>On-site installation &amp; commissioning</span><span><strong>'+xd+'</strong> / '+N+' done</span></div></div>';
}
var PAL=['#3b82f6','#a855f7','#14b8a6','#ec4899','#f97316','#64748b','#eab308','#22c55e','#ef4444'];
var SC={'Completed':'#22c55e','In Progress':'#14b8a6','On Hold':'#f97316','Not Started':'#64748b','Cancelled':'#ef4444'};
var MC={'Delivered':'#22c55e','Partially Delivered':'#eab308','Ordered':'#f97316','Not Ordered':'#ef4444'};
var PC={'High':'#ef4444','Medium':'#eab308','Low':'#22c55e'};
function donut(id,lbls,vals,clrs){
  dc(id);
  charts[id]=new Chart(document.getElementById(id),{type:'doughnut',
    data:{labels:lbls,datasets:[{data:vals,backgroundColor:clrs,borderWidth:2,borderColor:'#1e293b'}]},
    options:{cutout:'60%',animation:{duration:500},
      plugins:{legend:{position:'bottom',labels:{padding:8,font:{size:11},color:'#94a3b8',boxWidth:12}},
        tooltip:{callbacks:{label:function(ctx){var t=ctx.dataset.data.reduce(function(a,b){return a+b;},0);
          return ' '+ctx.label+': '+ctx.raw+' ('+Math.round(ctx.raw/t*100)+'%)';}}}}}}
  );
}
function buildCharts(rows){
  var stO=['Completed','In Progress','On Hold','Not Started','Cancelled'];
  donut('cSt',stO,stO.map(function(s){return rows.filter(function(r){return r.status===s;}).length;}),stO.map(function(s){return SC[s];}));
  var mtO=['Delivered','Partially Delivered','Ordered','Not Ordered'];
  donut('cMt',mtO,mtO.map(function(s){return rows.filter(function(r){return r.material_status===s;}).length;}),mtO.map(function(s){return MC[s];}));
  var prO=['High','Medium','Low'];
  donut('cPr',prO,prO.map(function(s){return rows.filter(function(r){return r.priority===s;}).length;}),prO.map(function(s){return PC[s];}));
  dc('cRg');
  var rgM={};rows.forEach(function(r){rgM[r.region]=(rgM[r.region]||0)+1;});
  var rgE=Object.entries(rgM).sort(function(a,b){return b[1]-a[1];});
  charts['cRg']=new Chart(document.getElementById('cRg'),{type:'bar',
    data:{labels:rgE.map(function(x){return x[0];}),datasets:[{data:rgE.map(function(x){return x[1];}),
      backgroundColor:rgE.map(function(_,i){return PAL[i%PAL.length];}),borderRadius:5,borderWidth:0}]},
    options:{animation:{duration:500},plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},x:{ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});
  dc('cCu');
  var csM={};rows.forEach(function(r){csM[r.customer]=(csM[r.customer]||0)+1;});
  var top=Object.entries(csM).sort(function(a,b){return b[1]-a[1];}).slice(0,10).reverse();
  charts['cCu']=new Chart(document.getElementById('cCu'),{type:'bar',
    data:{labels:top.map(function(x){return x[0];}),datasets:[{data:top.map(function(x){return x[1];}),
      backgroundColor:'#3b82f6',borderRadius:4,borderWidth:0}]},
    options:{indexAxis:'y',animation:{duration:500},plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},y:{ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});
  dc('cDi');
  var nI=DEL.length,done=new Array(nI).fill(0),part=new Array(nI).fill(0),ndone=new Array(nI).fill(0),naA=new Array(nI).fill(0);
  rows.forEach(function(r){(r.del_vals||[]).forEach(function(v,i){
    if(i>=nI)return;var u=(v||'').toUpperCase();
    if(u==='DONE')done[i]++;else if(u==='PART.DONE')part[i]++;else if(u==='N.DONE')ndone[i]++;else if(u==='N/A')naA[i]++;
  });});
  charts['cDi']=new Chart(document.getElementById('cDi'),{type:'bar',
    data:{labels:DEL,datasets:[
      {label:'Done',data:done,backgroundColor:'#22c55e',stack:'s',borderWidth:0},
      {label:'Partial',data:part,backgroundColor:'#eab308',stack:'s',borderWidth:0},
      {label:'Not Done',data:ndone,backgroundColor:'#ef4444',stack:'s',borderWidth:0},
      {label:'N/A',data:naA,backgroundColor:'#475569',stack:'s',borderWidth:0}]},
    options:{indexAxis:'y',animation:{duration:500},
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}}},
      scales:{x:{beginAtZero:true,stacked:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},
              y:{stacked:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});
  dc('cBk');
  var bk={'0-25%':0,'26-50%':0,'51-75%':0,'76-99%':0,'100%':0};
  rows.forEach(function(r){var p=r.progress;
    if(p<=25)bk['0-25%']++;else if(p<=50)bk['26-50%']++;else if(p<=75)bk['51-75%']++;else if(p<100)bk['76-99%']++;else bk['100%']++;});
  charts['cBk']=new Chart(document.getElementById('cBk'),{type:'bar',
    data:{labels:Object.keys(bk),datasets:[{data:Object.values(bk),
      backgroundColor:['#ef4444','#f97316','#eab308','#3b82f6','#22c55e'],borderRadius:6,borderWidth:0}]},
    options:{animation:{duration:500},plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true,ticks:{color:'#94a3b8'},grid:{color:'#263248'}},x:{ticks:{color:'#94a3b8'},grid:{color:'#263248'}}}}});
  dc('cPh');
  var so=rows.slice().sort(function(a,b){return b.progress-a.progress;}).slice(0,25);
  charts['cPh']=new Chart(document.getElementById('cPh'),{type:'bar',
    data:{labels:so.map(function(r){return '#'+r.s_no+' '+(r.project_name||r.customer).substring(0,30);}),
      datasets:[
        {label:'Engineering %',data:so.map(function(r){return r.eng_pct;}),backgroundColor:'#3b82f6',borderRadius:3,borderWidth:0},
        {label:'Delivery %',data:so.map(function(r){return r.del_pct;}),backgroundColor:'#a855f7',borderRadius:3,borderWidth:0},
        {label:'Execution %',data:so.map(function(r){return r.exec_pct;}),backgroundColor:'#f97316',borderRadius:3,borderWidth:0}]},
    options:{indexAxis:'y',animation:{duration:500},
      plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12,color:'#94a3b8'}},
        tooltip:{callbacks:{label:function(ctx){return ctx.dataset.label+': '+ctx.raw+'%';}}}},
      scales:{x:{beginAtZero:true,max:100,title:{display:true,text:'Percent Complete',color:'#94a3b8'},
                 ticks:{color:'#94a3b8'},grid:{color:'#263248'}},
              y:{ticks:{color:'#94a3b8',font:{size:11}},grid:{color:'#263248'}}}}});
}
function mbar(v,cls){var p=Math.max(0,Math.min(100,+v||0));
  return '<div class="bc"><div class="bb"><div class="bf '+cls+'" style="width:'+p+'%"></div></div><div class="bpct">'+p+'%</div></div>';}
var SB={'Completed':'sc','In Progress':'sp','On Hold':'sh','Cancelled':'sx','Not Started':'sn'};
var MB={'Delivered':'md','Partially Delivered':'mp','Ordered':'mo','Not Ordered':'mn'};
var PB={'High':'ph','Medium':'pm','Low':'pl'};
function buildTable(rows){
  document.getElementById('xRc').textContent=rows.length+' rows';
  document.getElementById('tB').innerHTML=rows.map(function(r){
    return '<tr><td>'+esc(r.s_no)+'</td><td>'+esc(r.customer)+'</td>'
      +'<td style="max-width:220px;white-space:normal;word-break:break-word">'+esc(r.project_name)+'</td>'
      +'<td>'+esc(r.region)+'</td>'
      +'<td><span class="badge '+(SB[r.status]||'sn')+'">'+esc(r.status)+'</span></td>'
      +'<td>'+mbar(r.eng_pct,'beng')+'</td><td>'+mbar(r.del_pct,'bdel')+'</td>'
      +'<td>'+mbar(r.exec_pct,'bexc')+'</td><td>'+mbar(r.progress,'ball')+'</td>'
      +'<td><span class="badge '+(MB[r.material_status]||'mn')+'">'+esc(r.material_status)+'</span></td>'
      +'<td><span class="badge '+(PB[r.priority]||'pm')+'">'+esc(r.priority)+'</span></td>'
      +'<td style="max-width:160px;white-space:normal;word-break:break-word;color:#64748b;font-size:11px">'+esc(r.work_status)+'</td>'
      +'</tr>';
  }).join('');
}
function fillDDs(){
  function fill(id,opts){var s=document.getElementById(id),cur=s.value;
    while(s.options.length>1)s.remove(1);
    opts.forEach(function(v){s.add(new Option(v,v));});if(cur)s.value=cur;}
  fill('fSt',['Not Started','In Progress','Completed','On Hold','Cancelled']);
  fill('fRg',uniq('region'));fill('fCu',uniq('customer'));
  fill('fMt',['Delivered','Partially Delivered','Ordered','Not Ordered']);
  fill('fPr',['High','Medium','Low']);
}
function getRows(){
  var base=activeQ?ALL.filter(function(r){return r.status===activeQ;}):ALL;
  var q=(document.getElementById('xSr').value||'').toLowerCase().trim();
  var fs=document.getElementById('fSt').value,fr=document.getElementById('fRg').value,
      fc=document.getElementById('fCu').value,fm=document.getElementById('fMt').value,fp=document.getElementById('fPr').value;
  return base.filter(function(r){
    if(fs&&r.status!==fs)return false;if(fr&&r.region!==fr)return false;
    if(fc&&r.customer!==fc)return false;if(fm&&r.material_status!==fm)return false;if(fp&&r.priority!==fp)return false;
    if(q){var h=[r.project_name,r.customer,r.location,r.job_ref,r.lpo_ref,r.work_status].join(' ').toLowerCase();if(h.indexOf(q)<0)return false;}
    return true;
  });
}
function redraw(){var rows=getRows();buildKPIs(rows);buildPhase(rows);buildCharts(rows);buildTable(rows);}
function qf(s){
  activeQ=s;
  ['qAll','qCo','qPr','qHo','qNS','qCa'].forEach(function(id){document.getElementById(id).classList.remove('active');});
  var m={'':'qAll','Completed':'qCo','In Progress':'qPr','On Hold':'qHo','Not Started':'qNS','Cancelled':'qCa'};
  if(m[s])document.getElementById(m[s]).classList.add('active');
  redraw();
}
function clearF(){activeQ='';document.getElementById('xSr').value='';
  ['fSt','fRg','fCu','fMt','fPr'].forEach(function(id){document.getElementById(id).value='';});qf('');}
['xSr','fSt','fRg','fCu','fMt','fPr'].forEach(function(id){
  document.getElementById(id).addEventListener('input',redraw);
  document.getElementById(id).addEventListener('change',redraw);
});
fillDDs();qf('');
"""

# Build the complete HTML by simple concatenation - no template substitution
html_parts = [
    "<!DOCTYPE html><html><head><meta charset=\"UTF-8\">",
    "<script src=\"https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js\"></script>",
    "<style>", css, "</style>",
    "</head><body>",
    body,
    "<script>",
    data_block,   # <-- actual JSON data injected here
    js,
    "</script>",
    "</body></html>"
]
HTML = "".join(html_parts)

components.html(HTML, height=3300, scrolling=True)
