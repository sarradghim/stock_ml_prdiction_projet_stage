import os
import json
import pandas as pd
import numpy as np

# ==============================
# CONFIG — CHEMINS
# ==============================
W1 = r"C:\Users\INFOTEC\OneDrive\Bureau\Pre_w1w2\Cross_Week_Results\Week1_Final.xlsx"
W2 = r"C:\Users\INFOTEC\OneDrive\Bureau\Pre_w2w3\Cross_Week_Results\Week2_Final.xlsx"
W3 = r"C:\Users\INFOTEC\OneDrive\Bureau\Pre_w2w3\Cross_Week_Results\Week3_Final.xlsx"
OUT_DIR = r"C:\Users\INFOTEC\OneDrive\Bureau\newton"
os.makedirs(OUT_DIR, exist_ok=True)

PAYS_LIST = ['Cyclam','Germany','India','Korea','Kunshan','Tianjin','USA','SAME','SCEET']

# ==============================
# HELPERS
# ==============================
def to_f(v):
    try:
        return float(str(v).replace(',', '.').replace(' ', ''))
    except:
        return 0.0

def newton_daily(inv_start, wu):
    xs = np.linspace(1, 3, 6)
    d1_n  = 1.01 - 0.99
    d2_n  = 0.99 - 1.01
    d12_n = (d2_n - d1_n) / 2
    factors = 0.99 + d1_n*(xs-1) + d12_n*(xs-1)*(xs-2)
    raw = (wu / 6) * factors
    if raw.sum() > 0:
        raw = raw * (wu / raw.sum())
    avant, apres = [], []
    s = inv_start
    for u in raw:
        avant.append(round(s, 2))
        s = max(0.0, s - u)
        apres.append(round(s, 2))
    usage = [round(a - b, 2) for a, b in zip(avant, apres)]
    return avant, usage, apres

# ==============================
# BUILD DATA
# ==============================
def build_data():
    pays_data = {}
    for pays in PAYS_LIST:
        print(f"  Loading {pays}...")
        try:
            w1 = pd.read_excel(W1, sheet_name=pays, header=0)
            w2 = pd.read_excel(W2, sheet_name=pays, header=0)
            w3 = pd.read_excel(W3, sheet_name=pays, header=0)
            for df in [w1, w2, w3]:
                df['Part Number'] = df['Part Number'].astype(str).str.strip()

            products = []
            for _, r in w1.iterrows():
                pn   = r['Part Number']
                desc = str(r.get('Description', ''))[:45]
                up   = to_f(r.get('Unit Price (EUR)', r.get('Unit Price (€)', 0)))

                r2 = w2[w2['Part Number'] == pn]
                r3 = w3[w3['Part Number'] == pn]

                inv1 = to_f(r.get('Real Inventory (Qty)', 0))
                wu1  = to_f(r.get('Weekly Usage (Qty)', 0))
                inv2 = to_f(r2.iloc[0].get('Real Inventory (Qty)', 0)) if not r2.empty else 0
                wu2  = to_f(r2.iloc[0].get('Weekly Usage (Qty)', 0))  if not r2.empty else 0
                inv3 = to_f(r3.iloc[0].get('Real Inventory (Qty)', 0)) if not r3.empty else 0
                wu3  = to_f(r3.iloc[0].get('Weekly Usage (Qty)', 0))  if not r3.empty else 0

                if wu1 == 0 and wu2 == 0 and wu3 == 0:
                    continue

                days_stock, days_usage, days_perturb = [], [], []
                for wk_inv, wk_wu in [(inv1, wu1), (inv2, wu2), (inv3, wu3)]:
                    av, us, ap = newton_daily(wk_inv, wk_wu)
                    wu6 = wk_wu / 6 if wk_wu != 0 else 0
                    for d in range(6):
                        days_stock.append(av[d])
                        days_usage.append(us[d])
                        p = ((us[d] - wu6) / wu6 * 100) if wu6 != 0 else 0
                        days_perturb.append(round(p, 2))

                products.append({
                    'pn':      pn,
                    'desc':    desc,
                    'up':      up,
                    'stock':   days_stock,
                    'usage':   days_usage,
                    'perturb': days_perturb,
                    'wu':      [wu1, wu2, wu3],
                    'inv':     [inv1, inv2, inv3]
                })

            pays_data[pays] = products
            print(f"    {pays}: {len(products)} produits avec mouvement")
        except Exception as e:
            print(f"    Skipped {pays}: {e}")
    return pays_data

# ==============================
# HTML TEMPLATE
# ==============================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Stock Dashboard 18 Jours</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root {
  --bg:#0a0e1a; --bg2:#111728; --bg3:#1a2235; --border:#1e2d4a;
  --accent:#00d4ff; --accent2:#7c3aed; --accent3:#10b981;
  --warn:#f59e0b; --danger:#ef4444;
  --text:#e2e8f0; --text2:#94a3b8; --text3:#475569;
  --w1:#3b82f6; --w2:#8b5cf6; --w3:#06b6d4;
  --font-mono:'Space Mono',monospace; --font-body:'DM Sans',sans-serif;
}
*{margin:0;padding:0;box-sizing:border-box;}
body{background:var(--bg);color:var(--text);font-family:var(--font-body);min-height:100vh;overflow-x:hidden;}
body::before{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(0,212,255,.03) 1px,transparent 1px),linear-gradient(90deg,rgba(0,212,255,.03) 1px,transparent 1px);background-size:40px 40px;pointer-events:none;z-index:0;}
.wrap{position:relative;z-index:1;max-width:1400px;margin:0 auto;padding:24px;}
header{display:flex;align-items:center;justify-content:space-between;padding:20px 28px;background:linear-gradient(135deg,var(--bg2) 0%,#0d1829 100%);border:1px solid var(--border);border-radius:16px;margin-bottom:24px;position:relative;overflow:hidden;}
header::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--accent),var(--accent2),var(--accent3));}
.logo{display:flex;align-items:center;gap:12px;}
.logo-icon{width:40px;height:40px;background:linear-gradient(135deg,var(--accent),var(--accent2));border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;}
.logo h1{font-family:var(--font-mono);font-size:16px;color:var(--text);letter-spacing:1px;}
.logo p{font-size:11px;color:var(--text2);margin-top:2px;}
.header-stats{display:flex;gap:20px;}
.hstat .val{font-family:var(--font-mono);font-size:20px;font-weight:700;color:var(--accent);text-align:right;}
.hstat .lbl{font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:1px;}
.controls{display:grid;grid-template-columns:200px 1fr;gap:16px;margin-bottom:24px;}
.pays-panel{background:var(--bg2);border:1px solid var(--border);border-radius:14px;overflow:hidden;}
.pays-title{padding:12px 16px;font-family:var(--font-mono);font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:2px;border-bottom:1px solid var(--border);background:var(--bg3);}
.pays-btn{display:block;width:100%;text-align:left;padding:10px 16px;background:transparent;border:none;cursor:pointer;color:var(--text2);font-family:var(--font-body);font-size:13px;border-bottom:1px solid rgba(30,45,74,.5);transition:all .15s;position:relative;}
.pays-btn:hover{background:var(--bg3);color:var(--text);}
.pays-btn.active{background:rgba(0,212,255,.08);color:var(--accent);font-weight:600;}
.pays-btn.active::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent);border-radius:0 2px 2px 0;}
.pays-count{float:right;background:var(--bg3);border:1px solid var(--border);border-radius:10px;padding:1px 7px;font-size:10px;font-family:var(--font-mono);color:var(--text3);}
.right-panel{display:flex;flex-direction:column;gap:16px;}
.search-row{display:flex;gap:12px;align-items:center;}
.search-wrap{flex:1;position:relative;}
.search-wrap svg{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--text3);}
input[type=search]{width:100%;padding:11px 14px 11px 42px;background:var(--bg2);border:1px solid var(--border);border-radius:10px;color:var(--text);font-family:var(--font-body);font-size:13px;outline:none;transition:border-color .2s;}
input[type=search]:focus{border-color:var(--accent);}
input[type=search]::placeholder{color:var(--text3);}
.sort-select{padding:11px 14px;background:var(--bg2);border:1px solid var(--border);border-radius:10px;color:var(--text);font-family:var(--font-body);font-size:13px;outline:none;cursor:pointer;min-width:160px;}
.product-list{background:var(--bg2);border:1px solid var(--border);border-radius:14px;overflow:hidden;max-height:320px;overflow-y:auto;}
.product-list::-webkit-scrollbar{width:5px;}
.product-list::-webkit-scrollbar-track{background:var(--bg);}
.product-list::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
.prod-item{display:flex;align-items:center;gap:12px;padding:10px 16px;border-bottom:1px solid rgba(30,45,74,.4);cursor:pointer;transition:all .15s;}
.prod-item:last-child{border-bottom:none;}
.prod-item:hover{background:var(--bg3);}
.prod-item.selected{background:rgba(0,212,255,.06);border-left:3px solid var(--accent);padding-left:13px;}
.prod-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;background:var(--text3);}
.prod-dot.has-data{background:var(--accent3);}
.prod-dot.high-perturb{background:var(--danger);}
.prod-pn{font-family:var(--font-mono);font-size:11px;color:var(--accent);min-width:100px;}
.prod-desc{font-size:12px;color:var(--text2);flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.prod-wu{font-family:var(--font-mono);font-size:11px;color:var(--text3);min-width:80px;text-align:right;}
.prod-badge{font-size:10px;padding:2px 8px;border-radius:10px;font-family:var(--font-mono);}
.badge-ok{background:rgba(16,185,129,.15);color:var(--accent3);}
.badge-ko{background:rgba(239,68,68,.15);color:var(--danger);}
.chart-area{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:24px;}
.chart-card{background:var(--bg2);border:1px solid var(--border);border-radius:14px;padding:20px;position:relative;overflow:hidden;animation:fadeIn .3s ease;}
.chart-card::before{content:'';position:absolute;top:0;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,var(--accent),transparent);opacity:.4;}
.chart-card.full{grid-column:1/-1;}
.card-title{font-family:var(--font-mono);font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:2px;margin-bottom:4px;}
.card-subtitle{font-size:12px;color:var(--text2);margin-bottom:16px;}
.card-subtitle span{color:var(--accent);font-weight:600;}
.chart-wrap{position:relative;height:220px;}
.chart-wrap-tall{position:relative;height:280px;}
.stat-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px;}
.stat-card{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px 18px;position:relative;overflow:hidden;}
.stat-card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;}
.stat-card.s1::after{background:var(--w1);}
.stat-card.s2::after{background:var(--accent3);}
.stat-card.s3::after{background:var(--warn);}
.stat-card.s4::after{background:var(--accent2);}
.stat-lbl{font-size:10px;text-transform:uppercase;letter-spacing:1px;color:var(--text3);margin-bottom:6px;}
.stat-val{font-family:var(--font-mono);font-size:22px;font-weight:700;color:var(--text);}
.stat-sub{font-size:11px;color:var(--text3);margin-top:4px;}
.stat-icon{position:absolute;right:14px;top:14px;font-size:22px;opacity:.3;}
.week-legend{display:flex;gap:16px;align-items:center;margin-bottom:12px;}
.wleg{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--text2);}
.wleg-dot{width:12px;height:3px;border-radius:2px;}
.perturb-grid{display:grid;grid-template-columns:repeat(18,1fr);gap:3px;margin-top:8px;}
.pg-day{text-align:center;border-radius:4px;padding:4px 2px;font-family:var(--font-mono);font-size:9px;cursor:default;transition:transform .1s;}
.pg-day:hover{transform:scale(1.2);z-index:1;position:relative;}
.pg-ok{background:rgba(16,185,129,.2);color:#4ade80;border:1px solid rgba(16,185,129,.3);}
.pg-ko{background:rgba(239,68,68,.2);color:#f87171;border:1px solid rgba(239,68,68,.3);}
.pg-neu{background:rgba(71,85,105,.2);color:#64748b;border:1px solid rgba(71,85,105,.2);}
.empty-state{text-align:center;padding:60px 20px;color:var(--text3);}
.empty-state .big{font-size:48px;margin-bottom:12px;}
html::-webkit-scrollbar{width:6px;}
html::-webkit-scrollbar-track{background:var(--bg);}
html::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:none;}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="logo">
      <div class="logo-icon">&#128202;</div>
      <div>
        <h1>STOCK DASHBOARD</h1>
        <p>Visualisation Newton -- 18 Jours x 9 Pays</p>
      </div>
    </div>
    <div class="header-stats">
      <div class="hstat"><div class="val" id="h-pays">9</div><div class="lbl">Pays</div></div>
      <div class="hstat"><div class="val" id="h-prods">--</div><div class="lbl">Produits actifs</div></div>
      <div class="hstat"><div class="val">18</div><div class="lbl">Jours (W1-W3)</div></div>
    </div>
  </header>

  <div class="controls">
    <div class="pays-panel">
      <div class="pays-title">Pays</div>
      <div id="pays-list"></div>
    </div>
    <div class="right-panel">
      <div class="search-row">
        <div class="search-wrap">
          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
          <input type="search" id="search" placeholder="Chercher Part Number ou description...">
        </div>
        <select class="sort-select" id="sort-select">
          <option value="wu_desc">Trier: Usage (haut)</option>
          <option value="wu_asc">Trier: Usage (bas)</option>
          <option value="pn_asc">Trier: PN A-Z</option>
          <option value="perturb_desc">Trier: Perturbation (haut)</option>
        </select>
      </div>
      <div class="product-list" id="product-list">
        <div class="empty-state"><div class="big">&#128072;</div><p>Choisissez un pays</p></div>
      </div>
    </div>
  </div>

  <div class="stat-row" id="stat-row" style="display:none">
    <div class="stat-card s1"><div class="stat-icon">&#128230;</div><div class="stat-lbl">Stock Initial W1</div><div class="stat-val" id="s-inv1">--</div><div class="stat-sub">Inventaire reel</div></div>
    <div class="stat-card s2"><div class="stat-icon">&#9881;&#65039;</div><div class="stat-lbl">Usage Total 18j</div><div class="stat-val" id="s-total-usage">--</div><div class="stat-sub">Unites consommees</div></div>
    <div class="stat-card s3"><div class="stat-icon">&#9888;&#65039;</div><div class="stat-lbl">Jours hors tolerance</div><div class="stat-val" id="s-perturb-ko">--</div><div class="stat-sub">Perturb% &gt; 2%</div></div>
    <div class="stat-card s4"><div class="stat-icon">&#128182;</div><div class="stat-lbl">Val. Stock Initial</div><div class="stat-val" id="s-val">--</div><div class="stat-sub">Unit Price x Inv W1</div></div>
  </div>

  <div id="chart-container" style="display:none">
    <div class="chart-area">
      <div class="chart-card full">
        <div class="card-title">Evolution du Stock -- 18 Jours</div>
        <div class="card-subtitle">W1 (J1-J6) - W2 (J7-J12) - W3 (J13-J18) -- Produit: <span id="chart-pn-title">--</span></div>
        <div class="week-legend">
          <div class="wleg"><div class="wleg-dot" style="background:#3b82f6"></div>Week 1</div>
          <div class="wleg"><div class="wleg-dot" style="background:#8b5cf6"></div>Week 2</div>
          <div class="wleg"><div class="wleg-dot" style="background:#06b6d4"></div>Week 3</div>
          <div class="wleg"><div class="wleg-dot" style="background:#10b981;border-top:2px dashed #10b981;height:0"></div>Usage</div>
        </div>
        <div class="chart-wrap-tall"><canvas id="chart-stock"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="card-title">Usage Journalier</div>
        <div class="card-subtitle">Quantite consommee chaque jour</div>
        <div class="chart-wrap"><canvas id="chart-usage"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="card-title">Perturbation %</div>
        <div class="card-subtitle">(Usage_Jd - WU/6) / (WU/6) x 100 -- <span style="color:#10b981">Vert =2%</span> / <span style="color:#ef4444">Rouge &gt;2%</span></div>
        <div class="chart-wrap"><canvas id="chart-perturb"></canvas></div>
      </div>
    </div>
    <div class="chart-card" style="margin-top:16px">
      <div class="card-title">Heatmap Perturbation -- 18 Jours</div>
      <div class="card-subtitle">Chaque case = 1 jour -- Vert =2% - Rouge &gt;2% - Gris = pas usage</div>
      <div id="heatmap-labels" style="display:flex;gap:3px;margin-bottom:6px;padding-top:4px"></div>
      <div class="perturb-grid" id="perturb-heatmap"></div>
      <div style="display:flex;gap:3px;margin-top:6px" id="heatmap-day-labels"></div>
    </div>
  </div>
</div>

<script>
var DATA = __INJECT_DATA__;

var DAYS_LBL = [];
['W1','W2','W3'].forEach(function(w,wi){
  ['Lu','Ma','Me','Je','Ve','Sa'].forEach(function(d,di){
    DAYS_LBL.push(w+'-J'+(di+1));
  });
});

var WK_COLORS = ['#3b82f6','#8b5cf6','#06b6d4'];
var USAGE_COLORS = ['rgba(59,130,246,.6)','rgba(139,92,246,.6)','rgba(6,182,212,.6)'];

var currentPays = null;
var currentProd = null;
var charts = {};
var filteredProds = [];

function fmt(n){
  if(n===undefined||n===null) return '--';
  if(Math.abs(n)>=1e6) return (n/1e6).toFixed(1)+'M';
  if(Math.abs(n)>=1e3) return (n/1e3).toFixed(1)+'k';
  return n.toFixed(2);
}

function initPays(){
  var container = document.getElementById('pays-list');
  container.innerHTML = '';
  var pays = Object.keys(DATA);
  document.getElementById('h-pays').textContent = pays.length;
  pays.forEach(function(p){
    var btn = document.createElement('button');
    btn.className = 'pays-btn';
    btn.innerHTML = p + ' <span class="pays-count">'+DATA[p].length+'</span>';
    btn.onclick = function(){ selectPays(p); };
    container.appendChild(btn);
  });
}

function selectPays(pays){
  currentPays = pays;
  currentProd = null;
  document.querySelectorAll('.pays-btn').forEach(function(b){
    b.classList.toggle('active', b.textContent.trim().startsWith(pays));
  });
  document.getElementById('h-prods').textContent = DATA[pays].length;
  renderProductList();
  hideCharts();
}

function renderProductList(){
  if(!currentPays) return;
  var search = document.getElementById('search').value.toLowerCase();
  var sort   = document.getElementById('sort-select').value;
  var prods  = DATA[currentPays].slice();

  if(search){
    prods = prods.filter(function(p){
      return p.pn.toLowerCase().includes(search) || p.desc.toLowerCase().includes(search);
    });
  }

  if(sort==='wu_desc')      prods.sort(function(a,b){ return Math.max.apply(null,b.wu)-Math.max.apply(null,a.wu); });
  else if(sort==='wu_asc')  prods.sort(function(a,b){ return Math.max.apply(null,a.wu)-Math.max.apply(null,b.wu); });
  else if(sort==='pn_asc')  prods.sort(function(a,b){ return a.pn.localeCompare(b.pn); });
  else if(sort==='perturb_desc') prods.sort(function(a,b){
    return Math.max.apply(null,b.perturb.map(Math.abs)) - Math.max.apply(null,a.perturb.map(Math.abs));
  });

  filteredProds = prods;
  var container = document.getElementById('product-list');
  if(!prods.length){
    container.innerHTML = '<div class="empty-state"><div class="big">&#128269;</div><p>Aucun produit trouve</p></div>';
    return;
  }

  container.innerHTML = prods.map(function(p,i){
    var maxP = Math.max.apply(null, p.perturb.map(Math.abs));
    var koCount = p.perturb.filter(function(x){ return Math.abs(x)>2; }).length;
    var dotClass = p.wu.some(function(w){ return w>0; }) ? (maxP>2?'has-data high-perturb':'has-data') : '';
    var badge = koCount>0
      ? '<span class="prod-badge badge-ko">'+koCount+' KO</span>'
      : '<span class="prod-badge badge-ok">OK</span>';
    var totalWU = p.wu.reduce(function(a,b){ return a+b; }, 0);
    return '<div class="prod-item'+(currentProd&&currentProd.pn===p.pn?' selected':'')+'" onclick="selectProduct('+i+')">'
      +'<div class="prod-dot '+dotClass+'"></div>'
      +'<div class="prod-pn">'+p.pn+'</div>'
      +'<div class="prod-desc">'+p.desc+'</div>'
      +'<div class="prod-wu">'+fmt(totalWU)+' u</div>'
      +badge
      +'</div>';
  }).join('');
}

function selectProduct(idx){
  currentProd = filteredProds[idx];
  renderProductList();
  renderCharts();
}

function hideCharts(){
  document.getElementById('chart-container').style.display = 'none';
  document.getElementById('stat-row').style.display = 'none';
  Object.keys(charts).forEach(function(k){ if(charts[k]) charts[k].destroy(); });
  charts = {};
}

function destroyChart(id){
  if(charts[id]){ charts[id].destroy(); delete charts[id]; }
}

function renderCharts(){
  if(!currentProd) return;
  var p = currentProd;

  document.getElementById('chart-container').style.display = 'block';
  document.getElementById('stat-row').style.display = 'grid';
  document.getElementById('chart-pn-title').textContent = p.pn + ' -- ' + p.desc;

  document.getElementById('s-inv1').textContent = fmt(p.inv[0]);
  var totalUsage = p.usage.reduce(function(a,b){ return a+b; }, 0);
  document.getElementById('s-total-usage').textContent = fmt(totalUsage);
  var koCount = p.perturb.filter(function(x){ return Math.abs(x)>2; }).length;
  document.getElementById('s-perturb-ko').textContent = koCount + ' / 18';
  document.getElementById('s-val').textContent = fmt(p.inv[0]*p.up) + ' EUR';

  // STOCK CHART
  destroyChart('chart-stock');
  var stockCtx = document.getElementById('chart-stock').getContext('2d');
  charts['chart-stock'] = new Chart(stockCtx, {
    type: 'line',
    data: {
      labels: DAYS_LBL,
      datasets: [
        {
          label: 'Stock Avant (unites)',
          data: p.stock,
          segment: { borderColor: function(ctx){ return WK_COLORS[Math.floor(ctx.p0DataIndex/6)]; } },
          backgroundColor: 'transparent',
          borderWidth: 2.5,
          pointRadius: 4,
          pointHoverRadius: 6,
          pointBackgroundColor: p.stock.map(function(_,i){ return WK_COLORS[Math.floor(i/6)]; }),
          pointBorderColor: 'transparent',
          tension: 0.35
        },
        {
          label: 'Usage Journalier',
          data: p.usage,
          borderColor: 'rgba(16,185,129,.7)',
          backgroundColor: 'rgba(16,185,129,.08)',
          borderWidth: 1.5,
          borderDash: [4,3],
          pointRadius: 3,
          pointBackgroundColor: 'rgba(16,185,129,.7)',
          tension: 0.35,
          fill: true,
          yAxisID: 'y2'
        }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { family: 'DM Sans', size: 11 } } },
        tooltip: { backgroundColor: '#1a2235', borderColor: '#1e2d4a', borderWidth: 1, titleColor: '#e2e8f0', bodyColor: '#94a3b8' }
      },
      scales: {
        x: { grid: { color: 'rgba(30,45,74,.5)' }, ticks: { color: '#475569', font: { family: 'Space Mono', size: 9 }, maxRotation: 45 } },
        y: { grid: { color: 'rgba(30,45,74,.5)' }, ticks: { color: '#94a3b8', font: { family: 'Space Mono', size: 9 } }, title: { display: true, text: 'Stock (unites)', color: '#475569', font: { size: 10 } } },
        y2: { position: 'right', grid: { drawOnChartArea: false }, ticks: { color: '#10b981', font: { family: 'Space Mono', size: 9 } }, title: { display: true, text: 'Usage', color: '#10b981', font: { size: 10 } } }
      }
    }
  });

  // USAGE CHART
  destroyChart('chart-usage');
  var usageCtx = document.getElementById('chart-usage').getContext('2d');
  charts['chart-usage'] = new Chart(usageCtx, {
    type: 'bar',
    data: {
      labels: DAYS_LBL,
      datasets: [{
        label: 'Usage',
        data: p.usage,
        backgroundColor: p.usage.map(function(_,i){ return USAGE_COLORS[Math.floor(i/6)]; }),
        borderColor:     p.usage.map(function(_,i){ return WK_COLORS[Math.floor(i/6)]; }),
        borderWidth: 1, borderRadius: 3
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { backgroundColor: '#1a2235', borderColor: '#1e2d4a', borderWidth: 1, titleColor: '#e2e8f0', bodyColor: '#94a3b8' } },
      scales: {
        x: { grid: { color: 'rgba(30,45,74,.5)' }, ticks: { color: '#475569', font: { family: 'Space Mono', size: 8 }, maxRotation: 45 } },
        y: { grid: { color: 'rgba(30,45,74,.5)' }, ticks: { color: '#94a3b8', font: { family: 'Space Mono', size: 9 } } }
      }
    }
  });

  // PERTURBATION CHART
  destroyChart('chart-perturb');
  var pCtx = document.getElementById('chart-perturb').getContext('2d');
  charts['chart-perturb'] = new Chart(pCtx, {
    type: 'bar',
    data: {
      labels: DAYS_LBL,
      datasets: [{
        label: 'Perturb%',
        data: p.perturb,
        backgroundColor: p.perturb.map(function(v){ return Math.abs(v)<=2?'rgba(16,185,129,.5)':'rgba(239,68,68,.5)'; }),
        borderColor:     p.perturb.map(function(v){ return Math.abs(v)<=2?'#10b981':'#ef4444'; }),
        borderWidth: 1, borderRadius: 3
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { backgroundColor: '#1a2235', borderColor: '#1e2d4a', borderWidth: 1, titleColor: '#e2e8f0', bodyColor: '#94a3b8', callbacks: { label: function(item){ return ' '+item.parsed.y.toFixed(2)+'%'; } } } },
      scales: {
        x: { grid: { color: 'rgba(30,45,74,.5)' }, ticks: { color: '#475569', font: { family: 'Space Mono', size: 8 }, maxRotation: 45 } },
        y: { grid: { color: 'rgba(30,45,74,.5)' }, ticks: { color: '#94a3b8', font: { family: 'Space Mono', size: 9 }, callback: function(v){ return v+'%'; } }, min: -15, max: 15 }
      }
    }
  });

  // HEATMAP
  var hm = document.getElementById('perturb-heatmap');
  var lbl = document.getElementById('heatmap-day-labels');
  var hlbl = document.getElementById('heatmap-labels');
  hlbl.innerHTML = ['Week 1 (J1-J6)','Week 2 (J7-J12)','Week 3 (J13-J18)'].map(function(w,wi){
    return '<div style="flex:6;text-align:center;font-size:10px;font-family:Space Mono,monospace;color:'+WK_COLORS[wi]+';border:1px solid '+WK_COLORS[wi]+'33;border-radius:4px;padding:2px 0">'+w+'</div>';
  }).join('');
  hm.innerHTML = '';
  lbl.innerHTML = '';
  p.perturb.forEach(function(v,i){
    var cls = v===0?'pg-neu':(Math.abs(v)<=2?'pg-ok':'pg-ko');
    var day = document.createElement('div');
    day.className = 'pg-day '+cls;
    day.textContent = v===0?'--':((v>0?'+':'')+v.toFixed(1));
    day.title = DAYS_LBL[i]+': '+v.toFixed(2)+'%';
    hm.appendChild(day);
    var l = document.createElement('div');
    l.style.cssText = 'font-size:8px;color:#334155;text-align:center;font-family:Space Mono,monospace;flex:1';
    l.textContent = 'J'+((i%6)+1);
    lbl.appendChild(l);
  });
}

document.getElementById('search').addEventListener('input', renderProductList);
document.getElementById('sort-select').addEventListener('change', renderProductList);
initPays();
var firstPays = Object.keys(DATA)[0];
if(firstPays) selectPays(firstPays);
</script>
</body>
</html>"""

# ==============================
# GENERATE HTML
# ==============================
def generate_html(pays_data):
    data_json = json.dumps(pays_data, separators=(',', ':'), ensure_ascii=True)
    html = HTML_TEMPLATE.replace('__INJECT_DATA__', data_json)
    return html

# ==============================
# MAIN
# ==============================
def main():
    print("Building data from Excel files...")
    pays_data = build_data()

    total_prods = sum(len(v) for v in pays_data.values())
    print(f"\nTotal produits avec mouvement: {total_prods}")

    print("\nGenerating HTML dashboard...")
    html = generate_html(pays_data)

    out_path = os.path.join(OUT_DIR, "dashboard_stock_18j.html")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(out_path) // 1024
    print(f"\nDone! Dashboard saved: {out_path}")
    print(f"File size: {size_kb} KB")
    print("\nOuvrir le fichier HTML dans votre navigateur (Chrome/Edge/Firefox)")

main()