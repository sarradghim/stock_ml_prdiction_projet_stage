"""
generate_dashboard_v3.py
========================
Genere un dashboard HTML interactif pour:
  - Stock 18 jours (Newton) -- tous les produits W1+W2+W3
  - Anomalies statistiques W1 + W2 + W3 (Isolation Forest + Rupture)

Usage:
    python generate_dashboard_v3.py

Sortie:
    dashboard_stock_18j.html  dans OUT_DIR
"""

import os
import json
import pandas as pd
import numpy as np
from sklearn.ensemble import IsolationForest

# ============================================================
# CONFIG
# ============================================================
W1 = r"C:\Users\INFOTEC\OneDrive\Bureau\Pre_w1w2\Cross_Week_Results\Week1_Final.xlsx"
W2 = r"C:\Users\INFOTEC\OneDrive\Bureau\Pre_w2w3\Cross_Week_Results\Week2_Final.xlsx"
W3 = r"C:\Users\INFOTEC\OneDrive\Bureau\pre_w3w4\Cross_Week_Results\Week3_Final.xlsx"

# Fichiers anomalies stat + metier (meme fichiers — contiennent anomaly col + Anomalie_Metier)
ANOM_W1 = r"C:\Users\INFOTEC\OneDrive\Bureau\anomalieW1\anomalie_metier\Week1_With_Anomalies_Metier.xlsx"
ANOM_W2 = r"C:\Users\INFOTEC\OneDrive\Bureau\anomalie_w2\anomalie_metier\Week2_With_Anomalies_Metier.xlsx"
ANOM_W3 = r"C:\Users\INFOTEC\OneDrive\Bureau\anomalie_w3\anomalie_metier\Week3_With_Anomalies_Metier.xlsx"


OUT_DIR  = r"C:\Users\INFOTEC\OneDrive\Bureau\newton"

PAYS_LIST = ['Cyclam', 'Germany', 'India', 'Korea',
             'Kunshan', 'Tianjin', 'USA', 'SAME', 'SCEET']

# ============================================================
# HELPERS
# ============================================================
def to_f(v):
    try:
        return float(str(v).replace(',', '.').replace(' ', '').replace('\xa0', ''))
    except Exception:
        return 0.0


def clean_str(s, max_len=50):
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace('`', "'")
            .replace('\\', '')
            .replace('\r', '')
            .replace('\n', ' ')
            .strip()[:max_len])


def newton_daily(inv_start, wu):
    xs = np.linspace(1, 3, 6)
    d1  =  1.01 - 0.99
    d2  =  0.99 - 1.01
    d12 = (d2 - d1) / 2
    factors = 0.99 + d1 * (xs - 1) + d12 * (xs - 1) * (xs - 2)
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


# ============================================================
# PARTIE 1 : STOCK 18J
# ============================================================
def build_stock_data():
    print("\n[STOCK] Calcul Newton 18 jours...")
    pays_data = {}

    for pays in PAYS_LIST:
        print(f"  {pays}...", end=" ")
        try:
            df1 = pd.read_excel(W1, sheet_name=pays, header=0)
            df2 = pd.read_excel(W2, sheet_name=pays, header=0)
            df3 = pd.read_excel(W3, sheet_name=pays, header=0)
            for df in [df1, df2, df3]:
                df['Part Number'] = df['Part Number'].astype(str).str.strip()

            products = []
            for _, row in df1.iterrows():
                pn   = clean_str(row['Part Number'])
                desc = clean_str(row.get('Description', ''))
                sup  = clean_str(row.get('Supplier', ''), 30)
                up   = to_f(row.get('Unit Price (EUR)', row.get('Unit Price (€)', 0)))

                r2 = df2[df2['Part Number'] == pn]
                r3 = df3[df3['Part Number'] == pn]

                inv1 = to_f(row.get('Real Inventory (Qty)', 0))
                wu1  = to_f(row.get('Weekly Usage (Qty)', 0))
                inv2 = to_f(r2.iloc[0]['Real Inventory (Qty)']) if not r2.empty else 0.0
                wu2  = to_f(r2.iloc[0]['Weekly Usage (Qty)'])   if not r2.empty else 0.0
                inv3 = to_f(r3.iloc[0]['Real Inventory (Qty)']) if not r3.empty else 0.0
                wu3  = to_f(r3.iloc[0]['Weekly Usage (Qty)'])   if not r3.empty else 0.0

                days_stock, days_usage, days_perturb = [], [], []
                for wk_inv, wk_wu in [(inv1, wu1), (inv2, wu2), (inv3, wu3)]:
                    av, us, _ = newton_daily(wk_inv, wk_wu)
                    ref = wk_wu / 6 if wk_wu != 0 else 0
                    for d in range(6):
                        days_stock.append(av[d])
                        days_usage.append(us[d])
                        p_pct = ((us[d] - ref) / ref * 100) if ref != 0 else 0.0
                        days_perturb.append(round(p_pct, 2))

                products.append({
                    'pn':      pn,
                    'desc':    desc,
                    'sup':     sup,
                    'up':      up,
                    'stock':   days_stock,
                    'usage':   days_usage,
                    'perturb': days_perturb,
                    'wu':      [wu1, wu2, wu3],
                    'inv':     [inv1, inv2, inv3],
                    'active':  1 if (wu1 > 0 or wu2 > 0 or wu3 > 0) else 0
                })

            pays_data[pays] = products
            active = sum(1 for p in products if p['active'] == 1)
            print(f"{len(products)} produits ({active} actifs)")

        except Exception as e:
            print(f"SKIP ({e})")

    return pays_data



# ============================================================
# PARTIE 2b : ANOMALIES METIER W1 + W2 + W3
# ============================================================
def build_metier_data():
    print("\n[METIER] Chargement anomalies metier W1+W2+W3...")
    all_metier = []

    week_configs = [
        ('W1', ANOM_W1),
        ('W2', ANOM_W2),
        ('W3', ANOM_W3),
    ]

    for week_label, met_file in week_configs:
        print(f"  Semaine {week_label}...")
        sheets = pd.read_excel(met_file, sheet_name=None)

        for pays in PAYS_LIST:
            if pays not in sheets:
                continue
            df = sheets[pays].copy()
            df.columns = df.columns.str.strip()

            am_col = next((c for c in df.columns if 'Anomalie_Metier' in c), None)
            if not am_col:
                continue

            cq  = next((c for c in df.columns if 'Real Inventory' in c), None)
            cv  = next((c for c in df.columns if 'Stock Value'    in c), None)
            cwu = next((c for c in df.columns if 'Weekly Usage'   in c), None)

            # Seulement les lignes avec anomalie reelle (pas Normal, pas vide)
            mask = df[am_col].notna() & (~df[am_col].isin(['Normal', '']))
            df_a = df[mask].copy()

            cnt = 0
            for _, row in df_a.iterrows():
                qty  = to_f(row[cq])  if cq  else 0.0
                val  = to_f(row[cv])  if cv  else 0.0
                wu   = to_f(row[cwu]) if cwu else 0.0
                pn   = clean_str(row.get('Part Number',  ''))
                desc = clean_str(row.get('Description',  ''))
                t    = clean_str(str(row[am_col]), 40)
                all_metier.append({
                    'pn': pn, 'desc': desc,
                    'qty': round(qty, 2), 'val': round(val, 2),
                    'wu':  round(wu,  2),
                    'type': t, 'pays': pays, 'week': week_label
                })
                cnt += 1
            print(f"    {pays}: {cnt} anomalies metier")

    print(f"  Total: {len(all_metier)} anomalies metier")
    return all_metier


# ============================================================
# PARTIE 2 : ANOMALIES W1 + W2 + W3
# ============================================================
def classify_anom(qty, val, q99_qty, q99_val):
    """Retourne le type d'anomalie, ou None si le produit est en rupture (qty<=0)."""
    if qty <= 0:
        return None              # rupture / negatif => ignorer
    elif val > q99_val:
        return 'Valeur extreme'
    elif qty > q99_qty:
        return 'Quantite extreme'
    else:
        return 'Anomalie statistique'


def build_anomaly_data():
    print("\n[ANOMALIES] Chargement W1+W2+W3...")
    all_anom = []

    week_configs = [
        ('W1', ANOM_W1, W1),
        ('W2', ANOM_W2, W2),
        ('W3', ANOM_W3, W3),
    ]

    for week_label, anom_file, orig_file in week_configs:
        print(f"\n  Semaine {week_label}...")

        sheets_anom = pd.read_excel(anom_file, sheet_name=None)
        sheets_orig = pd.read_excel(orig_file, sheet_name=None)

        for pays in PAYS_LIST:
            df_o = sheets_orig.get(pays, pd.DataFrame())
            if df_o.empty:
                continue
            df_o.columns = df_o.columns.str.strip()
            cq_o = next((c for c in df_o.columns if 'Real Inventory' in c), None)
            cv_o = next((c for c in df_o.columns if 'Stock Value'    in c), None)
            if not cq_o:
                continue

            df_o['_qty'] = df_o[cq_o].apply(to_f)
            df_o['_val'] = df_o[cv_o].apply(to_f) if cv_o else 0
            q99_qty = df_o['_qty'].quantile(0.99)
            q99_val = df_o['_val'].quantile(0.99)

            # --- Anomalies statistiques depuis fichier anomaly ---
            if pays in sheets_anom:
                df_a = sheets_anom[pays].copy()
                df_a.columns = df_a.columns.str.strip()
                cq = next((c for c in df_a.columns if 'Real Inventory' in c), None)
                cv = next((c for c in df_a.columns if 'Stock Value'    in c), None)
                # Filter: use existing anomaly col (W1) or re-run IsolationForest (W2/W3)
                if 'anomaly' in df_a.columns:
                    df_a = df_a[df_a['anomaly'] == -1].copy()
                else:
                    # Re-run IsolationForest to get equivalent detection
                    cq2 = next((c for c in df_o.columns if 'Real Inventory' in c), None)
                    cv2 = next((c for c in df_o.columns if 'Stock Value'    in c), None)
                    cwu2= next((c for c in df_o.columns if 'Weekly Usage'   in c), None)
                    if cq2:
                        Xfull = np.column_stack([
                            df_o[cq2].apply(to_f).values,
                            df_o[cv2].apply(to_f).values if cv2 else np.zeros(len(df_o)),
                            df_o[cwu2].apply(to_f).values if cwu2 else np.zeros(len(df_o)),
                        ])
                        if len(Xfull) >= 10:
                            from sklearn.ensemble import IsolationForest as IF
                            clf = IF(contamination=0.05, random_state=42)
                            preds = clf.fit_predict(Xfull)
                            scores= clf.score_samples(Xfull)
                            # Map pn -> score
                            pn_scores = {}
                            for i, row2 in df_o.reset_index(drop=True).iterrows():
                                pn_scores[str(row2.get('Part Number','')).strip()] = (preds[i], scores[i])
                            # Filter df_a to only anomaly==-1 rows
                            def is_anom(r):
                                pn2 = str(r.get('Part Number','')).strip()
                                return pn_scores.get(pn2, (1,0))[0] == -1
                            mask_a = df_a.apply(is_anom, axis=1)
                            df_a = df_a[mask_a].copy()
                            # Add score column
                            def get_score(r):
                                pn2 = str(r.get('Part Number','')).strip()
                                return pn_scores.get(pn2, (1,0))[1]
                            df_a['anomaly_score'] = df_a.apply(get_score, axis=1)

                for _, row in df_a.iterrows():
                    qty   = to_f(row.get(cq, 0))
                    val   = to_f(row.get(cv, 0))
                    score = to_f(row.get('anomaly_score', 0))
                    wu    = to_f(row.get('Weekly Usage (Qty)', 0))
                    pn    = clean_str(row.get('Part Number', ''))
                    desc  = clean_str(row.get('Description', ''))
                    t     = classify_anom(qty, val, q99_qty, q99_val)
                    if t is None:
                        continue   # skip: qty <= 0
                    all_anom.append({
                        'pn': pn, 'desc': desc,
                        'qty': round(qty, 2), 'val': round(val, 2),
                        'wu': round(wu, 2), 'score': round(score, 4),
                        'type': t, 'pays': pays, 'week': week_label
                    })
                pays_cnt = len([x for x in all_anom if x['week']==week_label and x['pays']==pays])
                print(f"    {pays}: {pays_cnt} anomalies")

    print(f"\n  Total: {len(all_anom)} anomalies")
    return all_anom


# ============================================================
# PARTIE 3 : TEMPLATE HTML
# ============================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Stock Dashboard 18 Jours</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root{--bg:#0a0e1a;--bg2:#111728;--bg3:#1a2235;--bd:#1e2d4a;--acc:#00d4ff;--acc2:#7c3aed;--acc3:#10b981;--warn:#f59e0b;--danger:#ef4444;--t:#e2e8f0;--t2:#94a3b8;--t3:#475569;--w1:#3b82f6;--w2:#8b5cf6;--w3:#06b6d4;--mono:'Space Mono',monospace;--body:'DM Sans',sans-serif;}
*{margin:0;padding:0;box-sizing:border-box;}
html{scrollbar-width:thin;scrollbar-color:#1e2d4a #0a0e1a;}
body{background:var(--bg);color:var(--t);font-family:var(--body);min-height:100vh;}
body::before{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(0,212,255,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(0,212,255,.025) 1px,transparent 1px);background-size:40px 40px;pointer-events:none;z-index:0;}
.wrap{position:relative;z-index:1;max-width:1400px;margin:0 auto;padding:20px;}
.hdr{display:flex;align-items:center;justify-content:space-between;padding:18px 24px;background:linear-gradient(135deg,#111728,#0d1829);border:1px solid #1e2d4a;border-radius:14px;margin-bottom:18px;position:relative;overflow:hidden;}
.hdr::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--acc),var(--acc2),var(--acc3));}
.hdr-logo{display:flex;align-items:center;gap:12px;}
.hdr-icon{width:38px;height:38px;background:linear-gradient(135deg,var(--acc),var(--acc2));border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:17px;}
.hdr-title{font-family:var(--mono);font-size:15px;letter-spacing:1px;}
.hdr-sub{font-size:11px;color:var(--t2);margin-top:2px;}
.hdr-stats{display:flex;gap:18px;}
.hstat .v{font-family:var(--mono);font-size:19px;font-weight:700;color:var(--acc);text-align:right;}
.hstat .l{font-size:10px;color:var(--t3);text-transform:uppercase;letter-spacing:1px;}
.tabs{display:flex;margin-bottom:18px;background:#111728;border:1px solid #1e2d4a;border-radius:12px;overflow:hidden;}
.tab-btn{flex:1;padding:13px 0;background:transparent;border:none;border-right:1px solid #1e2d4a;cursor:pointer;font-family:var(--mono);font-size:11px;font-weight:700;letter-spacing:1px;color:var(--t3);transition:all .2s;}
.tab-btn:last-child{border-right:none;}
.tab-btn.ts{background:rgba(0,212,255,.1);color:var(--acc);}
.tab-btn.ta{background:rgba(239,68,68,.1);color:var(--danger);}
.tab-btn.tm{background:rgba(245,158,11,.1);color:var(--warn);}
.controls{display:grid;grid-template-columns:190px 1fr;gap:14px;margin-bottom:18px;}
.pays-panel{background:#111728;border:1px solid #1e2d4a;border-radius:12px;overflow:hidden;}
.pays-hdr{padding:10px 14px;font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;border-bottom:1px solid #1e2d4a;background:#1a2235;}
.pays-btn{display:block;width:100%;text-align:left;padding:9px 14px;background:transparent;border:none;border-bottom:1px solid rgba(30,45,74,.4);cursor:pointer;color:var(--t2);font-family:var(--body);font-size:13px;transition:all .15s;position:relative;}
.pays-btn:hover{background:#1a2235;color:var(--t);}
.pays-btn.active{background:rgba(0,212,255,.07);color:var(--acc);font-weight:600;}
.pays-btn.active::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--acc);border-radius:0 2px 2px 0;}
.pc{float:right;background:#1a2235;border:1px solid #1e2d4a;border-radius:8px;padding:1px 6px;font-size:9px;font-family:var(--mono);color:var(--t3);}
.right-panel{display:flex;flex-direction:column;gap:12px;}
.srow{display:flex;gap:10px;align-items:center;}
.sw{flex:1;position:relative;}
.sw svg{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--t3);}
input[type=search]{width:100%;padding:10px 12px 10px 38px;background:#111728;border:1px solid #1e2d4a;border-radius:9px;color:var(--t);font-family:var(--body);font-size:13px;outline:none;transition:border-color .2s;}
input[type=search]:focus{border-color:var(--acc);}
input[type=search]::placeholder{color:var(--t3);}
select.csel{padding:10px 12px;background:#111728;border:1px solid #1e2d4a;border-radius:9px;color:var(--t);font-family:var(--body);font-size:13px;outline:none;cursor:pointer;}
button.cbtn{padding:10px 13px;background:#111728;border:1px solid #1e2d4a;border-radius:9px;color:var(--t3);font-family:var(--body);font-size:12px;cursor:pointer;white-space:nowrap;transition:all .2s;}
button.cbtn.on{background:rgba(0,212,255,.1);border-color:var(--acc);color:var(--acc);}
.pl{background:#111728;border:1px solid #1e2d4a;border-radius:12px;overflow:hidden;max-height:300px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:#1e2d4a #111728;}
.pi{display:flex;align-items:center;gap:10px;padding:9px 14px;border-bottom:1px solid rgba(30,45,74,.35);cursor:pointer;transition:all .15s;}
.pi:hover{background:#1a2235;}
.pi.sel{background:rgba(0,212,255,.05);border-left:3px solid var(--acc);padding-left:11px;}
.pi.ina{opacity:.35;}
.pi.ina .pn{color:var(--t3);}
.pdot{width:7px;height:7px;border-radius:50%;flex-shrink:0;background:var(--t3);}
.pdot.act{background:var(--acc3);}
.pdot.wrn{background:var(--danger);}
.pn{font-family:var(--mono);font-size:11px;color:var(--acc);min-width:95px;}
.pdesc{font-size:12px;color:var(--t2);flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.pwu{font-family:var(--mono);font-size:11px;color:var(--t3);min-width:72px;text-align:right;}
.pb{font-size:10px;padding:2px 7px;border-radius:8px;font-family:var(--mono);flex-shrink:0;}
.bok{background:rgba(16,185,129,.12);color:var(--acc3);}
.bko{background:rgba(239,68,68,.12);color:var(--danger);}
.sr{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px;}
.sc{background:#111728;border:1px solid #1e2d4a;border-radius:11px;padding:15px 16px;position:relative;overflow:hidden;}
.sc::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;}
.sc.c1::after{background:var(--w1)}.sc.c2::after{background:var(--acc3)}.sc.c3::after{background:var(--warn)}.sc.c4::after{background:var(--acc2)}
.slbl{font-size:9px;text-transform:uppercase;letter-spacing:1px;color:var(--t3);margin-bottom:5px;}
.sval{font-family:var(--mono);font-size:21px;font-weight:700;}
.ssub{font-size:10px;color:var(--t3);margin-top:3px;}
.sico{position:absolute;right:12px;top:12px;font-size:20px;opacity:.25;}
.ca{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;}
.ca.full{grid-template-columns:1fr;}
.cc{background:#111728;border:1px solid #1e2d4a;border-radius:12px;padding:18px;position:relative;overflow:hidden;}
.cc::before{content:'';position:absolute;top:0;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,var(--acc),transparent);opacity:.3;}
.ct{font-family:var(--mono);font-size:10px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:3px;}
.cs{font-size:11px;color:var(--t2);margin-bottom:14px;}
.h220{position:relative;height:220px;}.h280{position:relative;height:280px;}.h300{position:relative;height:300px;}
.wkl{display:flex;gap:14px;margin-bottom:10px;}
.wli{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--t2);}
.wld{width:11px;height:3px;border-radius:2px;}
.pgrid{display:grid;grid-template-columns:repeat(18,1fr);gap:2px;margin-top:6px;}
.pd{text-align:center;border-radius:3px;padding:4px 1px;font-family:var(--mono);font-size:8px;cursor:default;transition:transform .1s;}
.pd:hover{transform:scale(1.15);z-index:1;position:relative;}
.pok{background:rgba(16,185,129,.18);color:#4ade80;border:1px solid rgba(16,185,129,.28);}
.pko{background:rgba(239,68,68,.18);color:#f87171;border:1px solid rgba(239,68,68,.28);}
.pne{background:rgba(71,85,105,.15);color:#475569;border:1px solid rgba(71,85,105,.18);}
.ak{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px;}
.akc{background:#111728;border:1px solid #1e2d4a;border-radius:11px;padding:15px 16px;text-align:center;position:relative;overflow:hidden;}
.akv{font-family:var(--mono);font-size:22px;font-weight:700;margin:5px 0 3px;}
.akl{font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:1px;}
.tw{overflow-x:auto;max-height:400px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:#1e2d4a #111728;}
table{width:100%;border-collapse:collapse;font-size:12px;}
thead tr{background:#1a2235;position:sticky;top:0;z-index:2;}
th{padding:9px 12px;text-align:left;color:var(--t3);font-family:var(--mono);font-size:9px;letter-spacing:1px;white-space:nowrap;cursor:pointer;}
th:hover{color:var(--t);}
td{padding:7px 12px;border-bottom:1px solid rgba(30,45,74,.25);vertical-align:middle;}
tr:nth-child(even) td{background:rgba(26,34,53,.25);}
.bdg{font-size:10px;padding:2px 8px;border-radius:7px;font-family:var(--mono);white-space:nowrap;}
.emp{text-align:center;padding:50px 20px;color:var(--t3);}
.emp .ico{font-size:42px;margin-bottom:10px;}
</style>
</head>
<body>
<div class="wrap">

<div class="hdr">
  <div class="hdr-logo">
    <div class="hdr-icon">&#128202;</div>
    <div>
      <div class="hdr-title">STOCK DASHBOARD</div>
      <div class="hdr-sub">Newton 18J &nbsp;|&nbsp; 9 Pays &nbsp;|&nbsp; Anomalies Stat + Metier W1+W2+W3</div>
    </div>
  </div>
  <div class="hdr-stats">
    <div class="hstat"><div class="v" id="h-pays">9</div><div class="l">Pays</div></div>
    <div class="hstat"><div class="v" id="h-prods">--</div><div class="l">Produits</div></div>
    <div class="hstat"><div class="v">18</div><div class="l">Jours</div></div>
  </div>
</div>

<div class="tabs">
  <button class="tab-btn ts" id="tab-stock" onclick="switchTab('stock')">&#128202; STOCK 18 JOURS</button>
  <button class="tab-btn"    id="tab-anom"  onclick="switchTab('anom')">&#128308; ANOMALIES STAT</button>
  <button class="tab-btn"    id="tab-metier" onclick="switchTab('metier')">&#9888; ANOMALIES METIER</button>
</div>

<!-- ===================== STOCK TAB ===================== -->
<div id="view-stock">
  <div class="controls">
    <div class="pays-panel">
      <div class="pays-hdr">Pays</div>
      <div id="pays-list"></div>
    </div>
    <div class="right-panel">
      <div class="srow">
        <div class="sw">
          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
          <input type="search" id="search" placeholder="Chercher Part Number ou description...">
        </div>
        <select class="csel" id="sort-sel">
          <option value="wu_desc">Usage &#8595;</option>
          <option value="wu_asc">Usage &#8593;</option>
          <option value="pn_asc">PN A&#8594;Z</option>
          <option value="perturb_desc">Perturbation &#8595;</option>
        </select>
        <button class="cbtn" id="btn-ina" onclick="toggleIna()">&#128065; Tous: OFF</button>
      </div>
      <div class="pl" id="prod-list">
        <div class="emp"><div class="ico">&#128072;</div><p>Choisissez un pays</p></div>
      </div>
    </div>
  </div>
  <div class="sr" id="stat-row" style="display:none">
    <div class="sc c1"><div class="sico">&#128230;</div><div class="slbl">Stock Initial W1</div><div class="sval" id="s-inv">--</div><div class="ssub">Inventaire reel</div></div>
    <div class="sc c2"><div class="sico">&#9881;</div><div class="slbl">Usage Total 18j</div><div class="sval" id="s-use">--</div><div class="ssub">Unites consommees</div></div>
    <div class="sc c3"><div class="sico">&#9888;</div><div class="slbl">Hors tolerance</div><div class="sval" id="s-ko">--</div><div class="ssub">Perturb% &gt; 2%</div></div>
    <div class="sc c4"><div class="sico">&#128182;</div><div class="slbl">Valeur Stock W1</div><div class="sval" id="s-val">--</div><div class="ssub">Unit Price x Inv</div></div>
  </div>
  <div id="chart-box" style="display:none">
    <div class="ca full" style="margin-bottom:14px">
      <div class="cc">
        <div class="ct">Evolution du Stock -- 18 Jours</div>
        <div class="cs">W1 (J1-J6) W2 (J7-J12) W3 (J13-J18) &mdash; <span id="pn-title" style="color:var(--acc)">--</span></div>
        <div class="wkl">
          <div class="wli"><div class="wld" style="background:var(--w1)"></div>Week 1</div>
          <div class="wli"><div class="wld" style="background:var(--w2)"></div>Week 2</div>
          <div class="wli"><div class="wld" style="background:var(--w3)"></div>Week 3</div>
          <div class="wli"><div class="wld" style="background:var(--acc3)"></div>Usage</div>
        </div>
        <div class="h280"><canvas id="c-stock"></canvas></div>
      </div>
    </div>
    <div class="ca" style="margin-bottom:14px">
      <div class="cc"><div class="ct">Usage Journalier</div><div class="cs">Quantite consommee par jour</div><div class="h220"><canvas id="c-usage"></canvas></div></div>
      <div class="cc"><div class="ct">Perturbation %</div><div class="cs">(Usage_Jd - WU/6) / (WU/6) x 100</div><div class="h220"><canvas id="c-perturb"></canvas></div></div>
    </div>
    <div class="cc" style="margin-bottom:14px">
      <div class="ct">Heatmap Perturbation -- 18 Jours</div>
      <div class="cs">Vert ok (&le;2%) | Rouge hors tolerance (&gt;2%) | Gris sans usage</div>
      <div id="hm-wk" style="display:flex;gap:3px;margin-bottom:5px"></div>
      <div class="pgrid" id="hm-grid"></div>
      <div id="hm-days" style="display:flex;gap:2px;margin-top:5px"></div>
    </div>
  </div>
</div>

<!-- ===================== ANOMALIES STAT TAB ===================== -->
<div id="view-anom" style="display:none">
  <div style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:flex-end">
    <div>
      <div style="font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Semaine</div>
      <div style="display:flex;gap:6px">
        <button class="cbtn on" id="af-all" onclick="anomFilter('all')">Tout</button>
        <button class="cbtn" id="af-W1" onclick="anomFilter('W1')">W1</button>
        <button class="cbtn" id="af-W2" onclick="anomFilter('W2')">W2</button>
        <button class="cbtn" id="af-W3" onclick="anomFilter('W3')">W3</button>
      </div>
    </div>
    <div style="flex:1;min-width:160px">
      <div style="font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Pays</div>
      <select class="csel" id="af-pays" onchange="renderAnom()" style="width:100%"><option value="all">Tous les pays</option></select>
    </div>
    <div style="flex:1;min-width:160px">
      <div style="font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Type</div>
      <select class="csel" id="af-type" onchange="renderAnom()" style="width:100%"><option value="all">Tous les types</option></select>
    </div>
  </div>
  <div class="ak" id="anom-kpis"></div>
  <div class="ca" style="margin-bottom:14px">
    <div class="cc"><div class="ct">Anomalies par Pays</div><div class="h220"><canvas id="ac-pays"></canvas></div></div>
    <div class="cc"><div class="ct">Type d'Anomalie</div><div class="h220"><canvas id="ac-type"></canvas></div></div>
  </div>
  <div class="ca" style="margin-bottom:14px">
    <div class="cc"><div class="ct">Score -- Distribution</div><div class="h220"><canvas id="ac-score"></canvas></div></div>
    <div class="cc"><div class="ct">W1 vs W2 vs W3 par Pays</div><div class="h220"><canvas id="ac-compare"></canvas></div></div>
  </div>
  <div class="cc" style="margin-bottom:14px">
    <div class="ct">Scatter -- Quantite vs Valeur</div>
    <div class="h300"><canvas id="ac-scatter"></canvas></div>
  </div>
  <div class="cc">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <div class="ct">Detail Anomalies</div>
      <span id="anom-cnt" style="font-family:var(--mono);font-size:10px;color:var(--acc)"></span>
    </div>
    <div class="tw"><table>
      <thead><tr><th>SEM.</th><th>PAYS</th><th>PART NUMBER</th><th>DESCRIPTION</th>
        <th style="text-align:right">QTY</th><th style="text-align:right">VALEUR</th>
        <th style="text-align:right">SCORE</th><th>TYPE</th></tr></thead>
      <tbody id="anom-tbody"></tbody>
    </table></div>
  </div>
</div>

<!-- ===================== ANOMALIES METIER TAB ===================== -->
<div id="view-metier" style="display:none">
  <div style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:flex-end">
    <div>
      <div style="font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Semaine</div>
      <div style="display:flex;gap:6px">
        <button class="cbtn on" id="mf-all" onclick="metierFilter('all')">Tout</button>
        <button class="cbtn" id="mf-W1" onclick="metierFilter('W1')">W1</button>
        <button class="cbtn" id="mf-W2" onclick="metierFilter('W2')">W2</button>
        <button class="cbtn" id="mf-W3" onclick="metierFilter('W3')">W3</button>
      </div>
    </div>
    <div style="flex:1;min-width:160px">
      <div style="font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Pays</div>
      <select class="csel" id="mf-pays" onchange="renderMetier()" style="width:100%"><option value="all">Tous les pays</option></select>
    </div>
    <div style="flex:1;min-width:160px">
      <div style="font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Type</div>
      <select class="csel" id="mf-type" onchange="renderMetier()" style="width:100%"><option value="all">Tous les types</option></select>
    </div>
    <div style="min-width:220px">
      <div style="font-family:var(--mono);font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Recherche PN</div>
      <input type="search" id="mf-search" placeholder="Part Number..." oninput="renderMetier()" style="width:100%;padding:10px 12px;background:#111728;border:1px solid #1e2d4a;border-radius:9px;color:var(--t);font-family:var(--body);font-size:13px;outline:none;">
    </div>
  </div>

  <div class="ak" id="metier-kpis"></div>

  <div class="ca" style="margin-bottom:14px">
    <div class="cc"><div class="ct">Anomalies Metier par Pays</div><div class="h220"><canvas id="mc-pays"></canvas></div></div>
    <div class="cc"><div class="ct">Distribution par Type</div><div class="h220"><canvas id="mc-type"></canvas></div></div>
  </div>
  <div class="ca" style="margin-bottom:14px">
    <div class="cc"><div class="ct">W1 vs W2 vs W3 par Pays</div><div class="h220"><canvas id="mc-compare"></canvas></div></div>
    <div class="cc"><div class="ct">Evolution par Type (W1 -> W3)</div><div class="h220"><canvas id="mc-trend"></canvas></div></div>
  </div>

  <div class="cc">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <div class="ct">Detail Anomalies Metier</div>
      <span id="metier-cnt" style="font-family:var(--mono);font-size:10px;color:var(--warn)"></span>
    </div>
    <div class="tw"><table>
      <thead><tr><th>SEM.</th><th>PAYS</th><th>PART NUMBER</th><th>DESCRIPTION</th>
        <th style="text-align:right">QTY</th><th style="text-align:right">VALEUR (EUR)</th>
        <th style="text-align:right">WU</th><th>TYPE ANOMALIE</th></tr></thead>
      <tbody id="metier-tbody"></tbody>
    </table></div>
  </div>
</div>

</div>
<script>
const DATA = __DATA__;
const ANOM  = __ANOM__;
const METIER= __METIER__;

const DL=[];
['W1','W2','W3'].forEach(w=>{['Lu','Ma','Me','Je','Ve','Sa'].forEach((d,i)=>DL.push(w+'-J'+(i+1)));});
const WK =['#3b82f6','#8b5cf6','#06b6d4'];
const WKA=['rgba(59,130,246,.55)','rgba(139,92,246,.55)','rgba(6,182,212,.55)'];
const TCOL={'Valeur extreme':'#ef4444','Quantite extreme':'#f59e0b','Anomalie statistique':'#8b5cf6','Incoherence stock/valeur':'#06b6d4','Quantite negative':'#ec4899'};
const MCOL={'Rupture de stock':'#ef4444','Sur-stock critique':'#f97316','Rotation trop rapide':'#06b6d4','Sous-stock':'#f59e0b'};
const TT={backgroundColor:'#1a2235',borderColor:'#1e2d4a',borderWidth:1,titleColor:'#e2e8f0',bodyColor:'#94a3b8'};
const SCx={grid:{color:'rgba(30,45,74,.45)'},ticks:{color:'#475569',font:{family:'Space Mono',size:8},maxRotation:45}};
const SCy={grid:{color:'rgba(30,45,74,.45)'},ticks:{color:'#94a3b8',font:{family:'Space Mono',size:9}}};

let curPays=null,curProd=null,curProds=[],showAll=false,anomWk='all',metierWk='all';
const SC={},AC={},MC={};

function fmt(n){if(n==null)return'--';const a=Math.abs(n);if(a>=1e6)return(n/1e6).toFixed(1)+'M';if(a>=1e3)return(n/1e3).toFixed(1)+'k';return n.toFixed(2);}
function kill(s,id){if(s[id]){try{s[id].destroy();}catch(e){}delete s[id];}}

function switchTab(tab){
  ['stock','anom','metier'].forEach(t=>{
    document.getElementById('view-'+t).style.display=t===tab?'block':'none';
  });
  document.getElementById('tab-stock').className='tab-btn'+(tab==='stock'?' ts':'');
  document.getElementById('tab-anom').className='tab-btn'+(tab==='anom'?' ta':'');
  document.getElementById('tab-metier').className='tab-btn'+(tab==='metier'?' tm':'');
  if(tab==='anom'){initAnomFilters();renderAnom();}
  if(tab==='metier'){initMetierFilters();renderMetier();}
}

/* ---- PAYS ---- */
function initPays(){
  const c=document.getElementById('pays-list');c.innerHTML='';
  document.getElementById('h-pays').textContent=Object.keys(DATA).length;
  Object.keys(DATA).forEach(p=>{
    const b=document.createElement('button');b.className='pays-btn';
    b.innerHTML=p+'<span class="pc">'+DATA[p].length+'</span>';
    b.onclick=()=>selectPays(p);c.appendChild(b);
  });
}
function selectPays(p){
  curPays=p;curProd=null;
  document.querySelectorAll('.pays-btn').forEach(b=>b.classList.toggle('active',b.textContent.trim().startsWith(p)));
  const ac=DATA[p].filter(x=>x.active===1).length;
  document.getElementById('h-prods').textContent=DATA[p].length+' ('+ac+' actifs)';
  renderList();hideCharts();
}

/* ---- PRODUCT LIST ---- */
function renderList(){
  if(!curPays)return;
  const q=document.getElementById('search').value.toLowerCase();
  const s=document.getElementById('sort-sel').value;
  let pr=DATA[curPays].slice();
  if(!showAll)pr=pr.filter(p=>p.active===1);
  if(q)pr=pr.filter(p=>p.pn.toLowerCase().includes(q)||p.desc.toLowerCase().includes(q));
  if(s==='wu_desc')pr.sort((a,b)=>Math.max(...b.wu)-Math.max(...a.wu));
  else if(s==='wu_asc')pr.sort((a,b)=>Math.max(...a.wu)-Math.max(...b.wu));
  else if(s==='pn_asc')pr.sort((a,b)=>a.pn.localeCompare(b.pn));
  else if(s==='perturb_desc')pr.sort((a,b)=>Math.max(...b.perturb.map(Math.abs))-Math.max(...a.perturb.map(Math.abs)));
  curProds=pr;
  const c=document.getElementById('prod-list');
  if(!pr.length){c.innerHTML='<div class="emp"><div class="ico">&#128269;</div><p>Aucun produit</p></div>';return;}
  c.innerHTML=pr.map((p,i)=>{
    const mx=Math.max(...p.perturb.map(Math.abs));
    const ko=p.perturb.filter(x=>Math.abs(x)>2).length;
    const dot=p.active?(mx>2?'wrn':'act'):'';
    const bdg=ko>0?'<span class="pb bko">'+ko+' KO</span>':'<span class="pb bok">OK</span>';
    const sel=curProd&&curProd.pn===p.pn?' sel':'';
    const ina=p.active===0?' ina':'';
    return '<div class="pi'+sel+ina+'" onclick="selProd('+i+')">'
      +'<div class="pdot '+dot+'"></div><div class="pn">'+p.pn+'</div>'
      +'<div class="pdesc">'+p.desc+'</div><div class="pwu">'+fmt(p.wu.reduce((a,b)=>a+b,0))+'u</div>'+bdg+'</div>';
  }).join('');
}
function toggleIna(){showAll=!showAll;const b=document.getElementById('btn-ina');b.className='cbtn'+(showAll?' on':'');b.innerHTML=showAll?'&#128065; Tous: ON':'&#128065; Tous: OFF';renderList();}
function selProd(i){curProd=curProds[i];renderList();renderCharts();}

/* ---- STOCK CHARTS ---- */
function hideCharts(){document.getElementById('chart-box').style.display='none';document.getElementById('stat-row').style.display='none';['stock','usage','perturb'].forEach(id=>kill(SC,id));}
function renderCharts(){
  if(!curProd)return;
  const p=curProd;
  document.getElementById('chart-box').style.display='block';
  document.getElementById('stat-row').style.display='grid';
  document.getElementById('pn-title').textContent=p.pn+' -- '+p.desc;
  const tu=p.usage.reduce((a,b)=>a+b,0);
  const ko=p.perturb.filter(x=>Math.abs(x)>2).length;
  document.getElementById('s-inv').textContent=fmt(p.inv[0]);
  document.getElementById('s-use').textContent=fmt(tu);
  document.getElementById('s-ko').textContent=ko+' / 18';
  document.getElementById('s-val').textContent=fmt(p.inv[0]*p.up)+' EUR';
  kill(SC,'stock');
  SC['stock']=new Chart(document.getElementById('c-stock').getContext('2d'),{type:'line',data:{labels:DL,datasets:[{label:'Stock',data:p.stock,segment:{borderColor:ctx=>WK[Math.floor(ctx.p0DataIndex/6)]},pointBackgroundColor:p.stock.map((_,i)=>WK[Math.floor(i/6)]),pointBorderColor:'transparent',backgroundColor:'transparent',borderWidth:2.5,pointRadius:4,pointHoverRadius:6,tension:0.35},{label:'Usage',data:p.usage,borderColor:'rgba(16,185,129,.7)',backgroundColor:'rgba(16,185,129,.07)',borderWidth:1.5,borderDash:[4,3],pointRadius:3,pointBackgroundColor:'rgba(16,185,129,.7)',tension:0.35,fill:true,yAxisID:'y2'}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:11}}},tooltip:TT},scales:{x:SCx,y:{...SCy,title:{display:true,text:'Stock',color:'#475569',font:{size:10}}},y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:'#10b981',font:{family:'Space Mono',size:9}},title:{display:true,text:'Usage',color:'#10b981',font:{size:10}}}}}});
  kill(SC,'usage');
  SC['usage']=new Chart(document.getElementById('c-usage').getContext('2d'),{type:'bar',data:{labels:DL,datasets:[{label:'Usage',data:p.usage,backgroundColor:p.usage.map((_,i)=>WKA[Math.floor(i/6)]),borderColor:p.usage.map((_,i)=>WK[Math.floor(i/6)]),borderWidth:1,borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:TT},scales:{x:SCx,y:SCy}}});
  kill(SC,'perturb');
  SC['perturb']=new Chart(document.getElementById('c-perturb').getContext('2d'),{type:'bar',data:{labels:DL,datasets:[{label:'Perturb%',data:p.perturb,backgroundColor:p.perturb.map(v=>Math.abs(v)<=2?'rgba(16,185,129,.5)':'rgba(239,68,68,.5)'),borderColor:p.perturb.map(v=>Math.abs(v)<=2?'#10b981':'#ef4444'),borderWidth:1,borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{...TT,callbacks:{label:item=>' '+item.parsed.y.toFixed(2)+'%'}}},scales:{x:SCx,y:{...SCy,min:-15,max:15,ticks:{...SCy.ticks,callback:v=>v+'%'}}}}});
  const wk=document.getElementById('hm-wk'),gr=document.getElementById('hm-grid'),dl=document.getElementById('hm-days');
  wk.innerHTML=['W1 (J1-J6)','W2 (J7-J12)','W3 (J13-J18)'].map((w,i)=>'<div style="flex:6;text-align:center;font-size:9px;font-family:Space Mono,monospace;color:'+WK[i]+';border:1px solid '+WK[i]+'33;border-radius:4px;padding:1px 0">'+w+'</div>').join('');
  gr.innerHTML='';dl.innerHTML='';
  p.perturb.forEach((v,i)=>{
    const d=document.createElement('div');d.className='pd '+(v===0?'pne':Math.abs(v)<=2?'pok':'pko');
    d.textContent=v===0?'--':(v>0?'+':'')+v.toFixed(1);d.title=DL[i]+': '+v.toFixed(2)+'%';gr.appendChild(d);
    const l=document.createElement('div');l.style.cssText='font-size:7px;color:#334155;text-align:center;font-family:Space Mono,monospace;flex:1';l.textContent='J'+((i%6)+1);dl.appendChild(l);
  });
}

/* ---- ANOMALIE STAT ---- */
function initAnomFilters(){
  const pays=[...new Set(ANOM.map(d=>d.pays))].sort();
  const types=[...new Set(ANOM.map(d=>d.type))].sort();
  document.getElementById('af-pays').innerHTML='<option value="all">Tous les pays</option>'+pays.map(p=>'<option value="'+p+'">'+p+'</option>').join('');
  document.getElementById('af-type').innerHTML='<option value="all">Tous les types</option>'+types.map(t=>'<option value="'+t+'">'+t+'</option>').join('');
}
function anomFilter(w){anomWk=w;['all','W1','W2','W3'].forEach(id=>{const b=document.getElementById('af-'+id);if(b)b.className='cbtn'+(id===w?' on':'');});renderAnom();}
function getAnom(){const pays=document.getElementById('af-pays').value;const type=document.getElementById('af-type').value;return ANOM.filter(d=>(anomWk==='all'||d.week===anomWk)&&(pays==='all'||d.pays===pays)&&(type==='all'||d.type===type));}
function renderAnom(){
  const data=getAnom();
  const pc={},tc={};
  data.forEach(d=>{pc[d.pays]=(pc[d.pays]||0)+1;tc[d.type]=(tc[d.type]||0)+1;});
  const tv=data.reduce((s,d)=>s+d.val,0);
  document.getElementById('anom-kpis').innerHTML=[{ico:'&#128308;',v:data.length,l:'Anomalies',c:'#ef4444'},{ico:'&#127758;',v:Object.keys(pc).length,l:'Pays impactes',c:'#06b6d4'},{ico:'&#9881;',v:Object.keys(tc).length,l:'Types detectes',c:'#f59e0b'},{ico:'&#128182;',v:Math.round(tv).toLocaleString()+' EUR',l:'Valeur anorm.',c:'#8b5cf6',r:1}].map(k=>'<div class="akc"><div style="font-size:26px">'+k.ico+'</div><div class="akv" style="color:'+k.c+'">'+(k.r?k.v:Number(k.v).toLocaleString())+'</div><div class="akl">'+k.l+'</div><div style="position:absolute;bottom:0;left:0;right:0;height:2px;background:'+k.c+'"></div></div>').join('');
  const PL=Object.keys(pc).sort((a,b)=>pc[b]-pc[a]);
  const TL=Object.keys(tc).sort((a,b)=>tc[b]-tc[a]);
  kill(AC,'pays');AC['pays']=new Chart(document.getElementById('ac-pays').getContext('2d'),{type:'bar',data:{labels:PL,datasets:[{label:'Anomalies',data:PL.map(p=>pc[p]),backgroundColor:PL.map((_,i)=>'rgba(239,68,68,'+(0.35+i*0.05)+')'),borderColor:'#ef4444',borderWidth:1,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:TT},scales:{x:SCx,y:SCy}}});
  kill(AC,'type');AC['type']=new Chart(document.getElementById('ac-type').getContext('2d'),{type:'doughnut',data:{labels:TL,datasets:[{data:TL.map(t=>tc[t]),backgroundColor:TL.map(t=>TCOL[t]||'#64748b'),borderColor:'#111728',borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:10},padding:8},position:'bottom'},tooltip:TT}}});
  const scores=data.map(d=>d.score).sort((a,b)=>a-b);
  kill(AC,'score');AC['score']=new Chart(document.getElementById('ac-score').getContext('2d'),{type:'bar',data:{labels:scores.map(s=>s.toFixed(3)),datasets:[{data:scores,backgroundColor:scores.map(s=>s<-0.05?'rgba(239,68,68,.65)':'rgba(245,158,11,.55)'),borderWidth:0,borderRadius:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:TT},scales:{x:{display:false},y:SCy}}});
  const w1c={},w2c={},w3c={};
  data.forEach(d=>{if(d.week==='W1')w1c[d.pays]=(w1c[d.pays]||0)+1;else if(d.week==='W2')w2c[d.pays]=(w2c[d.pays]||0)+1;else w3c[d.pays]=(w3c[d.pays]||0)+1;});
  const allP=[...new Set([...Object.keys(w1c),...Object.keys(w2c),...Object.keys(w3c)])].sort();
  kill(AC,'compare');AC['compare']=new Chart(document.getElementById('ac-compare').getContext('2d'),{type:'bar',data:{labels:allP,datasets:[{label:'W1',data:allP.map(p=>w1c[p]||0),backgroundColor:'rgba(59,130,246,.6)',borderColor:'#3b82f6',borderWidth:1,borderRadius:3},{label:'W2',data:allP.map(p=>w2c[p]||0),backgroundColor:'rgba(139,92,246,.6)',borderColor:'#8b5cf6',borderWidth:1,borderRadius:3},{label:'W3',data:allP.map(p=>w3c[p]||0),backgroundColor:'rgba(6,182,212,.6)',borderColor:'#06b6d4',borderWidth:1,borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:11}}},tooltip:TT},scales:{x:SCx,y:SCy}}});
  const PC=['#3b82f6','#ef4444','#10b981','#f59e0b','#8b5cf6','#06b6d4','#ec4899','#84cc16','#f97316'];
  const dsm={};
  data.forEach(d=>{if(!dsm[d.pays])dsm[d.pays]={label:d.pays,data:[],backgroundColor:PC[Object.keys(dsm).length%PC.length],pointRadius:5,pointHoverRadius:8};dsm[d.pays].data.push({x:d.qty,y:d.val,pn:d.pn,type:d.type,week:d.week});});
  kill(AC,'scatter');AC['scatter']=new Chart(document.getElementById('ac-scatter').getContext('2d'),{type:'scatter',data:{datasets:Object.values(dsm)},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:10}},position:'right'},tooltip:{...TT,callbacks:{label:item=>{const d=item.raw;return[item.dataset.label+' '+d.week,'PN: '+d.pn,'Qty: '+d.x,'Val: '+d.y+' EUR','Type: '+d.type];}}}},scales:{x:{...SCx,title:{display:true,text:'Quantite (Qty)',color:'#475569',font:{size:10}}},y:{...SCy,title:{display:true,text:'Valeur Stock (EUR)',color:'#475569',font:{size:10}}}}}});
  document.getElementById('anom-cnt').textContent=data.length+' anomalies';
  const sorted=data.slice().sort((a,b)=>a.score-b.score);
  document.getElementById('anom-tbody').innerHTML=sorted.map(d=>{
    const wc=d.week==='W1'?'#3b82f6':d.week==='W2'?'#8b5cf6':'#06b6d4';
    const tc_=TCOL[d.type]||'#64748b';const sc=d.score<-0.05?'#ef4444':'#f59e0b';
    return '<tr><td><span class="bdg" style="background:'+wc+'22;color:'+wc+';border:1px solid '+wc+'44">'+d.week+'</span></td><td style="color:var(--t2)">'+d.pays+'</td><td style="font-family:var(--mono);font-size:11px;color:var(--acc)">'+d.pn+'</td><td style="color:var(--t2);max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+d.desc+'</td><td style="text-align:right;font-family:var(--mono);font-size:11px">'+d.qty.toLocaleString()+'</td><td style="text-align:right;font-family:var(--mono);font-size:11px">'+d.val.toLocaleString()+'</td><td style="text-align:right;font-family:var(--mono);font-size:11px;color:'+sc+'">'+d.score.toFixed(4)+'</td><td><span class="bdg" style="background:'+tc_+'22;color:'+tc_+';border:1px solid '+tc_+'44">'+d.type+'</span></td></tr>';
  }).join('');
}

/* ---- ANOMALIE METIER ---- */
function initMetierFilters(){
  const pays=[...new Set(METIER.map(d=>d.pays))].sort();
  const types=[...new Set(METIER.map(d=>d.type))].sort();
  document.getElementById('mf-pays').innerHTML='<option value="all">Tous les pays</option>'+pays.map(p=>'<option value="'+p+'">'+p+'</option>').join('');
  document.getElementById('mf-type').innerHTML='<option value="all">Tous les types</option>'+types.map(t=>'<option value="'+t+'">'+t+'</option>').join('');
}
function metierFilter(w){metierWk=w;['all','W1','W2','W3'].forEach(id=>{const b=document.getElementById('mf-'+id);if(b)b.className='cbtn'+(id===w?' on':'');});renderMetier();}
function getMetier(){
  const pays=document.getElementById('mf-pays').value;
  const type=document.getElementById('mf-type').value;
  const q=(document.getElementById('mf-search').value||'').toLowerCase();
  return METIER.filter(d=>(metierWk==='all'||d.week===metierWk)&&(pays==='all'||d.pays===pays)&&(type==='all'||d.type===type)&&(!q||d.pn.toLowerCase().includes(q)||d.desc.toLowerCase().includes(q)));
}
function renderMetier(){
  const data=getMetier();
  const pc={},tc={};
  data.forEach(d=>{pc[d.pays]=(pc[d.pays]||0)+1;tc[d.type]=(tc[d.type]||0)+1;});
  const tv=data.reduce((s,d)=>s+d.val,0);
  const pct=data.length>0?((Object.keys(pc).length/9)*100).toFixed(0):0;

  document.getElementById('metier-kpis').innerHTML=[
    {ico:'&#9888;',v:data.length,l:'Anomalies Metier',c:'#f59e0b'},
    {ico:'&#127758;',v:Object.keys(pc).length,l:'Pays impactes',c:'#ef4444'},
    {ico:'&#128202;',v:Object.keys(tc).length,l:'Types detectes',c:'#8b5cf6'},
    {ico:'&#128182;',v:Math.round(tv).toLocaleString()+' EUR',l:'Valeur concernee',c:'#06b6d4',r:1}
  ].map(k=>'<div class="akc"><div style="font-size:26px">'+k.ico+'</div><div class="akv" style="color:'+k.c+'">'+(k.r?k.v:Number(k.v).toLocaleString())+'</div><div class="akl">'+k.l+'</div><div style="position:absolute;bottom:0;left:0;right:0;height:2px;background:'+k.c+'"></div></div>').join('');

  const PL=Object.keys(pc).sort((a,b)=>pc[b]-pc[a]);
  const TL=Object.keys(tc).sort((a,b)=>tc[b]-tc[a]);

  kill(MC,'pays');MC['pays']=new Chart(document.getElementById('mc-pays').getContext('2d'),{type:'bar',data:{labels:PL,datasets:[{label:'Anomalies',data:PL.map(p=>pc[p]),backgroundColor:PL.map((_,i)=>'rgba(245,158,11,'+(0.35+i*0.06)+')'),borderColor:'#f59e0b',borderWidth:1,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:TT},scales:{x:SCx,y:SCy}}});

  kill(MC,'type');MC['type']=new Chart(document.getElementById('mc-type').getContext('2d'),{type:'doughnut',data:{labels:TL,datasets:[{data:TL.map(t=>tc[t]),backgroundColor:TL.map(t=>MCOL[t]||'#64748b'),borderColor:'#111728',borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:10},padding:8},position:'bottom'},tooltip:TT}}});

  // Compare W1 vs W2 vs W3
  const w1c={},w2c={},w3c={};
  data.forEach(d=>{if(d.week==='W1')w1c[d.pays]=(w1c[d.pays]||0)+1;else if(d.week==='W2')w2c[d.pays]=(w2c[d.pays]||0)+1;else w3c[d.pays]=(w3c[d.pays]||0)+1;});
  const allP=[...new Set([...Object.keys(w1c),...Object.keys(w2c),...Object.keys(w3c)])].sort();
  kill(MC,'compare');MC['compare']=new Chart(document.getElementById('mc-compare').getContext('2d'),{type:'bar',data:{labels:allP,datasets:[{label:'W1',data:allP.map(p=>w1c[p]||0),backgroundColor:'rgba(59,130,246,.6)',borderColor:'#3b82f6',borderWidth:1,borderRadius:3},{label:'W2',data:allP.map(p=>w2c[p]||0),backgroundColor:'rgba(139,92,246,.6)',borderColor:'#8b5cf6',borderWidth:1,borderRadius:3},{label:'W3',data:allP.map(p=>w3c[p]||0),backgroundColor:'rgba(6,182,212,.6)',borderColor:'#06b6d4',borderWidth:1,borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:11}}},tooltip:TT},scales:{x:SCx,y:SCy}}});

  // Trend: type evolution W1->W2->W3
  const allTypes=['Rupture de stock','Sur-stock critique','Rotation trop rapide','Sous-stock'];
  const trendData=allTypes.map(t=>{
    return{label:t,data:['W1','W2','W3'].map(w=>data.filter(d=>d.week===w&&d.type===t).length),borderColor:MCOL[t]||'#64748b',backgroundColor:(MCOL[t]||'#64748b')+'33',borderWidth:2,pointRadius:5,tension:0.3,fill:false};
  }).filter(d=>d.data.some(v=>v>0));
  kill(MC,'trend');MC['trend']=new Chart(document.getElementById('mc-trend').getContext('2d'),{type:'line',data:{labels:['W1','W2','W3'],datasets:trendData},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:10}}},tooltip:TT},scales:{x:SCx,y:SCy}}});

  document.getElementById('metier-cnt').textContent=data.length+' anomalies metier';

  // Table — show max 500 rows for performance
  const show=data.slice(0,500);
  document.getElementById('metier-tbody').innerHTML=show.map(d=>{
    const wc=d.week==='W1'?'#3b82f6':d.week==='W2'?'#8b5cf6':'#06b6d4';
    const tc_=MCOL[d.type]||'#f59e0b';
    return '<tr>'
      +'<td><span class="bdg" style="background:'+wc+'22;color:'+wc+';border:1px solid '+wc+'44">'+d.week+'</span></td>'
      +'<td style="color:var(--t2)">'+d.pays+'</td>'
      +'<td style="font-family:var(--mono);font-size:11px;color:var(--acc)">'+d.pn+'</td>'
      +'<td style="color:var(--t2);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+d.desc+'">'+d.desc+'</td>'
      +'<td style="text-align:right;font-family:var(--mono);font-size:11px">'+d.qty.toLocaleString()+'</td>'
      +'<td style="text-align:right;font-family:var(--mono);font-size:11px">'+d.val.toLocaleString()+'</td>'
      +'<td style="text-align:right;font-family:var(--mono);font-size:11px;color:var(--t3)">'+d.wu.toLocaleString()+'</td>'
      +'<td><span class="bdg" style="background:'+tc_+'22;color:'+tc_+';border:1px solid '+tc_+'44">'+d.type+'</span></td>'
      +'</tr>';
  }).join('');
  if(data.length>500){
    document.getElementById('metier-tbody').innerHTML+='<tr><td colspan="8" style="text-align:center;color:var(--t3);padding:12px;font-family:var(--mono);font-size:10px">... '+(data.length-500)+' lignes supplementaires (filtrer pour voir plus)</td></tr>';
  }
}

document.getElementById('search').addEventListener('input',renderList);
document.getElementById('sort-sel').addEventListener('change',renderList);
initPays();
const fp=Object.keys(DATA)[0];if(fp)selectPays(fp);
</script>
</body>
</html>"""


# ============================================================
# PARTIE 4 : GENERATION HTML
# ============================================================
def generate_html(stock_data, anom_data, metier_data):
    data_json = json.dumps(stock_data, separators=(',', ':'), ensure_ascii=True)
    anom_json = json.dumps(anom_data,  separators=(',', ':'), ensure_ascii=True)

    assert data_json.count('`') == 0, "ERREUR: backtick dans data_json!"
    metier_json = json.dumps(metier_data, separators=(',', ':'), ensure_ascii=True)
    assert anom_json.count('`') == 0, "ERREUR: backtick dans anom_json!"
    assert metier_json.count('`') == 0, "ERREUR: backtick dans metier_json!"

    html = HTML_TEMPLATE.replace('__DATA__', data_json).replace('__ANOM__', anom_json).replace('__METIER__', metier_json)
    return html


# ============================================================
# MAIN
# ============================================================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    stock_data  = build_stock_data()
    anom_data   = build_anomaly_data()
    metier_data = build_metier_data()

    print("\n[HTML] Generation du dashboard...")
    html = generate_html(stock_data, anom_data, metier_data)

    out_path = os.path.join(OUT_DIR, "dashboard_stock_18j.html")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb  = os.path.getsize(out_path) // 1024
    total    = sum(len(v) for v in stock_data.values())
    active   = sum(p['active'] for pays in stock_data.values() for p in pays)
    print(f"""
============================================================
  DASHBOARD GENERE AVEC SUCCES
============================================================
  Fichier   : {out_path}
  Taille    : {size_kb} KB
  Produits  : {total} total | {active} actifs
  Anomalies stat  : {len(anom_data)} (W1+W2+W3)
  Anomalies metier: {len(metier_data)} (W1+W2+W3)
============================================================
  >> Ouvrir dans Chrome / Edge / Firefox
============================================================
""")


if __name__ == '__main__':
    main()