# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os, json
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HISTORY_FILE = "kpi_history.json"

def load_kpi_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_kpi_history(history):
    try:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(history, f, indent=2, ensure_ascii=False)
    except:
        pass

def save_current_kpis(ckdf, qk, pk, pscores, qscores, pa, qa):
    history = load_kpi_history()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    MAX_H = 20
    def add_entry(key, val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return
        fv = round(float(val), 2)
        if key not in history:
            history[key] = []
        if len(history[key]) > 0 and history[key][-1]["date"][:16] == now_str[:16] and abs(history[key][-1]["value"] - fv) < 0.01:
            return
        history[key].append({"date": now_str, "value": fv})
        if len(history[key]) > MAX_H:
            history[key] = history[key][-MAX_H:]
    for poste in ckdf.index:
        for k in qk + pk:
            if k in ckdf.columns:
                v = ckdf.loc[poste, k]
                if pd.notna(v):
                    add_entry(f"{poste}__{k}", v)
        if poste in pscores:
            add_entry(f"{poste}__Score Performance", pscores[poste])
        if poste in qscores:
            add_entry(f"{poste}__Score Qualite", qscores[poste])
    for k in qk:
        add_entry(f"__total__{k}", pa.get(k))
    for k in pk:
        add_entry(f"__total__{k}", qa.get(k))
    add_entry("__total__Score Performance", float(np.mean(list(pscores.values()))) if pscores else 0.0)
    add_entry("__total__Score Qualite", float(np.mean(list(qscores.values()))) if qscores else 0.0)
    save_kpi_history(history)
    return history

def generate_sparkline(values):
    if not values or len(values) < 2:
        return ""
    blocks = ['\u2581', '\u2582', '\u2583', '\u2584', '\u2585', '\u2586', '\u2587', '\u2588']
    vals = [float(v) for v in values if v is not None]
    if len(vals) < 2:
        return ""
    dv = vals[-8:]
    mn, mx = min(dv), max(dv)
    if mx == mn:
        return '\u2584' * len(dv)
    norm = [(v - mn) / (mx - mn) * 7 for v in dv]
    return ''.join(blocks[min(7, max(0, int(round(n))))] for n in norm)

def get_trend_info(values, lower_better=False):
    if len(values) < 2:
        return "", ""
    prev, curr = float(values[-2]), float(values[-1])
    diff = curr - prev
    if abs(diff) < 0.5:
        return "\u27a1\ufe0f", "stable"
    if lower_better:
        return ("\U0001f4c8", "amelioration") if diff < 0 else ("\U0001f4c9", "degradation")
    else:
        return ("\U0001f4c8", "amelioration") if diff > 0 else ("\U0001f4c9", "degradation")

def build_kpi_cell(value_str, hist_entries, lower_better=False):
    if not hist_entries or len(hist_entries) < 2:
        return value_str
    values = [h["value"] for h in hist_entries if h.get("value") is not None]
    if len(values) < 2:
        return value_str
    spark = generate_sparkline(values)
    trend_icon, trend_label = get_trend_info(values, lower_better)
    curr, prev = hist_entries[-1], hist_entries[-2]
    cv, pv = curr.get("value", "N/A"), prev.get("value", "N/A")
    try:
        diff = float(cv) - float(pv)
    except:
        diff = 0
    lines = [
        "Valeur actuelle : %s" % cv,
        "Valeur precedente : %s" % pv,
        "Ecart : %+.2f" % diff,
        "Date actuelle : %s" % curr.get('date', 'N/A'),
        "Date precedente : %s" % prev.get('date', 'N/A'),
        "Tendance : %s" % trend_label,
        "Historique : %s point(s)" % len(hist_entries)
    ]
    tooltip = "&#10;".join(lines)
    return '<span title="%s" class="kc"><span class="kt">%s</span><span class="ks">%s</span><span class="kv">%s</span></span>' % (tooltip, trend_icon, spark, value_str)

def inject_custom_css():
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--p:#1e3a5f;--pl:#2c5282;--b:#e2e8f0;--r:10px}
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#edf2f7;font-family:'Inter',sans-serif}
    .main .block-container{max-width:100%!important;padding-left:0!important;padding-right:0!important;padding-top:0!important;padding-bottom:0!important}
    section.main > div{padding-left:0!important;padding-right:0!important}
    .mh,.cr,.ca,.tw{width:100%!important;margin-left:0!important;margin-right:0!important}
    .stTabs,.stTabs>div,.stTabs [data-baseweb="tab-list"]{width:100%!important;max-width:100%!important}
    .mh{background:linear-gradient(135deg,var(--p),var(--pl));padding:10px 16px;border-radius:var(--r);margin-bottom:4px;box-shadow:0 6px 20px rgba(0,0,0,.1);overflow:hidden}
    .mh h1{color:#fff;font-size:16px;font-weight:800;margin:0;display:inline}
    .mh .db{float:right;background:rgba(255,255,255,.15);padding:2px 10px;border-radius:14px;color:#fff;font-size:10px;font-weight:500;border:1px solid rgba(255,255,255,.2);margin-top:2px}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;margin-bottom:4px}
    .cc{background:#fff;border-radius:var(--r);padding:8px 10px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid var(--b);text-align:center}
    .cc .cv{font-size:22px;font-weight:900;line-height:1}
    .cc .cl{font-size:7px;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:1px}
    .cc.c1{border-top:3px solid #3182ce}.cc.c1 .cv{color:#2b6cb0}
    .cc.c2{border-top:3px solid #38a169}.cc.c2 .cv{color:#276749}
    .cc.c3{border-top:3px solid #805ad5}.cc.c3 .cv{color:#6b46c1}
    .cc.c4{border-top:3px solid #e53e3e}.cc.c4 .cv{color:#c53030}
    .stl{font-size:11px;font-weight:700;color:var(--p);margin:4px 0 1px 0;padding-left:8px;border-left:3px solid var(--pl)}
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}.stl.ev{border-left-color:#dd6b20}
    .tw{border-collapse:collapse;font-family:'Inter',sans-serif;font-size:8px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:7px;text-transform:uppercase;letter-spacing:.3px;padding:3px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw.ct thead th{background:linear-gradient(135deg,#553c9a,#805ad5)}
    .tw.evt thead th{background:linear-gradient(135deg,#c05621,#dd6b20)}
    .tw tbody td{padding:2px 3px;border-bottom:1px solid #edf2f7;white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important;font-size:8px!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important;font-size:8px!important}
    .kc{display:inline-flex;align-items:center;gap:1px;white-space:nowrap;cursor:help}
    .kc .kt{font-size:8px;line-height:1}.kc .ks{font-size:5px;letter-spacing:-1px;line-height:1;opacity:.85}.kc .kv{font-size:7px;font-weight:700}
    .stTabs [data-baseweb="tab-list"]{gap:2px;background:#e2e8f0;padding:2px;border-radius:6px;margin-bottom:3px}
    .stTabs [data-baseweb="tab"]{border-radius:5px;padding:5px 10px;font-weight:600;font-size:10px}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--p)!important;box-shadow:0 2px 5px rgba(0,0,0,.07)}
    .sr{display:flex;align-items:center;padding:4px 8px;background:#fff;border-radius:5px;margin-bottom:1px;border:1px solid var(--b);font-size:9px}
    .sr .sn{font-weight:700;color:var(--p);min-width:200px;font-size:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .sc{padding:2px 7px;border-radius:12px;font-weight:800;font-size:10px;min-width:40px;text-align:center;margin:0 6px;color:#fff}
    .sr .sa{color:#718096;font-size:8px;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .stg{font-size:7px;color:#718096;min-width:50px;text-align:center;white-space:nowrap}
    .sr .sb{font-size:7px;font-weight:700;padding:1px 5px;border-radius:3px;white-space:nowrap}
    .ca{background:#fff;border-radius:var(--r);padding:8px;margin-top:2px;border:1px solid var(--b);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .ca .ct{font-size:10px;font-weight:700;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid var(--b)}
    .car{display:flex;align-items:center;margin-bottom:3px;font-size:8px}
    .car:last-child{margin-bottom:0}
    .car .cal{width:160px;font-weight:600;color:var(--p);text-align:right;padding-right:6px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .car .cab{flex:1;height:22px;background:#edf2f7;border-radius:4px;overflow:hidden}
    .car .caf{height:100%;border-radius:4px;transition:width .3s}
    .car .cav-out{font-size:8px;font-weight:800;color:#1a202c;min-width:50px;text-align:right;padding-left:4px}
    .gbr{display:flex;align-items:center;padding:2px 0;font-size:8px;border-bottom:1px solid #f7fafc}
    .gbr:last-child{border:none}
    .gbr-l{width:140px;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:7px}
    .gbr-g{display:flex;align-items:center;gap:3px;flex:1}
    .gbr-w{flex:1;height:18px;background:#edf2f7;border-radius:3px;overflow:hidden}
    .gbr-f{height:100%;border-radius:3px}
    .gb-p{background:linear-gradient(90deg,#2b6cb0,#4299e1)}
    .gb-q{background:linear-gradient(90deg,#276749,#48bb78)}
    .gbr-v{font-size:7px;font-weight:800;min-width:42px;text-align:right;color:#1a202c}
    .gbr-legend{display:flex;gap:12px;margin-bottom:4px;font-size:8px;font-weight:700}
    .gbr-legend span{display:flex;align-items:center;gap:4px}
    .gbr-legend i{display:inline-block;width:12px;height:12px;border-radius:2px}
    .cg{display:grid;grid-template-columns:1fr 1fr;gap:4px}
    .cg>div{background:#fff;border-radius:var(--r);padding:6px 8px;border:1px solid var(--b)}
    .cg .ct{font-size:9px;font-weight:700;margin-bottom:2px;padding-bottom:2px;border-bottom:1px solid var(--b)}
    .cgr{display:flex;align-items:center;padding:2px 0;font-size:8px;border-bottom:1px solid #f7fafc}
    .cgr:last-child{border:none}
    .cgr .rk{width:14px;font-weight:800;text-align:center}
    .cgr .pn{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cgr .ps{font-weight:800;min-width:45px;text-align:right}
    .es{text-align:center;padding:10px;color:#718096;font-size:10px}
    .rec-card{background:#fff;border-radius:8px;padding:8px 10px;margin-bottom:4px;border-left:4px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,.04);font-size:8px}
    .rec-card.critique{border-left-color:#e53e3e;background:#fff5f5}
    .rec-card.elevee{border-left-color:#ed8936;background:#fffaf0}
    .rec-card.moyenne{border-left-color:#ecc94b;background:#fffff0}
    .rec-card.faible{border-left-color:#48bb78;background:#f0fff4}
    .rec-card .rec-hdr{display:flex;align-items:center;gap:6px;margin-bottom:3px}
    .rec-card .rec-badge{padding:1px 6px;border-radius:10px;font-size:7px;font-weight:800;color:#fff}
    .rec-card .rec-poste{font-weight:700;color:var(--p);font-size:8px}
    .rec-card .rec-kpi{color:#718096;font-size:7px}
    .rec-card .rec-msg{color:#2d3748;font-size:8px;line-height:1.4;margin-top:2px}
    .rec-card .rec-meta{display:flex;gap:8px;margin-top:3px;color:#a0aec0;font-size:7px}
    .find-card{background:#fff;border-radius:6px;padding:6px 8px;margin-bottom:3px;border:1px solid var(--b);font-size:8px}
    .find-card .find-type{font-weight:700;font-size:7px;color:#805ad5;text-transform:uppercase;letter-spacing:.5px}
    .find-card .find-msg{font-weight:600;color:#1a202c;margin:1px 0}
    .find-card .find-det{color:#718096;font-size:7px;line-height:1.3}
    ::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:2px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:9px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:2px 6px;margin-bottom:2px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.cg{grid-template-columns:1fr}}
    </style>""", unsafe_allow_html=True)

# ===================== RECOMMENDATION ENGINE =====================

def get_kpi_rec(kpi, poste, val, target, delta, worsening):
    msgs = {
        "TAUX_REALISATION_CORRECTIF/PT": "Identifier les OT correctifs anciens, eliminer les blocages (pieces, ressources, permis) et accelerer les clotures.",
        "OT préparation <1 mois": "Renforcer l'equipe de preparation, prioriser les OT recemment crees et reduire les delais d'analyse technique.",
        "OT préparation >3 mois": "Organiser des sessions de rattrapage pour les OT anciens. Revoir l'affectation des prepareurs.",
        "OT planification <1 mois": "Ameliorer la reactivite de la planification et la disponibilite des ressources.",
        "OT planification >3 mois": "Revoir le processus de planification, les delais d'approvisionnement et les dependances entre OT.",
        "OT exécution <1 mois": "Optimiser l'affectation des equipes d'execution et lever les obstacles terrain.",
        "OT exécution >3 mois": "Identifier les OT bloques en execution, resoudre les problemes d'approvisionnement ou de permis.",
        "OT préparation 1mois< <3mois": "Reducer cette tranche en accelerant le traitement.",
        "OT planification 1mois< <3mois": "Accelerer la planification de ces OT.",
        "OT exécution 1mois< <3mois": "Suivre et accelerer ces OT en cours.",
        "appel avis approuvé": "Creer des OT pour les avis sans ordre associe et relancer le circuit d'approbation.",
        "OT LANC ESTIME": "Exiger l'estimation des couts budgetes avant tout lancement d'OT.",
        "Backlog préparation caractérisé": "Caracteriser immediatement tous les OT en attente de preparation.",
        "Backlog planification caractérisé": "Caracteriser tous les OT avant planification.",
        "OT CONFIME": "Confirmer systematiquement les OT termines dans le systeme SAP.",
        "OT_COR_EGAL": "Rapprocher les couts et corriger les ecarts significatifs."
    }
    base = msgs.get(kpi, "Analyser et corriger.")
    trend_txt = " Tendance a la degradation (Delta=%+.1f)." % delta if worsening else ""
    return "Au poste %s, %s est a %.1f%% (cible %s%%).%s %s" % (poste, kpi, val, target, trend_txt, base)

def build_evolution_table(ckdf, kpi_history, all_kpis, cible, vp):
    rows = []
    for poste in vp:
        if poste not in ckdf.index:
            continue
        for kpi in all_kpis:
            if kpi not in ckdf.columns:
                continue
            val = ckdf.loc[poste, kpi]
            if pd.isna(val):
                continue
            current = round(float(val), 1)
            key = f"{poste}__{kpi}"
            hist = kpi_history.get(key, [])
            lb = is_lb(kpi)
            target = cible.get(kpi, 100)
            if len(hist) >= 2:
                prev = round(hist[-2]["value"], 1)
                delta = round(current - prev, 1)
                if abs(delta) <= 0.5:
                    trend, ticon, tcol = "Stable", "\u27a1\ufe0f", "#718096"
                elif (delta > 0.5 and not lb) or (delta < -0.5 and lb):
                    trend, ticon, tcol = "Amelioration", "\U0001f4c8", "#38a169"
                else:
                    trend, ticon, tcol = "Degradation", "\U0001f4c9", "#e53e3e"
            else:
                prev = None
                delta = None
                trend, ticon, tcol = "Nouvelle", "\U0001f195", "#805ad5"
            met = (current <= target) if lb else (current >= target)
            rows.append({
                "Poste de Travail": poste, "KPI": kpi, "Valeur Actuelle": current,
                "Valeur Precedente": prev if prev is not None else "N/A",
                "Delta": delta if delta is not None else "N/A",
                "Delta_Raw": delta, "Tendance": "%s %s" % (ticon, trend),
                "Tendance_Color": tcol, "Cible": target, "Atteint": met, "LB": lb
            })
    rows.sort(key=lambda r: (r["Delta_Raw"] is None, r["Delta_Raw"] if r["Delta_Raw"] is not None else 0), reverse=True)
    return rows

def generate_recommendations(ckdf, cible, vp, kpi_history, all_kpis):
    recs = []
    for poste in vp:
        if poste not in ckdf.index:
            continue
        row = ckdf.loc[poste]
        for kpi in all_kpis:
            if kpi not in ckdf.columns:
                continue
            val = row[kpi]
            if pd.isna(val):
                continue
            target = cible.get(kpi, 100)
            lb = is_lb(kpi)
            met = val <= target if lb else val >= target
            if met:
                continue
            key = f"{poste}__{kpi}"
            hist = kpi_history.get(key, [])
            delta = None
            worsening = False
            if len(hist) >= 2:
                delta = round(hist[-1]["value"] - hist[-2]["value"], 1)
                if lb and delta > 0.5:
                    worsening = True
                if not lb and delta < -0.5:
                    worsening = True
            gap = round(abs(val - target), 1)
            if worsening:
                prio, pcol, psort = "Critique", "#e53e3e", 0
            elif gap >= 30:
                prio, pcol, psort = "Critique", "#e53e3e", 0
            elif gap >= 15:
                prio, pcol, psort = "Elevee", "#ed8936", 1
            elif gap >= 5:
                prio, pcol, psort = "Moyenne", "#ecc94b", 2
            else:
                prio, pcol, psort = "Faible", "#48bb78", 3
            msg = get_kpi_rec(kpi, poste, val, target, delta, worsening)
            cat = "Performance" if kpi in [
                "TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois",
                "OT préparation 1mois< <3mois", "OT planification <1 mois", "OT planification >3 mois",
                "OT planification 1mois< <3mois", "OT exécution <1 mois", "OT exécution >3 mois",
                "OT exécution 1mois< <3mois"
            ] else "Qualite"
            recs.append({
                "Priorite": prio, "Priorite_Color": pcol, "Priorite_Sort": psort,
                "Categorie": cat, "Poste": poste, "KPI": kpi,
                "Actuel": round(val, 1), "Cible": target, "Ecart": gap,
                "Delta": delta if delta is not None else "N/A", "Recommandation": msg
            })
    recs.sort(key=lambda r: (r["Priorite_Sort"], -r["Ecart"]))
    return recs

def analyze_backlog_auto(dfp, avf, vp):
    findings = []
    total_ot = len(dfp)
    if total_ot == 0:
        findings.append({"Type": "Information", "Priorite": "Info",
            "Message": "Aucun OT dans les filtres selectionnes",
            "Detail": "Modifier les filtres pour obtenir des resultats."})
        return findings
    if "Statut OT" in dfp.columns:
        sc = dfp["Statut OT"].value_counts()
        for st, cnt in sc.items():
            pct = round(cnt / total_ot * 100, 1)
            findings.append({"Type": "OT - Repartition", "Priorite": "Info",
                "Message": "%s: %d OT (%.1f%%)" % (st, cnt, pct),
                "Detail": "Sur %d OT, %d en statut %s." % (total_ot, cnt, st)})
        for st in ["CRÉÉ", "LANC"]:
            if st in sc and sc[st] / total_ot * 100 > 60:
                findings.append({"Type": "OT - Concentration", "Priorite": "Critique",
                    "Message": "Concentration anormale en %s: %.1f%%" % (st, sc[st] / total_ot * 100),
                    "Detail": "Plus de 60%% des OT en statut %s. Goulot d'etranglement detecte." % st})
    for phase, col_a, sf in [
        ("Preparation", "ap", dfp["Statut OT"] == "CRÉÉ"),
        ("Planification", "alp", (dfp["Statut OT"] == "LANC") & (dfp["Contient SOPL"] == 0)),
        ("Execution", "aex", (dfp["Statut OT"] == "LANC") & (dfp["Contient SOPL"] == 1))
    ]:
        if col_a not in dfp.columns:
            continue
        sub = dfp[sf]
        old = sub[sub[col_a] == ">3 mois"]
        if len(old) > 0:
            bp = old.groupby("Poste travail princ.").size().sort_values(ascending=False).head(5)
            pl = ", ".join(["%s (%d)" % (p, n) for p, n in bp.items()])
            pr = "Critique" if len(old) > 10 else "Elevee"
            findings.append({"Type": "OT - %s >3 mois" % phase, "Priorite": pr,
                "Message": "%d OT en %s depuis >3 mois" % (len(old), phase.lower()),
                "Detail": "Postes: %s. Traitement prioritaire requis." % pl})
    for phase, col_c, sf in [
        ("Preparation", "Backlog preparation", dfp["Statut OT"] == "CRÉÉ"),
        ("Planification", "Backlog planification", (dfp["Statut OT"] == "LANC") & (dfp["Contient SOPL"] == 0))
    ]:
        if col_c not in dfp.columns:
            continue
        sub = dfp[sf]
        nc = sub[sub[col_c] == "NON CARACTERISE"]
        if len(nc) > 0:
            bp = nc.groupby("Poste travail princ.").size().sort_values(ascending=False).head(5)
            pl = ", ".join(["%s (%d)" % (p, n) for p, n in bp.items()])
            pn = round(len(nc) / len(sub) * 100, 1) if len(sub) > 0 else 0
            pr = "Critique" if pn > 30 else ("Elevee" if pn > 10 else "Moyenne")
            findings.append({"Type": "OT - %s Non Caracterise" % phase, "Priorite": pr,
                "Message": "%d OT non caracterises en %s (%.1f%%)" % (len(nc), phase.lower(), pn),
                "Detail": "Postes: %s. Caracteriser pour debloquer." % pl})
    if avf is not None and not avf.empty:
        ta = len(avf)
        findings.append({"Type": "AVIS - Volume", "Priorite": "Info",
            "Message": "%d avis sans OT associe" % ta,
            "Detail": "Ces avis necessitent la creation d'OT."})
        if "Statut utilisateur" in avf.columns:
            as_ = avf["Statut utilisateur"].value_counts()
            if "REJT" in as_ and as_["REJT"] > 0:
                findings.append({"Type": "AVIS - Rejetes", "Priorite": "Elevee",
                    "Message": "%d avis rejetes (REJT)" % as_["REJT"],
                    "Detail": "Analyser les causes de rejet et corriger."})
    sopl = dfp[dfp["Statut système"].str.contains("SOPL", na=False)]
    if not sopl.empty:
        cl = sopl[sopl["Statut système"].str.contains("CLOT|TCLO", na=False)]
        cr = round(len(cl) / len(sopl) * 100, 1)
        if cr < 70:
            findings.append({"Type": "OT - Execution SOPL", "Priorite": "Critique",
                "Message": "Taux cloture SOPL: %.1f%% (%d/%d)" % (cr, len(cl), len(sopl)),
                "Detail": "Taux critique. Identifier les OT bloques."})
        elif cr < 85:
            findings.append({"Type": "OT - Execution SOPL", "Priorite": "Moyenne",
                "Message": "Taux cloture SOPL: %.1f%% (%d/%d)" % (cr, len(cl), len(sopl)),
                "Detail": "Sous la cible. Accelerer le traitement."})
    return findings

def html_evolution_table(rows):
    cols = ["Poste de Travail", "KPI", "Valeur Actuelle", "Valeur Precedente", "Delta", "Tendance", "Cible", "Statut"]
    h = '<table class="tw evt"><thead><tr>'
    for c in cols:
        h += '<th>%s</th>' % c
    h += '</tr></thead><tbody>'
    for r in rows:
        if r["Delta"] != "N/A":
            d = r["Delta"]
            ds = "%+.1f" % d
            dc = "#38a169" if ((d > 0.5 and not r["LB"]) or (d < -0.5 and r["LB"])) else (
                "#e53e3e" if ((d < -0.5 and not r["LB"]) or (d > 0.5 and r["LB"])) else "#718096")
        else:
            ds = "N/A"
            dc = "#a0aec0"
        stat_s = "ATTEINT" if r["Atteint"] else "NON ATTEINT"
        stat_c = "#c6efce" if r["Atteint"] else "#ffc7ce"
        h += '<tr><td>%s</td><td>%s</td><td style="font-weight:700">%s</td><td>%s</td>' % (r["Poste de Travail"], r["KPI"], r["Valeur Actuelle"], r["Valeur Precedente"])
        h += '<td style="color:%s;font-weight:700">%s</td>' % (dc, ds)
        h += '<td style="color:%s;font-weight:600">%s</td>' % (r["Tendance_Color"], r["Tendance"])
        h += '<td>%s</td>' % r["Cible"]
        h += '<td style="background:%s;color:#1a202c;font-weight:700;text-align:center">%s</td></tr>' % (stat_c, stat_s)
    h += '</tbody></table>'
    return h

def html_rec_cards(recs):
    if not recs:
        return '<div class="es">Aucune recommandation - Tous les KPIs sont atteints</div>'
    h = ''
    cls_map = {0: "critique", 1: "elevee", 2: "moyenne", 3: "faible"}
    for r in recs:
        cc = cls_map[r["Priorite_Sort"]]
        cat_bg = "#ebf8ff" if r["Categorie"] == "Performance" else "#f0fff4"
        cat_fg = "#2b6cb0" if r["Categorie"] == "Performance" else "#276749"
        h += '<div class="rec-card %s">' % cc
        h += '<div class="rec-hdr">'
        h += '<span class="rec-badge" style="background:%s">%s</span>' % (r["Priorite_Color"], r["Priorite"])
        h += '<span class="rec-badge" style="background:%s;color:%s">%s</span>' % (cat_bg, cat_fg, r["Categorie"])
        h += '<span class="rec-poste">%s</span>' % r["Poste"]
        h += '<span class="rec-kpi">%s</span>' % r["KPI"]
        h += '</div>'
        h += '<div class="rec-msg">%s</div>' % r["Recommandation"]
        h += '<div class="rec-meta">Actuel: %s | Cible: %s | Ecart: %s | Delta: %s</div>' % (r["Actuel"], r["Cible"], r["Ecart"], r["Delta"])
        h += '</div>'
    return h

def html_findings(findings):
    if not findings:
        return '<div class="es">Aucun finding</div>'
    h = ''
    pcol = {"Critique": "#e53e3e", "Elevee": "#ed8936", "Moyenne": "#ecc94b", "Info": "#718096"}
    for f in findings:
        bc = pcol.get(f["Priorite"], "#718096")
        h += '<div class="find-card" style="border-left:3px solid %s">' % bc
        h += '<div class="find-type" style="color:%s">%s - %s</div>' % (bc, f["Type"], f["Priorite"])
        h += '<div class="find-msg">%s</div>' % f["Message"]
        h += '<div class="find-det">%s</div>' % f["Detail"]
        h += '</div>'
    return h

def create_analyse_excel(prep_tbl, plan_tbl, exec_tbl, evol_df, recs_df, findings_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not prep_tbl.empty:
            prep_tbl.to_excel(writer, sheet_name='Backlog Preparation', index=False)
        if not plan_tbl.empty:
            plan_tbl.to_excel(writer, sheet_name='Backlog Planification', index=False)
        if not exec_tbl.empty:
            exec_tbl.to_excel(writer, sheet_name='Backlog Execution', index=False)
        if evol_df is not None and not evol_df.empty:
            evol_df.to_excel(writer, sheet_name='Analyse Evolutive', index=False)
        if recs_df is not None and not recs_df.empty:
            recs_df.to_excel(writer, sheet_name='Recommandations', index=False)
        if findings_df is not None and not findings_df.empty:
            findings_df.to_excel(writer, sheet_name='Analyse Backlog', index=False)
        hdr_fill = PatternFill(start_color="1e3a5f", end_color="1e3a5f", fill_type="solid")
        hdr_font = Font(bold=True, color="FFFFFF", size=10)
        tot_fill = PatternFill(start_color="e2e8f0", end_color="e2e8f0", fill_type="solid")
        tot_font = Font(bold=True, size=10)
        thin = Border(
            left=Side(style='thin', color='d0d0d0'), right=Side(style='thin', color='d0d0d0'),
            top=Side(style='thin', color='d0d0d0'), bottom=Side(style='thin', color='d0d0d0'))
        for sn in writer.sheets:
            ws = writer.sheets[sn]
            for cell in ws[1]:
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal='center')
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = thin
                    cell.alignment = Alignment(horizontal='center', wrap_text=True)
            last_r = ws.max_row
            if last_r > 1:
                for cell in ws[last_r]:
                    cell.fill = tot_fill
                    cell.font = tot_font
            for col in ws.columns:
                mx = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(mx + 4, 50)
    output.seek(0)
    return output

# ===================== HELPERS =====================

def contient_mot(t, lm):
    t = str(t)
    return any(m in t for l in lm for m in l.split())

def cat_age(a):
    if pd.isna(a):
        return "Inconnu"
    if a <= 1:
        return "<1 mois"
    elif a >= 3:
        return ">3 mois"
    return "1 mois < <3 mois"

def ckpi(n, d, sz=100):
    return np.where(d == 0, sz, (n / d) * 100)

def cpiv(df, f, c, p):
    return pd.pivot_table(df[f], index="Poste travail princ.", columns=c, values="Ordre", aggfunc="count", fill_value=0).reindex(p, fill_value=0)

def excr(df):
    if "Poste travail princ." in df.columns:
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)].copy()
    return df

def get_atelier(p):
    p = str(p).upper()
    if "PS" in p: return "Sulfurique"
    if "PP" in p: return "Phosphorique"
    if "TSP" in p or "REX" in p: return "Engrais"
    if "MCP" in p or "DCP" in p: return "Feed"
    return "Centrale et Utilites"

def is_lb(k):
    return k in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois",
                 "OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]

def ks(v, c):
    try:
        val = float(v)
    except:
        return ""
    if c in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
        return "background:#c6efce;color:#006100;font-weight:600" if val >= 80 else (
            "background:#ffeb9c;color:#9c6500;font-weight:600" if val >= 75 else "background:#ffc7ce;color:#9c0006;font-weight:600")
    if c in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
        return "background:#c6efce;color:#006100;font-weight:600" if val <= 15 else "background:#ffc7ce;color:#9c0006;font-weight:600"
    if c in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
        return "background:#c6efce;color:#006100;font-weight:600" if val <= 5 else "background:#ffc7ce;color:#9c0006;font-weight:600"
    if c == "TAUX_REALISATION_CORRECTIF/PT":
        return "background:#c6efce;color:#006100;font-weight:600" if val >= 85 else (
            "background:#ffeb9c;color:#9c6500;font-weight:600" if val >= 80 else "background:#ffc7ce;color:#9c0006;font-weight:600")
    if c == "appel avis approuvé":
        return "background:#c6efce;color:#006100;font-weight:600" if val >= 95 else (
            "background:#ffeb9c;color:#9c6500;font-weight:600" if val >= 90 else "background:#ffc7ce;color:#9c0006;font-weight:600")
    if c in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
        return "background:#c6efce;color:#006100;font-weight:600" if val >= 100 else (
            "background:#ffeb9c;color:#9c6500;font-weight:600" if val >= 95 else "background:#ffc7ce;color:#9c0006;font-weight:600")
    return ""

def cs(v):
    try:
        val = float(str(v).replace(' %', '').strip())
    except:
        return ""
    return "background:#c6efce;color:#006100;font-weight:700" if val >= 90 else (
        "background:#ffeb9c;color:#9c6500;font-weight:700" if val >= 80 else "background:#ffc7ce;color:#9c0006;font-weight:700")

def kas(v):
    try:
        val = int(v)
    except:
        return ""
    if val == 0: return "color:#cbd5e0"
    if val <= 3: return "background:#ffeb9c;color:#9c6500;font-weight:600"
    if val <= 10: return "background:#fed7d7;color:#c53030;font-weight:600"
    return "background:#fc8181;color:#742a2a;font-weight:800"

def gscore(k, a, t):
    if pd.isna(a) or pd.isna(t):
        return 0
    if k in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
        return 1 if a >= 75 else 0
    if k in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
        return 1 if a <= 15 else 0
    if k in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
        return 1 if a <= 5 else 0
    if k == "TAUX_REALISATION_CORRECTIF/PT":
        return 1 if a >= 80 else 0
    if k == "appel avis approuvé":
        return 1 if a >= 90 else 0
    if k in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
        return 1 if a >= 95 else 0
    return 0

def calc_kpis(df_i, av_i, now, posts):
    res = {}
    df = df_i.copy()
    av = av_i.copy()
    mp = ["CRPR ATPD", "CRPR ATMR", "CRPR ATER", "CRPR ATRS", "CRPR ATMO", "ATPD", "ATMR", "ATER", "ATRS", "ATMO"]
    df["Backlog preparation"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mp)), "CARACTERISE", "NON CARACTERISE")
    mplan = ["ATPL ATEI", "ATPL ATAL", "ATPL ATER", "ATPL AGAR", "ATPL ATHS", "ATEI", "ATAL", "ATAS", "AGAR", "ATHS"]
    df["Backlog planification"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mplan)), "CARACTERISE", "NON CARACTERISE")
    for dc, am, ac in [('Créé le', "amp", "ap"), ('Date de début planifiée', "amlp", "alp"), ('Date de début planifiée', "amex", "aex")]:
        if dc in df.columns:
            df[dc] = pd.to_datetime(df[dc], errors='coerce')
            df[am] = ((now.year - df[dc].dt.year) * 12 + (now.month - df[dc].dt.month)).round(2)
            df[ac] = df[am].apply(cat_age)
        else:
            df[am] = np.nan
            df[ac] = "Inconnu"
    df["OT CONFIME"] = np.where(df["Statut système"].str.contains("CLO", na=False) & df["Statut système"].str.contains("CONF", na=False), "OUI", "NON")
    df["Contient SOPL"] = df["Statut utilisateur"].str.contains("SOPL", na=False).map({True: 1, False: 0})
    if "Total coûts budgétés" in df.columns:
        df["OT LANC ESTIME"] = np.where(df["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
    else:
        df["OT LANC ESTIME"] = "NON"
    if "Total coûts budgétés" in df.columns and "Total coûts réels" in df.columns:
        df["OT_COR_EGAL"] = np.where((df["Total coûts budgétés"].fillna(0) - df["Total coûts réels"].fillna(0)) == 0, "OUI", "NON")
    else:
        df["OT_COR_EGAL"] = "NON"
    res['dfp'] = df
    an = cpiv(df, df["Nº appel pl.entret."].fillna(0) == 0, "Statut OT", posts)
    for c in ["CLOT", "CRÉÉ", "LANC", "TCLO"]:
        an[c] = an.get(c, 0)
    an["Total"] = an[["CLOT", "CRÉÉ", "LANC", "TCLO"]].sum(axis=1)
    an["TAUX_REALISATION_CORRECTIF/PT"] = ckpi(an["TCLO"], an["Total"])
    pr = cpiv(df, df["Statut OT"] == "CRÉÉ", "ap", posts)
    for c in ["<1 mois", ">3 mois", "1 mois < <3 mois"]:
        pr[c] = pr.get(c, 0)
    pr["Total"] = pr[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    pr["OT préparation <1 mois"] = ckpi(pr["<1 mois"], pr["Total"])
    pr["OT préparation >3 mois"] = ckpi(pr[">3 mois"], pr["Total"], 0)
    pr["OT préparation 1mois< <3mois"] = ckpi(pr["1 mois < <3 mois"], pr["Total"], 0)
    pl = cpiv(df, (df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 0), "alp", posts)
    for c in ["<1 mois", ">3 mois", "1 mois < <3 mois"]:
        pl[c] = pl.get(c, 0)
    pl["Total"] = pl[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    pl["OT planification <1 mois"] = ckpi(pl["<1 mois"], pl["Total"])
    pl["OT planification >3 mois"] = ckpi(pl[">3 mois"], pl["Total"], 0)
    pl["OT planification 1mois< <3mois"] = ckpi(pl["1 mois < <3 mois"], pl["Total"], 0)
    ex = cpiv(df, (df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 1), "aex", posts)
    for c in ["<1 mois", ">3 mois", "1 mois < <3 mois"]:
        ex[c] = ex.get(c, 0)
    ex["Total"] = ex[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    ex["OT exécution <1 mois"] = ckpi(ex["<1 mois"], ex["Total"])
    ex["OT exécution >3 mois"] = ckpi(ex[">3 mois"], ex["Total"], 0)
    ex["OT exécution 1mois< <3mois"] = ckpi(ex["1 mois < <3 mois"], ex["Total"], 0)
    la = pd.pivot_table(df[df["Statut OT"] == "LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["OUI", "NON"]:
        la[c] = la.get(c, 0)
    la["Total"] = la["OUI"] + la["NON"]
    la["OT LANC ESTIME"] = ckpi(la["OUI"], la["Total"])
    pc = pd.pivot_table(df[df["Statut OT"] == "CRÉÉ"], index="Poste travail princ.", columns="Backlog preparation", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["CARACTERISE", "NON CARACTERISE"]:
        pc[c] = pc.get(c, 0)
    pc["Total"] = pc["CARACTERISE"] + pc["NON CARACTERISE"]
    pc["Backlog préparation caractérisé"] = ckpi(pc["CARACTERISE"], pc["Total"])
    plc = pd.pivot_table(df[df["Statut OT"] == "LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["CARACTERISE", "NON CARACTERISE"]:
        plc[c] = plc.get(c, 0)
    plc["Total"] = plc["CARACTERISE"] + plc["NON CARACTERISE"]
    plc["Backlog planification caractérisé"] = ckpi(plc["CARACTERISE"], plc["Total"])
    for kn, cn in [("OT CONFIME", "OT CONFIME"), ("OT_COR_EGAL", "OT_COR_EGAL")]:
        pv = pd.pivot_table(df, index="Poste travail princ.", columns=cn, values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["OUI", "NON"]:
            pv[c] = pv.get(c, 0)
        pv["Total"] = pv["OUI"] + pv["NON"]
        pv[cn] = ckpi(pv["OUI"], pv["Total"])
        res[kn.lower().replace(" ", "_")] = pv
    avf = av[(av["Ordre"].isna()) | (av["Ordre"].astype(str).str.strip() == "")].copy()
    res['avf'] = avf
    tca = pd.pivot_table(avf, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["APRQ", "APRV", "APRV AVAU", "REJT"]:
        tca[c] = tca.get(c, 0)
    tca["Total"] = tca[["APRQ", "APRV", "APRV AVAU", "REJT"]].sum(axis=1)
    tca["appel avis approuvé"] = ckpi(tca["APRV"], tca["Total"])
    res['ckdf'] = pd.DataFrame({
        "TAUX_REALISATION_CORRECTIF/PT": an["TAUX_REALISATION_CORRECTIF/PT"],
        "OT préparation <1 mois": pr["OT préparation <1 mois"],
        "OT préparation >3 mois": pr["OT préparation >3 mois"],
        "OT préparation 1mois< <3mois": pr["OT préparation 1mois< <3mois"],
        "OT planification <1 mois": pl["OT planification <1 mois"],
        "OT planification >3 mois": pl["OT planification >3 mois"],
        "OT planification 1mois< <3mois": pl["OT planification 1mois< <3mois"],
        "OT exécution <1 mois": ex["OT exécution <1 mois"],
        "OT exécution >3 mois": ex["OT exécution >3 mois"],
        "OT exécution 1mois< <3mois": ex["OT exécution 1mois< <3mois"],
        "appel avis approuvé": tca["appel avis approuvé"],
        "OT LANC ESTIME": la["OT LANC ESTIME"],
        "Backlog préparation caractérisé": pc["Backlog préparation caractérisé"],
        "Backlog planification caractérisé": plc["Backlog planification caractérisé"],
        "OT CONFIME": res['ot_confime']["OT CONFIME"],
        "OT_COR_EGAL": res['ot_cor_egal']["OT_COR_EGAL"]
    })
    return res

def html_synth(kpi_list, actuals, targets, act_map, accent):
    h = ''
    for k in kpi_list:
        av = actuals.get(k, 0)
        tv = targets.get(k, 100)
        met = av <= tv if is_lb(k) else av >= tv
        sbg = "#c6efce" if met else "#ffc7ce"
        sclr = "#006100" if met else "#9c0006"
        scbg = accent if met else "#e53e3e"
        act = "Objectif atteint" if met else act_map.get(k, "")
        h += '<div class="sr"><div class="sn">%s</div><div class="sc" style="background:%s">%.1f%%</div><div class="stg">Cible: %s%%</div><div class="sb" style="color:%s;background:%s">%s</div><div class="sa">%s</div></div>' % (k, scbg, av, tv, sclr, sbg, "ATTEINT" if met else "NON ATTEINT", act)
    return h

def html_classement(scores, accent, threshold=90):
    sp = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    met_p = [(p, s) for p, s in sp if s >= threshold]
    not_p = [(p, s) for p, s in sp if s < threshold]
    t5 = met_p[:5]
    b5 = not_p[-5:] if len(not_p) > 5 else not_p
    h = '<div class="cg"><div><div class="ct" style="color:#38a169">Top 5 - Atteint (>=%.0f%%)</div>' % threshold
    if t5:
        for i, (p, s) in enumerate(t5):
            h += '<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>' % (accent, i + 1, p, cs("%.2f" % s), s)
    else:
        h += '<div style="padding:4px;font-size:8px;color:#718096">Aucun</div>'
    h += '</div><div><div class="ct" style="color:#e53e3e">Bottom 5 - Non Atteint (<%.0f%%)</div>' % threshold
    if b5:
        for i, (p, s) in enumerate(reversed(b5)):
            h += '<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>' % (len(b5) - i, p, cs("%.2f" % s), s)
    else:
        h += '<div style="padding:4px;font-size:8px;color:#38a169">Tous atteints</div>'
    h += '</div></div>'
    return h

def html_kpi_bars(kpi_list, actuals, targets, title, color_ok, color_fail, kpi_history=None):
    h = '<div class="ca"><div class="ct" style="color:%s">%s</div>' % (color_ok, title)
    for k in kpi_list:
        av = actuals.get(k, 0)
        tv = targets.get(k, 100)
        met = av <= tv if is_lb(k) else av >= tv
        bw = min(max(av, 0), 100)
        bg = color_ok if met else color_fail
        sp = ""
        if kpi_history:
            hist = kpi_history.get("__total__%s" % k, [])
            if len(hist) >= 2:
                sp = build_kpi_cell("", hist, is_lb(k))
        h += '<div class="car"><div class="cal" style="width:240px">%s %s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>' % (k, sp, bw, bg, av)
    h += '</div>'
    return h

def html_grouped_bars(posts, pscores, qscores, title, threshold=90):
    h = '<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>' % title
    h += '<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span><span><i style="background:#e53e3e;width:2px;height:18px"></i> Cible %.0f%%</span></div>' % threshold
    sp2 = sorted(posts, key=lambda x: (pscores.get(x, 0) + qscores.get(x, 0)) / 2, reverse=True)
    for p in sp2:
        pv = pscores.get(p, 0)
        qv = qscores.get(p, 0)
        pw = min(max(pv, 0), 100)
        qw = min(max(qv, 0), 100)
        h += '<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w" style="position:relative"><div class="gbr-f gb-p" style="width:%s%%"></div><div style="position:absolute;left:%s%%;top:-2px;bottom:-2px;width:2px;background:#e53e3e"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w" style="position:relative"><div class="gbr-f gb-q" style="width:%s%%"></div><div style="position:absolute;left:%s%%;top:-2px;bottom:-2px;width:2px;background:#e53e3e"></div></div><div class="gbr-v">%.1f%%</div></div></div>' % (p, pw, threshold, pv, qw, threshold, qv)
    h += '</div>'
    return h

# ===================== MAIN =====================

def main():
    try:
        locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_ALL, 'fr_FR')
        except:
            pass
    inject_custom_css()

    consignes = [
        "Port obligatoire des EPI avant toute intervention.",
        "Port obligatoire du casque de securite.",
        "Port obligatoire des lunettes de protection.",
        "Ne jamais intervenir sur un equipement en marche.",
        "Baliser et securiser la zone de travail.",
        "Respecter la procedure de consignation et deconsignation.",
        "Verifier l'absence de tension avant toute intervention electrique.",
        "Utiliser les protections auditives dans les zones bruyantes.",
        "Ne jamais contourner une procedure HSE.",
        "Aucun travail n'est plus urgent que la securite."
    ]

    if "hse_affiche" not in st.session_state:
        st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        c = random.choice(consignes)
        st.markdown('''<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:42px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SECURITE</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:18px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Securite - Sante - Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:28px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:32px;font-weight:900">Aucun travail n'est plus urgent que la securite</h2>
        </div>''' % c, unsafe_allow_html=True)
        time.sleep(6)
        st.session_state.hse_affiche = True
        st.rerun()
        st.stop()

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown('<div style="padding:10px 0 4px 0"><div style="font-size:18px;margin-bottom:2px">⚙️</div><div style="font-size:12px;font-weight:800;color:white">Filtres & Parametres</div></div>', unsafe_allow_html=True)
        st.markdown("---")
        
        unf = st.toggle("📁 Charger nouveaux fichiers", value=False, key="tf")
        ot_f = None
        av_f = None
        apm = []
        df_dt = datetime.now().strftime("%d/%m/%Y")
        
        if unf:
            ot_f = st.file_uploader("Fichier OT", type=["xlsx"], key="uot")
            av_f = st.file_uploader("Fichier AVIS", type=["xlsx"], key="uav")
        else:
            if os.path.exists("ot.xlsx"):
                try:
                    df_dt = datetime.fromtimestamp(os.path.getmtime("ot.xlsx")).strftime("%d/%m/%Y")
                    _t = excr(pd.read_excel("ot.xlsx"))
                    if "Poste travail princ." in _t.columns:
                        apm = sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
                except Exception as e:
                    st.warning("Erreur lecture ot.xlsx: %s" % str(e))
            st.markdown('<div style="background:rgba(255,255,255,.1);padding:5px 8px;border-radius:6px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:7px;color:rgba(255,255,255,.5);text-transform:uppercase">Donnees</div><div style="font-size:10px;color:white;font-weight:600;margin-top:1px">📅 %s</div></div>' % df_dt, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("**🎯 Postes**")
        sp = st.multiselect("Poste", ["All"] + apm, ["All"], key="sp")
        st.markdown("**🏭 Atelier**")
        sa = st.multiselect("Atelier", ["All", "Sulfurique (PS)", "Phosphorique (PP)", "Engrais (TSP/REX)", "Feed (MCP/DCP)", "Centrale et Utilites"], ["All"], key="sa")
        st.markdown("**🏢 Division**")
        sd = st.multiselect("Division", ["All", "SF1", "SF2"], ["All"], key="sd")
        st.markdown("---")
        st.markdown("**📅 Periode**")
        dr = st.date_input("Date debut planifiee", value=(datetime(2025, 1, 1).date(), datetime.today().date()), format="DD/MM/YYYY", key="dr")

    # ===================== CHARGEMENT DONNEES =====================
    if unf:
        if ot_f is None or av_f is None:
            st.info("Veuillez charger les fichiers OT et AVIS.")
            st.stop()
    
    try:
        if unf:
            raw_ot = pd.read_excel(ot_f)
            raw_av = pd.read_excel(av_f)
            df_dt = datetime.now().strftime("%d/%m/%Y")
        else:
            if not os.path.exists("ot.xlsx") or not os.path.exists("avis.xlsx"):
                st.error("Fichiers ot.xlsx et avis.xlsx non trouves. Activez le toggle pour les charger.")
                st.stop()
            raw_ot = pd.read_excel("ot.xlsx")
            raw_av = pd.read_excel("avis.xlsx")
        
        raw_ot = excr(raw_ot)
        raw_av = excr(raw_av)
        
        # Conversion dates
        for c in ["Créé le", "Date de début planifiée", "Date de clôture", "Début réel", "Fin réelle"]:
            if c in raw_ot.columns:
                raw_ot[c] = pd.to_datetime(raw_ot[c], errors="coerce")
        for c in ["Créé le", "Début souhaité", "Date de la clôture"]:
            if c in raw_av.columns:
                raw_av[c] = pd.to_datetime(raw_av[c], errors="coerce")
        
        # Colonnes requises
        required_ot = ["Poste travail princ.", "Statut OT", "Statut système", "Statut utilisateur", "Ordre"]
        missing_ot = [c for c in required_ot if c not in raw_ot.columns]
        if missing_ot:
            st.error("Colonnes manquantes dans OT: %s" % ", ".join(missing_ot))
            st.stop()
        
        if "Poste travail princ." not in raw_av.columns:
            raw_av["Poste travail princ."] = ""
        
        # Postes disponibles
        if not apm:
            apm = sorted(raw_ot[raw_ot["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
        
        # Filtres
        if "All" in sp or not sp:
            sp_filtered = apm
        else:
            sp_filtered = sp
        
        if "All" in sa or not sa:
            sa_filtered = ["All"]
        else:
            sa_filtered = sa
        
        if "All" in sd or not sd:
            sd_filtered = ["All"]
        else:
            sd_filtered = sd
        
        sdt = pd.to_datetime(dr[0]) if len(dr) == 2 else pd.to_datetime(datetime(2025, 1, 1))
        edt = pd.to_datetime(dr[1]) if len(dr) == 2 else pd.to_datetime(datetime.today())

        def mf(poste):
            p = str(poste).upper()
            # Filtre atelier
            if "All" not in sa_filtered:
                m = False
                if "Sulfurique (PS)" in sa_filtered and "PS" in p: m = True
                if "Phosphorique (PP)" in sa_filtered and "PP" in p: m = True
                if "Engrais (TSP/REX)" in sa_filtered and ("TSP" in p or "REX" in p): m = True
                if "Feed (MCP/DCP)" in sa_filtered and ("MCP" in p or "DCP" in p): m = True
                if "Centrale et Utilites" in sa_filtered and not any(x in p for x in ["PS", "PP", "TSP", "REX", "MCP", "DCP"]): m = True
                if not m:
                    return False
            # Filtre division
            if "All" not in sd_filtered:
                d = False
                if "SF1" in sd_filtered and "SF1" in p: d = True
                if "SF2" in sd_filtered and "SF2" in p: d = True
                if not d:
                    return False
            # Filtre poste
            if sp_filtered and p not in [str(x).upper() for x in sp_filtered]:
                return False
            return True

        # Filtrage OT
        mask_ot = raw_ot["Poste travail princ."].apply(mf)
        if "Date de début planifiée" in raw_ot.columns:
            mask_ot = mask_ot & (raw_ot["Date de début planifiée"] >= sdt) & (raw_ot["Date de début planifiée"] <= edt)
        df_filtered = raw_ot[mask_ot].copy()
        
        # Filtrage AVIS
        mask_av = raw_av["Poste travail princ."].apply(mf) if "Poste travail princ." in raw_av.columns else pd.Series(True, index=raw_av.index)
        av_filtered = raw_av[mask_av].copy()
        
        # Postes visibles
        vp = sorted(df_filtered["Poste travail princ."].dropna().unique().tolist()) if not df_filtered.empty else []
        
        if not vp:
            st.warning("Aucun poste trouvé avec les filtres sélectionnés. Modifiez vos critères.")
            st.stop()
        
        # ===================== CALCUL KPIs =====================
        now = datetime.now()
        results = calc_kpis(df_filtered, av_filtered, now, vp)
        dfp = results['dfp']
        ckdf = results['ckdf']
        avf = results.get('avf', pd.DataFrame())
        
        # Définitions KPIs
        pk = [
            "TAUX_REALISATION_CORRECTIF/PT",
            "OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois",
            "OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois",
            "OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"
        ]
        qk = [
            "appel avis approuvé", "OT LANC ESTIME",
            "Backlog préparation caractérisé", "Backlog planification caractérisé",
            "OT CONFIME", "OT_COR_EGAL"
        ]
        all_kpis = pk + qk
        
        cible = {
            "TAUX_REALISATION_CORRECTIF/PT": 85,
            "OT préparation <1 mois": 80, "OT préparation >3 mois": 5, "OT préparation 1mois< <3mois": 15,
            "OT planification <1 mois": 80, "OT planification >3 mois": 5, "OT planification 1mois< <3mois": 15,
            "OT exécution <1 mois": 80, "OT exécution >3 mois": 5, "OT exécution 1mois< <3mois": 15,
            "appel avis approuvé": 95, "OT LANC ESTIME": 100,
            "Backlog préparation caractérisé": 100, "Backlog planification caractérisé": 100,
            "OT CONFIME": 100, "OT_COR_EGAL": 100
        }
        
        # Calcul scores
        pscores = {}
        for poste in vp:
            if poste in ckdf.index:
                pts = sum(gscore(k, ckdf.loc[poste, k], cible.get(k, 100)) for k in pk if k in ckdf.columns and pd.notna(ckdf.loc[poste, k]))
                pscores[poste] = (pts / len(pk)) * 100 if pk else 0
        
        qscores = {}
        for poste in vp:
            if poste in ckdf.index:
                pts = sum(gscore(k, ckdf.loc[poste, k], cible.get(k, 100)) for k in qk if k in ckdf.columns and pd.notna(ckdf.loc[poste, k]))
                qscores[poste] = (pts / len(qk)) * 100 if qk else 0
        
        # Moyennes globales
        pa = {}
        for k in pk:
            if k in ckdf.columns:
                pa[k] = round(float(ckdf[k].mean()), 2)
        qa = {}
        for k in qk:
            if k in ckdf.columns:
                qa[k] = round(float(ckdf[k].mean()), 2)
        
        p_global = round(float(np.mean(list(pscores.values()))) if pscores else 0, 1)
        q_global = round(float(np.mean(list(qscores.values()))) if qscores else 0, 1)
        
        # Sauvegarde historique
        kpi_history = save_current_kpis(ckdf, qk, pk, pscores, qscores, pa, qa)
        
        # ===================== AFFICHAGE =====================
        # En-tête
        st.markdown('<div class="mh"><h1>📊 TABLEAU DE BORD MAINTENANCE</h1><div class="db">📅 %s | %d postes | %d OT</div></div>' % (df_dt, len(vp), len(dfp)), unsafe_allow_html=True)
        
        # Cartes résumé
        nb_non_atteints = sum(1 for k in all_kpis if k in pa and ((is_lb(k) and pa[k] > cible.get(k, 100)) or (not is_lb(k) and pa[k] < cible.get(k, 100)))) + sum(1 for k in all_kpis if k in qa and ((is_lb(k) and qa[k] > cible.get(k, 100)) or (not is_lb(k) and qa[k] < cible.get(k, 100))))
        st.markdown('''<div class="cr">
            <div class="cc c1"><div class="cv">%.1f%%</div><div class="cl">Score Performance</div></div>
            <div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Qualité</div></div>
            <div class="cc c3"><div class="cv">%d</div><div class="cl">Postes Suivis</div></div>
            <div class="cc c4"><div class="cv">%d</div><div class="cl">KPI Non Atteints</div></div>
        </div>''' % (p_global, q_global, len(vp), nb_non_atteints), unsafe_allow_html=True)
        
        # Onglets
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📋 Synthèse", "📈 Performance", "✅ Qualité", "📊 Analyse", "🔍 Évolution", "💡 Recommandations"])
        
        with tab1:
            st.markdown('<div class="stl p">INDICATEURS PERFORMANCE</div>', unsafe_allow_html=True)
            st.markdown(html_synth(pk, pa, cible, {
                "TAUX_REALISATION_CORRECTIF/PT": "Taux de réalisation insuffisant",
                "OT préparation <1 mois": "Trop d'OT récents en préparation",
                "OT préparation >3 mois": "Trop d'OT >3 mois en préparation",
                "OT planification <1 mois": "Trop d'OT récents en planification",
                "OT planification >3 mois": "Trop d'OT >3 mois en planification",
                "OT exécution <1 mois": "Trop d'OT récents en exécution",
                "OT exécution >3 mois": "Trop d'OT >3 mois en exécution"
            }, "#2b6cb0"), unsafe_allow_html=True)
            
            st.markdown('<div class="stl q">INDICATEURS QUALITÉ</div>', unsafe_allow_html=True)
            st.markdown(html_synth(qk, qa, cible, {
                "appel avis approuvé": "Avis sans OT associé",
                "OT LANC ESTIME": "OT lancés sans estimation",
                "Backlog préparation caractérisé": "OT non caractérisés en préparation",
                "Backlog planification caractérisé": "OT non caractérisés en planification",
                "OT CONFIME": "OT non confirmés dans le système",
                "OT_COR_EGAL": "Écarts de coûts non corrigés"
            }, "#276749"), unsafe_allow_html=True)
            
            st.markdown('<div class="stl c">CLASSEMENT COMBINÉ</div>', unsafe_allow_html=True)
            combined_scores = {p: round((pscores.get(p, 0) + qscores.get(p, 0)) / 2, 2) for p in vp}
            st.markdown(html_classement(combined_scores, "#805ad5"), unsafe_allow_html=True)
            
            st.markdown('<div class="stl ev">COMPARAISON PERFORMANCE vs QUALITÉ</div>', unsafe_allow_html=True)
            st.markdown(html_grouped_bars(vp, pscores, qscores, "Performance vs Qualité par Poste", 90), unsafe_allow_html=True)
        
        with tab2:
            st.markdown('<div class="stl p">BARRES DE PROGRESSION PERFORMANCE</div>', unsafe_allow_html=True)
            st.markdown(html_kpi_bars(pk, pa, cible, "Taux de réalisation Performance", "#3182ce", "#e53e3e", kpi_history), unsafe_allow_html=True)
            
            # Graphique Plotly
            if pa:
                fig_p = go.Figure()
                for k in pk:
                    fig_p.add_trace(go.Bar(x=[k], y=[pa.get(k, 0)], name=k, marker_color='#3182ce' if ((not is_lb(k) and pa.get(k, 0) >= cible.get(k, 100)) or (is_lb(k) and pa.get(k, 0) <= cible.get(k, 100))) else '#e53e3e'))
                fig_p.add_trace(go.Scatter(x=pk, y=[cible.get(k, 100) for k in pk], mode='lines+markers', name='Cible', line=dict(color='#e53e3e', dash='dash'), marker=dict(size=8)))
                fig_p.update_layout(height=400, margin=dict(l=200, r=20, t=40, b=80), title="Performance par KPI", xaxis_tickangle=-45, showlegend=False)
                st.plotly_chart(fig_p, use_container_width=True)
        
        with tab3:
            st.markdown('<div class="stl q">BARRES DE PROGRESSION QUALITÉ</div>', unsafe_allow_html=True)
            st.markdown(html_kpi_bars(qk, qa, cible, "Taux de réalisation Qualité", "#38a169", "#e53e3e", kpi_history), unsafe_allow_html=True)
            
            if qa:
                fig_q = go.Figure()
                for k in qk:
                    fig_q.add_trace(go.Bar(x=[k], y=[qa.get(k, 0)], name=k, marker_color='#38a169' if ((not is_lb(k) and qa.get(k, 0) >= cible.get(k, 100)) or (is_lb(k) and qa.get(k, 0) <= cible.get(k, 100))) else '#e53e3e'))
                fig_q.add_trace(go.Scatter(x=qk, y=[cible.get(k, 100) for k in qk], mode='lines+markers', name='Cible', line=dict(color='#e53e3e', dash='dash'), marker=dict(size=8)))
                fig_q.update_layout(height=400, margin=dict(l=200, r=20, t=40, b=80), title="Qualité par KPI", xaxis_tickangle=-45, showlegend=False)
                st.plotly_chart(fig_q, use_container_width=True)
        
        with tab4:
            st.markdown('<div class="stl ev">ANALYSE AUTOMATIQUE DU BACKLOG</div>', unsafe_allow_html=True)
            findings = analyze_backlog_auto(dfp, avf, vp)
            st.markdown(html_findings(findings), unsafe_allow_html=True)
            
            if findings:
                findings_df = pd.DataFrame(findings)
                st.download_button("📥 Exporter Analyse Backlog", data=findings_df.to_csv(index=False).encode(), "analyse_backlog.csv", "text/csv")
        
        with tab5:
            st.markdown('<div class="stl ev">TABLEAU ÉVOLUTIF DES KPIs</div>', unsafe_allow_html=True)
            evol_rows = build_evolution_table(ckdf, kpi_history, all_kpis, cible, vp)
            st.markdown(html_evolution_table(evol_rows), unsafe_allow_html=True)
            
            if evol_rows:
                evol_df = pd.DataFrame([{k: v for k, v in r.items() if k not in ["Tendance_Color", "Delta_Raw", "LB"]} for r in evol_rows])
                st.download_button("📥 Exporter Évolution", data=evol_df.to_csv(index=False).encode(), "evolution_kpi.csv", "text/csv")
        
        with tab6:
            st.markdown('<div class="stl a">RECOMMANDATIONS PRIORISÉES</div>', unsafe_allow_html=True)
            recs = generate_recommendations(ckdf, cible, vp, kpi_history, all_kpis)
            st.markdown(html_rec_cards(recs), unsafe_allow_html=True)
            
            if recs:
                recs_df = pd.DataFrame([{k: v for k, v in r.items() if k not in ["Priorite_Color", "Priorite_Sort"]} for r in recs])
                st.download_button("📥 Exporter Recommandations", data=recs_df.to_csv(index=False).encode(), "recommandations.csv", "text/csv")
        
        # ===================== EXPORT COMPLET =====================
        st.markdown("---")
        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            st.markdown("**📥 Export Complet Excel**")
        with col_exp2:
            evol_export = pd.DataFrame([{k: v for k, v in r.items() if k not in ["Tendance_Color", "Delta_Raw", "LB"]} for r in evol_rows]) if evol_rows else None
            recs_export = pd.DataFrame([{k: v for k, v in r.items() if k not in ["Priorite_Color", "Priorite_Sort"]} for r in recs]) if recs else None
            findings_export = pd.DataFrame(findings) if findings else None
            
            excel_data = create_analyse_excel(
                pd.DataFrame(), pd.DataFrame(), pd.DataFrame(),
                evol_export, recs_export, findings_export
            )
            st.download_button("📥 Télécharger Analyse Complète", data=excel_data.getvalue(), file_name="analyse_maintenance.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error("❌ Erreur lors du traitement: %s" % str(e))
        import traceback
        with st.expander("Détails de l'erreur"):
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
