# -*- coding: utf-8 -*-
"""app.py — KPI Dashboard MC & FEED — Version optimisée"""
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, altair as alt, random, time, os
from datetime import datetime

def inject_custom_css():
    st.markdown("""<style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        :root{--primary:#1e3a5f;--primary-light:#2c5282;--accent:#ed8936;--success:#38a169;--danger:#e53e3e;--warning:#d69e2e;--border:#e2e8f0;--radius:12px;--shadow:0 2px 12px rgba(0,0,0,0.06)}
        .stApp{background:linear-gradient(135deg,#f0f4f8 0%,#e2e8f0 50%,#f0f4f8 100%);font-family:'Inter',sans-serif}
        .main .block-container{padding-left:.75rem!important;padding-right:.75rem!important;padding-top:1rem!important;max-width:100%!important}
        [data-testid="stHorizontalBlock"]{gap:.6rem!important;align-items:stretch}
        [data-testid="stVerticalBlock"]{gap:.15rem!important}
        .stVerticalBlock{gap:.15rem!important}
        .main-header{background:linear-gradient(135deg,var(--primary) 0%,var(--primary-light) 100%);padding:18px 24px;border-radius:var(--radius);margin-bottom:12px;box-shadow:0 6px 24px rgba(0,0,0,0.1);position:relative;overflow:hidden}
        .main-header h1{color:#fff;font-size:24px;font-weight:800;margin:0}
        .main-header .subtitle{color:rgba(255,255,255,.7);font-size:12px;margin-top:2px}
        .main-header .date-badge{position:absolute;top:18px;right:24px;background:rgba(255,255,255,.15);padding:5px 12px;border-radius:30px;color:#fff;font-size:11px;border:1px solid rgba(255,255,255,.2)}
        .section-title{font-size:15px;font-weight:700;color:var(--primary);margin:10px 0 6px 0;padding-left:10px;border-left:4px solid var(--accent)}
        .chart-container{background:#fff;border-radius:var(--radius);padding:12px;box-shadow:var(--shadow);border:1px solid var(--border);margin-bottom:4px}
        .empty-state{text-align:center;padding:24px 16px;color:#718096}.empty-state .icon{font-size:32px;margin-bottom:8px}.empty-state h3{color:#1a202c;font-size:14px;font-weight:600;margin-bottom:4px}
        div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--primary) 0%,#0f2744 100%);width:240px!important}
        div[data-testid="stSidebar"] *{color:rgba(255,255,255,.9)!important}
        div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:10px;text-transform:uppercase}
        div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:8px;padding:2px 8px;margin-bottom:3px;border:1px solid rgba(255,255,255,.1)}
        div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:6px}
        .dv{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:9px;display:block;overflow-x:auto}
        .dv thead th{background:var(--primary)!important;color:#fff!important;font-weight:700!important;font-size:8px;text-transform:uppercase;padding:5px 4px!important;border:none!important;white-space:nowrap;position:sticky;top:0;z-index:10}
        .dv tbody td{padding:3px 4px!important;border-bottom:1px solid var(--border)!important;white-space:nowrap}
        .dv tbody tr:nth-child(even) td{background:#f9fbfd!important}.dv tbody tr:hover td{background:#edf2f7!important}
        .stTabs [data-baseweb="tab-list"]{gap:2px;background:#edf2f7;padding:3px;border-radius:8px;margin-bottom:6px}
        .stTabs [data-baseweb="tab"]{border-radius:6px;padding:6px 12px;font-weight:600;font-size:11px}
        .stTabs [aria-selected="true"]{background:#fff!important;color:var(--primary)!important;box-shadow:0 2px 6px rgba(0,0,0,.08)}
        .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--primary) 0%,var(--primary-light) 100%);border:none;border-radius:8px;padding:8px 16px;font-weight:700;font-size:12px;width:100%}
        .stDownloadButton>button{background:linear-gradient(135deg,var(--success) 0%,#276749 100%);border:none;border-radius:8px;padding:8px 16px;font-weight:700;font-size:12px;color:#fff;width:100%}
        ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1;border-radius:3px}::-webkit-scrollbar-thumb{background:#c1c1c1;border-radius:3px}
        [data-testid="stRadio"]{margin-bottom:4px}
        [data-testid="stRadio"]>div>div>label{padding:4px 10px!important;font-size:11px}
        .ano-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}
        .ano-header span{font-size:13px;font-weight:700;color:var(--primary);padding-left:10px;border-left:4px solid var(--accent)}
    </style>""", unsafe_allow_html=True)

# ============================================================
# CONSTANTES
# ============================================================
PERF_KPIS = ["TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois", "OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois", "OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]
QUAL_KPIS = ["appel avis approuvé", "OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT COR_EGAL"]

PERF_SHORT = {"TAUX_REALISATION_CORRECTIF/PT":"Taux Réal.","OT préparation <1 mois":"Prép. <1m","OT préparation >3 mois":"Prép. >3m","OT préparation 1mois< <3mois":"Prép. 1-3m","OT planification <1 mois":"Plan. <1m","OT planification >3 mois":"Plan. >3m","OT planification 1mois< <3mois":"Plan. 1-3m","OT exécution <1 mois":"Exéc. <1m","OT exécution >3 mois":"Exéc. >3m","OT exécution 1mois< <3mois":"Exéc. 1-3m"}
QUAL_SHORT = {"appel avis approuvé":"Appel Avis","OT LANC ESTIME":"OT Estimé","Backlog préparation caractérisé":"Backlog Prép.","Backlog planification caractérisé":"Backlog Plan.","OT CONFIME":"OT Confirmé","OT COR_EGAL":"Coûts ="}

# ============================================================
# UTILITAIRES
# ============================================================
def exclure_cresseurs(df):
    if "Poste travail princ." not in df.columns: return df
    return df[~df["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)]

def contient_mot(t, l):
    t = str(t); return any(m in t for s in l for m in s.split())

def categorie_age(a):
    return "<1 mois" if a <= 1 else (">3 mois" if a >= 3 else "1 mois < <3 mois")

def calcul_kpi(n, d, s=100):
    return np.where(d == 0, s, (n / d) * 100)

def creer_groupby(df_filt, col_cat, postes):
    """Remplacement de pivot_table par groupby + unstack"""
    g = df_filt.groupby(["Poste travail princ.", col_cat])["Ordre"].count().unstack(fill_value=0)
    return g.reindex(postes, fill_value=0)

def rename_safe(df, old, new):
    m = {o: n for o, n in zip(old, new) if o in df.columns}
    return df.rename(columns=m) if m else df

def get_kpi_score(k, a, t):
    if pd.isna(a) or pd.isna(t): return 0
    if k in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return 1 if a >= 75 else 0
    if k in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return 1 if a <= 15 else 0
    if k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return 1 if a <= 5 else 0
    if k == "TAUX_REALISATION_CORRECTIF/PT": return 1 if a >= 80 else 0
    if k == "appel avis approuvé": return 1 if a >= 90 else 0
    if k in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT COR_EGAL"]: return 1 if a >= 95 else 0
    return 0

def get_metier(p):
    p = str(p).upper()
    if "E" in p: return "Électrique"
    if "M" in p: return "Mécanique"
    if "R" in p: return "Instrumentation"
    if "G" in p: return "Génie Civil"
    return "Autre"

def get_atelier(p):
    p = str(p).upper()
    if any(x in p for x in ["CENT","CC","CALD","CHAUD","TURB"]): return "Centrale"
    if any(x in p for x in ["UTIL","UTI","AIR","EAU","STA"]): return "Utilitaires"
    if "PS" in p: return "Sulfurique"
    if "PP" in p: return "Phosphorique"
    if "TSP" in p or "REX" in p: return "Engrais"
    if "MCP" in p or "DCP" in p: return "Feed"
    return "Autre"

def get_chart_color(val, kpi):
    try: v = float(val)
    except: return "#cbd5e0"
    if kpi in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return "#38a169" if v >= 80 else ("#ecc94b" if v >= 75 else "#e53e3e")
    if kpi in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return "#38a169" if v <= 15 else "#e53e3e"
    if kpi in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return "#38a169" if v <= 5 else "#e53e3e"
    if kpi == "TAUX_REALISATION_CORRECTIF/PT": return "#38a169" if v >= 85 else ("#ecc94b" if v >= 80 else "#e53e3e")
    if kpi == "appel avis approuvé": return "#38a169" if v >= 95 else ("#ecc94b" if v >= 90 else "#e53e3e")
    if kpi in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT COR_EGAL"]: return "#38a169" if v >= 100 else ("#ecc94b" if v >= 95 else "#e53e3e")
    return "#cbd5e0"

def score_color(v):
    if v >= 80: return "#38a169"
    if v >= 60: return "#ecc94b"
    return "#e53e3e"

# ============================================================
# CHARTS
# ============================================================
def render_total_vertical_chart(kpi_names, values, short_map):
    """Barres verticales pour Total Général avec valeur en haut"""
    vals = [float(v) if pd.notna(v) else 0 for v in values]
    df = pd.DataFrame({"KPI": kpi_names, "Valeur": vals})
    df["Label"] = df["KPI"].map(short_map).fillna(df["KPI"])
    df["Couleur"] = df.apply(lambda r: get_chart_color(r["Valeur"], r["KPI"]), axis=1)
    h = max(200, 28 * len(df))
    bars = alt.Chart(df).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X('Label:O', title='', axis=alt.Axis(labelAngle=-50, labelFontSize=10, labelPadding=4)),
        y=alt.Y('Valeur:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title='%'),
        color=alt.Color('Couleur:N', scale=None, legend=None)
    )
    texts = alt.Chart(df).mark_text(align='center', baseline='bottom', dy=-3, fontSize=9, fontWeight='700', color='#1a202c').encode(
        x=alt.X('Label:O'), y=alt.Y('Valeur:Q'), text=alt.Text('Valeur:Q', format='.1f')
    )
    return (bars + texts).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent').properties(height=h)

def render_grouped_h_chart(df_group, label_col, value_col):
    """Barres horizontales par groupe avec valeur à l'extérieur (droite)"""
    df = df_group[[label_col, value_col]].dropna().copy()
    if df.empty: return None
    df = df.sort_values(value_col, ascending=True)
    df["Couleur"] = df[value_col].apply(score_color)
    h = max(100, 24 * len(df))
    bars = alt.Chart(df).mark_bar(height=18, cornerRadiusTopLeft=4, cornerRadiusBottomLeft=4).encode(
        x=alt.X(f'{value_col}:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title='%'),
        y=alt.Y(f'{label_col}:O', title='', axis=alt.Axis(labelFontSize=11)),
        color=alt.Color('Couleur:N', scale=None, legend=None)
    )
    texts = alt.Chart(df).mark_text(align='left', baseline='middle', dx=6, fontSize=11, fontWeight='700', color='#1a202c').encode(
        x=alt.X(f'{value_col}:Q'), y=alt.Y(f'{label_col}:O'), text=alt.Text(f'{value_col}:Q', format='.1f')
    )
    return (bars + texts).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent').properties(height=h)

# ============================================================
# CHARGEMENT & TRAITEMENT (CACHÉ)
# ============================================================
@st.cache_data
def load_raw_data():
    df = pd.read_excel("ot.xlsx")
    av = pd.read_excel("avis.xlsx")
    df = exclure_cresseurs(df)
    av = exclure_cresseurs(av)
    for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
        if c in df.columns: df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in ["Créé le","Début souhaité","Date de la clôture"]:
        if c in av.columns: av[c] = pd.to_datetime(av[c], errors="coerce")
    return df, av

@st.cache_data
def process_data(_unused, valid_postes_tuple, start_str, end_str):
    df_ot_raw, avis_df_raw = load_raw_data()
    valid_postes = list(valid_postes_tuple)
    start_date = pd.to_datetime(start_str)
    end_date = pd.to_datetime(end_str)
    now = pd.Timestamp.now()

    df = df_ot_raw[
        (df_ot_raw["Poste travail princ."].isin(valid_postes)) &
        (df_ot_raw["Date de début planifiée"].between(start_date, end_date))
    ].copy()
    avis_df = avis_df_raw[avis_df_raw["Poste travail princ."].isin(valid_postes)].copy()

    mask_sf = df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)
    df = exclure_cresseurs(df[mask_sf].drop_duplicates(subset=["Ordre"], keep="first"))
    avis_df = exclure_cresseurs(avis_df[(avis_df["Ordre"].isna()) | (avis_df["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates())

    if "Statut système" in df.columns:
        df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

    res = {}
    df_t = df

    mp = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
    df_t["Backlog préparation"] = np.where(df_t["Statut utilisateur"].apply(lambda x: contient_mot(x, mp)), "CARACTERISE", "NON CARACTERISE")

    mplan = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
    df_t["Backlog planification"] = np.where(df_t["Statut utilisateur"].apply(lambda x: contient_mot(x, mplan)), "CARACTERISE", "NON CARACTERISE")

    for dc, am, ac in [('Créé le',"Age mois préparation","Age préparation"),('Date de début planifiée',"Age mois planification","Age planification"),('Date de début planifiée',"Age mois exécution","Age exécution")]:
        if dc in df_t.columns:
            df_t[dc] = pd.to_datetime(df_t[dc], errors='coerce')
            df_t[am] = ((now.year - df_t[dc].dt.year)*12 + (now.month - df_t[dc].dt.month)).round(2)
            df_t[ac] = df_t[am].apply(categorie_age)
        else:
            df_t[am] = np.nan; df_t[ac] = "Inconnu"

    df_t["OT CONFIME"] = np.where(df_t["Statut système"].str.contains("CLO", na=False) & df_t["Statut système"].str.contains("CONF", na=False), "OUI", "NON")
    df_t["Contient SOPL"] = df_t["Statut utilisateur"].str.contains("SOPL", na=False).astype(int)
    df_t["OT LANC ESTIME"] = np.where(df_t["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
    df_t["OT COR_EGAL"] = np.where((df_t["Total coûts budgétés"].fillna(0) - df_t["Total coûts réels"].fillna(0)) == 0, "OUI", "NON")

    res['df_processed'] = df_t

    # --- Anomalie (correctif) ---
    an = creer_groupby(df_t[df_t["Nº appel pl.entret."].fillna(0) == 0], "Statut OT", valid_postes)
    for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c] = an.get(c, 0)
    an["Total"] = an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
    an["TAUX_REALISATION_CORRECTIF/PT"] = calcul_kpi(an["TCLO"], an["Total"])

    # --- Préparation ---
    pr = creer_groupby(df_t[df_t["Statut OT"] == "CRÉÉ"], "Age préparation", valid_postes)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pr[c] = pr.get(c, 0)
    pr["Total"] = pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    pr["OT préparation <1 mois"] = calcul_kpi(pr["<1 mois"], pr["Total"])
    pr["OT préparation >3 mois"] = calcul_kpi(pr[">3 mois"], pr["Total"], 0)
    pr["OT préparation 1mois< <3mois"] = calcul_kpi(pr["1 mois < <3 mois"], pr["Total"], 0)

    # --- Planification ---
    pl = creer_groupby(df_t[(df_t["Statut OT"] == "LANC") & (df_t["Contient SOPL"] == 0)], "Age planification", valid_postes)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pl[c] = pl.get(c, 0)
    pl["Total"] = pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    pl["OT planification <1 mois"] = calcul_kpi(pl["<1 mois"], pl["Total"])
    pl["OT planification >3 mois"] = calcul_kpi(pl[">3 mois"], pl["Total"], 0)
    pl["OT planification 1mois< <3mois"] = calcul_kpi(pl["1 mois < <3 mois"], pl["Total"], 0)

    # --- Exécution ---
    ex = creer_groupby(df_t[(df_t["Statut OT"] == "LANC") & (df_t["Contient SOPL"] == 1)], "Age exécution", valid_postes)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: ex[c] = ex.get(c, 0)
    ex["Total"] = ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    ex["OT exécution <1 mois"] = calcul_kpi(ex["<1 mois"], ex["Total"])
    ex["OT exécution >3 mois"] = calcul_kpi(ex[">3 mois"], ex["Total"], 0)
    ex["OT exécution 1mois< <3mois"] = calcul_kpi(ex["1 mois < <3 mois"], ex["Total"], 0)

    # --- OT LANC ESTIME ---
    la = creer_groupby(df_t[df_t["Statut OT"] == "LANC"], "OT LANC ESTIME", valid_postes)
    for c in ["OUI","NON"]: la[c] = la.get(c, 0)
    la["Total"] = la["OUI"] + la["NON"]; la["OT LANC ESTIME"] = calcul_kpi(la["OUI"], la["Total"])

    # --- Backlog préparation ---
    pc = creer_groupby(df_t[df_t["Statut OT"] == "CRÉÉ"], "Backlog préparation", valid_postes)
    for c in ["CARACTERISE","NON CARACTERISE"]: pc[c] = pc.get(c, 0)
    pc["Total"] = pc["CARACTERISE"] + pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"] = calcul_kpi(pc["CARACTERISE"], pc["Total"])

    # --- Backlog planification ---
    plc = creer_groupby(df_t[df_t["Statut OT"] == "LANC"], "Backlog planification", valid_postes)
    for c in ["CARACTERISE","NON CARACTERISE"]: plc[c] = plc.get(c, 0)
    plc["Total"] = plc["CARACTERISE"] + plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"] = calcul_kpi(plc["CARACTERISE"], plc["Total"])

    # --- OT CONFIME & OT COR_EGAL ---
    for kn, cn in [("ot_confime","OT CONFIME"),("ot_cor_egal","OT COR_EGAL")]:
        pv = creer_groupby(df_t, cn, valid_postes)
        for c in ["OUI","NON"]: pv[c] = pv.get(c, 0)
        pv["Total"] = pv["OUI"] + pv["NON"]; pv[cn] = calcul_kpi(pv["OUI"], pv["Total"]); res[kn] = pv

    # --- Avis ---
    avf = avis_df.copy(); res['avis_df_filtered'] = avf
    tca = creer_groupby(avf, "Statut utilisateur", valid_postes)
    for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c] = tca.get(c, 0)
    tca["Total"] = tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["appel avis approuvé"] = calcul_kpi(tca["APRV"], tca["Total"])

    # --- Consolidation KPI ---
    calculated_kpis_df = pd.concat([
        an[["TAUX_REALISATION_CORRECTIF/PT"]],
        pr[["OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois"]],
        pl[["OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois"]],
        ex[["OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois"]],
        tca[["appel avis approuvé"]],
        la[["OT LANC ESTIME"]],
        pc[["Backlog préparation caractérisé"]],
        plc[["Backlog planification caractérisé"]],
        res['ot_confime'][["OT CONFIME"]],
        res['ot_cor_egal'][["OT COR_EGAL"]]
    ], axis=1)

    cible = pd.DataFrame([{k: 85 if k == "TAUX_REALISATION_CORRECTIF/PT" else (80 if "<1 mois" in k else (5 if ">3 mois" in k else (15 if "1mois" in k else (95 if k == "appel avis approuvé" else 100)))) for k in PERF_KPIS + QUAL_KPIS}], index=["CIBLE"])
    df_processed = df_t

    # --- Anomalies ---
    ano_perf, ano_qual = [], []
    checks = [
        ("TAUX_REALISATION_CORRECTIF/PT","TAUX_REALISATION_CORRECTIF/PT",(df_processed["Nº appel pl.entret."].fillna(0)==0)&(~df_processed["Statut OT"].isin(["CLOT","TCLO"])),"Améliorer le taux de réalisation."),
        ("OT préparation <1 mois","OT préparation <1 mois",(df_processed["Statut OT"]=="CRÉÉ")&(df_processed["Age préparation"]!="<1 mois"),"Réduire l'âge de préparation."),
        ("OT préparation >3 mois","OT préparation >3 mois",(df_processed["Statut OT"]=="CRÉÉ")&(df_processed["Age préparation"]==">3 mois"),"Traiter les OT > 3 mois."),
        ("OT planification <1 mois","OT planification <1 mois",(df_processed["Statut OT"]=="LANC")&(df_processed["Contient SOPL"]==0)&(df_processed["Age planification"]!="<1 mois"),"Réduire l'âge de planification."),
        ("OT planification >3 mois","OT planification >3 mois",(df_processed["Statut OT"]=="LANC")&(df_processed["Contient SOPL"]==0)&(df_processed["Age planification"]==">3 mois"),"Traiter OT planif. > 3 mois."),
        ("OT exécution <1 mois","OT exécution <1 mois",(df_processed["Statut OT"]=="LANC")&(df_processed["Contient SOPL"]==1)&(df_processed["Age exécution"]!="<1 mois"),"Réduire l'âge d'exécution."),
        ("OT exécution >3 mois","OT exécution >3 mois",(df_processed["Statut OT"]=="LANC")&(df_processed["Contient SOPL"]==1)&(df_processed["Age exécution"]==">3 mois"),"Traiter OT exéc. > 3 mois."),
        ("OT LANC ESTIME","OT LANC ESTIME",(df_processed["Statut OT"]=="LANC")&(df_processed["OT LANC ESTIME"]=="NON"),"Estimer les coûts."),
        ("Backlog préparation caractérisé","Backlog préparation caractérisé",(df_processed["Statut OT"]=="CRÉÉ")&(df_processed["Backlog préparation"]=="NON CARACTERISE"),"Caractériser backlog prépa."),
        ("Backlog planification caractérisé","Backlog planification caractérisé",(df_processed["Statut OT"]=="LANC")&(df_processed["Backlog planification"]=="NON CARACTERISE"),"Caractériser backlog planif."),
        ("OT CONFIME","OT CONFIME",df_processed["OT CONFIME"]=="NON","Confirmer les OT terminés."),
        ("OT COR_EGAL","OT COR_EGAL",df_processed["OT COR_EGAL"]=="NON","Rapprocher les coûts.")
    ]

    for poste in valid_postes:
        if poste not in df_processed["Poste travail princ."].values: continue
        dp = df_processed[df_processed["Poste travail princ."] == poste]
        for kn, cc, cond, act in checks:
            vk = calculated_kpis_df.loc[poste, kn] if poste in calculated_kpis_df.index else 100
            if pd.notna(vk) and vk < cible.loc['CIBLE', cc]:
                # CORRECTION : cond.sum() compte les True de la Series booléenne
                # au lieu de dp[cond].sum() qui essayait de sommer toutes les colonnes
                # y compris les colonnes datetime
                cnt = int(cond.sum())
                if cnt > 0:
                    (ano_perf if kn in PERF_KPIS else ano_qual).append({"Poste travail princ.": poste, "KPI": kn, "Nb OT impactés": cnt, "Action Suggérée": act})
        va = calculated_kpis_df.loc[poste, "appel avis approuvé"] if poste in calculated_kpis_df.index else 100
        if pd.notna(va) and va < cible.loc['CIBLE', "appel avis approuvé"]:
            cnt = len(avf[avf["Poste travail princ."] == poste])
            if cnt > 0: ano_qual.append({"Poste travail princ.": poste, "KPI": "appel avis approuvé", "Nb OT impactés": cnt, "Action Suggérée": "Créer un OT pour les avis sans ordre."})

    def prep_ano(df_ano):
        if df_ano.empty: return pd.DataFrame(columns=["Poste de travail"])
        df_ot = df_ano[df_ano["KPI"] != "appel avis approuvé"].groupby(["Poste travail princ.","KPI"])["Nb OT impactés"].sum().unstack(fill_value=0)
        df_av = df_ano[df_ano["KPI"] == "appel avis approuvé"].groupby(["Poste travail princ."])["Nb OT impactés"].sum().to_frame("Nb Avis sans ordre")
        dash = df_ot.join(df_av, how='outer').fillna(0).astype(int)
        if not dash.empty:
            dash["Total éléments impactés"] = dash.sum(axis=1)
            dash.loc["Total général"] = dash.sum()
        return dash.rename_axis("Poste de travail").reset_index()

    def prep_kpi_table(kpi_list):
        tg = calculated_kpis_df[kpi_list].mean().to_frame().T; tg.index = ["Total général"]
        final = pd.concat([cible[kpi_list], calculated_kpis_df[kpi_list], tg]).round(2).rename_axis("Poste de travail").reset_index()
        return final

    def prep_class_table(kpi_list, col_name):
        rows = []
        for p in calculated_kpis_df.index:
            sc = sum(get_kpi_score(k, calculated_kpis_df.loc[p, k], cible.loc['CIBLE', k]) for k in kpi_list if k in calculated_kpis_df.columns) / len(kpi_list) * 100
            rows.append({"Poste travail princ.": p, col_name: sc})
        df_cl = pd.DataFrame(rows)
        mean_val = df_cl[col_name].mean()
        tg_cl = pd.DataFrame([{"Poste travail princ.": "Total général", col_name: mean_val}])
        df_cl = pd.concat([df_cl, tg_cl], ignore_index=True)
        df_cl[col_name] = df_cl[col_name].apply(lambda x: f"{x:.2f} %")
        return df_cl

    # --- Classement par poste ---
    df_class = pd.DataFrame([{
        "Poste travail princ.": p,
        "Indicateur de Performance": sum(get_kpi_score(k, calculated_kpis_df.loc[p, k], cible.loc['CIBLE', k]) for k in PERF_KPIS if k in calculated_kpis_df.columns) / len(PERF_KPIS) * 100,
        "Indicateur de Qualité": sum(get_kpi_score(k, calculated_kpis_df.loc[p, k], cible.loc['CIBLE', k]) for k in QUAL_KPIS if k in calculated_kpis_df.columns) / len(QUAL_KPIS) * 100
    } for p in calculated_kpis_df.index])
    df_class["Atelier"] = df_class["Poste travail princ."].apply(get_atelier)
    df_class["Métier"] = df_class["Poste travail princ."].apply(get_metier)

    # --- Agrégation par Atelier ---
    by_atelier = df_class.groupby("Atelier").agg(
        Performance=("Indicateur de Performance","mean"),
        Qualité=("Indicateur de Qualité","mean")
    ).reset_index().round(2)

    # --- Agrégation par Métier ---
    by_metier = df_class.groupby("Métier").agg(
        Performance=("Indicateur de Performance","mean"),
        Qualité=("Indicateur de Qualité","mean")
    ).reset_index().round(2)

    return {
        'kpi_perf': prep_kpi_table(PERF_KPIS), 'kpi_qual': prep_kpi_table(QUAL_KPIS),
        'class_perf': prep_class_table(PERF_KPIS, "Indicateur de Performance"),
        'class_qual': prep_class_table(QUAL_KPIS, "Indicateur de Qualité"),
        'ano_perf': prep_ano(pd.DataFrame(ano_perf)), 'ano_qual': prep_ano(pd.DataFrame(ano_qual)),
        'df_class': df_class, 'df_processed': df_processed,
        'df_ano_perf': pd.DataFrame(ano_perf), 'df_ano_qual': pd.DataFrame(ano_qual),
        'cible': cible, 'avis_df_filtered': avf,
        'by_atelier': by_atelier, 'by_metier': by_metier
    }

# ============================================================
# STYLES DE LIGNES
# ============================================================
def style_kpi_rows(row):
    if row["Poste de travail"] == "CIBLE": return ["padding:5px 4px;background-color:#1e3a5f;color:#fff;font-weight:700;font-size:9px;"]*len(row)
    if row["Poste de travail"] == "Total général": return ["padding:5px 4px;background-color:#e2e8f0;color:#1a202c;font-weight:800;font-size:9px;border-top:2px solid #1e3a5f;"]*len(row)
    return ["padding:4px 4px;font-size:9px;"]*len(row)

def style_class_rows(row):
    if row["Poste travail princ."] == "Total général": return ["padding:5px 4px;background-color:#e2e8f0;color:#1a202c;font-weight:800;font-size:9px;border-top:2px solid #1e3a5f;"]*len(row)
    return ["padding:4px 4px;font-size:9px;"]*len(row)

def style_ano_rows(row):
    if row["Poste de travail"] == "Total général": return ["padding:5px 4px;background-color:#1e3a5f;color:#fff;font-weight:800;font-size:9px;"]*len(row)
    return ["padding:4px 4px;font-size:9px;"]*len(row)

# ============================================================
# MAIN
# ============================================================
def main():
    try: locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except: pass
    inject_custom_css()

    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        c = random.choice(["Port obligatoire des EPI avant toute intervention.","Aucun travail n'est plus urgent que la sécurité.","Ne jamais intervenir sur un équipement en marche.","Baliser et sécuriser la zone de travail."])
        st.markdown(f"""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d 0%,#2d3748 50%,#1a365d 100%);padding:40px;">
            <div style="font-size:56px;margin-bottom:16px;">🦺</div>
            <h1 style="color:#fff;font-weight:900;font-size:36px;margin:0;">HSE - CONSIGNE DE SÉCURITÉ</h1>
            <div style="background:linear-gradient(135deg,#f6e05e 0%,#ed8936 100%);padding:28px 40px;border-radius:16px;font-size:24px;font-weight:700;text-align:center;margin:30px 0;color:#1a202c;max-width:700px;box-shadow:0 16px 48px rgba(0,0,0,0.3);">⚠️ {c}</div>
            <h2 style="color:#48bb78;font-size:26px;font-weight:900;">Aucun travail n'est plus urgent que la sécurité</h2>
            <div style="margin-top:30px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden;">
                <div style="width:100%;height:100%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 1s ease-in-out forwards;"></div>
            </div>
            <style>@keyframes ld{{from{{width:0%}}to{{width:100%}}}}</style>
        </div>""", unsafe_allow_html=True)
        time.sleep(1)
        st.session_state.hse_affiche = True; st.rerun(); st.stop()

    # ========== SIDEBAR ==========
    with st.sidebar:
        st.markdown("""<div style="padding:12px 0 6px 0;"><div style="font-size:22px;margin-bottom:2px;">⚙️</div><div style="font-size:14px;font-weight:800;color:white;">Filtres & Paramètres</div></div>""", unsafe_allow_html=True)
        st.markdown("---")
        use_new_files = st.toggle("📁 Charger de nouveaux fichiers", value=False, key="toggle_files")
        ot_file = avis_file = None; all_postes_master_list = []

        if use_new_files:
            ot_file = st.file_uploader("Fichier OT", type=["xlsx"], key="up_ot")
            avis_file = st.file_uploader("Fichier AVIS", type=["xlsx"], key="up_avis")
        else:
            date_f = datetime.now().strftime("%d/%m/%Y")
            if os.path.exists("ot.xlsx"):
                try:
                    date_f = datetime.fromtimestamp(os.path.getmtime("ot.xlsx")).strftime("%d/%m/%Y")
                    _tmp = exclure_cresseurs(load_raw_data()[0])
                    mask_tmp = _tmp["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)
                    all_postes_master_list = sorted(_tmp[mask_tmp]["Poste travail princ."].dropna().unique().tolist())
                except: pass
            st.markdown(f"""<div style="background:rgba(255,255,255,.1);padding:8px 12px;border-radius:8px;border:1px solid rgba(255,255,255,.15);"><div style="font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;">Données</div><div style="font-size:12px;color:white;font-weight:600;margin-top:2px;">📅 {date_f}</div></div>""", unsafe_allow_html=True)

        st.markdown("---")
        selected_postes = st.multiselect("Poste de travail", ["All"] + all_postes_master_list, ["All"], key="sel_postes")
        selected_ateliers = st.multiselect("Atelier", ["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)","Centrale","Utilitaires"], ["All"], key="sel_ateliers")
        selected_divisions = st.multiselect("Division", ["All","SF1","SF2"], ["All"], key="sel_div")
        st.markdown("---")
        date_range = st.date_input("Date début planifiée", value=(datetime(2025,1,1).date(), datetime.today().date()), format="DD/MM/YYYY", key="date_range")

    if not use_new_files or (ot_file is not None and avis_file is not None):
        try:
            if use_new_files:
                df_ot_raw = pd.read_excel(ot_file); avis_df_raw = pd.read_excel(avis_file); date_fichier = datetime.now().strftime("%d/%m/%Y")
            else:
                df_ot_raw, avis_df_raw = load_raw_data(); date_fichier = datetime.fromtimestamp(os.path.getmtime("ot.xlsx")).strftime("%d/%m/%Y")

            if not all_postes_master_list:
                _tmp_raw = exclure_cresseurs(df_ot_raw)
                mask_raw = _tmp_raw["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)
                all_postes_master_list = sorted(_tmp_raw[mask_raw]["Poste travail princ."].dropna().unique().tolist())

            if "All" in selected_postes or not selected_postes: selected_postes = all_postes_master_list
            if "All" in selected_ateliers or not selected_ateliers: selected_ateliers = ["All"]
            if "All" in selected_divisions or not selected_divisions: selected_divisions = ["All"]
            start_date = pd.to_datetime(date_range[0]) if len(date_range) == 2 else pd.to_datetime(datetime(2025,1,1))
            end_date = pd.to_datetime(date_range[1]) if len(date_range) == 2 else pd.to_datetime(datetime.today())

            def match_filters(poste):
                p = str(poste).upper()
                if "All" not in selected_ateliers:
                    m = False
                    if "Sulfurique (PS)" in selected_ateliers and "PS" in p: m = True
                    if "Phosphorique (PP)" in selected_ateliers and "PP" in p: m = True
                    if "Engrais (TSP/REX)" in selected_ateliers and ("TSP" in p or "REX" in p): m = True
                    if "Feed (MCP/DCP)" in selected_ateliers and ("MCP" in p or "DCP" in p): m = True
                    if "Centrale" in selected_ateliers and any(x in p for x in ["CENT","CC","CALD","CHAUD","TURB"]): m = True
                    if "Utilitaires" in selected_ateliers and any(x in p for x in ["UTIL","UTI","AIR","EAU","STA"]): m = True
                    if not m: return False
                if "All" not in selected_divisions:
                    m = False
                    if "SF1" in selected_divisions and "SF1" in p: m = True
                    if "SF2" in selected_divisions and "SF2" in p: m = True
                    if not m: return False
                return True

            valid_postes = [p for p in all_postes_master_list if match_filters(p) and p in selected_postes]
            if not valid_postes: valid_postes = all_postes_master_list[:1]
            data = process_data("run", tuple(valid_postes), start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))

            # ========== HEADER ==========
            st.markdown(f"""<div class="main-header"><h1>📊 KPI Dashboard MC & FEED</h1><div class="subtitle">Maintenance Conditionnelle • Suivi des indicateurs de performance</div><div class="date-badge">📅 {date_fichier}</div></div>""", unsafe_allow_html=True)

            # ========== 1. INDICATEURS DE PERFORMANCE ==========
            st.markdown('<p class="section-title">📊 Indicateurs de Performance</p>', unsafe_allow_html=True)

            c1, c2 = st.columns(2)
            with c1:
                st.markdown(data['kpi_perf'].style.apply(style_kpi_rows, axis=1).to_html(index=False, classes="dv"), unsafe_allow_html=True)
            with c2:
                st.markdown(data['class_perf'].style.apply(style_class_rows, axis=1).to_html(index=False, classes="dv"), unsafe_allow_html=True)

            # --- Total Général Performance : barres verticales, valeur en haut ---
            total_perf_row = data['kpi_perf'][data['kpi_perf']['Poste de travail'] == 'Total général']
            if not total_perf_row.empty:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div style="font-size:12px;font-weight:700;color:#1e3a5f;margin-bottom:6px;">📈 Total Général — Performance</div>', unsafe_allow_html=True)
                st.altair_chart(
                    render_total_vertical_chart(PERF_KPIS, [total_perf_row[k].values[0] for k in PERF_KPIS], PERF_SHORT),
                    use_container_width=True
                )
                st.markdown('</div>', unsafe_allow_html=True)

            # --- Vue Par Atelier / Par Métier / Anomalies ---
            vue_perf = st.radio("Vue Performance", ["🏭 Par Atelier", "🔧 Par Métier", "🚨 Anomalies"], key="radio_perf", horizontal=True, label_visibility="collapsed")
            if vue_perf == "🏭 Par Atelier":
                ch = render_grouped_h_chart(data['by_atelier'], "Atelier", "Performance")
                if ch:
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown('<div style="font-size:12px;font-weight:700;color:#1e3a5f;margin-bottom:4px;">Performance par Atelier</div>', unsafe_allow_html=True)
                    st.altair_chart(ch, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
            elif vue_perf == "🔧 Par Métier":
                ch = render_grouped_h_chart(data['by_metier'], "Métier", "Performance")
                if ch:
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown('<div style="font-size:12px;font-weight:700;color:#1e3a5f;margin-bottom:4px;">Performance par Métier</div>', unsafe_allow_html=True)
                    st.altair_chart(ch, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                if not data['ano_perf'].empty:
                    st.markdown('<div class="ano-header"><span>🚨 Anomalies Performance</span></div>', unsafe_allow_html=True)
                    st.markdown(data['ano_perf'].style.apply(style_ano_rows, axis=1).to_html(index=False, classes="dv"), unsafe_allow_html=True)
                else:
                    st.markdown('<div class="empty-state"><div class="icon">✅</div><h3>Aucune anomalie de performance</h3></div>', unsafe_allow_html=True)

            st.markdown("<hr style='border:1px solid #e2e8f0;margin:12px 0;'>", unsafe_allow_html=True)

            # ========== 2. INDICATEURS DE QUALITÉ ==========
            st.markdown('<p class="section-title">🎯 Indicateurs de Qualité</p>', unsafe_allow_html=True)

            c3, c4 = st.columns(2)
            with c3:
                st.markdown(data['kpi_qual'].style.apply(style_kpi_rows, axis=1).to_html(index=False, classes="dv"), unsafe_allow_html=True)
            with c4:
                st.markdown(data['class_qual'].style.apply(style_class_rows, axis=1).to_html(index=False, classes="dv"), unsafe_allow_html=True)

            # --- Total Général Qualité : barres verticales, valeur en haut ---
            total_qual_row = data['kpi_qual'][data['kpi_qual']['Poste de travail'] == 'Total général']
            if not total_qual_row.empty:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div style="font-size:12px;font-weight:700;color:#1e3a5f;margin-bottom:6px;">🎯 Total Général — Indicateurs Qualité</div>', unsafe_allow_html=True)
                st.altair_chart(
                    render_total_vertical_chart(QUAL_KPIS, [total_qual_row[k].values[0] for k in QUAL_KPIS], QUAL_SHORT),
                    use_container_width=True
                )
                st.markdown('</div>', unsafe_allow_html=True)

            # --- Vue Par Atelier / Par Métier / Anomalies ---
            vue_qual = st.radio("Vue Qualité", ["🏭 Par Atelier", "🔧 Par Métier", "🚨 Anomalies"], key="radio_qual", horizontal=True, label_visibility="collapsed")
            if vue_qual == "🏭 Par Atelier":
                ch = render_grouped_h_chart(data['by_atelier'], "Atelier", "Qualité")
                if ch:
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown('<div style="font-size:12px;font-weight:700;color:#1e3a5f;margin-bottom:4px;">Qualité par Atelier</div>', unsafe_allow_html=True)
                    st.altair_chart(ch, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
            elif vue_qual == "🔧 Par Métier":
                ch = render_grouped_h_chart(data['by_metier'], "Métier", "Qualité")
                if ch:
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.markdown('<div style="font-size:12px;font-weight:700;color:#1e3a5f;margin-bottom:4px;">Qualité par Métier</div>', unsafe_allow_html=True)
                    st.altair_chart(ch, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                if not data['ano_qual'].empty:
                    st.markdown('<div class="ano-header"><span>🚨 Anomalies Qualité</span></div>', unsafe_allow_html=True)
                    st.markdown(data['ano_qual'].style.apply(style_ano_rows, axis=1).to_html(index=False, classes="dv"), unsafe_allow_html=True)
                else:
                    st.markdown('<div class="empty-state"><div class="icon">✅</div><h3>Aucune anomalie de qualité</h3></div>', unsafe_allow_html=True)

            # ========== 3. EXPORT ==========
            st.markdown("<hr style='border:1px solid #e2e8f0;margin:12px 0;'>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">💾 Export des Plans d\'Action</p>', unsafe_allow_html=True)

            all_postes_ano = list(set(
                data['ano_perf'][data['ano_perf']['Poste de travail'] != 'Total général']['Poste de travail'].tolist() +
                data['ano_qual'][data['ano_qual']['Poste de travail'] != 'Total général']['Poste de travail'].tolist()
            ))

            if all_postes_ano:
                ce1, ce2 = st.columns([1,1])
                with ce1:
                    sel_exp = st.selectbox("Poste de travail :", options=["📌 Tous les postes"] + all_postes_ano, key="sel_exp")
                with ce2:
                    st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
                    gen_btn = st.button("📥 Générer le fichier Excel", type="primary", key="btn_exp", use_container_width=True)

                if gen_btn:
                    with st.spinner("Génération en cours..."):
                        out = io.BytesIO()
                        with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                            df_processed = data['df_processed']
                            df_ano_perf = data['df_ano_perf']
                            df_ano_qual = data['df_ano_qual']
                            avis_filtered = data['avis_df_filtered']

                            for pe in (all_postes_ano if sel_exp == "📌 Tous les postes" else [sel_exp]):
                                kds = list(set(
                                    df_ano_perf[df_ano_perf["Poste travail princ."] == pe]["KPI"].tolist() +
                                    df_ano_qual[df_ano_qual["Poste travail princ."] == pe]["KPI"].tolist()
                                ))
                                for kpi in kds:
                                    sd = pd.DataFrame()
                                    if kpi != "appel avis approuvé":
                                        dpf = df_processed[df_processed["Poste travail princ."] == pe]
                                        cmap = {
                                            "TAUX_REALISATION_CORRECTIF/PT":(dpf["Nº appel pl.entret."].fillna(0)==0)&(~dpf["Statut OT"].isin(["CLOT","TCLO"])),
                                            "OT préparation <1 mois":(dpf["Statut OT"]=="CRÉÉ")&(dpf["Age préparation"]!="<1 mois"),
                                            "OT préparation >3 mois":(dpf["Statut OT"]=="CRÉÉ")&(dpf["Age préparation"]==">3 mois"),
                                            "OT planification <1 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==0)&(dpf["Age planification"]!="<1 mois"),
                                            "OT planification >3 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==0)&(dpf["Age planification"]==">3 mois"),
                                            "OT exécution <1 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==1)&(dpf["Age exécution"]!="<1 mois"),
                                            "OT exécution >3 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==1)&(dpf["Age exécution"]==">3 mois"),
                                            "OT LANC ESTIME":(dpf["Statut OT"]=="LANC")&(dpf["OT LANC ESTIME"]=="NON"),
                                            "Backlog préparation caractérisé":(dpf["Statut OT"]=="CRÉÉ")&(dpf["Backlog préparation"]=="NON CARACTERISE"),
                                            "Backlog planification caractérisé":(dpf["Statut OT"]=="LANC")&(dpf["Backlog planification"]=="NON CARACTERISE"),
                                            "OT CONFIME":dpf["OT CONFIME"]=="NON",
                                            "OT COR_EGAL":dpf["OT COR_EGAL"]=="NON"
                                        }
                                        sub = dpf[cmap.get(kpi, pd.Series(False, index=dpf.index))]
                                        if not sub.empty:
                                            sub = rename_safe(sub, ["Ordre","Désignation","Emplacement technique","Poste travail princ.","Statut système","Statut utilisateur","Date de début planifiée","Type d'ordre","Backlog préparation","Backlog planification"], ["Ordre de travail","Désignation","Poste technique","Poste de travail principal","Statut système","Statut utilisateur","Date de début planifiée","Type d'ordre","Caractérisation backlog Préparation","Caractérisation backlog Planification"])
                                            sub["KPI impacté"] = kpi; sub["Action recommandée"] = f"Corriger l'indicateur {kpi}."
                                            sd = pd.concat([sd, sub])

                                    if kpi == "appel avis approuvé":
                                        sa = avis_filtered[avis_filtered["Poste travail princ."] == pe]
                                        if not sa.empty:
                                            sa = rename_safe(sa, ["Avis","Désignation texte","Emplacement technique","Poste travail princ.","Statut utilisateur","Créé le"], ["Avis","Désignation","Poste technique","Poste de travail principal","Statut","Date de création"])
                                            sa["KPI impacté"] = kpi; sa["Action recommandée"] = "Créer un OT pour cet Avis."
                                            sd = pd.concat([sd, sa])

                                    if not sd.empty:
                                        sn = f"{pe.replace(' ','_').replace('/','_')[:20]}_{kpi.replace('/','_').replace(' ','_')[:10]}"[:31]
                                        sd.to_excel(w, sheet_name=sn, index=False)

                        out.seek(0)
                        nf = "Plan_Action_Tous_Postes.xlsx" if sel_exp == "📌 Tous les postes" else f"Plan_Action_{sel_exp.replace(' ','_')}.xlsx"
                        st.download_button(label="✅ Télécharger le fichier Excel", data=out.getvalue(), file_name=nf, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.markdown('<div class="empty-state" style="padding:24px;"><div class="icon">🎉</div><h3>Aucun export nécessaire</h3></div>', unsafe_allow_html=True)

            st.markdown(f"""<div style="text-align:center;padding:16px 0 6px 0;color:#a0aec0;font-size:10px;">KPI Dashboard MC & FEED • Maintenance Conditionnelle • {date_fichier}</div>""", unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
            import traceback; st.code(traceback.format_exc())

if __name__ == "__main__":
    try: locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except: pass
    main()
