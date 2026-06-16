# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os, hashlib, json
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
st.set_page_config(layout="wide", page_title="Dashboard KPI")
# ============================================================

QK = ["TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois",
      "OT préparation 1mois< <3mois", "OT planification <1 mois", "OT planification >3 mois",
      "OT planification 1mois< <3mois", "OT exécution <1 mois", "OT exécution >3 mois",
      "OT exécution 1mois< <3mois",
      "Performance Graissage", "Performance Inspection", "Performance Appels Systématiques"]
PK = ["appel avis approuvé", "OT LANC ESTIME", "Backlog préparation caractérisé",
      "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL",
      "OT Fiabilité", "Total Avis de Panne"]
ALL_KPI = QK + PK

CIBLE = {"TAUX_REALISATION_CORRECTIF/PT": 85, "OT préparation <1 mois": 80, "OT préparation >3 mois": 5,
         "OT préparation 1mois< <3mois": 15, "OT planification <1 mois": 80, "OT planification >3 mois": 5,
         "OT planification 1mois< <3mois": 15, "OT exécution <1 mois": 80, "OT exécution >3 mois": 5,
         "OT exécution 1mois< <3mois": 15, "appel avis approuvé": 95, "OT LANC ESTIME": 100,
         "Backlog préparation caractérisé": 100, "Backlog planification caractérisé": 100,
         "OT CONFIME": 100, "OT_COR_EGAL": 100,
         "Performance Graissage": 95, "Performance Inspection": 95, "Performance Appels Systématiques": 95,
         "OT Fiabilité": 100, "Total Avis de Panne": 100}

ACT_MAP = {"TAUX_REALISATION_CORRECTIF/PT": "Ameliorer le taux de realisation des OT.",
           "OT préparation <1 mois": "Reduire l'age de preparation des OT (< 1 mois).",
           "OT préparation >3 mois": "Traiter les OT avec preparation > 3 mois.",
           "OT planification <1 mois": "Reduire l'age de planification des OT (< 1 mois).",
           "OT planification >3 mois": "Traiter les OT avec planification > 3 mois.",
           "OT exécution <1 mois": "Reduire l'age d'execution des OT (< 1 mois).",
           "OT exécution >3 mois": "Traiter les OT avec execution > 3 mois.",
           "OT LANC ESTIME": "Estimer les couts des OT lances.",
           "Backlog préparation caractérisé": "Caracteriser le backlog de preparation.",
           "Backlog planification caractérisé": "Caracteriser le backlog de planification.",
           "OT CONFIME": "Confirmer les OT termines.",
           "OT_COR_EGAL": "Rapprocher les couts reels et budgetes.",
           "appel avis approuvé": "Creer un OT pour les avis sans ordre.",
           "OT préparation 1mois< <3mois": "Reduire les OT entre 1 et 3 mois.",
           "OT planification 1mois< <3mois": "Reduire les OT entre 1 et 3 mois.",
           "OT exécution 1mois< <3mois": "Reduire les OT entre 1 et 3 mois.",
           "Performance Graissage": "Ameliorer le taux de realisation des OT de graissage (Type 350).",
           "Performance Inspection": "Ameliorer le taux de realisation des OT d'inspection (Types 290,300,310).",
           "Performance Appels Systématiques": "Ameliorer le taux de realisation des appels systematiques (Type 360).",
           "OT Fiabilité": "Maintenir la fiabilite des OT a 100%.",
           "Total Avis de Panne": "Maintenir le suivi des avis de panne a 100%."}

LOWER_BETTER = ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois",
                "OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]

MP_KW = ["CRPR ATPD", "CRPR ATMR", "CRPR ATER", "CRPR ATRS", "CRPR ATMO",
         "ATPD", "ATMR", "ATER", "ATRS", "ATMO"]
MPLAN_KW = ["ATPL ATEI", "ATPL ATAL", "ATPL ATER", "ATPL AGAR", "ATPL ATHS",
            "ATEI", "ATAL", "ATAS", "AGAR", "ATHS"]

CHANGELOG = [
    {"version": "2.1", "date": "2025-06-18", "changes": [
        "Deplacement KPI Graissage/Inspection/Systematiques de Qualite vers Performance",
        "Nouveau tableau OT OMS par Poste et Statut OT avec 2 Pie charts",
        "Nouveau tableau OT Thermographie par Poste et Statut OT avec 2 Pie charts",
        "Nouveau tableau Tous les OT par Poste et Statut OT avec 2 Pie charts",
        "Page Anomalies simplifiee : resume KPI x Poste avec coloriage",
        "Page Suivi & Evolution : synthese entre deux dates par poste",
        "Nouvelle page Backlog : OMS, Thermographie, Statuts OT, Caracterisation",
        "Cartes KPI separees par groupe de postes (SF1, SF2, Autres)",
        "Classement des anomalies par nombre decroissant",
        "Uniformisation des couleurs et mise en forme des tableaux"
    ]},
    {"version": "2.0", "date": "2025-06-15", "changes": [
        "KPI Taux realisation correctif/PT : ajout filtre SOPL=1, numerateur CLOT+TCLO, total=0 => 100%",
        "KPI Age backlog preparation : filtre Statut OT=CRE + Statut utilisateur contient CRPR",
        "KPI Age backlog planification : filtre Statut OT=LANC + Statut utilisateur contient ATPL",
        "Nouveau KPI Performance Graissage (Type 350) - Seuils V/J/R",
        "Nouveau KPI Performance Inspection (Types 290,300,310) - Exclusion dates futures - Seuils V/J/R",
        "Nouveau KPI Performance Appels Systematiques (Type 360) - Exclusion dates futures - Seuils V/J/R",
        "Nouveaux KPI Qualite Appels : OT Fiabilite (100%), Total Avis de Panne (100%)",
        "Mise en place mecanisme de cache pour eviter les recalculs systematiques",
        "Activation du suivi des ameliorations et evolutions (changelog)"
    ]}
]

CONSIGNES_HSE = [
    "Port obligatoire des EPI avant toute intervention.", "Port obligatoire du casque de securite.",
    "Port obligatoire des lunettes de protection.", "Port obligatoire des gants adaptes au travail.",
    "Utiliser les protections auditives dans les zones bruyantes.", "Verifier l'absence de tension avant toute intervention electrique.",
    "Respecter la procedure de consignation et deconsignation.", "Ne jamais intervenir sur un equipement en marche.",
    "Baliser et securiser la zone de travail.", "Maintenir le poste de travail propre et ordonne.",
    "Verifier l'etat des outils avant utilisation.", "Utiliser uniquement du materiel homologue.",
    "Respecter les permis de travail en vigueur.", "Identifier les risques avant de commencer une tache.",
    "Signaler immediatement toute situation dangereuse.", "Signaler tout incident ou presque accident.",
    "Ne jamais neutraliser un dispositif de securite.", "Verifier les detecteurs de gaz avant utilisation.",
    "Verifier la bonne ventilation des zones de travail.", "Respecter les regles des espaces confines.",
    "Controler l'atmosphere avant d'entrer dans un espace confine.", "Utiliser les points d'ancrage pour les travaux en hauteur.",
    "Verifier l'etat des echafaudages avant utilisation.", "Securiser les outils lors des travaux en hauteur.",
    "Ne pas travailler seul lors d'operations a risque.", "Controler les elingues avant chaque levage.",
    "Respecter les limites de charge des equipements.", "Verifier l'etat des appareils de levage.",
    "Maintenir les voies de circulation degagees.", "Respecter la signalisation de securite.",
    "Verifier les extincteurs a proximite du chantier.", "Connaitre les issues de secours les plus proches.",
    "Respecter les procedures d'arret d'urgence.", "Verifier les flexibles et raccords avant mise en service.",
    "Controler les fuites avant demarrage d'un equipement.", "Respecter les distances de securite.",
    "Ne jamais contourner une procedure HSE.", "Porter les EPI adaptes au risque identifie.",
    "Prevenir son responsable avant toute intervention particuliere.", "Analyser les risques avant chaque demarrage de chantier.",
    "Verifier la stabilite des equipements.", "Utiliser les bons outils pour la bonne tache.",
    "Respecter les consignes specifiques du chantier.", "Ne jamais prendre de raccourci au detriment de la securite.",
    "Arreter immediatement les travaux en cas de danger.", "Proteger l'environnement lors des interventions.",
    "Collecter et trier correctement les dechets.", "Eviter toute pollution accidentelle.",
    "Respecter les consignes de stockage des produits dangereux.", "Lire les fiches de securite avant manipulation.",
    "Verifier les equipements avant chaque prise de poste.", "S'assurer de la disponibilite des moyens de secours.",
    "Communiquer clairement avec l'equipe avant intervention.", "Respecter les regles de circulation des engins.",
    "Garder une vigilance permanente sur son environnement.", "Prendre le temps d'effectuer le travail en securite.",
    "La securite est l'affaire de tous.", "Chaque incident peut etre evite par la prevention.",
    "Aucun travail n'est plus urgent que la securite.", "Zero accident commence par un comportement sur."]


# ============================================================
# FONCTIONS UTILITAIRES GLOBALES
# ============================================================
def compute_cache_key(file_date, filters_dict, now_str):
    key_data = {"file_date": file_date, "filters": filters_dict, "now": now_str}
    return hashlib.md5(json.dumps(key_data, sort_keys=True, default=str).encode()).hexdigest()


def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt", "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass
    return datetime.now().strftime("%d/%m/%Y")


def save_kpis_to_excel(prows, pcols, qrows, qcols, ano_p_r, ano_p_c, ano_q_r, ano_q_c, sheet_name):
    kpis_dir = "kpis"
    os.makedirs(kpis_dir, exist_ok=True)
    filepath = os.path.join(kpis_dir, "indicateurs_kpis.xlsx")
    sn = str(sheet_name).replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "")[:31]
    hf = Font(bold=True, color="FFFFFF", size=10)
    hfl = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    tf = Font(bold=True, size=12, color="1E3A5F")
    tb = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    try:
        wb = load_workbook(filepath)
    except Exception:
        wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    if sn in wb.sheetnames:
        del wb[sn]
    ws = wb.create_sheet(sn)
    rn = 1

    def ws_sec(title, cols, rows, sr):
        ws.cell(row=sr, column=1, value=title).font = tf
        sr += 1
        for j, c in enumerate(cols, 1):
            cl = ws.cell(row=sr, column=j, value=c)
            cl.font = hf
            cl.fill = hfl
            cl.alignment = Alignment(horizontal='center')
            cl.border = tb
        sr += 1
        for r in rows:
            for j, c in enumerate(cols, 1):
                cl = ws.cell(row=sr, column=j, value=r.get(c, ""))
                cl.border = tb
                cl.alignment = Alignment(horizontal='center')
            sr += 1
        return sr + 1

    rn = ws_sec("INDICATEURS DE PERFORMANCE", pcols, prows, rn)
    if ano_p_c and ano_p_r:
        rn = ws_sec("ANOMALIES PERFORMANCE", ano_p_c, ano_p_r, rn)
    rn = ws_sec("INDICATEURS DE QUALITE", qcols, qrows, rn)
    if ano_q_c and ano_q_r:
        rn = ws_sec("ANOMALIES QUALITE", ano_q_c, ano_q_r, rn)
    try:
        wb.save(filepath)
    except Exception:
        pass


def load_historical_kpis(filepath):
    if not os.path.exists(filepath):
        return pd.DataFrame()
    try:
        wb = load_workbook(filepath, read_only=True, data_only=True)
    except Exception:
        return pd.DataFrame()
    records = []
    section = None
    headers = None
    for sheet_name in wb.sheetnames:
        try:
            ws = wb[sheet_name]
            rows_data = list(ws.iter_rows(values_only=True))
            for row in rows_data:
                cell0 = str(row[0]).strip() if row[0] else ""
                if "INDICATEURS DE PERFORMANCE" in cell0.upper():
                    section = "perf"
                    headers = None
                    continue
                elif "INDICATEURS DE QUALITE" in cell0.upper():
                    section = "qual"
                    headers = None
                    continue
                elif "ANOMALIES" in cell0.upper():
                    section = None
                    continue
                if section and headers is None and cell0:
                    headers = [str(c).strip() if c else "" for c in row]
                    continue
                if section and headers and cell0 and cell0 not in ("CIBLE", "Total general", ""):
                    entry = {"Date": sheet_name}
                    for j, h in enumerate(headers):
                        if j < len(row):
                            entry[h] = row[j]
                    entry["_section"] = section
                    records.append(entry)
        except Exception:
            continue
    wb.close()
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df["Date_parsed"] = pd.to_datetime(df["Date"].str.replace("-", "/"), format="%d/%m/%Y", errors="coerce")
    return df.sort_values("Date_parsed").reset_index(drop=True)


def calculate_variations(hist_df):
    if hist_df.empty or "Date" not in hist_df.columns:
        return pd.DataFrame()
    dates = sorted(hist_df["Date"].unique())
    if len(dates) < 2:
        return pd.DataFrame()
    perf_df = hist_df[hist_df["_section"] == "perf"].copy()
    qual_df = hist_df[hist_df["_section"] == "qual"].copy()
    variations = []
    for i in range(1, len(dates)):
        prev_date, curr_date = dates[i - 1], dates[i]
        prev_perf = perf_df[perf_df["Date"] == prev_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        curr_perf = perf_df[perf_df["Date"] == curr_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        prev_qual = qual_df[qual_df["Date"] == prev_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        curr_qual = qual_df[qual_df["Date"] == curr_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        for sec_name, prev_d, curr_d, kpi_list in [("Performance", prev_perf, curr_perf, QK + ["Score Performance"]), ("Qualite", prev_qual, curr_qual, PK + ["Score Qualite"])]:
            for poste in set(prev_d.index) & set(curr_d.index):
                for kpi in kpi_list:
                    if kpi not in prev_d.columns or kpi not in curr_d.columns:
                        continue
                    try:
                        pv = float(prev_d.loc[poste, kpi])
                    except Exception:
                        continue
                    try:
                        cv = float(curr_d.loc[poste, kpi])
                    except Exception:
                        continue
                    diff = cv - pv
                    pct = (diff / pv * 100) if pv != 0 else (100 if cv != 0 else 0)
                    if abs(diff) <= 0.5:
                        trend = "stabilite"
                    elif diff > 0.5:
                        trend = "hausse"
                    else:
                        trend = "baisse"
                    variations.append({"Date precedente": prev_date, "Date actuelle": curr_date, "Poste": poste,
                                       "Type": sec_name, "KPI": kpi, "Valeur precedente": round(pv, 2), "Valeur actuelle": round(cv, 2),
                                       "Ecart": round(diff, 2), "Ecart %": round(pct, 2), "Tendance": trend})
    return pd.DataFrame(variations)


def generate_journal(var_df):
    if var_df.empty:
        return pd.DataFrame()
    j = var_df.copy()
    j["Significatif"] = j["Ecart %"].abs() >= 5
    j = j[j["Significatif"]].copy()
    j["Sens"] = j.apply(lambda r: "Amelioration" if ((r["Tendance"] == "hausse" and r["KPI"] not in LOWER_BETTER) or (r["Tendance"] == "baisse" and r["KPI"] in LOWER_BETTER)) else "Degradation", axis=1)
    return j.sort_values(["Date actuelle", "Sens", "Ecart %"], ascending=[True, False, False])


def calculate_rankings(var_df):
    if var_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    scores = {}
    for poste in var_df["Poste"].unique():
        pv = var_df[var_df["Poste"] == poste].copy()
        scores[poste] = sum((-r["Ecart %"] if r["KPI"] in LOWER_BETTER else r["Ecart %"]) for _, r in pv.iterrows())
    ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    return pd.DataFrame(ranked[:5], columns=["Poste", "Score variation"]), pd.DataFrame(ranked[-5:][::-1], columns=["Poste", "Score variation"])


# ============================================================
# CSS MODIFIEE - uniformisation, premiere colonne, total, score
# ============================================================
def inject_custom_css():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{width:250px!important}
    section[data-testid="stSidebar"][aria-expanded="false"]{width:0px!important}
    .main .block-container{max-width:100%!important;width:100%!important;padding-left:0.5rem!important;padding-right:0.5rem!important}
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--p:#1e3a5f;--pl:#2c5282;--b:#e2e8f0;--r:10px}
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#edf2f7;font-family:'Inter',sans-serif}
    .main .block-container{padding-top:.8rem;padding-bottom:.8rem}
    .stTabs,.stTabs>div,.stTabs [data-baseweb="tab-list"]{width:100%!important;max-width:100%!important}
    .mh{background:linear-gradient(135deg,var(--p),var(--pl));padding:12px 20px;border-radius:var(--r);margin-bottom:6px;box-shadow:0 6px 20px rgba(0,0,0,.1);overflow:hidden}
    .mh h1{color:#fff;font-size:20px;font-weight:800;margin:0;display:inline}
    .mh .db{float:right;background:rgba(255,255,255,.15);padding:3px 12px;border-radius:14px;color:#fff;font-size:14px;font-weight:500;border:1px solid rgba(255,255,255,.2);margin-top:2px}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:6px}
    .cc{background:#fff;border-radius:var(--r);padding:10px 12px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid var(--b);text-align:center}
    .cc .cv{font-size:26px;font-weight:900;line-height:1}
    .cc .cl{font-size:11px;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
    .cc.c1{border-top:3px solid #3182ce}.cc.c1 .cv{color:#2b6cb0}
    .cc.c2{border-top:3px solid #38a169}.cc.c2 .cv{color:#276749}
    .cc.c3{border-top:3px solid #805ad5}.cc.c3 .cv{color:#6b46c1}
    .cc.c4{border-top:3px solid #e53e3e}.cc.c4 .cv{color:#c53030}
    .stl{font-size:15px;font-weight:700;color:var(--p);margin:6px 0 2px 0;padding-left:10px;border-left:3px solid var(--pl)}
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}.stl.s{border-left-color:#d69e2e}.stl.bl{border-left-color:#7c3aed}

    /* ===== TABLEAUX - UNIFORMISATION COMPLETE ===== */
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    /* En-tetes uniformes : meme gradient sur toutes les pages */
    .tw thead th{background:linear-gradient(135deg,#1e3a5f,#2c5282);color:#fff;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.3px;padding:6px 8px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    /* Premiere colonne : fond bleu clair distinct pour meilleure visibilite */
    .tw tbody td.fc{background:#e0f2fe!important;color:#0c4a6e!important;font-weight:700!important;text-align:left!important;min-width:140px;position:sticky;left:0;z-index:3;box-shadow:2px 0 4px rgba(0,0,0,.04)}
    /* Colonne Score : bordure verte distincte */
    .tw tbody td.sc{border-left:3px solid #059669!important;background:#f0fdf4!important;font-weight:800!important}
    .tw thead th.sc-head{border-left:3px solid #10b981;background:linear-gradient(135deg,#134e4a,#0f766e)!important}
    .tw thead th.fc-head{background:linear-gradient(135deg,#0c4a6e,#1a6d9e)!important;text-align:left}
    /* Ligne Total General : mise en evidence verte */
    .tw tbody tr.total-row td{background:#d1fae5!important;color:#065f46!important;font-weight:800!important;font-size:12.5px!important}
    .tw tbody tr.total-row td.fc{background:#a7f3d0!important;color:#064e3b!important;box-shadow:2px 0 6px rgba(0,0,0,.06)}
    .tw tbody tr.total-row td.sc{background:#6ee7b7!important;border-left-color:#047857!important;color:#064e3b!important}
    .tw tbody tr.cible-row td{background:#fef3c7!important;color:#92400e!important;font-weight:600!important}
    .tw tbody tr:nth-child(even) td:not(.fc):not(.sc){background:#f7fafc}
    .tw tbody tr:hover td:not(.fc):not(.sc){background:#ebf8ff!important}

    /* Variante anomalies */
    .tw.at thead th{background:linear-gradient(135deg,#991b1b,#dc2626)!important}
    .tw.at thead th.fc-head{background:linear-gradient(135deg,#7f1d1d,#b91c1c)!important}
    .tw.at thead th.sc-head{border-left-color:#f87171;background:linear-gradient(135deg,#7f1d1d,#b91c1c)!important}
    .tw.at tbody td.fc{background:#fef2f2!important;color:#7f1d1d!important}

    /* Variante backlog */
    .tw.bl thead th{background:linear-gradient(135deg,#4c1d95,#7c3aed)!important}
    .tw.bl thead th.fc-head{background:linear-gradient(135deg,#3b0764,#6d28d9)!important}
    .tw.bl thead th.sc-head{border-left-color:#a78bfa;background:linear-gradient(135deg,#3b0764,#6d28d9)!important}

    /* Cellules de couleur KPI */
    .cg{background:#c6efce!important;color:#006100!important;font-weight:600!important}
    .cw{background:#ffeb9c!important;color:#9c6500!important;font-weight:600!important}
    .cb{background:#ffc7ce!important;color:#9c0006!important;font-weight:600!important}

    /* Cellules anomalie */
    .a0{background:#c6efce!important;color:#006100!important;font-weight:600!important}
    .a1{background:#ffeb9c!important;color:#9c6500!important;font-weight:600!important}
    .a2{background:#fed7d7!important;color:#c53030!important;font-weight:600!important}
    .a3{background:#fecaca!important;color:#991b1b!important;font-weight:800!important}

    .stTabs [data-baseweb="tab-list"]{gap:3px;background:#e2e8f0;padding:3px;border-radius:6px;margin-bottom:4px}
    .stTabs [data-baseweb="tab"]{border-radius:5px;padding:6px 14px;font-weight:600;font-size:14px}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--p)!important;box-shadow:0 2px 5px rgba(0,0,0,.07)}
    .ca{background:#fff;border-radius:var(--r);padding:10px;margin-top:4px;border:1px solid var(--b);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .ca .ct{font-size:14px;font-weight:700;margin-bottom:6px;padding-bottom:4px;border-bottom:1px solid var(--b)}
    .car{display:flex;align-items:center;margin-bottom:4px;font-size:12px}
    .car:last-child{margin-bottom:0}
    .car .cal{width:260px;font-weight:600;color:var(--p);text-align:right;padding-right:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .car .cab{flex:1;height:24px;background:#edf2f7;border-radius:4px;overflow:hidden}
    .car .caf{height:100%;border-radius:4px;transition:width .3s}
    .car .cav-out{font-size:12px;font-weight:800;color:#1a202c;min-width:55px;text-align:right;padding-left:6px}
    .gbr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f7fafc}
    .gbr:last-child{border:none}
    .gbr-l{width:160px;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px}
    .gbr-g{display:flex;align-items:center;gap:4px;flex:1}
    .gbr-w{flex:1;height:20px;background:#edf2f7;border-radius:3px;overflow:hidden}
    .gbr-f{height:100%;border-radius:3px}
    .gb-p{background:linear-gradient(90deg,#2b6cb0,#4299e1)}.gb-q{background:linear-gradient(90deg,#276749,#48bb78)}
    .gbr-v{font-size:11px;font-weight:800;min-width:48px;text-align:right;color:#1a202c}
    .gbr-legend{display:flex;gap:14px;margin-bottom:6px;font-size:12px;font-weight:700}
    .gbr-legend span{display:flex;align-items:center;gap:5px}
    .gbr-legend i{display:inline-block;width:14px;height:14px;border-radius:2px}
    .cg-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .cg-grid>div{background:#fff;border-radius:var(--r);padding:8px 10px;border:1px solid var(--b)}
    .cg-grid .ct{font-size:13px;font-weight:700;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid var(--b)}
    .cgr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f7fafc}
    .cgr:last-child{border:none}
    .cgr .rk{width:18px;font-weight:800;text-align:center}
    .cgr .pn{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cgr .ps{font-weight:800;min-width:55px;text-align:right}
    .dgrid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:8px 14px;font-weight:700;font-size:15px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label,div[data-testid="stSidebar"] .stCheckbox label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:3px 8px;margin-bottom:3px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:14px;color:#718096;font-size:14px}

    /* Groupe label */
    .grp-label{display:inline-flex;align-items:center;gap:8px;background:linear-gradient(135deg,var(--p),var(--pl));padding:6px 16px;border-radius:8px;color:#fff;font-size:13px;font-weight:700;margin:10px 0 6px 0;letter-spacing:.3px}
    .grp-label .grp-dot{width:8px;height:8px;border-radius:50%;background:#5eead4}

    /* Mini barres groupe */
    .mb-row{display:flex;align-items:center;gap:6px;padding:3px 0;font-size:12px}
    .mb-row .mb-name{min-width:50px;font-weight:700;color:#0c4a6e}
    .mb-row .mb-bars{flex:1;display:flex;gap:4px;align-items:center}
    .mb-row .mb-bar{flex:1;height:18px;background:#edf2f7;border-radius:3px;overflow:hidden}
    .mb-row .mb-fill{height:100%;border-radius:3px}
    .mb-row .mb-val{min-width:50px;text-align:right;font-weight:800;font-size:11px}

    .evol-timeline{border-left:3px solid #2c5282;margin-left:12px;padding-left:20px}
    .evol-item{position:relative;padding-bottom:18px}
    .evol-item::before{content:'';position:absolute;left:-27px;top:4px;width:12px;height:12px;border-radius:50%;background:#2c5282;border:2px solid #fff;box-shadow:0 0 0 2px #2c5282}
    .evol-ver{font-size:13px;font-weight:800;color:#2c5282;margin-bottom:2px}
    .evol-date{font-size:11px;color:#718096;margin-bottom:4px}
    .evol-change{font-size:12px;color:#4a5568;padding:2px 0;padding-left:14px;position:relative}
    .evol-change::before{content:'\\2022';position:absolute;left:0;color:#38a169;font-weight:800}
    .synth-tbl{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px}
    .synth-tbl thead th{background:linear-gradient(135deg,#1e3a5f,#2c5282);color:#fff;font-weight:700;font-size:11px;padding:5px 8px;border:none;white-space:nowrap;position:sticky;top:0}
    .synth-tbl tbody td{padding:4px 8px;border-bottom:1px solid #edf2f7;text-align:center}
    .synth-tbl tbody tr:nth-child(even) td{background:#f7fafc}
    .synth-tbl tbody tr:hover td{background:#ebf8ff!important}
    .synth-tbl .poste-cell{text-align:left;font-weight:700;white-space:nowrap;min-width:140px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:17px}.cg-grid,.dgrid{grid-template-columns:1fr}.car .cal{width:120px}.gbr-l{width:100px}}
    </style>""", unsafe_allow_html=True)


# ============================================================
# FONCTION PRINCIPALE
# ============================================================
def main():
    try:
        locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except Exception:
        try:
            locale.setlocale(locale.LC_ALL, 'fr_FR')
        except Exception:
            pass
    inject_custom_css()
    fichier_date = get_date_from_file()

    # --- Ecran HSE ---
    if "hse_affiche" not in st.session_state:
        st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        c = random.choice(CONSIGNES_HSE)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:46px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SECURITE</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:22px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Securite - Sante - Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:32px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:36px;font-weight:900">Aucun travail n'est plus urgent que la securite</h2>
        <div style="margin-top:40px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden"><div style="width:100%%;height:100%%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 5.5s ease-in-out forwards"></div></div>
        <style>@keyframes ld{from{width:0}to{width:100%%}}</style></div>""" % c, unsafe_allow_html=True)
        time.sleep(6)
        st.session_state.hse_affiche = True
        st.rerun()
        st.stop()

    # --- Fonctions locales (compatibilite avec le code original) ---
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

    def get_text_col(df):
        for c in ["Désignation", "Designation", "Désignation OT", "Texte ordre", "Texte", "Description", "Libellé", "Libelle"]:
            if c in df.columns:
                return c
        for c in df.columns:
            if df[c].dtype == 'object' and any(kw in str(c).lower() for kw in ['sign', 'text', 'desc', 'libell']):
                return c
        return None

    def build_statut_pivot(df_sub, posts):
        if df_sub.empty:
            return pd.DataFrame(index=posts, columns=["CRÉÉ", "LANC", "CLOT", "TCLO", "Total"]).fillna(0).astype(int)
        piv = pd.pivot_table(df_sub, index="Poste travail princ.", columns="Statut OT", values="Ordre", aggfunc="count", fill_value=0)
        for s in ["CRÉÉ", "LANC", "CLOT", "TCLO"]:
            if s not in piv.columns:
                piv[s] = 0
        piv["Total"] = piv[["CRÉÉ", "LANC", "CLOT", "TCLO"]].sum(axis=1)
        return piv.reindex(posts, fill_value=0).fillna(0).astype(int)

    # --- Fonctions de coloriage (meme logique que le code original) ---
    def ks(v, c):
        try:
            val = float(v)
        except Exception:
            return ""
        if c in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
            return "cg" if val >= 80 else ("cw" if val >= 75 else "cb")
        if c in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
            return "cg" if val <= 15 else "cb"
        if c in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
            return "cg" if val <= 5 else "cb"
        if c == "TAUX_REALISATION_CORRECTIF/PT":
            return "cg" if val >= 85 else ("cw" if val >= 80 else "cb")
        if c == "appel avis approuvé":
            return "cg" if val >= 95 else ("cw" if val >= 90 else "cb")
        if c in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
            return "cg" if val >= 100 else ("cw" if val >= 95 else "cb")
        if c in ["Performance Graissage", "Performance Inspection", "Performance Appels Systématiques"]:
            return "cg" if val >= 95 else ("cw" if val > 90 else "cb")
        if c in ["OT Fiabilité", "Total Avis de Panne"]:
            return "cg" if val >= 100 else ("cw" if val >= 95 else "cb")
        return ""

    def cs(v):
        try:
            val = float(str(v).replace(' %', '').strip())
        except Exception:
            return ""
        return "cg" if val >= 90 else ("cw" if val >= 80 else "cb")

    def kas(v):
        try:
            val = int(v)
        except Exception:
            return ""
        if val == 0:
            return "a0"
        if val <= 1:
            return "a1"
        if val <= 3:
            return "a2"
        return "a3"

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
        if k in ["Performance Graissage", "Performance Inspection", "Performance Appels Systématiques"]:
            return 1 if a >= 95 else 0
        if k in ["OT Fiabilité", "Total Avis de Panne"]:
            return 1 if a >= 100 else 0
        return 0

    def is_lb(k):
        return k in LOWER_BETTER

    # --- Fonctions HTML de rendu des tableaux ---

    def html_table(rows, cols, tc, sc_col=None):
        h = '<table class="tw %s"><thead><tr>' % tc
        for i, c in enumerate(cols):
            cls = 'fc-head' if i == 0 else ('sc-head' if sc_col and c in sc_col else '')
            h += '<th class="%s">%s</th>' % (cls, c)
        h += '</tr></thead><tbody>'
        for r in rows:
            rc = "cible-row" if r.get("_t") == "cible" else ("total-row" if r.get("_t") == "total" else "")
            h += '<tr class="%s">' % rc
            for i, c in enumerate(cols):
                v = r.get(c, "")
                fc_cls = 'fc' if i == 0 else ('sc' if sc_col and c in sc_col else '')
                if r.get("_t") == "cible":
                    h += '<td class="%s">%s</td>' % (fc_cls, v)
                else:
                    s = cs(v) if sc_col and c in sc_col else ks(v, c)
                    h += '<td class="%s %s">%s</td>' % (fc_cls, s or "", v)
            h += '</tr>'
        return h + '</tbody></table>'

    def html_kpi_bars(kpi_list, actuals, targets, title, color_ok, color_fail):
        h = '<div class="ca"><div class="ct" style="color:%s">%s</div>' % (color_ok, title)
        for k in kpi_list:
            av = actuals.get(k, 0)
            tv = targets.get(k, 100)
            met = av <= tv if is_lb(k) else av >= tv
            bw = min(max(av, 0), 100)
            bg = color_ok if met else color_fail
            h += '<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>' % (k, bw, bg, av)
        return h + '</div>'

    def html_classement(scores, accent):
        sp = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        met_p = [(p, s) for p, s in sp if s >= 80]
        not_p = [(p, s) for p, s in sp if s < 80]
        t5 = met_p[:5]
        b5 = not_p[-5:] if len(not_p) > 5 else not_p
        h = '<div class="cg-grid"><div><div class="ct" style="color:#38a169">Top 5 - Objectif Atteint</div>'
        if t5:
            for i, (p, s) in enumerate(t5):
                h += '<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>' % (accent, i + 1, p, cs("%.2f" % s), s)
        else:
            h += '<div style="padding:6px;font-size:12px;color:#718096">Aucun poste</div>'
        h += '</div><div><div class="ct" style="color:#e53e3e">Bottom 5 - Non Atteint</div>'
        if b5:
            for i, (p, s) in enumerate(reversed(b5)):
                h += '<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>' % (len(b5) - i, p, cs("%.2f" % s), s)
        else:
            h += '<div style="padding:6px;font-size:12px;color:#38a169">Tous atteints</div>'
        h += '</div></div>'
        return h

    def html_grouped_bars(posts, pscores, qscores, title):
        h = '<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>' % title
        h += '<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span></div>'
        for p in sorted(posts, key=lambda x: (pscores.get(x, 0) + qscores.get(x, 0)) / 2, reverse=True):
            pv, qv = pscores.get(p, 0), qscores.get(p, 0)
            h += '<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f gb-p" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w"><div class="gbr-f gb-q" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div></div></div>' % (p, min(max(pv, 0), 100), pv, min(max(qv, 0), 100), qv)
        return h + '</div>'

    def html_actions_table(kpi_list, actuals, targets, act_map):
        h = '<table class="tw at"><thead><tr>'
        h += '<th class="fc-head">KPI</th><th>Valeur Actuelle</th><th>Cible</th><th>Ecart</th><th>Statut</th><th>Action Recommandee</th>'
        h += '</tr></thead><tbody>'
        for k in kpi_list:
            av = actuals.get(k, 0)
            tv = targets.get(k, 100)
            diff = av - tv
            met = av <= tv if is_lb(k) else av >= tv
            status = "ATTEINT" if met else "NON ATTEINT"
            st_s = "cg" if met else "cb"
            ec_clr = "#276749" if met else "#c53030"
            action = "Objectif atteint" if met else act_map.get(k, "")
            h += '<tr><td class="fc" style="text-align:left;font-weight:600">%s</td><td>%.1f%%</td><td>%.0f%%</td><td style="color:%s;font-weight:700">%+.1f%%</td><td class="%s" style="text-align:center">%s</td><td style="text-align:left;color:#4a5568;font-size:11px">%s</td></tr>' % (k, av, tv, ec_clr, diff, st_s, status, action)
        return h + '</tbody></table>'

    def show_pie_pair(piv_df, title_prefix):
        global_counts = piv_df[["CRÉÉ", "LANC", "CLOT", "TCLO"]].sum()
        global_counts = global_counts[global_counts > 0]
        realised = global_counts.get("CLOT", 0) + global_counts.get("TCLO", 0)
        not_realised = global_counts.sum() - realised
        c1, c2 = st.columns(2)
        with c1:
            if not global_counts.empty:
                fig1 = px.pie(global_counts, names=global_counts.index, values=global_counts.values,
                              title="%s - Par Statut OT" % title_prefix,
                              color_discrete_sequence=["#e53e3e", "#d69e2e", "#38a169", "#3182ce"])
                fig1.update_traces(textposition='inside', textinfo='percent+value', textfont_size=11)
                fig1.update_layout(margin=dict(t=50, b=10, l=10, r=10), height=340,
                                   legend=dict(font_size=10, orientation="h", yanchor="bottom", y=-0.1))
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True)
        with c2:
            if global_counts.sum() > 0:
                pie2_data = pd.DataFrame({"Statut": ["Realises (CLOT+TCLO)", "Non Realises"], "Nombre": [realised, not_realised]})
                fig2 = px.pie(pie2_data, names="Statut", values="Nombre",
                              title="%s - Realises vs Non Realises" % title_prefix,
                              color="Statut", color_discrete_map={"Realises (CLOT+TCLO)": "#38a169", "Non Realises": "#e53e3e"})
                fig2.update_traces(textposition='inside', textinfo='percent+value', textfont_size=11)
                fig2.update_layout(margin=dict(t=50, b=10, l=10, r=10), height=340,
                                   legend=dict(font_size=10, orientation="h", yanchor="bottom", y=-0.1))
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.markdown('<div class="es">Aucune donnee</div>', unsafe_allow_html=True)

    def html_statut_pivot(piv_df, table_class):
        cols = ["Poste de travail", "CRÉÉ", "LANC", "CLOT", "TCLO", "Total"]
        h = '<table class="tw %s"><thead><tr>' % table_class
        for i, c in enumerate(cols):
            cls = 'fc-head' if i == 0 else ('sc-head' if c == "Total" else '')
            h += '<th class="%s">%s</th>' % (cls, c)
        h += '</tr></thead><tbody>'
        for poste, row in piv_df.iterrows():
            h += '<tr><td class="fc">%s</td>' % poste
            for c in ["CRÉÉ", "LANC", "CLOT", "TCLO"]:
                v = int(row.get(c, 0))
                cls = "cb" if c == "CRÉÉ" else ("cw" if c == "LANC" else "cg")
                h += '<td class="%s" style="text-align:center">%d</td>' % (cls, v)
            h += '<td class="sc" style="text-align:center;font-weight:800">%d</td>' % int(row.get("Total", 0))
            h += '</tr>'
        h += '<tr class="total-row"><td class="fc">Total General</td>'
        for c in ["CRÉÉ", "LANC", "CLOT", "TCLO"]:
            h += '<td style="text-align:center">%d</td>' % int(piv_df[c].sum())
        h += '<td class="sc" style="text-align:center">%d</td>' % int(piv_df["Total"].sum())
        h += '</tr></tbody></table>'
        return h

    # --- Calcul des KPIs ---
    def calc_kpis(df_i, av_i, now, posts):
        res = {}
        df = df_i.copy()
        av = av_i.copy()
        df["Backlog preparation"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, MP_KW)), "CARACTERISE", "NON CARACTERISE")
        df["Backlog planification"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, MPLAN_KW)), "CARACTERISE", "NON CARACTERISE")
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
        df["OT LANC ESTIME"] = np.where(df["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
        df["OT_COR_EGAL"] = np.where((df["Total coûts budgétés"].fillna(0) - df["Total coûts réels"].fillna(0)) == 0, "OUI", "NON")
        df["_tw_num"] = pd.to_numeric(df.get("Type de travail", pd.Series(dtype=float)), errors="coerce")
        res['dfp'] = df

        # Taux realisation correctif
        filt_corr = (df["Nº appel pl.entret."].fillna(0) == 0) & (df["Contient SOPL"] == 1)
        an = cpiv(df, filt_corr, "Statut OT", posts)
        for c in ["CLOT", "CRÉÉ", "LANC", "TCLO"]:
            an[c] = an.get(c, 0)
        an["OT_CLOTURES"] = an["CLOT"] + an["TCLO"]
        an["TOTAL_OT"] = an[["CLOT", "CRÉÉ", "LANC", "TCLO"]].sum(axis=1)
        an["TAUX_REALISATION_CORRECTIF/PT"] = np.where(an["TOTAL_OT"] == 0, 100.0, ckpi(an["OT_CLOTURES"], an["TOTAL_OT"]))

        # Age preparation
        pr = cpiv(df, (df["Statut OT"] == "CRÉÉ") & (df["Statut utilisateur"].str.contains("CRPR", na=False)), "ap", posts)
        for c in ["<1 mois", ">3 mois", "1 mois < <3 mois", "Inconnu"]:
            pr[c] = pr.get(c, 0)
        pr["Total"] = pr[["<1 mois", "1 mois < <3 mois", ">3 mois", "Inconnu"]].sum(axis=1)
        pr["OT préparation <1 mois"] = ckpi(pr["<1 mois"], pr["Total"])
        pr["OT préparation >3 mois"] = ckpi(pr[">3 mois"], pr["Total"], 0)
        pr["OT préparation 1mois< <3mois"] = ckpi(pr["1 mois < <3 mois"], pr["Total"], 0)
        res['pr'] = pr

        # Age planification
        pl = cpiv(df, (df["Statut OT"] == "LANC") & (df["Statut utilisateur"].str.contains("ATPL", case=False, na=False)), "alp", posts)
        for c in ["<1 mois", ">3 mois", "1 mois < <3 mois", "Inconnu"]:
            pl[c] = pl.get(c, 0)
        pl["Total"] = pl[["<1 mois", "1 mois < <3 mois", ">3 mois", "Inconnu"]].sum(axis=1)
        pl["OT planification <1 mois"] = ckpi(pl["<1 mois"], pl["Total"])
        pl["OT planification >3 mois"] = ckpi(pl[">3 mois"], pl["Total"], 0)
        pl["OT planification 1mois< <3mois"] = ckpi(pl["1 mois < <3 mois"], pl["Total"], 0)
        res['pl'] = pl

        # Age execution
        ex = cpiv(df, (df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 1), "aex", posts)
        for c in ["<1 mois", ">3 mois", "1 mois < <3 mois", "Inconnu"]:
            ex[c] = ex.get(c, 0)
        ex["Total"] = ex[["<1 mois", "1 mois < <3 mois", ">3 mois", "Inconnu"]].sum(axis=1)
        ex["OT exécution <1 mois"] = ckpi(ex["<1 mois"], ex["Total"])
        ex["OT exécution >3 mois"] = ckpi(ex[">3 mois"], ex["Total"], 0)
        ex["OT exécution 1mois< <3mois"] = ckpi(ex["1 mois < <3 mois"], ex["Total"], 0)
        res['ex'] = ex

        # OT LANC ESTIME
        la = pd.pivot_table(df[df["Statut OT"] == "LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["OUI", "NON"]:
            la[c] = la.get(c, 0)
        la["Total"] = la["OUI"] + la["NON"]
        la["OT LANC ESTIME"] = ckpi(la["OUI"], la["Total"])

        # Backlog preparation caracterise
        pc = pd.pivot_table(df[df["Statut OT"] == "CRÉÉ"], index="Poste travail princ.", columns="Backlog preparation", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["CARACTERISE", "NON CARACTERISE"]:
            pc[c] = pc.get(c, 0)
        pc["Total"] = pc["CARACTERISE"] + pc["NON CARACTERISE"]
        pc["Backlog préparation caractérisé"] = ckpi(pc["CARACTERISE"], pc["Total"])
        res['pc'] = pc

        # Backlog planification caracterise
        plc = pd.pivot_table(df[df["Statut OT"] == "LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["CARACTERISE", "NON CARACTERISE"]:
            plc[c] = plc.get(c, 0)
        plc["Total"] = plc["CARACTERISE"] + plc["NON CARACTERISE"]
        plc["Backlog planification caractérisé"] = ckpi(plc["CARACTERISE"], plc["Total"])
        res['plc'] = plc

        # OT CONFIME et OT_COR_EGAL
        for kn, cn in [("OT CONFIME", "OT CONFIME"), ("OT_COR_EGAL", "OT_COR_EGAL")]:
            pv = pd.pivot_table(df, index="Poste travail princ.", columns=cn, values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
            for c in ["OUI", "NON"]:
                pv[c] = pv.get(c, 0)
            pv["Total"] = pv["OUI"] + pv["NON"]
            pv[cn] = ckpi(pv["OUI"], pv["Total"])
            res[kn.lower().replace(" ", "_")] = pv

        # Appel avis approuve
        avf = av[(av["Ordre"].isna()) | (av["Ordre"].astype(str).str.strip() == "")].copy()
        res['avf'] = avf
        tca = pd.pivot_table(avf, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["APRQ", "APRV", "APRV AVAU", "REJT"]:
            tca[c] = tca.get(c, 0)
        tca["Total"] = tca[["APRQ", "APRV", "APRV AVAU", "REJT"]].sum(axis=1)
        tca["appel avis approuvé"] = ckpi(tca["APRV"], tca["Total"])

        # Performance Graissage (Type 350)
        g_num = df[(df["Statut OT"].isin(["CLOT", "TCLO"])) & (df["_tw_num"] == 350)].groupby("Poste travail princ.")["Ordre"].count()
        g_den = df[(df["Contient SOPL"] == 1) & (df["_tw_num"] == 350)].groupby("Poste travail princ.")["Ordre"].count()
        g_df = pd.DataFrame({"_n": g_num, "_d": g_den}).reindex(posts, fill_value=0)
        g_df["Performance Graissage"] = np.where(g_df["_d"] == 0, 100.0, (g_df["_n"] / g_df["_d"]) * 100)

        # Performance Inspection (Types 290,300,310)
        ins_types = [290, 300, 310]
        ins_base = (df["_tw_num"].isin(ins_types)) & (df["Date de début planifiée"].notna()) & (df["Date de début planifiée"] <= now)
        ins_num = df[(df["Statut OT"].isin(["CLOT", "TCLO"])) & ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_den = df[(df["Contient SOPL"] == 1) & ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_df = pd.DataFrame({"_n": ins_num, "_d": ins_den}).reindex(posts, fill_value=0)
        ins_df["Performance Inspection"] = np.where(ins_df["_d"] == 0, 100.0, (ins_df["_n"] / ins_df["_d"]) * 100)

        # Performance Appels Systematiques (Type 360)
        sys_base = (df["_tw_num"] == 360) & (df["Date de début planifiée"].notna()) & (df["Date de début planifiée"] <= now)
        sys_num = df[(df["Statut OT"].isin(["CLOT", "TCLO"])) & sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_den = df[(df["Contient SOPL"] == 1) & sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_df = pd.DataFrame({"_n": sys_num, "_d": sys_den}).reindex(posts, fill_value=0)
        sys_df["Performance Appels Systématiques"] = np.where(sys_df["_d"] == 0, 100.0, (sys_df["_n"] / sys_df["_d"]) * 100)

        # Fiabilite et Avis de Panne : 0 si aucune donnee (pas de cellule vide)
        fiab_s = pd.Series(100.0, index=posts)
        avpan_s = pd.Series(100.0, index=posts)

        # Assemblage du DataFrame complet
        res['ckdf'] = pd.DataFrame({
            "TAUX_REALISATION_CORRECTIF/PT": an["TAUX_REALISATION_CORRECTIF/PT"],
            "OT préparation <1 mois": pr["OT préparation <1 mois"], "OT préparation >3 mois": pr["OT préparation >3 mois"], "OT préparation 1mois< <3mois": pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois": pl["OT planification <1 mois"], "OT planification >3 mois": pl["OT planification >3 mois"], "OT planification 1mois< <3mois": pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois": ex["OT exécution <1 mois"], "OT exécution >3 mois": ex["OT exécution >3 mois"], "OT exécution 1mois< <3mois": ex["OT exécution 1mois< <3mois"],
            "Performance Graissage": g_df["Performance Graissage"], "Performance Inspection": ins_df["Performance Inspection"], "Performance Appels Systématiques": sys_df["Performance Appels Systématiques"],
            "appel avis approuvé": tca["appel avis approuvé"], "OT LANC ESTIME": la["OT LANC ESTIME"],
            "Backlog préparation caractérisé": pc["Backlog préparation caractérisé"], "Backlog planification caractérisé": plc["Backlog planification caractérisé"],
            "OT CONFIME": res['ot_confime']["OT CONFIME"], "OT_COR_EGAL": res['ot_cor_egal']["OT_COR_EGAL"],
            "OT Fiabilité": fiab_s, "Total Avis de Panne": avpan_s
        })
        return res

    # --- Groupement des postes de travail ---
    def build_group_map(posts):
        sf1 = [p for p in posts if "SF1" in str(p).upper()]
        sf2 = [p for p in posts if "SF2" in str(p).upper()]
        autres = [p for p in posts if p not in sf1 and p not in sf2]
        gmap = {}
        if sf1:
            gmap["SF1"] = sf1
        if sf2:
            gmap["SF2"] = sf2
        if autres:
            gmap["Autres"] = autres
        return gmap

    def calc_group_score(ckdf, group_posts, kpi_list):
        if not group_posts:
            return 0.0
        sub = ckdf.loc[ckdf.index.isin(group_posts)]
        if sub.empty:
            return 0.0
        met_count = 0
        total_count = 0
        for _, row in sub.iterrows():
            for k in kpi_list:
                if k in row.index:
                    total_count += 1
                    if gscore(k, row[k], CIBLE.get(k, 100)):
                        met_count += 1
        return (met_count / total_count * 100) if total_count > 0 else 100.0

    # --- Rendu des cartes groupees ---
    def render_grouped_cards(ckdf, group_map):
        for grp_name, grp_posts in group_map.items():
            p_sc = calc_group_score(ckdf, grp_posts, QK)
            q_sc = calc_group_score(ckdf, grp_posts, PK)
            st.markdown('<div class="grp-label"><span class="grp-dot"></span>%s <span style="opacity:.5;font-weight:400;font-size:11px">%s</span></div>' % (grp_name, " / ".join(grp_posts)), unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                st.markdown('<div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Performance - %s</div></div>' % (p_sc, grp_name), unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="cc c1"><div class="cv">%.1f%%</div><div class="cl">Score Qualite - %s</div></div>' % (q_sc, grp_name), unsafe_allow_html=True)
            if len(grp_posts) > 1:
                mb_h = '<div class="ca" style="margin-top:4px;margin-bottom:8px"><div class="ct" style="font-size:12px">Detail par poste - %s</div>' % grp_name
                for p in grp_posts:
                    if p in ckdf.index:
                        ps = calc_group_score(ckdf, [p], QK)
                        qs = calc_group_score(ckdf, [p], PK)
                        mb_h += '<div class="mb-row"><span class="mb-name">%s</span><div class="mb-bars"><div class="mb-bar"><div class="mb-fill" style="width:%.1f%%;background:linear-gradient(90deg,#276749,#48bb78)"></div></div><div class="mb-bar"><div class="mb-fill" style="width:%.1f%%;background:linear-gradient(90deg,#2b6cb0,#4299e1)"></div></div></div><span class="mb-val" style="color:#276749">%.1f%%</span><span class="mb-val" style="color:#2b6cb0">%.1f%%</span></div>' % (p, min(ps, 100), min(qs, 100), ps, qs)
                mb_h += '</div>'
                st.markdown(mb_h, unsafe_allow_html=True)

    # --- Page Anomalies avec classement ---
    def render_anomalies(ckdf, posts):
        st.markdown('<div class="stl a"><i class="fas fa-exclamation-triangle"></i> Anomalies par Indicateur</div>', unsafe_allow_html=True)

        for section_name, kpi_list in [("Performance", QK), ("Qualite", PK)]:
            # Calcul du nombre d'anomalies par KPI x Poste
            ano_data = {}
            kpi_ano_total = {}
            for k in kpi_list:
                ano_data[k] = {}
                total_ano = 0
                for p in posts:
                    if p in ckdf.index and k in ckdf.columns:
                        # Pour les indicateurs de fiabilite bases sur les avis : 0 si aucune donnee
                        if k in ["OT Fiabilité", "Total Avis de Panne"]:
                            ano_data[k][p] = 0
                        else:
                            val = ckdf.loc[p, k]
                            ano_data[k][p] = 0 if gscore(k, val, CIBLE.get(k, 100)) else 1
                    else:
                        ano_data[k][p] = 0
                    total_ano += ano_data[k][p]
                kpi_ano_total[k] = total_ano

            # Classement par nombre d'anomalies decroissant
            sorted_kpis = sorted(kpi_list, key=lambda k: kpi_ano_total[k], reverse=True)

            # Construction du tableau HTML
            tc = "at"
            h = '<table class="tw %s"><thead><tr><th class="fc-head">Indicateur KPI</th>' % tc
            for p in posts:
                h += '<th>%s</th>' % p
            h += '<th class="sc-head">Total Anomalies</th></tr></thead><tbody>'
            for k in sorted_kpis:
                h += '<tr><td class="fc" style="text-align:left;white-space:normal;max-width:280px">%s</td>' % k
                tot = 0
                for p in posts:
                    n = ano_data[k][p]
                    tot += n
                    h += '<td class="%s" style="text-align:center">%d</td>' % (kas(n), n)
                h += '<td class="sc %s" style="text-align:center;font-weight:800">%d</td>' % (kas(tot), tot)
                h += '</tr>'
            # Total General
            g_tot = sum(kpi_ano_total.values())
            h += '<tr class="total-row"><td class="fc">Total General</td>'
            for p in posts:
                p_tot = sum(ano_data[k][p] for k in kpi_list)
                h += '<td style="text-align:center;font-weight:800">%d</td>' % p_tot
            h += '<td class="sc" style="text-align:center;font-weight:800">%d</td>' % g_tot
            h += '</tr></tbody></table>'
            st.markdown(h, unsafe_allow_html=True)

        # Resume par poste
        st.markdown('<div class="stl a"><i class="fas fa-map-marker-alt"></i> Resume des Anomalies par Poste</div>', unsafe_allow_html=True)
        poste_ano = {}
        for p in posts:
            poste_ano[p] = 0
            for k in ALL_KPI:
                if k in ["OT Fiabilité", "Total Avis de Panne"]:
                    continue
                if p in ckdf.index and k in ckdf.columns:
                    if not gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100)):
                        poste_ano[p] += 1
        sorted_postes = sorted(posts, key=lambda p: poste_ano[p], reverse=True)
        h2 = '<table class="tw at"><thead><tr><th class="fc-head">Poste de travail</th><th>Nb Anomalies</th><th>Indicateurs en anomalie</th><th class="sc-head">Severite</th></tr></thead><tbody>'
        for p in sorted_postes:
            n = poste_ano[p]
            kpis_ano = [k for k in ALL_KPI if k not in ["OT Fiabilité", "Total Avis de Panne"] and p in ckdf.index and k in ckdf.columns and not gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100))]
            sev = "Aucune" if n == 0 else ("Faible" if n <= 3 else ("Moderee" if n <= 6 else "Critique"))
            h2 += '<tr><td class="fc">%s</td><td class="%s" style="text-align:center;font-weight:800">%d</td><td style="text-align:left;font-size:11px;color:#4a5568;white-space:normal;max-width:450px">%s</td><td class="sc %s" style="text-align:center">%s</td></tr>' % (p, kas(n), n, ("; ".join(kpis_ano) if kpis_ano else '<span style="color:#059669;font-weight:600">Aucun</span>'), kas(n), sev)
        h2 += '</tbody></table>'
        st.markdown(h2, unsafe_allow_html=True)

    # --- Page Backlog ---
    def render_backlog(res, df, posts, group_map, text_col):
        st.markdown('<div class="stl bl"><i class="fas fa-layer-group"></i> OT OMS par Poste et Statut OT</div>', unsafe_allow_html=True)

        # Filtrage OMS
        if text_col:
            oms_df = df[df[text_col].astype(str).str.contains("OMS", case=False, na=False)]
        else:
            oms_df = pd.DataFrame()
        oms_piv = build_statut_pivot(oms_df, posts)
        st.markdown(html_statut_pivot(oms_piv, "bl"), unsafe_allow_html=True)
        if not oms_df.empty:
            show_pie_pair(oms_piv, "OMS")

        # Filtrage Thermographie
        st.markdown('<div class="stl bl"><i class="fas fa-temperature-high"></i> OT Thermographie par Poste et Statut OT</div>', unsafe_allow_html=True)
        if text_col:
            thermo_df = df[df[text_col].astype(str).str.contains("THERMO|THERMOGRAPHIE", case=False, na=False)]
        else:
            thermo_df = pd.DataFrame()
        thermo_piv = build_statut_pivot(thermo_df, posts)
        st.markdown(html_statut_pivot(thermo_piv, "bl"), unsafe_allow_html=True)
        if not thermo_df.empty:
            show_pie_pair(thermo_piv, "Thermographie")

        # Tous les OT par statut
        st.markdown('<div class="stl bl"><i class="fas fa-chart-bar"></i> Statistiques des Statuts OT - Tous Types</div>', unsafe_allow_html=True)
        all_piv = build_statut_pivot(df, posts)
        st.markdown(html_statut_pivot(all_piv, "bl"), unsafe_allow_html=True)
        if not df.empty:
            show_pie_pair(all_piv, "Tous OT")

        # Caracterisation Backlog Preparation
        st.markdown('<div class="stl bl"><i class="fas fa-tags"></i> Caracterisation Backlog Preparation</div>', unsafe_allow_html=True)
        pr = res.get('pr')
        if pr is not None:
            pr_display = pr[["<1 mois", "1 mois < <3 mois", ">3 mois", "Inconnu"]].copy()
            pr_display.columns = ["<1 mois", "1-3 mois", ">3 mois", "Inconnu"]
            st.markdown(html_statut_pivot(pr_display, "bl").replace("CRÉÉ", "<1 mois").replace("LANC", "1-3 mois").replace("CLOT", ">3 mois").replace("TCLO", "Inconnu").replace("Total General", "Total General"), unsafe_allow_html=True)
            # Graphique
            try:
                fig_prep = px.bar(pr_display, barmode='stack', title="Backlog Preparation - Repartition par Age",
                                  color_discrete_map={"<1 mois": "#059669", "1-3 mois": "#d97706", ">3 mois": "#dc2626", "Inconnu": "#94a3b8"})
                fig_prep.update_layout(margin=dict(t=40, b=10, l=10, r=10), height=350, legend=dict(orientation="h", yanchor="bottom", y=-0.15))
                st.plotly_chart(fig_prep, use_container_width=True)
            except Exception:
                pass

        # Caracterisation Backlog Planification
        st.markdown('<div class="stl bl"><i class="fas fa-tags"></i> Caracterisation Backlog Planification</div>', unsafe_allow_html=True)
        pl = res.get('pl')
        if pl is not None:
            pl_display = pl[["<1 mois", "1 mois < <3 mois", ">3 mois", "Inconnu"]].copy()
            pl_display.columns = ["<1 mois", "1-3 mois", ">3 mois", "Inconnu"]
            st.markdown(html_statut_pivot(pl_display, "bl"), unsafe_allow_html=True)
            try:
                fig_plan = px.bar(pl_display, barmode='stack', title="Backlog Planification - Repartition par Age",
                                  color_discrete_map={"<1 mois": "#059669", "1-3 mois": "#d97706", ">3 mois": "#dc2626", "Inconnu": "#94a3b8"})
                fig_plan.update_layout(margin=dict(t=40, b=10, l=10, r=10), height=350, legend=dict(orientation="h", yanchor="bottom", y=-0.15))
                st.plotly_chart(fig_plan, use_container_width=True)
            except Exception:
                pass

        # Caracterisation Backlog Execution
        st.markdown('<div class="stl bl"><i class="fas fa-tags"></i> Caracterisation Backlog Execution</div>', unsafe_allow_html=True)
        ex = res.get('ex')
        if ex is not None:
            ex_display = ex[["<1 mois", "1 mois < <3 mois", ">3 mois", "Inconnu"]].copy()
            ex_display.columns = ["<1 mois", "1-3 mois", ">3 mois", "Inconnu"]
            st.markdown(html_statut_pivot(ex_display, "bl"), unsafe_allow_html=True)
            try:
                fig_exec = px.bar(ex_display, barmode='stack', title="Backlog Execution - Repartition par Age",
                                  color_discrete_map={"<1 mois": "#059669", "1-3 mois": "#d97706", ">3 mois": "#dc2626", "Inconnu": "#94a3b8"})
                fig_exec.update_layout(margin=dict(t=40, b=10, l=10, r=10), height=350, legend=dict(orientation="h", yanchor="bottom", y=-0.15))
                st.plotly_chart(fig_exec, use_container_width=True)
            except Exception:
                pass

        # Graphique de synthese backlog
        st.markdown('<div class="stl bl"><i class="fas fa-chart-bar"></i> Synthese Backlog - Comparaison par Poste</div>', unsafe_allow_html=True)
        try:
            synth_data = []
            for p in posts:
                for label, src in [("Preparation", pr), ("Planification", pl), ("Execution", ex)]:
                    if src is not None and p in src.index:
                        synth_data.append({"Poste": p, "Type": label, "Nombre": int(src.loc[p, "Total"])})
            if synth_data:
                synth_df = pd.DataFrame(synth_data)
                fig_synth = px.bar(synth_df, x="Poste", y="Nombre", color="Type", barmode="group",
                                   title="Comparaison du Backlog par Poste et Type",
                                   color_discrete_map={"Preparation": "#0d9488", "Planification": "#7c3aed", "Execution": "#d97706"})
                fig_synth.update_layout(margin=dict(t=40, b=10, l=10, r=10), height=380, legend=dict(orientation="h", yanchor="bottom", y=-0.15))
                st.plotly_chart(fig_synth, use_container_width=True)
        except Exception:
            pass

    # --- Page Suivi (code conserve, non appele) ---
    def render_suivi(posts, ckdf, hist_df):
        st.markdown('<div class="stl s"><i class="fas fa-tasks"></i> Suivi des Ameliorations et Evolutions</div>', unsafe_allow_html=True)
        st.markdown('<div class="ca"><div class="ct">Journal des Evolutions</div>', unsafe_allow_html=True)
        h = '<div class="evol-timeline">'
        for entry in CHANGELOG:
            h += '<div class="evol-item"><div class="evol-ver">Version %s</div><div class="evol-date">%s</div>' % (entry["version"], entry["date"])
            for change in entry["changes"]:
                h += '<div class="evol-change">%s</div>' % change
            h += '</div>'
        h += '</div></div>'
        st.markdown(h, unsafe_allow_html=True)

    # ============================================================
    # UPLOAD FICHIERS ET TRAITEMENT
    # ============================================================
    uploaded_file = st.sidebar.file_uploader("Fichier OT (Excel)", type=["xlsx", "xls"], key="ot_file")
    avis_file = st.sidebar.file_uploader("Fichier Avis (Excel)", type=["xlsx", "xls"], key="av_file")

    if uploaded_file and avis_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            av_raw = pd.read_excel(avis_file)
        except Exception as e:
            st.error("Erreur de lecture des fichiers : %s" % str(e))
            st.stop()

        df = excr(df_raw)
        if "Poste travail princ." not in df.columns:
            st.error("Colonne 'Poste travail princ.' introuvable dans le fichier OT.")
            st.stop()

        posts = sorted(df["Poste travail princ."].dropna().unique().tolist())
        if not posts:
            st.error("Aucun poste de travail trouve.")
            st.stop()

        now = datetime.now()
        res = calc_kpis(df, av_raw, now, posts)
        ckdf = res['ckdf']
        group_map = build_group_map(posts)
        text_col = get_text_col(df)

        # Calcul des scores par poste
        perf_scores = {}
        qual_scores = {}
        for p in posts:
            if p in ckdf.index:
                met_p = sum(1 for k in QK if k in ckdf.columns and gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100)))
                met_q = sum(1 for k in PK if k in ckdf.columns and gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100)))
                perf_scores[p] = (met_p / len(QK)) * 100 if QK else 100
                qual_scores[p] = (met_q / len(PK)) * 100 if PK else 100

        # ===== ONGLETS =====
        # NOTE: L'onglet Suivi est masque mais le code est conserve ci-dessous
        tab_dash, tab_ano, tab_bl = st.tabs([
            "Performance & Qualite",
            "Anomalies",
            "Backlog"
        ])

        # ===== ONGLET DASHBOARD =====
        with tab_dash:
            st.markdown('<div class="mh"><h1><i class="fas fa-tachometer-alt"></i> Performance & Qualite</h1><span class="db"><i class="fas fa-calendar-alt"></i> %s</span></div>' % fichier_date, unsafe_allow_html=True)

            # Cartes resume global
            global_perf = calc_group_score(ckdf, posts, QK)
            global_qual = calc_group_score(ckdf, posts, PK)
            total_ot = len(df)
            taux_moyen = ckdf["TAUX_REALISATION_CORRECTIF/PT"].mean() if "TAUX_REALISATION_CORRECTIF/PT" in ckdf.columns else 0

            st.markdown('<div class="cr"><div class="cc c4"><div class="cv">%d</div><div class="cl">Total OT</div></div><div class="cc c4"><div class="cv">%.1f%%</div><div class="cl">Taux Real. Moyen</div></div><div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Performance Global</div></div><div class="cc c1"><div class="cv">%.1f%%</div><div class="cl">Score Qualite Global</div></div></div>' % (total_ot, taux_moyen, global_perf, global_qual), unsafe_allow_html=True)

            # Cartes groupees par groupe de postes
            render_grouped_cards(ckdf, group_map)

            # Tableau Performance
            st.markdown('<div class="stl p"><i class="fas fa-tachometer-alt"></i> Indicateurs de Performance</div>', unsafe_allow_html=True)
            p_cols = ["Poste de travail"] + QK + ["Score Performance"]
            p_rows = []
            for p in posts:
                if p in ckdf.index:
                    row = {"Poste de travail": p}
                    met_count = 0
                    for k in QK:
                        if k in ckdf.columns:
                            row[k] = "%.1f%%" % ckdf.loc[p, k]
                            if gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100)):
                                met_count += 1
                        else:
                            row[k] = "N/A"
                    sc_val = (met_count / len(QK)) * 100 if QK else 100
                    row["Score Performance"] = "%.2f%%" % sc_val
                    p_rows.append(row)
            # Ligne Cible
            cible_row = {"Poste de travail": "CIBLE", "_t": "cible"}
            for k in QK:
                cible_row[k] = "%.0f%%" % CIBLE.get(k, 100)
            cible_row["Score Performance"] = "100%"
            p_rows.append(cible_row)
            # Ligne Total General
            total_row = {"Poste de travail": "Total General", "_t": "total"}
            for k in QK:
                if k in ckdf.columns:
                    total_row[k] = "%.1f%%" % ckdf[k].mean()
                else:
                    total_row[k] = "N/A"
            avg_sc = np.mean([float(r["Score Performance"].replace("%", "")) for r in p_rows if r.get("Poste de travail") != "CIBLE" and r.get("Poste de travail") != "Total General"])
            total_row["Score Performance"] = "%.2f%%" % avg_sc
            p_rows.append(total_row)
            st.markdown(html_table(p_rows, p_cols, "", sc_col=["Score Performance"]), unsafe_allow_html=True)

            # Tableau Qualite
            st.markdown('<div class="stl q"><i class="fas fa-check-circle"></i> Indicateurs de Qualite</div>', unsafe_allow_html=True)
            q_cols = ["Poste de travail"] + PK + ["Score Qualite"]
            q_rows = []
            for p in posts:
                if p in ckdf.index:
                    row = {"Poste de travail": p}
                    met_count = 0
                    for k in PK:
                        if k in ckdf.columns:
                            row[k] = "%.1f%%" % ckdf.loc[p, k]
                            if gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100)):
                                met_count += 1
                        else:
                            row[k] = "N/A"
                    sc_val = (met_count / len(PK)) * 100 if PK else 100
                    row["Score Qualite"] = "%.2f%%" % sc_val
                    q_rows.append(row)
            cible_row_q = {"Poste de travail": "CIBLE", "_t": "cible"}
            for k in PK:
                cible_row_q[k] = "%.0f%%" % CIBLE.get(k, 100)
            cible_row_q["Score Qualite"] = "100%"
            q_rows.append(cible_row_q)
            total_row_q = {"Poste de travail": "Total General", "_t": "total"}
            for k in PK:
                if k in ckdf.columns:
                    total_row_q[k] = "%.1f%%" % ckdf[k].mean()
                else:
                    total_row_q[k] = "N/A"
            avg_sc_q = np.mean([float(r["Score Qualite"].replace("%", "")) for r in q_rows if r.get("Poste de travail") != "CIBLE" and r.get("Poste de travail") != "Total General"])
            total_row_q["Score Qualite"] = "%.2f%%" % avg_sc_q
            q_rows.append(total_row_q)
            st.markdown(html_table(q_rows, q_cols, "", sc_col=["Score Qualite"]), unsafe_allow_html=True)

            # Barres de progression
            st.markdown('<div class="stl p"><i class="fas fa-chart-bar"></i> Progression Performance</div>', unsafe_allow_html=True)
            perf_actuals = {k: ckdf[k].mean() for k in QK if k in ckdf.columns}
            st.markdown(html_kpi_bars(QK, perf_actuals, CIBLE, "Indicateurs de Performance - Moyenne Globale", "#38a169", "#e53e3e"), unsafe_allow_html=True)

            st.markdown('<div class="stl q"><i class="fas fa-chart-bar"></i> Progression Qualite</div>', unsafe_allow_html=True)
            qual_actuals = {k: ckdf[k].mean() for k in PK if k in ckdf.columns}
            st.markdown(html_kpi_bars(PK, qual_actuals, CIBLE, "Indicateurs de Qualite - Moyenne Globale", "#3182ce", "#e53e3e"), unsafe_allow_html=True)

            # Classement
            all_scores = {}
            for p in posts:
                all_scores[p] = (perf_scores.get(p, 0) + qual_scores.get(p, 0)) / 2
            st.markdown(html_classement(all_scores, "#2b6cb0"), unsafe_allow_html=True)

            # Barres groupees
            st.markdown(html_grouped_bars(posts, perf_scores, qual_scores, "Comparaison Performance / Qualite par Poste"), unsafe_allow_html=True)

            # Actions recommandees
            st.markdown('<div class="stl a"><i class="fas fa-lightbulb"></i> Actions Recommandees</div>', unsafe_allow_html=True)
            st.markdown(html_actions_table(ALL_KPI, perf_actuals.update(qual_actuals) or perf_actuals, CIBLE, ACT_MAP), unsafe_allow_html=True)

            # Export Excel
            p_rows_export = [{"Poste de travail": r["Poste de travail"], **{k: r.get(k, "") for k in QK}, "Score Performance": r.get("Score Performance", "")} for r in p_rows if r.get("_t") != "cible"]
            q_rows_export = [{"Poste de travail": r["Poste de travail"], **{k: r.get(k, "") for k in PK}, "Score Qualite": r.get("Score Qualite", "")} for r in q_rows if r.get("_t") != "cible"]
            # Anomalies pour export
            ano_p_rows, ano_p_cols = [], ["Poste de travail"] + QK
            ano_q_rows, ano_q_cols = [], ["Poste de travail"] + PK
            for p in posts:
                if p in ckdf.index:
                    row_p = {"Poste de travail": p}
                    row_q = {"Poste de travail": p}
                    for k in QK:
                        row_p[k] = 0 if (k in ["OT Fiabilité", "Total Avis de Panne"]) else (0 if gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100)) else 1)
                    for k in PK:
                        row_q[k] = 0 if (k in ["OT Fiabilité", "Total Avis de Panne"]) else (0 if gscore(k, ckdf.loc[p, k], CIBLE.get(k, 100)) else 1)
                    ano_p_rows.append(row_p)
                    ano_q_rows.append(row_q)

            if st.sidebar.button("Exporter Excel", key="export_btn"):
                save_kpis_to_excel(p_rows_export, p_cols, q_rows_export, q_cols,
                                   ano_p_rows, ano_p_cols, ano_q_rows, ano_q_cols, fichier_date)
                st.sidebar.success("Fichier exporte dans kpis/indicateurs_kpis.xlsx")

        # ===== ONGLET ANOMALIES =====
        with tab_ano:
            st.markdown('<div class="mh" style="background:linear-gradient(135deg,#991b1b,#dc2626)"><h1><i class="fas fa-exclamation-triangle"></i> Anomalies</h1><span class="db"><i class="fas fa-calendar-alt"></i> %s</span></div>' % fichier_date, unsafe_allow_html=True)
            render_anomalies(ckdf, posts)

        # ===== ONGLET BACKLOG =====
        with tab_bl:
            st.markdown('<div class="mh" style="background:linear-gradient(135deg,#4c1d95,#7c3aed)"><h1><i class="fas fa-layer-group"></i> Backlog</h1><span class="db"><i class="fas fa-calendar-alt"></i> %s</span></div>' % fichier_date, unsafe_allow_html=True)
            render_backlog(res, df, posts, group_map, text_col)

        # ====================================================================
        # SUIVI DES AMELIORATIONS - CODE CONSERVE POUR REACTIVATION ULTERIEURE
        # Pour reactiver, decommenter le bloc ci-dessous et ajouter l'onglet
        # dans st.tabs() ci-dessus :
        #   tab_dash, tab_ano, tab_bl, tab_suivi = st.tabs([...,"Suivi"])
        # ====================================================================
        # with tab_suivi:
        #     st.markdown('<div class="mh" style="background:linear-gradient(135deg,#975a16,#d69e2e)"><h1><i class="fas fa-tasks"></i> Suivi des Ameliorations et Evolutions</h1><span class="db"><i class="fas fa-calendar-alt"></i> %s</span></div>' % fichier_date, unsafe_allow_html=True)
        #     render_suivi(posts, ckdf, hist_df)
        # ====================================================================

    else:
        st.info("Veuillez charger les deux fichiers Excel (OT et Avis) dans la barre laterale pour afficher le dashboard.")


if __name__ == "__main__":
    main()
