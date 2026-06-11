# -*- coding: utf-8 -*-
"""app.py - KPI Dashboard MC et FEED"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import locale
import altair as alt
import random
import os
from datetime import datetime

# ==================================================
# CONFIGURATION DE LA PAGE (FULL WIDTH)
# ==================================================
st.set_page_config(
    page_title="KPI Dashboard MC et FEED",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================================================
# CONFIGURATION LOCALE
# ==================================================
def set_french_locale():
    for loc in ['fr_FR.UTF-8', 'fr_FR', 'french']:
        try:
            locale.setlocale(locale.LC_ALL, loc)
            return
        except locale.Error:
            continue

# ==================================================
# STYLES CSS PERSONNALISÉS (FULL WIDTH)
# ==================================================
CUSTOM_CSS = """
<style>
    .main .block-container {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    .header-kpi {
        background: linear-gradient(135deg, #1e3a5f 0%, #2980b9 100%);
        color: white;
        padding: 18px 30px;
        border-radius: 12px;
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin: 15px 0;
        box-shadow: 0 4px 15px rgba(41, 128, 185, 0.3);
        letter-spacing: 1px;
    }
    .header-anomalie {
        background: linear-gradient(135deg, #922b21 0%, #e74c3c 100%);
        color: white;
        padding: 18px 30px;
        border-radius: 12px;
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        margin: 15px 0;
        box-shadow: 0 4px 15px rgba(231, 76, 60, 0.3);
        letter-spacing: 1px;
    }
    .hse-splash {
        background: linear-gradient(180deg, #f8f9fa 0%, #dee2e6 100%);
        padding: 50px;
        border-radius: 20px;
        text-align: center;
        margin: 30px auto;
        max-width: 900px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
    }
    .hse-warning-box {
        background: linear-gradient(135deg, #fff9e6 0%, #fff3cd 100%);
        border-left: 12px solid #ffc107;
        padding: 40px;
        border-radius: 0 15px 15px 0;
        font-size: 30px;
        font-weight: bold;
        text-align: center;
        margin: 25px 0 35px 0;
        box-shadow: 0 4px 15px rgba(255, 193, 7, 0.2);
    }
    .hse-motto {
        color: #198754;
        font-size: 38px;
        font-weight: 800;
        text-align: center;
        margin-top: 30px;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }
    .section-title {
        font-size: 18px;
        font-weight: 600;
        color: #2c3e50;
        padding: 10px 0 5px 0;
        border-bottom: 2px solid #3498db;
        margin-bottom: 15px;
    }
    div[data-testid="stTable"] {
        font-size: 13px;
    }
    div[data-testid="stDataFrame"] {
        font-size: 13px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 10px 25px;
        background-color: #f1f3f5;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2980b9 !important;
        color: white !important;
    }
    .synthese-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 10px;
        padding: 15px;
        margin: 8px 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 4px solid #2980b9;
    }
</style>
"""

# ==================================================
# CONSIGNES HSE
# ==================================================
CONSIGNES_HSE = [
    "Port obligatoire des EPI avant toute intervention.",
    "Port obligatoire du casque de sécurité.",
    "Port obligatoire des lunettes de protection.",
    "Port obligatoire des gants adaptés au travail.",
    "Utiliser les protections auditives dans les zones bruyantes.",
    "Vérifier l'absence de tension avant toute intervention électrique.",
    "Respecter la procédure de consignation et déconsignation.",
    "Ne jamais intervenir sur un équipement en marche.",
    "Baliser et sécuriser la zone de travail.",
    "Maintenir le poste de travail propre et ordonné.",
    "Vérifier l'état des outils avant utilisation.",
    "Utiliser uniquement du matériel homologué.",
    "Respecter les permis de travail en vigueur.",
    "Identifier les risques avant de commencer une tâche.",
    "Signaler immédiatement toute situation dangereuse.",
    "Signaler tout incident ou presque accident.",
    "Ne jamais neutraliser un dispositif de sécurité.",
    "Vérifier les détecteurs de gaz avant utilisation.",
    "Vérifier la bonne ventilation des zones de travail.",
    "Respecter les règles des espaces confinés.",
    "Contrôler l'atmosphère avant d'entrer dans un espace confiné.",
    "Utiliser les points d'ancrage pour les travaux en hauteur.",
    "Vérifier l'état des échafaudages avant utilisation.",
    "Sécuriser les outils lors des travaux en hauteur.",
    "Ne pas travailler seul lors d'opérations à risque.",
    "Contrôler les élingues avant chaque levage.",
    "Respecter les limites de charge des équipements.",
    "Vérifier l'état des appareils de levage.",
    "Maintenir les voies de circulation dégagées.",
    "Respecter la signalisation de sécurité.",
    "Vérifier les extincteurs à proximité du chantier.",
    "Connaître les issues de secours les plus proches.",
    "Respecter les procédures d'arrêt d'urgence.",
    "Vérifier les flexibles et raccords avant mise en service.",
    "Contrôler les fuites avant démarrage d'un équipement.",
    "Respecter les distances de sécurité.",
    "Ne jamais contourner une procédure HSE.",
    "Porter les EPI adaptés au risque identifié.",
    "Prévenir son responsable avant toute intervention particulière.",
    "Analyser les risques avant chaque démarrage de chantier.",
    "Vérifier la stabilité des équipements.",
    "Utiliser les bons outils pour la bonne tâche.",
    "Respecter les consignes spécifiques du chantier.",
    "Ne jamais prendre de raccourci au détriment de la sécurité.",
    "Arrêter immédiatement les travaux en cas de danger.",
    "Protéger l'environnement lors des interventions.",
    "Collecter et trier correctement les déchets.",
    "Éviter toute pollution accidentelle.",
    "Respecter les consignes de stockage des produits dangereux.",
    "Lire les fiches de sécurité avant manipulation.",
    "Vérifier les équipements avant chaque prise de poste.",
    "S'assurer de la disponibilité des moyens de secours.",
    "Communiquer clairement avec l'équipe avant intervention.",
    "Respecter les règles de circulation des engins.",
    "Garder une vigilance permanente sur son environnement.",
    "Prendre le temps d'effectuer le travail en sécurité.",
    "La sécurité est l'affaire de tous.",
    "Chaque incident peut être évité par la prévention.",
    "Aucun travail n'est plus urgent que la sécurité.",
    "Zéro accident commence par un comportement sûr."
]

# ==================================================
# CONSTANTES
# ==================================================
MOTS_PREPARATION = ["CRPR ATPD", "CRPR ATMR", "CRPR ATER", "CRPR ATRS", "CRPR ATMO", 
                    "ATPD", "ATMR", "ATER", "ATRS", "ATMO"]
MOTS_PLANIFICATION = ["ATPL ATEI", "ATPL ATAL", "ATPL ATER", "ATPL AGAR", "ATPL ATHS", 
                      "ATEI", "ATAL", "ATAS", "AGAR", "ATHS"]

CIBLES_KPI = {
    "TAUX_REALISATION_CORRECTIF/PT": 85,
    "OT préparation <1 mois": 80,
    "OT préparation >3 mois": 5,
    "OT préparation 1mois< <3mois": 15,
    "OT planification <1 mois": 80,
    "OT planification >3 mois": 5,
    "OT planification 1mois< <3mois": 15,
    "OT exécution <1 mois": 80,
    "OT exécution >3 mois": 5,
    "OT exécution 1mois< <3mois": 15,
    "appel avis approuvé": 95,
    "OT LANC ESTIME": 100,
    "Backlog préparation caractérisé": 100,
    "Backlog planification caractérisé": 100,
    "OT CONFIME": 100,
    "OT_COR_EGAL": 100,
}

QTY_KPIS = ["TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois", 
            "OT préparation 1mois< <3mois", "OT planification <1 mois", "OT planification >3 mois", 
            "OT planification 1mois< <3mois", "OT exécution <1 mois", "OT exécution >3 mois", 
            "OT exécution 1mois< <3mois"]

QUAL_KPIS = ["appel avis approuvé", "OT LANC ESTIME", "Backlog préparation caractérisé", 
             "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]

# ==================================================
# FONCTIONS UTILITAIRES
# ==================================================
def rename_safe(df, old_names, new_names):
    mapping = {old: new for old, new in zip(old_names, new_names) if old in df.columns}
    return df.rename(columns=mapping)


def contient_mot(texte, liste_mots):
    texte = str(texte)
    return any(mot in texte for element in liste_mots for mot in element.split())


def categorie_age(age):
    if pd.isna(age):
        return "Inconnu"
    if age <= 1:
        return "<1 mois"
    elif age >= 3:
        return ">3 mois"
    else:
        return "1 mois < <3 mois"


def calcul_kpi(numerateur, denominateur, si_zero=100):
    return np.where(denominateur == 0, si_zero, (numerateur / denominateur) * 100)


def creer_pivot(dataframe, filtre, colonne, postes_to_reindex):
    pivot = pd.pivot_table(
        dataframe[filtre], 
        index="Poste travail princ.", 
        columns=colonne, 
        values="Ordre", 
        aggfunc="count", 
        fill_value=0
    )
    return pivot.reindex(postes_to_reindex, fill_value=0)


def exclure_cresseurs(df, colonne="Poste travail princ."):
    if df.empty:
        return df
    return df[~df[colonne].astype(str).str.upper().str.contains("CRESSEUR", na=False)]


def get_groupe_metier(poste):
    p = str(poste).upper()
    if "E" in p: return "Électrique"
    elif "M" in p: return "Mécanique"
    elif "R" in p: return "Instrumentation"
    elif "G" in p: return "Génie Civil"
    return "Autre"


def get_groupe_atelier(poste):
    p = str(poste).upper()
    if "PS" in p: return "Sulfurique"
    elif "PP" in p: return "Phosphorique"
    elif "TSP" in p or "REX" in p: return "Engrais"
    elif "MCP" in p or "DCP" in p: return "Feed"
    return "Autre"


def get_groupe_division(poste):
    p = str(poste).upper()
    if "SF1" in p: return "SF1"
    elif "SF2" in p: return "SF2"
    return "Autre"


def get_kpi_score(kpi_name, actual_value, target_value):
    if pd.isna(actual_value) or pd.isna(target_value):
        return 0
    rules = {
        "OT préparation <1 mois": lambda a, t: a >= 75,
        "OT planification <1 mois": lambda a, t: a >= 75,
        "OT exécution <1 mois": lambda a, t: a >= 75,
        "OT préparation 1mois< <3mois": lambda a, t: a <= 15,
        "OT planification 1mois< <3mois": lambda a, t: a <= 15,
        "OT exécution 1mois< <3mois": lambda a, t: a <= 15,
        "OT préparation >3 mois": lambda a, t: a <= 5,
        "OT planification >3 mois": lambda a, t: a <= 5,
        "OT exécution >3 mois": lambda a, t: a <= 5,
        "TAUX_REALISATION_CORRECTIF/PT": lambda a, t: a >= 80,
        "appel avis approuvé": lambda a, t: a >= 90,
        "OT LANC ESTIME": lambda a, t: a >= 95,
        "Backlog préparation caractérisé": lambda a, t: a >= 95,
        "Backlog planification caractérisé": lambda a, t: a >= 95,
        "OT CONFIME": lambda a, t: a >= 95,
        "OT_COR_EGAL": lambda a, t: a >= 95,
    }
    if kpi_name in rules:
        return 1 if rules[kpi_name](actual_value, target_value) else 0
    return 0


# ==================================================
# CALCUL DES KPIs
# ==================================================
def calculate_all_kpis_and_intermediate_dfs(df_input, avis_df_input, now_timestamp, all_postes_list):
    results = {}
    df_temp = df_input.copy()
    avis_df_temp = avis_df_input.copy()

    df_temp["Backlog préparation"] = np.where(
        df_temp["Statut utilisateur"].apply(lambda x: contient_mot(x, MOTS_PREPARATION)),
        "CARACTERISE", "NON CARACTERISE"
    )
    df_temp["Backlog planification"] = np.where(
        df_temp["Statut utilisateur"].apply(lambda x: contient_mot(x, MOTS_PLANIFICATION)),
        "CARACTERISE", "NON CARACTERISE"
    )

    age_configs = [
        ('Créé le', "Age mois préparation", "Age préparation"),
        ('Date de début planifiée', "Age mois planification", "Age planification"),
        ('Date de début planifiée', "Age mois exécution", "Age exécution"),
    ]
    
    for date_col, age_mois_col, age_col in age_configs:
        if date_col in df_temp.columns:
            df_temp[date_col] = pd.to_datetime(df_temp[date_col], errors='coerce')
            df_temp[age_mois_col] = (
                (now_timestamp.year - df_temp[date_col].dt.year) * 12 + 
                (now_timestamp.month - df_temp[date_col].dt.month)
            ).round(2)
            df_temp[age_col] = df_temp[age_mois_col].apply(categorie_age)
        else:
            df_temp[age_mois_col] = np.nan
            df_temp[age_col] = "Inconnu"

    df_temp["OT CONFIME"] = np.where(
        df_temp["Statut système"].str.contains("CLO", na=False) & 
        df_temp["Statut système"].str.contains("CONF", na=False),
        "OUI", "NON"
    )
    df_temp["Contient SOPL"] = df_temp["Statut utilisateur"].str.contains("SOPL", na=False).map({True: 1, False: 0})
    df_temp["OT LANC ESTIME"] = np.where(df_temp["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
    df_temp["OT_COR_EGAL"] = np.where(
        (df_temp["Total coûts budgétés"].fillna(0) - df_temp["Total coûts réels"].fillna(0)) == 0,
        "OUI", "NON"
    )
    
    results['df_processed'] = df_temp

    analyse = creer_pivot(df_temp, df_temp["Nº appel pl.entret."].fillna(0) == 0, "Statut OT", all_postes_list)
    for col in ["CLOT", "CRÉÉ", "LANC", "TCLO"]:
        analyse[col] = analyse.get(col, 0)
    analyse["Total"] = analyse[["CLOT", "CRÉÉ", "LANC", "TCLO"]].sum(axis=1)
    analyse["TAUX_REALISATION_CORRECTIF/PT"] = calcul_kpi(analyse["TCLO"], analyse["Total"])

    prep = creer_pivot(df_temp, df_temp["Statut OT"] == "CRÉÉ", "Age préparation", all_postes_list)
    for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]:
        prep[col] = prep.get(col, 0)
    prep["Total"] = prep[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    prep["OT préparation <1 mois"] = calcul_kpi(prep["<1 mois"], prep["Total"])
    prep["OT préparation >3 mois"] = calcul_kpi(prep[">3 mois"], prep["Total"], 0)
    prep["OT préparation 1mois< <3mois"] = calcul_kpi(prep["1 mois < <3 mois"], prep["Total"], 0)

    planif = creer_pivot(
        df_temp, 
        (df_temp["Statut OT"] == "LANC") & (df_temp["Contient SOPL"] == 0), 
        "Age planification", 
        all_postes_list
    )
    for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]:
        planif[col] = planif.get(col, 0)
    planif["Total"] = planif[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    planif["OT planification <1 mois"] = calcul_kpi(planif["<1 mois"], planif["Total"])
    planif["OT planification >3 mois"] = calcul_kpi(planif[">3 mois"], planif["Total"], 0)
    planif["OT planification 1mois< <3mois"] = calcul_kpi(planif["1 mois < <3 mois"], planif["Total"], 0)

    execu = creer_pivot(
        df_temp, 
        (df_temp["Statut OT"] == "LANC") & (df_temp["Contient SOPL"] == 1), 
        "Age exécution", 
        all_postes_list
    )
    for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]:
        execu[col] = execu.get(col, 0)
    execu["Total"] = execu[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    execu["OT exécution <1 mois"] = calcul_kpi(execu["<1 mois"], execu["Total"])
    execu["OT exécution >3 mois"] = calcul_kpi(execu[">3 mois"], execu["Total"], 0)
    execu["OT exécution 1mois< <3mois"] = calcul_kpi(execu["1 mois < <3 mois"], execu["Total"], 0)

    lanc = pd.pivot_table(
        df_temp[df_temp["Statut OT"] == "LANC"], 
        index="Poste travail princ.", 
        columns="OT LANC ESTIME", 
        values="Ordre", 
        aggfunc="count", 
        fill_value=0
    ).reindex(all_postes_list, fill_value=0)
    for col in ["OUI", "NON"]:
        lanc[col] = lanc.get(col, 0)
    lanc["Total"] = lanc["OUI"] + lanc["NON"]
    lanc["OT LANC ESTIME"] = calcul_kpi(lanc["OUI"], lanc["Total"])

    prep_carac = pd.pivot_table(
        df_temp[df_temp["Statut OT"] == "CRÉÉ"], 
        index="Poste travail princ.", 
        columns="Backlog préparation", 
        values="Ordre", 
        aggfunc="count", 
        fill_value=0
    ).reindex(all_postes_list, fill_value=0)
    for col in ["CARACTERISE", "NON CARACTERISE"]:
        prep_carac[col] = prep_carac.get(col, 0)
    prep_carac["Total"] = prep_carac["CARACTERISE"] + prep_carac["NON CARACTERISE"]
    prep_carac["Backlog préparation caractérisé"] = calcul_kpi(prep_carac["CARACTERISE"], prep_carac["Total"])

    planif_carac = pd.pivot_table(
        df_temp[df_temp["Statut OT"] == "LANC"], 
        index="Poste travail princ.", 
        columns="Backlog planification", 
        values="Ordre", 
        aggfunc="count", 
        fill_value=0
    ).reindex(all_postes_list, fill_value=0)
    for col in ["CARACTERISE", "NON CARACTERISE"]:
        planif_carac[col] = planif_carac.get(col, 0)
    planif_carac["Total"] = planif_carac["CARACTERISE"] + planif_carac["NON CARACTERISE"]
    planif_carac["Backlog planification caractérisé"] = calcul_kpi(planif_carac["CARACTERISE"], planif_carac["Total"])

    for kpi_name, col_name in [("OT CONFIME", "OT CONFIME"), ("OT_COR_EGAL", "OT_COR_EGAL")]:
        df_pivot = pd.pivot_table(
            df_temp, 
            index="Poste travail princ.", 
            columns=col_name, 
            values="Ordre", 
            aggfunc="count", 
            fill_value=0
        ).reindex(all_postes_list, fill_value=0)
        for c in ["OUI", "NON"]:
            df_pivot[c] = df_pivot.get(c, 0)
        df_pivot["Total"] = df_pivot["OUI"] + df_pivot["NON"]
        df_pivot[col_name] = calcul_kpi(df_pivot["OUI"], df_pivot["Total"])
        results[kpi_name.lower().replace(" ", "_")] = df_pivot

    avis_df_filtered_local = avis_df_temp[
        (avis_df_temp["Ordre"].isna()) | (avis_df_temp["Ordre"].astype(str).str.strip() == "")
    ].copy()
    results['avis_df_filtered'] = avis_df_filtered_local

    tableau_croise_avis = pd.pivot_table(
        avis_df_filtered_local, 
        index="Poste travail princ.", 
        columns="Statut utilisateur", 
        values="Avis", 
        aggfunc="count", 
        fill_value=0
    ).reindex(all_postes_list, fill_value=0)
    for col in ["APRQ", "APRV", "APRV AVAU", "REJT"]:
        tableau_croise_avis[col] = tableau_croise_avis.get(col, 0)
    tableau_croise_avis["Total"] = tableau_croise_avis[["APRQ", "APRV", "APRV AVAU", "REJT"]].sum(axis=1)
    tableau_croise_avis["appel avis approuvé"] = calcul_kpi(tableau_croise_avis["APRV"], tableau_croise_avis["Total"])

    calculated_kpis_df = pd.concat([
        analyse[["TAUX_REALISATION_CORRECTIF/PT"]],
        prep[["OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois"]],
        planif[["OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois"]],
        execu[["OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]],
        tableau_croise_avis[["appel avis approuvé"]],
        lanc[["OT LANC ESTIME"]],
        prep_carac[["Backlog préparation caractérisé"]],
        planif_carac[["Backlog planification caractérisé"]],
        results['ot_confime'][["OT CONFIME"]],
        results['ot_cor_egal'][["OT_COR_EGAL"]]
    ], axis=1)

    results['calculated_kpis_df'] = calculated_kpis_df
    return results


# ==================================================
# MISE EN FORME CONDITIONNELLE
# ==================================================
def highlight_kpis(row):
    styles = []
    if row.name == 'CIBLE':
        return ['padding: 6px; background-color: #d6eaf8; color: #1a5276; font-weight: bold; text-align: center;'] * len(row)
    if row.name == 'Total général':
        return ['padding: 6px; background-color: #e8daef; color: #4a235a; font-weight: bold; text-align: center;'] * len(row)
    
    for col in row.index:
        current_style = 'padding: 6px; text-align: center;'
        try:
            value = float(row[col])
        except (ValueError, TypeError):
            styles.append(current_style)
            continue
        
        if col in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
            if value >= 80: current_style += 'background-color:#c6efce; color:#006100; font-weight:600;'
            elif value >= 75: current_style += 'background-color:#ffeb9c; color:#9c6500; font-weight:600;'
            else: current_style += 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
        elif col in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
            if value <= 15: current_style += 'background-color:#c6efce; color:#006100; font-weight:600;'
            else: current_style += 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
        elif col in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
            if value <= 5: current_style += 'background-color:#c6efce; color:#006100; font-weight:600;'
            else: current_style += 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
        elif col == "TAUX_REALISATION_CORRECTIF/PT":
            if value >= 85: current_style += 'background-color:#c6efce; color:#006100; font-weight:600;'
            elif value >= 80: current_style += 'background-color:#ffeb9c; color:#9c6500; font-weight:600;'
            else: current_style += 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
        elif col == "appel avis approuvé":
            if value >= 95: current_style += 'background-color:#c6efce; color:#006100; font-weight:600;'
            elif value >= 90: current_style += 'background-color:#ffeb9c; color:#9c6500; font-weight:600;'
            else: current_style += 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
        elif col in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
            if value >= 100: current_style += 'background-color:#c6efce; color:#006100; font-weight:600;'
            elif value >= 95: current_style += 'background-color:#ffeb9c; color:#9c6500; font-weight:600;'
            else: current_style += 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
        
        styles.append(current_style)
    return styles


def highlight_classification_table_kpis(row):
    styles = [''] * len(row)
    for i, col in enumerate(row.index):
        if col in ["Score KPIs Quantité", "Score KPIs Qualité", "Total performance "]:
            try:
                s = str(row[col])
                value = float(s.replace(' %', '').strip()) if ' %' in s else float(s)
                if value >= 90: styles[i] = 'background-color: #c6efce; color: #006100; font-weight: 600;'
                elif value >= 80: styles[i] = 'background-color: #ffeb9c; color: #9c6500; font-weight: 600;'
                else: styles[i] = 'background-color: #ffc7ce; color: #9c0006; font-weight: 600;'
            except (ValueError, TypeError):
                pass
        elif col == "Poste travail princ.":
            styles[i] = 'font-weight: 600;'
    return styles


def highlight_synthese(row):
    styles = [''] * len(row)
    for i, col in enumerate(row.index):
        styles[i] = 'text-align: center; padding: 8px;'
        if col == "KPI":
            styles[i] += 'font-weight: 600; text-align: left; padding-left: 15px;'
        elif col == "Action Suggérée":
            styles[i] += 'text-align: left; padding-left: 15px;'
        elif col == "Total OT / Avis impactés":
            styles[i] += 'font-weight: 700; font-size: 16px; color: #c0392b;'
        elif col == "Cible (%)":
            styles[i] += 'font-weight: 600; color: #2980b9;'
    return styles


# ==================================================
# ÉCRAN HSE
# ==================================================
def afficher_ecran_hse():
    consigne = random.choice(CONSIGNES_HSE)
    
    st.markdown("""
    <div class="hse-splash">
        <h1 style="font-size: 48px; color: #0f172a; margin-bottom: 10px;">
            🦺 HSE - CONSIGNE DE SÉCURITÉ
        </h1>
        <h2 style="color: #64748b; font-weight: normal; font-size: 22px;">
            Sécurité • Santé • Environnement
        </h2>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="hse-warning-box">
        ⚠️ {consigne}
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="hse-motto">
        Aucun travail n'est plus urgent que la sécurité
    </div>
    """, unsafe_allow_html=True)
    
    col_center = st.columns([1, 2, 1])[1]
    with col_center:
        if st.button("✅ J'ai lu la consigne - Accéder au Dashboard", type="primary", use_container_width=True):
            st.session_state.hse_affiche = True
            st.rerun()


# ==================================================
# DÉTECTION DES ANOMALIES
# ==================================================
def detecter_anomalies(df_processed, calculated_kpis_df, valid_postes, cible_df, avis_df_filtered):
    anomalies_ot_records = []
    anomalies_avis_records = []

    for poste in valid_postes:
        if poste not in df_processed["Poste travail princ."].values:
            continue
        
        df_poste = df_processed[df_processed["Poste travail princ."] == poste]
        avis_poste = avis_df_filtered[avis_df_filtered["Poste travail princ."] == poste]

        # TAUX_REALISATION
        val_real = calculated_kpis_df.loc[poste, "TAUX_REALISATION_CORRECTIF/PT"] if poste in calculated_kpis_df.index else 100
        if pd.notna(val_real) and val_real < cible_df.loc['CIBLE', "TAUX_REALISATION_CORRECTIF/PT"]:
            count_anom = len(df_poste[(df_poste["Nº appel pl.entret."].fillna(0) == 0) & (~df_poste["Statut OT"].isin(["CLOT", "TCLO"]))])
            if count_anom > 0:
                anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "TAUX_REALISATION_CORRECTIF/PT", "Nb OT impactés": count_anom, "Action Suggérée": "Améliorer le taux de réalisation des OT."})

        # KPIs âge
        kpi_age_checks = [
            ("OT préparation <1 mois", "CRÉÉ", None, "Age préparation", "!=", "<1 mois", "Réduire l'âge de préparation des OT."),
            ("OT préparation >3 mois", "CRÉÉ", None, "Age préparation", "==", ">3 mois", "Traiter les OT ayant un âge de préparation > 3 mois."),
            ("OT planification <1 mois", "LANC", 0, "Age planification", "!=", "<1 mois", "Réduire l'âge de planification des OT."),
            ("OT planification >3 mois", "LANC", 0, "Age planification", "==", ">3 mois", "Traiter les OT ayant un âge de planification > 3 mois."),
            ("OT exécution <1 mois", "LANC", 1, "Age exécution", "!=", "<1 mois", "Réduire l'âge d'exécution des OT."),
            ("OT exécution >3 mois", "LANC", 1, "Age exécution", "==", ">3 mois", "Traiter les OT ayant un âge d'exécution > 3 mois."),
        ]

        for kpi_name, statut, sopl_value, age_col, operator, target_value, action in kpi_age_checks:
            val_kpi = calculated_kpis_df.loc[poste, kpi_name] if poste in calculated_kpis_df.index else 100
            if pd.notna(val_kpi):
                cible_val = cible_df.loc['CIBLE', kpi_name]
                is_problem = False
                if "<1 mois" in kpi_name and val_kpi < cible_val: is_problem = True
                elif ">3 mois" in kpi_name and val_kpi > cible_val: is_problem = True
                
                if is_problem:
                    condition = (df_poste["Statut OT"] == statut) if sopl_value is None else (df_poste["Statut OT"] == statut) & (df_poste["Contient SOPL"] == sopl_value)
                    if operator == "!=": condition = condition & (df_poste[age_col] != target_value)
                    else: condition = condition & (df_poste[age_col] == target_value)
                    
                    count_anom = len(df_poste[condition])
                    if count_anom > 0:
                        anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": kpi_name, "Nb OT impactés": count_anom, "Action Suggérée": action})

        # Qualité
        quality_checks = [
            ("OT LANC ESTIME", (df_poste["Statut OT"] == "LANC") & (df_poste["OT LANC ESTIME"] == "NON"), "Estimer les coûts des OT lancés."),
            ("Backlog préparation caractérisé", (df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Backlog préparation"] == "NON CARACTERISE"), "Caractériser le backlog de préparation."),
            ("Backlog planification caractérisé", (df_poste["Statut OT"] == "LANC") & (df_poste["Backlog planification"] == "NON CARACTERISE"), "Caractériser le backlog de planification."),
            ("OT CONFIME", df_poste["OT CONFIME"] == "NON", "Confirmer les OT terminés."),
            ("OT_COR_EGAL", df_poste["OT_COR_EGAL"] == "NON", "Rapprocher les coûts réels et budgétés."),
        ]

        for kpi_name, condition, action in quality_checks:
            val_kpi = calculated_kpis_df.loc[poste, kpi_name] if poste in calculated_kpis_df.index else 100
            if pd.notna(val_kpi) and val_kpi < cible_df.loc['CIBLE', kpi_name]:
                count_anom = len(df_poste[condition])
                if count_anom > 0:
                    anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": kpi_name, "Nb OT impactés": count_anom, "Action Suggérée": action})

        # Avis
        val_avis = calculated_kpis_df.loc[poste, "appel avis approuvé"] if poste in calculated_kpis_df.index else 100
        if pd.notna(val_avis) and val_avis < cible_df.loc['CIBLE', "appel avis approuvé"]:
            count_avis_anom = len(avis_poste)
            if count_avis_anom > 0:
                anomalies_avis_records.append({"Poste travail princ.": poste, "KPI": "appel avis approuvé", "Nb Avis impactés": count_avis_anom, "Action Suggérée": "Créer un OT pour les avis sans ordre."})

    return pd.DataFrame(anomalies_ot_records), pd.DataFrame(anomalies_avis_records)


# ==================================================
# GÉNÉRATION EXCEL
# ==================================================
def generer_excel_export(df_processed, df_anomalies_ot, pivot_avis, results, selected_poste_export, postes_avec_anomalies):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        postes_a_traiter = postes_avec_anomalies if selected_poste_export == "All" else [selected_poste_export]

        for poste_export in postes_a_traiter:
            kpis_en_defaut = df_anomalies_ot[df_anomalies_ot["Poste travail princ."] == poste_export]["KPI"].unique().tolist()
            
            if not pivot_avis.empty and "Nb Avis sans ordre" in pivot_avis.columns and poste_export in pivot_avis.index:
                if pivot_avis.loc[poste_export, "Nb Avis sans ordre"] > 0:
                    kpis_en_defaut.append("appel avis approuvé")

            for kpi in kpis_en_defaut:
                sheet_data = pd.DataFrame()

                if kpi != "appel avis approuvé":
                    df_poste_filtered = df_processed[df_processed["Poste travail princ."] == poste_export].copy()
                    
                    kpi_conditions = {
                        "TAUX_REALISATION_CORRECTIF/PT": lambda df: df[(df["Nº appel pl.entret."].fillna(0) == 0) & (~df["Statut OT"].isin(["CLOT", "TCLO"]))],
                        "OT préparation <1 mois": lambda df: df[(df["Statut OT"] == "CRÉÉ") & (df["Age préparation"] != "<1 mois")],
                        "OT préparation >3 mois": lambda df: df[(df["Statut OT"] == "CRÉÉ") & (df["Age préparation"] == ">3 mois")],
                        "OT planification <1 mois": lambda df: df[(df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 0) & (df["Age planification"] != "<1 mois")],
                        "OT planification >3 mois": lambda df: df[(df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 0) & (df["Age planification"] == ">3 mois")],
                        "OT exécution <1 mois": lambda df: df[(df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 1) & (df["Age exécution"] != "<1 mois")],
                        "OT exécution >3 mois": lambda df: df[(df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 1) & (df["Age exécution"] == ">3 mois")],
                        "OT LANC ESTIME": lambda df: df[(df["Statut OT"] == "LANC") & (df["OT LANC ESTIME"] == "NON")],
                        "Backlog préparation caractérisé": lambda df: df[(df["Statut OT"] == "CRÉÉ") & (df["Backlog préparation"] == "NON CARACTERISE")],
                        "Backlog planification caractérisé": lambda df: df[(df["Statut OT"] == "LANC") & (df["Backlog planification"] == "NON CARACTERISE")],
                        "OT CONFIME": lambda df: df[df["OT CONFIME"] == "NON"],
                        "OT_COR_EGAL": lambda df: df[df["OT_COR_EGAL"] == "NON"],
                    }

                    if kpi in kpi_conditions:
                        subset_ot = kpi_conditions[kpi](df_poste_filtered)
                        if not subset_ot.empty:
                            old_cols = ["Ordre", "Désignation", "Emplacement technique", "Poste travail princ.", 
                                       "Statut système", "Statut utilisateur", "Date de début planifiée", 
                                       "Type d'ordre", "Backlog préparation", "Backlog planification"]
                            new_cols = ["Ordre de travail", "Désignation", "Poste technique", "Poste de travail principal", 
                                       "Statut système", "Statut utilisateur", "Date de début planifiée", 
                                       "Type d'ordre", "Caractérisation backlog Préparation", "Caractérisation backlog Planification"]
                            subset_ot = rename_safe(subset_ot, old_cols, new_cols)
                            subset_ot = exclure_cresseurs(subset_ot, "Poste de travail principal")
                            subset_ot["KPI impacté"] = kpi
                            subset_ot["Action recommandée"] = f"Corriger l'indicateur {kpi}."
                            sheet_data = pd.concat([sheet_data, subset_ot])

                if kpi == "appel avis approuvé":
                    subset_avis = results['avis_df_filtered'][results['avis_df_filtered']["Poste travail princ."] == poste_export].copy()
                    if not subset_avis.empty:
                        subset_avis = exclure_cresseurs(subset_avis)
                        old_cols_avis = ["Avis", "Désignation texte", "Emplacement technique", "Poste travail princ.", "Statut utilisateur", "Créé le"]
                        new_cols_avis = ["Avis", "Désignation", "Poste technique", "Poste de travail principal", "Statut", "Date de création"]
                        subset_avis = rename_safe(subset_avis, old_cols_avis, new_cols_avis)
                        subset_avis["KPI impacté"] = kpi
                        subset_avis["Action recommandée"] = "Créer un Ordre de Travail pour cet Avis ou clarifier son statut."
                        sheet_data = pd.concat([sheet_data, subset_avis])

                if not sheet_data.empty:
                    base_name = poste_export.replace(" ", "_").replace("/", "_")[:20]
                    kpi_name = kpi.replace("/", "_").replace(" ", "_")[:10]
                    sheet_name = f"{base_name}_{kpi_name}"[:31]
                    sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output


# ==================================================
# FONCTION PRINCIPALE
# ==================================================
def main():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    
    if "hse_affiche" not in st.session_state:
        st.session_state.hse_affiche = False

    if not st.session_state.hse_affiche:
        afficher_ecran_hse()
        st.stop()

    # ==================================================
    # DASHBOARD PRINCIPAL - FULL WIDTH
    # ==================================================
    st.markdown("# 📊 KPI Dashboard MC et FEED")

    # Toggle fichiers
    col_toggle, col_info = st.columns([3, 2])
    with col_toggle:
        use_new_files = st.toggle("Charger de nouveaux fichiers OT et AVIS", value=False)
    with col_info:
        if not use_new_files:
            date_fichier = datetime.now().strftime("%d/%m/%Y")
            if os.path.exists("ot.xlsx"):
                timestamp = os.path.getmtime("ot.xlsx")
                date_fichier = datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y")
            st.info(f"📁 Version : V1.0 — {date_fichier}")

    ot_file = None
    avis_file = None

    if use_new_files:
        col1, col2 = st.columns(2)
        with col1:
            ot_file = st.file_uploader("📂 Charger le fichier OT", type=["xlsx"])
        with col2:
            avis_file = st.file_uploader("📂 Charger le fichier AVIS", type=["xlsx"])
    
    fichiers_prets = False
    if use_new_files:
        fichiers_prets = (ot_file is not None and avis_file is not None)
    else:
        fichiers_prets = os.path.exists("ot.xlsx") and os.path.exists("avis.xlsx")

    if not fichiers_prets:
        if use_new_files:
            st.warning("⚠️ Veuillez charger les deux fichiers (OT et AVIS).")
        else:
            st.error("❌ Fichiers ot.xlsx et/ou avis.xlsx introuvables.")
        st.stop()

    try:
        if use_new_files:
            df_ot_raw = pd.read_excel(ot_file)
            avis_df_raw = pd.read_excel(avis_file)
        else:
            df_ot_raw = pd.read_excel("ot.xlsx")
            avis_df_raw = pd.read_excel("avis.xlsx")

        # Conversion dates
        for col in ["Créé le", "Date de début planifiée", "Date de clôture", "Début réel", "Fin réelle"]:
            if col in df_ot_raw.columns:
                df_ot_raw[col] = pd.to_datetime(df_ot_raw[col], errors="coerce")
        for col in ["Créé le", "Début souhaité", "Date de la clôture"]:
            if col in avis_df_raw.columns:
                avis_df_raw[col] = pd.to_datetime(avis_df_raw[col], errors="coerce")

        # Exclusion cresseurs DÈS LE DÉPART sur les raw
        df_ot_raw = exclure_cresseurs(df_ot_raw)
        avis_df_raw = exclure_cresseurs(avis_df_raw)

        all_postes_master_list = sorted(
            df_ot_raw[df_ot_raw["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist()
        )

        if not all_postes_master_list:
            st.warning("⚠️ Aucun poste de travail trouvé.")
            st.stop()

        # ==================================================
        # FILTRES - LIGNE COMPLÈTE
        # ==================================================
        col_f1, col_f2, col_f3, col_f4 = st.columns([3, 2, 2, 2])
        
        with col_f1:
            selected_postes = st.multiselect("📍 Poste de travail", ["All"] + all_postes_master_list, ["All"])
            if "All" in selected_postes or len(selected_postes) == 0:
                selected_postes = all_postes_master_list

        with col_f2:
            ateliers_options = ["All", "Sulfurique (PS)", "Phosphorique (PP)", "Engrais (TSP/REX)", "Feed (MCP/DCP)"]
            selected_ateliers = st.multiselect("🏭 Atelier", ateliers_options, ["All"])
            if "All" in selected_ateliers or len(selected_ateliers) == 0:
                selected_ateliers = ["All"]

        with col_f3:
            divisions_options = ["All", "SF1", "SF2"]
            selected_divisions = st.multiselect("🏢 Division", divisions_options, ["All"])
            if "All" in selected_divisions or len(selected_divisions) == 0:
                selected_divisions = ["All"]

        with col_f4:
            default_start = datetime(2025, 1, 1).date()
            default_end = datetime.today().date()
            date_range = st.date_input("📅 Période", value=(default_start, default_end), format="DD/MM/YYYY")

        if len(date_range) == 2:
            start_date = pd.to_datetime(date_range[0])
            end_date = pd.to_datetime(date_range[1])
        else:
            start_date = pd.to_datetime(default_start)
            end_date = pd.to_datetime(default_end)

        def match_filters(poste):
            p = str(poste).upper()
            if "All" not in selected_ateliers:
                match_at = False
                if "Sulfurique (PS)" in selected_ateliers and "PS" in p: match_at = True
                if "Phosphorique (PP)" in selected_ateliers and "PP" in p: match_at = True
                if "Engrais (TSP/REX)" in selected_ateliers and ("TSP" in p or "REX" in p): match_at = True
                if "Feed (MCP/DCP)" in selected_ateliers and ("MCP" in p or "DCP" in p): match_at = True
                if not match_at: return False
            if "All" not in selected_divisions:
                match_div = False
                if "SF1" in selected_divisions and "SF1" in p: match_div = True
                if "SF2" in selected_divisions and "SF2" in p: match_div = True
                if not match_div: return False
            return True

        valid_postes = [p for p in all_postes_master_list if match_filters(p) and p in selected_postes]

        if not valid_postes:
            st.warning("⚠️ Aucun poste ne correspond aux filtres.")
            st.stop()

        # Application filtres + exclusion cresseurs (double sécurité)
        df = df_ot_raw[
            (df_ot_raw["Poste travail princ."].isin(valid_postes)) & 
            (df_ot_raw["Date de début planifiée"].between(start_date, end_date))
        ].copy()
        df = exclure_cresseurs(df)
        df = df[df["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)].drop_duplicates()

        avis_df = avis_df_raw[avis_df_raw["Poste travail princ."].isin(valid_postes)].copy()
        avis_df = exclure_cresseurs(avis_df)
        avis_df = avis_df[(avis_df["Ordre"].isna()) | (avis_df["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates()

        if "Statut système" in df.columns:
            df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

        # Indicateur de sélection
        is_all_postes = (len(selected_postes) == len(all_postes_master_list))
        
        st.success(f"📊 **{len(df)} OT** filtrés | **{len(avis_df)} AVIS** filtrés | **{len(valid_postes)} postes** sélectionnés")

        # ==================================================
        # CALCUL DES KPIs
        # ==================================================
        now = pd.Timestamp.now()
        results = calculate_all_kpis_and_intermediate_dfs(df, avis_df, now, valid_postes)
        calculated_kpis_df = results['calculated_kpis_df']
        df_processed = results['df_processed']

        cible = pd.DataFrame([CIBLES_KPI], index=["CIBLE"])

        # ==================================================
        # DÉTECTION DES ANOMALIES
        # ==================================================
        df_anomalies_ot, df_anomalies_avis = detecter_anomalies(
            df_processed, calculated_kpis_df, valid_postes, cible, results['avis_df_filtered']
        )

        if not df_anomalies_ot.empty:
            pivot_ot = df_anomalies_ot.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb OT impactés", aggfunc="sum", fill_value=0)
        else:
            pivot_ot = pd.DataFrame()

        if not df_anomalies_avis.empty:
            pivot_avis = df_anomalies_avis.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb Avis impactés", aggfunc="sum", fill_value=0)
            if "appel avis approuvé" in pivot_avis.columns:
                pivot_avis = pivot_avis.rename(columns={"appel avis approuvé": "Nb Avis sans ordre"})
        else:
            pivot_avis = pd.DataFrame()

        anomalies_dashboard = pivot_ot.join(pivot_avis, how='outer').fillna(0).astype(int)
        if not anomalies_dashboard.empty:
            anomalies_dashboard["Total éléments impactés"] = anomalies_dashboard.sum(axis=1)
            total_row = pd.DataFrame(anomalies_dashboard.sum()).T
            total_row.index = ["Total général"]
            anomalies_dashboard = pd.concat([anomalies_dashboard, total_row])
            anomalies_dashboard = exclure_cresseurs(anomalies_dashboard.reset_index()).set_index("Poste travail princ.") if "Poste travail princ." in anomalies_dashboard.reset_index().columns else anomalies_dashboard

        # ==================================================
        # CLASSIFICATION
        # ==================================================
        class_results = []
        for poste in calculated_kpis_df.index:
            row = calculated_kpis_df.loc[poste]
            s_qty = sum(get_kpi_score(k, row[k], cible.loc['CIBLE', k]) for k in QTY_KPIS if k in row.index)
            s_qual = sum(get_kpi_score(k, row[k], cible.loc['CIBLE', k]) for k in QUAL_KPIS if k in row.index)
            p_qty = (s_qty / len(QTY_KPIS) * 100) if QTY_KPIS else 0
            p_qual = (s_qual / len(QUAL_KPIS) * 100) if QUAL_KPIS else 0
            class_results.append({
                "Poste travail princ.": poste,
                "Score KPIs Quantité": p_qty,
                "Score KPIs Qualité": p_qual,
                "Total performance ": (p_qty + p_qual) / 2
            })
        df_class = pd.DataFrame(class_results)

        # ==================================================
        # ONGLETS PRINCIPAUX
        # ==================================================
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📋 Synthèse Actions",
            "📊 Tableau KPIs",
            "🚨 Tableau Anomalies",
            "🏆 Classement & Graphiques",
            "📥 Export"
        ])

        # ==================================================
        # ONGLET 1 : SYNTHÈSE DES ACTIONS
        # ==================================================
        with tab1:
            if is_all_postes:
                # Vue consolidée : KPI + Total OT + Action
                st.markdown('<div class="section-title">📋 Synthèse Consolidée des Actions KPI</div>', unsafe_allow_html=True)
                
                if not df_anomalies_ot.empty:
                    synthese_all = df_anomalies_ot.groupby(["KPI", "Action Suggérée"]).agg(
                        Total_OT=("Nb OT impactés", "sum"),
                        Nb_Postes=("Poste travail princ.", "nunique")
                    ).reset_index()
                    synthese_all = synthese_all.rename(columns={
                        "KPI": "KPI",
                        "Total_OT": "Total OT impactés",
                        "Nb_Postes": "Nb Postes concernés",
                        "Action Suggérée": "Action Suggérée"
                    })
                    synthese_all["Cible (%)"] = synthese_all["KPI"].map(CIBLES_KPI)
                    synthese_all = synthese_all[["KPI", "Cible (%)", "Total OT impactés", "Nb Postes concernés", "Action Suggérée"]]
                    synthese_all = synthese_all.sort_values("Total OT impactés", ascending=False)
                    
                    st.dataframe(
                        synthese_all.style.apply(highlight_synthese, axis=1),
                        use_container_width=True,
                        height=max(400, len(synthese_all) * 45 + 40)
                    )
                else:
                    st.info("✅ Tous les KPIs atteignent leurs cibles. Aucune action requise.")

                # Ajout synthèse avis si anomalies
                if not df_anomalies_avis.empty:
                    st.markdown("---")
                    st.markdown('<div class="section-title">📨 Synthèse des Avis sans Ordre</div>', unsafe_allow_html=True)
                    synthese_avis_all = df_anomalies_avis.groupby("Action Suggérée").agg(
                        Total_Avis=("Nb Avis impactés", "sum"),
                        Nb_Postes=("Poste travail princ.", "nunique")
                    ).reset_index()
                    synthese_avis_all.columns = ["Action Suggérée", "Total Avis impactés", "Nb Postes concernés"]
                    st.dataframe(synthese_avis_all, use_container_width=True)

            else:
                # Vue par poste
                st.markdown('<div class="section-title">📋 Actions KPI par Poste de Travail</div>', unsafe_allow_html=True)
                
                if not df_anomalies_ot.empty:
                    display_actions = df_anomalies_ot[["Poste travail princ.", "KPI", "Nb OT impactés", "Action Suggérée"]]
                    display_actions = display_actions.sort_values(["Poste travail princ.", "KPI"])
                    st.dataframe(display_actions, use_container_width=True, height=max(400, len(display_actions) * 35 + 40))
                else:
                    st.info("✅ Tous les KPIs atteignent leurs cibles.")

                if not df_anomalies_avis.empty:
                    st.markdown("---")
                    st.markdown('<div class="section-title">📨 Avis sans Ordre par Poste</div>', unsafe_allow_html=True)
                    st.dataframe(df_anomalies_avis, use_container_width=True)

        # ==================================================
        # ONGLET 2 : TABLEAU DE BORD DES KPIs
        # ==================================================
        with tab2:
            st.markdown('<div class="header-kpi">📊 TABLEAU DE BORD DES KPIs</div>', unsafe_allow_html=True)
            
            total_general_kpi = pd.DataFrame(calculated_kpis_df.mean()).T
            total_general_kpi.index = ["Total général"]
            final_kpi = pd.concat([cible, calculated_kpis_df, total_general_kpi]).round(2)
            
            st.dataframe(
                final_kpi.style.apply(highlight_kpis, axis=1).format("{:.2f}"),
                use_container_width=True,
                height=max(400, len(final_kpi) * 35 + 40)
            )

        # ==================================================
        # ONGLET 3 : TABLEAU DE BORD DES ANOMALIES
        # ==================================================
        with tab3:
            st.markdown('<div class="header-anomalie">🚨 TABLEAU DE BORD DES ANOMALIES</div>', unsafe_allow_html=True)
            
            if not anomalies_dashboard.empty:
                st.dataframe(
                    anomalies_dashboard,
                    use_container_width=True,
                    height=max(400, len(anomalies_dashboard) * 35 + 40)
                )
            else:
                st.info("✅ Aucune anomalie détectée.")

        # ==================================================
        # ONGLET 4 : CLASSEMENT & GRAPHIQUES
        # ==================================================
        with tab4:
            # Classement
            st.markdown('<div class="section-title">🏆 Classement des Postes par Qualité des KPIs</div>', unsafe_allow_html=True)

            df_class_display = df_class.copy()
            df_class_display["Score KPIs Quantité"] = df_class_display["Score KPIs Quantité"].apply(lambda x: f"{x:.2f} %")
            df_class_display["Score KPIs Qualité"] = df_class_display["Score KPIs Qualité"].apply(lambda x: f"{x:.2f} %")
            df_class_display["Total performance "] = df_class_display["Total performance "].apply(lambda x: f"{x:.2f} %")

            total_gen_class = pd.DataFrame([{
                "Poste travail princ.": "Total général",
                "Score KPIs Quantité": f"{df_class['Score KPIs Quantité'].mean():.2f} %",
                "Score KPIs Qualité": f"{df_class['Score KPIs Qualité'].mean():.2f} %",
                "Total performance ": f"{df_class['Total performance '].mean():.2f} %"
            }])
            df_class_display = pd.concat([df_class_display, total_gen_class], ignore_index=True)
            
            st.dataframe(
                df_class_display.style.apply(highlight_classification_table_kpis, axis=1),
                use_container_width=True,
                height=max(400, len(df_class_display) * 35 + 40)
            )

            # Top 5
            st.markdown("---")
            col_top1, col_top2, col_top3 = st.columns(3)
            
            with col_top1:
                st.markdown("#### ⬇️ Top 5 Quantité")
                top5 = df_class.nsmallest(5, "Score KPIs Quantité")[["Poste travail princ.", "Score KPIs Quantité"]].round(2)
                st.dataframe(top5.set_index("Poste travail princ."))
            with col_top2:
                st.markdown("#### ⬇️ Top 5 Qualité")
                top5 = df_class.nsmallest(5, "Score KPIs Qualité")[["Poste travail princ.", "Score KPIs Qualité"]].round(2)
                st.dataframe(top5.set_index("Poste travail princ."))
            with col_top3:
                st.markdown("#### ⬇️ Top 5 Performance")
                top5 = df_class.nsmallest(5, "Total performance ")[["Poste travail princ.", "Total performance "]].round(2)
                st.dataframe(top5.set_index("Poste travail princ."))

            # Graphiques avec % clairs
            st.markdown("---")
            st.markdown('<div class="section-title">📈 Performance par Catégorie</div>', unsafe_allow_html=True)
            
            df_class["Métier"] = df_class["Poste travail princ."].apply(get_groupe_metier)
            df_class["Atelier"] = df_class["Poste travail princ."].apply(get_groupe_atelier)
            df_class["Division"] = df_class["Poste travail princ."].apply(get_groupe_division)

            c1, c2, c3 = st.columns(3)

            chart_configs = [
                ("Métier", "#3498db", c1),
                ("Atelier", "#e74c3c", c2),
                ("Division", "#27ae60", c3),
            ]

            for group_col, color, container in chart_configs:
                with container:
                    df_group = df_class.groupby(group_col)["Total performance "].mean().reset_index()
                    df_group["Pourcentage"] = df_group["Total performance "].apply(lambda x: f"{x:.1f} %")
                    
                    base = alt.Chart(df_group).encode(
                        x=alt.X(f'{group_col}:O', axis=alt.Axis(labelAngle=0, labelFontSize=12)),
                        y=alt.Y('Total performance :Q', scale=alt.Scale(domain=[0, 100]), axis=alt.Axis(grid=False, ticks=False, labels=False, title=None))
                    )
                    
                    bars = base.mark_bar(
                        color=color, 
                        cornerRadiusTopLeft=5, 
                        cornerRadiusTopRight=5,
                        size=50
                    )
                    
                    labels = base.mark_text(
                        align='center',
                        baseline='bottom',
                        dy=-8,
                        fontSize=16,
                        fontWeight='bold',
                        color='#2c3e50'
                    ).encode(text='Pourcentage')
                    
                    chart = (bars + labels).properties(
                        height=220,
                        width=alt.Step(70)
                    ).configure_view(
                        stroke='transparent'
                    )
                    
                    st.altair_chart(chart, use_container_width=True)

        # ==================================================
        # ONGLET 5 : EXPORT
        # ==================================================
        with tab5:
            st.markdown('<div class="section-title">📥 Exporter les Plans d\'Action Détaillés</div>', unsafe_allow_html=True)
            
            postes_avec_anomalies = []
            if not anomalies_dashboard.empty:
                postes_avec_anomalies = [p for p in anomalies_dashboard.index if p != "Total général"]

            if postes_avec_anomalies:
                col_sel, col_btn = st.columns([3, 2])
                with col_sel:
                    selected_poste_export = st.selectbox(
                        "Sélectionnez le poste de travail :",
                        options=["All"] + postes_avec_anomalies,
                        format_func=lambda x: "📦 Tous les postes avec anomalies" if x == "All" else x
                    )
                with col_btn:
                    st.markdown("<br>", unsafe_allow_html=True)
                    generate_btn = st.button("📥 Générer et télécharger le fichier Excel", type="primary", use_container_width=True)

                if generate_btn:
                    with st.spinner("⏳ Génération en cours..."):
                        output = generer_excel_export(
                            df_processed, df_anomalies_ot, pivot_avis, results, 
                            selected_poste_export, postes_avec_anomalies
                        )
                        
                        nom_fichier = "Plan_Action_Tous_Postes.xlsx" if selected_poste_export == "All" else f"Plan_Action_{selected_poste_export.replace(' ', '_')}.xlsx"
                        
                        st.download_button(
                            label="✅ Cliquez ici pour télécharger",
                            data=output.getvalue(),
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            else:
                st.info("✅ Aucune anomalie détectée. L'export est désactivé.")

    except FileNotFoundError as e:
        st.error(f"❌ Fichier introuvable : {e}")
    except pd.errors.EmptyDataError:
        st.error("❌ Un des fichiers est vide ou corrompu.")
    except KeyError as e:
        st.error(f"❌ Colonne manquante : {e}")
    except Exception as e:
        st.error(f"❌ Erreur inattendue : {e}")
        with st.expander("🔍 Détails techniques", expanded=False):
            st.code(str(e))


# ==================================================
# POINT D'ENTRÉE
# ==================================================
if __name__ == "__main__":
    set_french_locale()
    main()
