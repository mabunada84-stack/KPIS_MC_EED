# -*- coding: utf-8 -*-
"""app.py"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import locale
import altair as alt
import random
from datetime import datetime
import os

# ==================================================
# CSS STYLES
# ==================================================
def inject_css():
    st.markdown("""
    <style>
        /* Page entière sans scroll */
        .main .block-container {
            padding-top: 1rem;
            padding-bottom: 1rem;
            max-width: 100%;
            padding-left: 2rem;
            padding-right: 2rem;
        }
        
        /* Masquer scrollbar */
        ::-webkit-scrollbar { display: none; }
        
        /* Tableaux compacts */
        .dataframe th, .dataframe td {
            padding: 4px 8px !important;
            font-size: 11px !important;
            white-space: nowrap;
        }
        
        /* Boutons toggle stylisés */
        .stButton > button {
            width: 100%;
            transition: all 0.3s ease;
        }
        
        /* Headers */
        .header-kpi {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            padding: 12px 20px;
            border-radius: 8px;
            color: white;
            font-size: 18px;
            font-weight: bold;
            text-align: center;
            margin: 10px 0;
        }
        .header-anomalie {
            background: linear-gradient(135deg, #c0392b 0%, #e74c3c 100%);
            padding: 12px 20px;
            border-radius: 8px;
            color: white;
            font-size: 18px;
            font-weight: bold;
            text-align: center;
            margin: 10px 0;
        }
        
        /* Splash screen */
        .splash-container {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 90vh;
        }
        
        /* Réduire espacement */
        .stMarkdown, .stAlert, .stDataFrame {
            margin-top: 0.2rem;
            margin-bottom: 0.2rem;
        }
        
        /* Metric cards */
        [data-testid="stMetric"] {
            background-color: #f8f9fa;
            padding: 10px;
            border-radius: 8px;
            border: 1px solid #e9ecef;
        }
        
        /* Boutons radio personnalisés */
        .tab-button {
            display: inline-block;
            padding: 10px 30px;
            margin: 0 5px;
            border-radius: 8px 8px 0 0;
            cursor: pointer;
            font-weight: bold;
            font-size: 14px;
            transition: all 0.3s;
        }
        .tab-active-kpi {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
        }
        .tab-active-anomalie {
            background: linear-gradient(135deg, #c0392b 0%, #e74c3c 100%);
            color: white;
        }
        .tab-inactive {
            background: #e9ecef;
            color: #6c757d;
        }
    </style>
    """, unsafe_allow_html=True)

# ==================================================
# CONSTANTES
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

MOTS_PREP = ["CRPR ATPD", "CRPR ATMR", "CRPR ATER", "CRPR ATRS", "CRPR ATMO", "ATPD", "ATMR", "ATER", "ATRS", "ATMO"]
MOTS_PLANIF = ["ATPL ATEI", "ATPL ATAL", "ATPL ATER", "ATPL AGAR", "ATPL ATHS", "ATEI", "ATAL", "ATAS", "AGAR", "ATHS"]

CIBLES = {
    "TAUX_REALISATION_CORRECTIF/PT": 85,
    "OT préparation <1 mois": 80, "OT préparation >3 mois": 5, "OT préparation 1mois< <3mois": 15,
    "OT planification <1 mois": 80, "OT planification >3 mois": 5, "OT planification 1mois< <3mois": 15,
    "OT exécution <1 mois": 80, "OT exécution >3 mois": 5, "OT exécution 1mois< <3mois": 15,
    "appel avis approuvé": 95, "OT LANC ESTIME": 100, "Backlog préparation caractérisé": 100,
    "Backlog planification caractérisé": 100, "OT CONFIME": 100, "OT_COR_EGAL": 100,
}

ACTIONS_KPI = {
    "TAUX_REALISATION_CORRECTIF/PT": "Améliorer le taux de réalisation des OT.",
    "OT préparation <1 mois": "Réduire l'âge de préparation des OT à <1 mois.",
    "OT préparation >3 mois": "Traiter les OT ayant un âge de préparation > 3 mois.",
    "OT préparation 1mois< <3mois": "Réduire les OT avec âge de préparation entre 1 et 3 mois.",
    "OT planification <1 mois": "Réduire l'âge de planification des OT à <1 mois.",
    "OT planification >3 mois": "Traiter les OT ayant un âge de planification > 3 mois.",
    "OT planification 1mois< <3mois": "Réduire les OT avec âge de planification entre 1 et 3 mois.",
    "OT exécution <1 mois": "Réduire l'âge d'exécution des OT à <1 mois.",
    "OT exécution >3 mois": "Traiter les OT ayant un âge d'exécution > 3 mois.",
    "OT exécution 1mois< <3mois": "Réduire les OT avec âge d'exécution entre 1 et 3 mois.",
    "OT LANC ESTIME": "Estimer les coûts des OT lancés.",
    "Backlog préparation caractérisé": "Caractériser le backlog de préparation.",
    "Backlog planification caractérisé": "Caractériser le backlog de planification.",
    "OT CONFIME": "Confirmer les OT terminés.",
    "OT_COR_EGAL": "Rapprocher les coûts réels et budgétés.",
    "appel avis approuvé": "Créer un OT pour les avis sans ordre.",
}

# ==================================================
# FONCTIONS UTILITAIRES
# ==================================================
def rename_safe(df, old_names, new_names):
    mapping = {old: new for old, new in zip(old_names, new_names) if old in df.columns}
    return df.rename(columns=mapping)

def contient_mot(texte, liste_mots):
    texte = str(texte)
    return any(mot in texte for lst in liste_mots for mot in lst.split())

def categorie_age(age):
    if age <= 1: return "<1 mois"
    elif age >= 3: return ">3 mois"
    else: return "1 mois < <3 mois"

def calcul_kpi(numerateur, denominateur, si_zero=100):
    return np.where(denominateur == 0, si_zero, (numerateur / denominateur) * 100)

def creer_pivot(dataframe, filtre, colonne, postes_to_reindex):
    pivot = pd.pivot_table(dataframe[filtre], index="Poste travail princ.", columns=colonne, values="Ordre", aggfunc="count", fill_value=0)
    return pivot.reindex(postes_to_reindex, fill_value=0)

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

def match_filters(poste, selected_ateliers, selected_divisions):
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

# ==================================================
# CALCUL DES KPIs
# ==================================================
def process_dataframe(df_temp, now_timestamp):
    df_temp["Backlog préparation"] = np.where(df_temp["Statut utilisateur"].apply(lambda x: contient_mot(x, MOTS_PREP)), "CARACTERISE", "NON CARACTERISE")
    df_temp["Backlog planification"] = np.where(df_temp["Statut utilisateur"].apply(lambda x: contient_mot(x, MOTS_PLANIF)), "CARACTERISE", "NON CARACTERISE")

    for date_col, age_mois_col, age_col in [('Créé le', "Age mois préparation", "Age préparation"), 
                                              ('Date de début planifiée', "Age mois planification", "Age planification"), 
                                              ('Date de début planifiée', "Age mois exécution", "Age exécution")]:
        if date_col in df_temp.columns:
            df_temp[date_col] = pd.to_datetime(df_temp[date_col], errors='coerce')
            df_temp[age_mois_col] = ((now_timestamp.year - df_temp[date_col].dt.year) * 12 + (now_timestamp.month - df_temp[date_col].dt.month)).round(2)
            df_temp[age_col] = df_temp[age_mois_col].apply(categorie_age)
        else:
            df_temp[age_mois_col] = np.nan
            df_temp[age_col] = "Inconnu"

    df_temp["OT CONFIME"] = np.where(df_temp["Statut système"].str.contains("CLO", na=False) & df_temp["Statut système"].str.contains("CONF", na=False), "OUI", "NON")
    df_temp["Contient SOPL"] = df_temp["Statut utilisateur"].str.contains("SOPL", na=False).map({True: 1, False: 0})
    df_temp["OT LANC ESTIME"] = np.where(df_temp["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
    df_temp["OT_COR_EGAL"] = np.where((df_temp["Total coûts budgétés"].fillna(0) - df_temp["Total coûts réels"].fillna(0)) == 0, "OUI", "NON")
    return df_temp

def calculate_kpis(df_temp, avis_df_temp, all_postes_list):
    results = {}
    
    # Analyse Correctif
    analyse = creer_pivot(df_temp, df_temp["Nº appel pl.entret."].fillna(0) == 0, "Statut OT", all_postes_list)
    for col in ["CLOT", "CRÉÉ", "LANC", "TCLO"]: analyse[col] = analyse.get(col, 0)
    analyse["Total"] = analyse[["CLOT", "CRÉÉ", "LANC", "TCLO"]].sum(axis=1)
    analyse["TAUX_REALISATION_CORRECTIF/PT"] = calcul_kpi(analyse["TCLO"], analyse["Total"])

    # Préparation
    prep = creer_pivot(df_temp, df_temp["Statut OT"] == "CRÉÉ", "Age préparation", all_postes_list)
    for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: prep[col] = prep.get(col, 0)
    prep["Total"] = prep[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    prep["OT préparation <1 mois"] = calcul_kpi(prep["<1 mois"], prep["Total"])
    prep["OT préparation >3 mois"] = calcul_kpi(prep[">3 mois"], prep["Total"], 0)
    prep["OT préparation 1mois< <3mois"] = calcul_kpi(prep["1 mois < <3 mois"], prep["Total"], 0)

    # Planification
    planif = creer_pivot(df_temp, (df_temp["Statut OT"] == "LANC") & (df_temp["Contient SOPL"] == 0), "Age planification", all_postes_list)
    for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: planif[col] = planif.get(col, 0)
    planif["Total"] = planif[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    planif["OT planification <1 mois"] = calcul_kpi(planif["<1 mois"], planif["Total"])
    planif["OT planification >3 mois"] = calcul_kpi(planif[">3 mois"], planif["Total"], 0)
    planif["OT planification 1mois< <3mois"] = calcul_kpi(planif["1 mois < <3 mois"], planif["Total"], 0)

    # Exécution
    execu = creer_pivot(df_temp, (df_temp["Statut OT"] == "LANC") & (df_temp["Contient SOPL"] == 1), "Age exécution", all_postes_list)
    for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: execu[col] = execu.get(col, 0)
    execu["Total"] = execu[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
    execu["OT exécution <1 mois"] = calcul_kpi(execu["<1 mois"], execu["Total"])
    execu["OT exécution >3 mois"] = calcul_kpi(execu[">3 mois"], execu["Total"], 0)
    execu["OT exécution 1mois< <3mois"] = calcul_kpi(execu["1 mois < <3 mois"], execu["Total"], 0)

    # OT LANC ESTIME
    lanc = pd.pivot_table(df_temp[df_temp["Statut OT"] == "LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
    for col in ["OUI", "NON"]: lanc[col] = lanc.get(col, 0)
    lanc["Total"] = lanc["OUI"] + lanc["NON"]
    lanc["OT LANC ESTIME"] = calcul_kpi(lanc["OUI"], lanc["Total"])

    # Backlog préparation caractérisé
    prep_carac = pd.pivot_table(df_temp[df_temp["Statut OT"] == "CRÉÉ"], index="Poste travail princ.", columns="Backlog préparation", values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
    for col in ["CARACTERISE", "NON CARACTERISE"]: prep_carac[col] = prep_carac.get(col, 0)
    prep_carac["Total"] = prep_carac["CARACTERISE"] + prep_carac["NON CARACTERISE"]
    prep_carac["Backlog préparation caractérisé"] = calcul_kpi(prep_carac["CARACTERISE"], prep_carac["Total"])

    # Backlog planification caractérisé
    planif_carac = pd.pivot_table(df_temp[df_temp["Statut OT"] == "LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
    for col in ["CARACTERISE", "NON CARACTERISE"]: planif_carac[col] = planif_carac.get(col, 0)
    planif_carac["Total"] = planif_carac["CARACTERISE"] + planif_carac["NON CARACTERISE"]
    planif_carac["Backlog planification caractérisé"] = calcul_kpi(planif_carac["CARACTERISE"], planif_carac["Total"])

    # OT CONFIME et OT_COR_EGAL
    for kpi_name, col_name in [("OT CONFIME", "OT CONFIME"), ("OT_COR_EGAL", "OT_COR_EGAL")]:
        df_pivot = pd.pivot_table(df_temp, index="Poste travail princ.", columns=col_name, values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
        for c in ["OUI", "NON"]: df_pivot[c] = df_pivot.get(c, 0)
        df_pivot["Total"] = df_pivot["OUI"] + df_pivot["NON"]
        df_pivot[col_name] = calcul_kpi(df_pivot["OUI"], df_pivot["Total"])
        results[kpi_name.lower().replace(" ", "_")] = df_pivot

    # Avis
    avis_df_filtered = avis_df_temp[(avis_df_temp["Ordre"].isna()) | (avis_df_temp["Ordre"].astype(str).str.strip() == "")].copy()
    results['avis_df_filtered'] = avis_df_filtered

    tableau_avis = pd.pivot_table(avis_df_filtered, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
    for col in ["APRQ", "APRV", "APRV AVAU", "REJT"]: tableau_avis[col] = tableau_avis.get(col, 0)
    tableau_avis["Total"] = tableau_avis[["APRQ", "APRV", "APRV AVAU", "REJT"]].sum(axis=1)
    tableau_avis["appel avis approuvé"] = calcul_kpi(tableau_avis["APRV"], tableau_avis["Total"])

    # Concaténation finale
    calculated_kpis_df = pd.concat([
        analyse[["TAUX_REALISATION_CORRECTIF/PT"]], 
        prep[["OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois"]],
        planif[["OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois"]], 
        execu[["OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]],
        tableau_avis[["appel avis approuvé"]], 
        lanc[["OT LANC ESTIME"]], 
        prep_carac[["Backlog préparation caractérisé"]],
        planif_carac[["Backlog planification caractérisé"]], 
        results['ot_confime'][["OT CONFIME"]], 
        results['ot_cor_egal'][["OT_COR_EGAL"]]
    ], axis=1)

    results['calculated_kpis_df'] = calculated_kpis_df
    return results

# ==================================================
# MISE EN FORME
# ==================================================
def highlight_kpis(row):
    styles = []
    if row.name == 'CIBLE': 
        return ['padding: 6px; background-color: #d6eaf8; color: #1a5276; font-weight: bold; font-size: 11px;'] * len(row)
    if row.name == 'Total général': 
        return ['padding: 6px; background-color: #f5f5f5; font-weight: bold; font-size: 11px;'] * len(row)
    for col in row.index:
        s = 'padding: 6px; font-size: 11px;'
        try: value = float(row[col])
        except (ValueError, TypeError): styles.append(s); continue
        if col in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
            if value >= 80: s += 'background-color:#c6efce; color:#006100;'
            elif value >= 75: s += 'background-color:#ffeb9c; color:#9c6500;'
            else: s += 'background-color:#ffc7ce; color:#9c0006;'
        elif col in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
            if value <= 15: s += 'background-color:#c6efce; color:#006100;'
            else: s += 'background-color:#ffc7ce; color:#9c0006;'
        elif col in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
            if value <= 5: s += 'background-color:#c6efce; color:#006100;'
            else: s += 'background-color:#ffc7ce; color:#9c0006;'
        elif col == "TAUX_REALISATION_CORRECTIF/PT":
            if value >= 85: s += 'background-color:#c6efce; color:#006100;'
            elif value >= 80: s += 'background-color:#ffeb9c; color:#9c6500;'
            else: s += 'background-color:#ffc7ce; color:#9c0006;'
        elif col == "appel avis approuvé":
            if value >= 95: s += 'background-color:#c6efce; color:#006100;'
            elif value >= 90: s += 'background-color:#ffeb9c; color:#9c6500;'
            else: s += 'background-color:#ffc7ce; color:#9c0006;'
        elif col in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
            if value >= 100: s += 'background-color:#c6efce; color:#006100;'
            elif value >= 95: s += 'background-color:#ffeb9c; color:#9c6500;'
            else: s += 'background-color:#ffc7ce; color:#9c0006;'
        styles.append(s)
    return styles

def highlight_classification(row):
    styles = ['padding: 6px; font-size: 11px;'] * len(row)
    for i, col in enumerate(row.index):
        if col in ["Score KPIs Quantité", "Score KPIs Qualité", "Total performance "]:
            try:
                s = str(row[col])
                value = float(s.replace(' %', '').strip()) if ' %' in s else float(s)
                if value >= 90: styles[i] += 'background-color: #c6efce; color: #006100;'
                elif value >= 80: styles[i] += 'background-color: #ffeb9c; color: #9c6500;'
                else: styles[i] += 'background-color: #ffc7ce; color: #9c0006;'
            except: pass
    return styles

def get_kpi_score(kpi_name, actual_value, target_value):
    if pd.isna(actual_value) or pd.isna(target_value): return 0
    if kpi_name in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]: return 1 if actual_value >= 75 else 0
    if kpi_name in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]: return 1 if actual_value <= 15 else 0
    if kpi_name in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]: return 1 if actual_value <= 5 else 0
    if kpi_name == "TAUX_REALISATION_CORRECTIF/PT": return 1 if actual_value >= 80 else 0
    if kpi_name == "appel avis approuvé": return 1 if actual_value >= 90 else 0
    if kpi_name in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]: return 1 if actual_value >= 95 else 0
    return 0

# ==================================================
# DÉTECTION ANOMALIES
# ==================================================
def detect_anomalies(df_processed, calculated_kpis_df, results, valid_postes, cible):
    anomalies_ot_records = []
    anomalies_avis_records = []

    for poste in valid_postes:
        if poste not in df_processed["Poste travail princ."].values: continue
        df_poste = df_processed[df_processed["Poste travail princ."] == poste]
        avis_poste = results['avis_df_filtered'][results['avis_df_filtered']["Poste travail princ."] == poste]

        # KPIs avec conditions spécifiques
        kpi_conditions = {
            "TAUX_REALISATION_CORRECTIF/PT": (df_poste["Nº appel pl.entret."].fillna(0) == 0) & (~df_poste["Statut OT"].isin(["CLOT", "TCLO"])),
            "OT préparation <1 mois": (df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Age préparation"] != "<1 mois"),
            "OT préparation >3 mois": (df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Age préparation"] == ">3 mois"),
            "OT préparation 1mois< <3mois": (df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Age préparation"] == "1 mois < <3 mois"),
            "OT planification <1 mois": (df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 0) & (df_poste["Age planification"] != "<1 mois"),
            "OT planification >3 mois": (df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 0) & (df_poste["Age planification"] == ">3 mois"),
            "OT planification 1mois< <3mois": (df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 0) & (df_poste["Age planification"] == "1 mois < <3 mois"),
            "OT exécution <1 mois": (df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 1) & (df_poste["Age exécution"] != "<1 mois"),
            "OT exécution >3 mois": (df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 1) & (df_poste["Age exécution"] == ">3 mois"),
            "OT exécution 1mois< <3mois": (df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 1) & (df_poste["Age exécution"] == "1 mois < <3 mois"),
            "OT LANC ESTIME": (df_poste["Statut OT"] == "LANC") & (df_poste["OT LANC ESTIME"] == "NON"),
            "Backlog préparation caractérisé": (df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Backlog préparation"] == "NON CARACTERISE"),
            "Backlog planification caractérisé": (df_poste["Statut OT"] == "LANC") & (df_poste["Backlog planification"] == "NON CARACTERISE"),
            "OT CONFIME": df_poste["OT CONFIME"] == "NON",
            "OT_COR_EGAL": df_poste["OT_COR_EGAL"] == "NON",
        }

        for kpi_name, condition in kpi_conditions.items():
            if poste not in calculated_kpis_df.index: continue
            val_kpi = calculated_kpis_df.loc[poste, kpi_name]
            
            # Logique de comparaison selon le KPI
            is_anomaly = False
            if kpi_name in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
                is_anomaly = pd.notna(val_kpi) and val_kpi > cible[kpi_name]
            elif kpi_name in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
                is_anomaly = pd.notna(val_kpi) and val_kpi > cible[kpi_name]
            else:
                is_anomaly = pd.notna(val_kpi) and val_kpi < cible[kpi_name]

            if is_anomaly:
                count_anom = len(df_poste[condition])
                if count_anom > 0:
                    anomalies_ot_records.append({
                        "Poste travail princ.": poste, 
                        "KPI": kpi_name, 
                        "Nb OT impactés": count_anom, 
                        "Action Suggérée": ACTIONS_KPI.get(kpi_name, "")
                    })

        # Anomalies Avis
        if poste in calculated_kpis_df.index:
            val_avis = calculated_kpis_df.loc[poste, "appel avis approuvé"]
            if pd.notna(val_avis) and val_avis < cible["appel avis approuvé"]:
                count_avis = len(avis_poste)
                if count_avis > 0:
                    anomalies_avis_records.append({
                        "Poste travail princ.": poste, 
                        "KPI": "appel avis approuvé", 
                        "Nb Avis impactés": count_avis, 
                        "Action Suggérée": ACTIONS_KPI["appel avis approuvé"]
                    })

    return pd.DataFrame(anomalies_ot_records), pd.DataFrame(anomalies_avis_records)

# ==================================================
# SPLASH SCREEN HSE
# ==================================================
def show_hse_splash():
    consigne = random.choice(CONSIGNES_HSE)
    
    st.markdown("""
    <div class="splash-container">
        <h1 style="text-align:center; font-size:48px; color:#0f172a; margin-bottom:10px;">
            🦺 HSE - CONSIGNE DE SÉCURITÉ
        </h1>
        <h2 style="text-align:center; color:#64748b; margin-bottom:30px;">
            Sécurité • Santé • Environnement
        </h2>
        <div style="background-color:#fff3cd; border-left:8px solid #ffc107; padding:30px 50px; border-radius:12px; font-size:32px; font-weight:bold; text-align:center; margin-bottom:30px; max-width:900px;">
            ⚠️ {consigne}
        </div>
        <h1 style="text-align:center; color:#198754; font-size:36px; font-weight:800; margin-bottom:40px;">
            Aucun travail n'est plus urgent que la sécurité
        </h1>
    </div>
    """.format(consigne=consigne), unsafe_allow_html=True)
    
    col_btn = st.columns([1, 2, 1])
    with col_btn[1]:
        if st.button("✅ J'ai lu la consigne - Accéder au dashboard", type="primary", use_container_width=True):
            st.session_state.hse_affiche = True
            st.rerun()

# ==================================================
# MAIN APPLICATION
# ==================================================
def main():
    inject_css()
    
    # Splash HSE
    if "hse_affiche" not in st.session_state:
        st.session_state.hse_affiche = False
    
    if not st.session_state.hse_affiche:
        show_hse_splash()
        st.stop()

    # ==================================================
    # TITRE ET CHARGEMENT
    # ==================================================
    st.markdown("# 📊 KPI Dashboard MC et FEED")
    
    use_new_files = st.toggle("Charger de nouveaux fichiers OT et AVIS", value=False)
    ot_file = avis_file = None

    if use_new_files:
        col1, col2 = st.columns(2)
        with col1: ot_file = st.file_uploader("📂 Fichier OT", type=["xlsx"], key="ot")
        with col2: avis_file = st.file_uploader("📂 Fichier AVIS", type=["xlsx"], key="avis")
    else:
        date_fichier = datetime.now().strftime("%d/%m/%Y")
        if os.path.exists("ot.xlsx"):
            date_fichier = datetime.fromtimestamp(os.path.getmtime("ot.xlsx")).strftime("%d/%m/%Y")
        st.caption(f"📁 Data Version : V1.0 – Mise à jour du : {date_fichier}")

    if not use_new_files or (ot_file is not None and avis_file is not None):
        try:
            # Chargement
            if use_new_files:
                df_ot_raw = pd.read_excel(ot_file)
                avis_df_raw = pd.read_excel(avis_file)
            else:
                df_ot_raw = pd.read_excel("ot.xlsx")
                avis_df_raw = pd.read_excel("avis.xlsx")

            for col in ["Créé le", "Date de début planifiée", "Date de clôture", "Début réel", "Fin réelle"]:
                if col in df_ot_raw.columns: df_ot_raw[col] = pd.to_datetime(df_ot_raw[col], errors="coerce")
            for col in ["Créé le", "Début souhaité", "Date de la clôture"]:
                if col in avis_df_raw.columns: avis_df_raw[col] = pd.to_datetime(avis_df_raw[col], errors="coerce")

            # Exclusion cresseurs
            cresseur_mask = df_ot_raw["Poste travail princ."].astype(str).str.contains("cresseur|vercilat|horizon", case=False, na=False)
            df_ot_clean = df_ot_raw[~cresseur_mask].copy()
            cresseur_mask_avis = avis_df_raw["Poste travail princ."].astype(str).str.contains("cresseur|vercilat|horizon", case=False, na=False)
            avis_df_clean = avis_df_raw[~cresseur_mask_avis].copy()

            # Liste postes valides (SF1/SF2 uniquement, sans cresseurs)
            all_postes_master_list = sorted(df_ot_clean[df_ot_clean["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())

            # ==================================================
            # FILTRES DATE (appliqué globalement)
            # ==================================================
            default_start = datetime(2025, 1, 1).date()
            default_end = datetime.today().date()
            date_range = st.date_input("📅 Filtre Date de début planifiée", value=(default_start, default_end), format="DD/MM/YYYY")
            start_date = pd.to_datetime(date_range[0]) if len(date_range) == 2 else pd.to_datetime(default_start)
            end_date = pd.to_datetime(date_range[1]) if len(date_range) == 2 else pd.to_datetime(default_end)

            # Données globales filtrées par date
            df_global = df_ot_clean[(df_ot_clean["Poste travail princ."].isin(all_postes_master_list)) & 
                                    (df_ot_clean["Date de début planifiée"].between(start_date, end_date))].copy()
            avis_global = avis_df_clean[avis_df_clean["Poste travail princ."].isin(all_postes_master_list)].copy()
            df_global = df_global[df_global["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)].drop_duplicates()
            avis_global = avis_global[(avis_global["Ordre"].isna()) | (avis_global["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates()
            if "Statut système" in df_global.columns:
                df_global["Statut OT"] = df_global["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            # Traitement et KPIs globaux
            now = pd.Timestamp.now()
            df_processed_global = process_dataframe(df_global, now)
            results_global = calculate_kpis(df_processed_global, avis_global, all_postes_master_list)

            # ==================================================
            # RÉSUMÉ EXÉCUTIF (sans filtres poste/atelier/division)
            # ==================================================
            st.markdown("---")
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            with col_m1:
                st.metric("Total OT", f"{len(df_global):,}")
            with col_m2:
                taux = results_global['calculated_kpis_df']["TAUX_REALISATION_CORRECTIF/PT"].mean()
                st.metric("Taux Réalisation", f"{taux:.1f}%", delta=f"{taux - CIBLES['TAUX_REALISATION_CORRECTIF/PT']:.1f}% vs cible")
            with col_m3:
                st.metric("Total Postes", len(all_postes_master_list))
            with col_m4:
                st.metric("Période", f"{start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m/%Y')}")

            # ==================================================
            # FILTRES POSTE/ATELIER/DIVISION (pour KPIs et Anomalies)
            # ==================================================
            st.markdown("---")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                selected_postes = st.multiselect("Poste de travail", ["All"] + all_postes_master_list, ["All"], key="poste")
            with col_f2:
                ateliers_options = ["All", "Sulfurique (PS)", "Phosphorique (PP)", "Engrais (TSP/REX)", "Feed (MCP/DCP)"]
                selected_ateliers = st.multiselect("Atelier", ateliers_options, ["All"], key="atelier")
            with col_f3:
                divisions_options = ["All", "SF1", "SF2"]
                selected_divisions = st.multiselect("Division", divisions_options, ["All"], key="division")

            # Postes filtrés
            if "All" in selected_postes or len(selected_postes) == 0:
                selected_postes_filtred = all_postes_master_list
            else:
                selected_postes_filtred = selected_postes

            valid_postes = [p for p in selected_postes_filtred if match_filters(p, selected_ateliers, selected_divisions)]

            # Recalcul KPIs filtrés
            df_filtered = df_processed_global[df_processed_global["Poste travail princ."].isin(valid_postes)].copy()
            avis_filtered = avis_global[avis_global["Poste travail princ."].isin(valid_postes)].copy()
            results_filtered = calculate_kpis(df_filtered, avis_filtered, valid_postes)
            calculated_kpis_df = results_filtered['calculated_kpis_df']
            df_processed = df_filtered

            # Cible
            cible = pd.DataFrame([CIBLES], index=["CIBLE"])

            # Détection anomalies
            df_anomalies_ot, df_anomalies_avis = detect_anomalies(df_processed, calculated_kpis_df, results_filtered, valid_postes, CIBLES)

            # Construction tableau anomalies
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

            # ==================================================
            # BOUTONS TOGGLE KPI / ANOMALIES
            # ==================================================
            st.markdown("---")
            
            if "view_mode" not in st.session_state:
                st.session_state.view_mode = "kpi"
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("📊 TABLEAU DE BORD DES KPIs", 
                           type="primary" if st.session_state.view_mode == "kpi" else "secondary",
                           use_container_width=True):
                    st.session_state.view_mode = "kpi"
            with col_btn2:
                if st.button("⚠️ TABLEAU DE BORD DES ANOMALIES", 
                           type="primary" if st.session_state.view_mode == "anomalie" else "secondary",
                           use_container_width=True):
                    st.session_state.view_mode = "anomalie"

            # ==================================================
            # AFFICHAGE CONDITIONNEL
            # ==================================================
            if st.session_state.view_mode == "kpi":
                st.markdown('<div class="header-kpi">TABLEAU DE BORD DES KPIs</div>', unsafe_allow_html=True)
                
                total_general_kpi = pd.DataFrame(calculated_kpis_df.mean()).T
                total_general_kpi.index = ["Total général"]
                final_kpi = pd.concat([cible, calculated_kpis_df, total_general_kpi]).round(2)
                st.table(final_kpi.style.apply(highlight_kpis, axis=1).format("{:.2f}"))

            else:  # anomalie
                st.markdown('<div class="header-anomalie">TABLEAU DE BORD DES ANOMALIES</div>', unsafe_allow_html=True)
                if not anomalies_dashboard.empty:
                    st.dataframe(anomalies_dashboard, use_container_width=True)
                else:
                    st.info("✅ Aucune anomalie détectée.")

            # ==================================================
            # SYNTHÈSE DES ACTIONS (par KPI quand ALL sélectionné)
            # ==================================================
            st.markdown("---")
            st.markdown("## 📋 Synthèse des Actions KPI")
            
            is_all_postes = "All" in selected_postes or len(selected_postes) == 0

            if is_all_postes:
                # Regroupement par KPI avec somme des OT impactés
                if not df_anomalies_ot.empty:
                    synthese = df_anomalies_ot.groupby("KPI").agg(
                        Nb_OT_impactés=("Nb OT impactés", "sum"),
                        Action_Suggérée=("Action Suggérée", "first")
                    ).reset_index()
                    synthese.columns = ["KPI", "Nombre OT impactés", "Action Suggérée"]
                    synthese = synthese.sort_values("Nombre OT impactés", ascending=False)
                    st.dataframe(synthese, use_container_width=True, hide_index=True)
                else:
                    st.info("✅ Tous les KPIs atteignent leurs cibles. Aucune action requise.")
                
                # Ajouter les avis si anomalie
                if not df_anomalies_avis.empty:
                    st.markdown("### Avis sans ordre")
                    synthese_avis = df_anomalies_avis.groupby("KPI").agg(
                        Nb_Avis_impactés=("Nb Avis impactés", "sum"),
                        Action_Suggérée=("Action Suggérée", "first")
                    ).reset_index()
                    synthese_avis.columns = ["KPI", "Nombre Avis impactés", "Action Suggérée"]
                    st.dataframe(synthese_avis, use_container_width=True, hide_index=True)
            else:
                # Affichage par poste
                if not df_anomalies_ot.empty:
                    display_actions = df_anomalies_ot[["KPI", "Nb OT impactés", "Action Suggérée"]].copy()
                    display_actions.columns = ["KPI", "Nombre OT impactés", "Action Suggérée"]
                    st.dataframe(display_actions, use_container_width=True, hide_index=True)
                else:
                    st.info("✅ Tous les KPIs atteignent leurs cibles. Aucune action requise.")

            # ==================================================
            # CLASSEMENT POSTES
            # ==================================================
            st.markdown("---")
            st.markdown("## 🏆 Classement des Postes par Performance")

            qty_kpis = ["TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois", 
                       "OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois", 
                       "OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]
            qual_kpis = ["appel avis approuvé", "OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]

            class_results = []
            for poste in calculated_kpis_df.index:
                row = calculated_kpis_df.loc[poste]
                s_qty = sum(get_kpi_score(k, row[k], CIBLES[k]) for k in qty_kpis if k in row.index)
                s_qual = sum(get_kpi_score(k, row[k], CIBLES[k]) for k in qual_kpis if k in row.index)
                p_qty = (s_qty / len(qty_kpis) * 100) if qty_kpis else 0
                p_qual = (s_qual / len(qual_kpis) * 100) if qual_kpis else 0
                class_results.append({
                    "Poste travail princ.": poste,
                    "Score KPIs Quantité": p_qty,
                    "Score KPIs Qualité": p_qual,
                    "Total performance ": (p_qty + p_qual) / 2
                })
            df_class = pd.DataFrame(class_results)

            df_class_display = df_class.copy()
            for col in ["Score KPIs Quantité", "Score KPIs Qualité", "Total performance "]:
                df_class_display[col] = df_class_display[col].apply(lambda x: f"{x:.2f} %")

            total_gen_class = pd.DataFrame([{
                "Poste travail princ.": "Total général",
                "Score KPIs Quantité": f"{df_class['Score KPIs Quantité'].mean():.2f} %",
                "Score KPIs Qualité": f"{df_class['Score KPIs Qualité'].mean():.2f} %",
                "Total performance ": f"{df_class['Total performance '].mean():.2f} %"
            }])
            df_class_display = pd.concat([df_class_display, total_gen_class], ignore_index=True)
            st.table(df_class_display.style.apply(highlight_classification, axis=1))

            # Top 5
            col_t1, col_t2, col_t3 = st.columns(3)
            with col_t1:
                st.markdown("#### ⬇️ Top 5 Quantité")
                top5 = df_class.nsmallest(5, "Score KPIs Quantité")[["Poste travail princ.", "Score KPIs Quantité"]].round(2)
                st.dataframe(top5.set_index("Poste travail princ."), use_container_width=True)
            with col_t2:
                st.markdown("#### ⬇️ Top 5 Qualité")
                top5 = df_class.nsmallest(5, "Score KPIs Qualité")[["Poste travail princ.", "Score KPIs Qualité"]].round(2)
                st.dataframe(top5.set_index("Poste travail princ."), use_container_width=True)
            with col_t3:
                st.markdown("#### ⬇️ Top 5 Performance")
                top5 = df_class.nsmallest(5, "Total performance ")[["Poste travail princ.", "Total performance "]].round(2)
                st.dataframe(top5.set_index("Poste travail princ."), use_container_width=True)

            # Graphiques
            st.markdown("---")
            df_class["Métier"] = df_class["Poste travail princ."].apply(get_groupe_metier)
            df_class["Atelier"] = df_class["Poste travail princ."].apply(get_groupe_atelier)
            df_class["Division"] = df_class["Poste travail princ."].apply(get_groupe_division)

            c1, c2, c3 = st.columns(3)
            with c1:
                df_m = df_class.groupby("Métier")["Total performance "].mean().reset_index()
                chart = alt.Chart(df_m).mark_bar(color='#3498db').encode(x='Métier:O', y='Total performance :Q')
                text = chart.mark_text(align='center', baseline='bottom', dy=-10, fontSize=14).encode(text=alt.Text('Total performance :Q', format='.1f'))
                st.altair_chart((chart + text).configure_axisY(labels=False, ticks=False, grid=False, domain=False).configure_view(stroke='transparent').properties(height=200), use_container_width=True)
            with c2:
                df_a = df_class.groupby("Atelier")["Total performance "].mean().reset_index()
                chart = alt.Chart(df_a).mark_bar(color='#e74c3c').encode(x='Atelier:O', y='Total performance :Q')
                text = chart.mark_text(align='center', baseline='bottom', dy=-10, fontSize=14).encode(text=alt.Text('Total performance :Q', format='.1f'))
                st.altair_chart((chart + text).configure_axisY(labels=False, ticks=False, grid=False, domain=False).configure_view(stroke='transparent').properties(height=200), use_container_width=True)
            with c3:
                df_d = df_class.groupby("Division")["Total performance "].mean().reset_index()
                chart = alt.Chart(df_d).mark_bar(color='#2ecc71').encode(x='Division:O', y='Total performance :Q')
                text = chart.mark_text(align='center', baseline='bottom', dy=-10, fontSize=14).encode(text=alt.Text('Total performance :Q', format='.1f'))
                st.altair_chart((chart + text).configure_axisY(labels=False, ticks=False, grid=False, domain=False).configure_view(stroke='transparent').properties(height=200), use_container_width=True)

            # ==================================================
            # EXPORT
            # ==================================================
            st.markdown("---")
            st.markdown("## 📥 Export des Plans d'Action")
            postes_avec_anomalies = anomalies_dashboard[anomalies_dashboard.index != "Total général"].index.tolist() if not anomalies_dashboard.empty else []

            if postes_avec_anomalies:
                selected_poste_export = st.selectbox("Poste à exporter :", options=["All"] + postes_avec_anomalies, key="export")
                
                if st.button("📥 Générer le fichier Excel", type="primary", use_container_width=True):
                    with st.spinner("Génération en cours..."):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            postes_a_traiter = postes_avec_anomalies if selected_poste_export == "All" else [selected_poste_export]

                            for poste_export in postes_a_traiter:
                                kpis_en_defaut = df_anomalies_ot[df_anomalies_ot["Poste travail princ."] == poste_export]["KPI"].unique().tolist()
                                if not df_anomalies_avis.empty and poste_export in df_anomalies_avis["Poste travail princ."].values:
                                    kpis_en_defaut.append("appel avis approuvé")

                                for kpi in kpis_en_defaut:
                                    sheet_data = pd.DataFrame()

                                    if kpi != "appel avis approuvé":
                                        df_poste_exp = df_processed[df_processed["Poste travail princ."] == poste_export].copy()
                                        conditions_map = {
                                            "TAUX_REALISATION_CORRECTIF/PT": (df_poste_exp["Nº appel pl.entret."].fillna(0) == 0) & (~df_poste_exp["Statut OT"].isin(["CLOT", "TCLO"])),
                                            "OT préparation <1 mois": (df_poste_exp["Statut OT"] == "CRÉÉ") & (df_poste_exp["Age préparation"] != "<1 mois"),
                                            "OT préparation >3 mois": (df_poste_exp["Statut OT"] == "CRÉÉ") & (df_poste_exp["Age préparation"] == ">3 mois"),
                                            "OT préparation 1mois< <3mois": (df_poste_exp["Statut OT"] == "CRÉÉ") & (df_poste_exp["Age préparation"] == "1 mois < <3 mois"),
                                            "OT planification <1 mois": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["Contient SOPL"] == 0) & (df_poste_exp["Age planification"] != "<1 mois"),
                                            "OT planification >3 mois": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["Contient SOPL"] == 0) & (df_poste_exp["Age planification"] == ">3 mois"),
                                            "OT planification 1mois< <3mois": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["Contient SOPL"] == 0) & (df_poste_exp["Age planification"] == "1 mois < <3 mois"),
                                            "OT exécution <1 mois": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["Contient SOPL"] == 1) & (df_poste_exp["Age exécution"] != "<1 mois"),
                                            "OT exécution >3 mois": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["Contient SOPL"] == 1) & (df_poste_exp["Age exécution"] == ">3 mois"),
                                            "OT exécution 1mois< <3mois": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["Contient SOPL"] == 1) & (df_poste_exp["Age exécution"] == "1 mois < <3 mois"),
                                            "OT LANC ESTIME": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["OT LANC ESTIME"] == "NON"),
                                            "Backlog préparation caractérisé": (df_poste_exp["Statut OT"] == "CRÉÉ") & (df_poste_exp["Backlog préparation"] == "NON CARACTERISE"),
                                            "Backlog planification caractérisé": (df_poste_exp["Statut OT"] == "LANC") & (df_poste_exp["Backlog planification"] == "NON CARACTERISE"),
                                            "OT CONFIME": df_poste_exp["OT CONFIME"] == "NON",
                                            "OT_COR_EGAL": df_poste_exp["OT_COR_EGAL"] == "NON",
                                        }
                                        if kpi in conditions_map:
                                            subset_ot = df_poste_exp[conditions_map[kpi]]
                                            if not subset_ot.empty:
                                                old_cols = ["Ordre", "Désignation", "Emplacement technique", "Poste travail princ.", "Statut système", "Statut utilisateur", "Date de début planifiée", "Type d'ordre", "Backlog préparation", "Backlog planification"]
                                                new_cols = ["Ordre de travail", "Désignation", "Poste technique", "Poste de travail principal", "Statut système", "Statut utilisateur", "Date de début planifiée", "Type d'ordre", "Caractérisation backlog Préparation", "Caractérisation backlog Planification"]
                                                subset_ot = rename_safe(subset_ot, old_cols, new_cols)
                                                subset_ot["KPI impacté"] = kpi
                                                subset_ot["Action recommandée"] = ACTIONS_KPI.get(kpi, "")
                                                sheet_data = pd.concat([sheet_data, subset_ot])

                                    if kpi == "appel avis approuvé":
                                        subset_avis = results_filtered['avis_df_filtered'][results_filtered['avis_df_filtered']["Poste travail princ."] == poste_export].copy()
                                        if not subset_avis.empty:
                                            old_cols_avis = ["Avis", "Désignation texte", "Emplacement technique", "Poste travail princ.", "Statut utilisateur", "Créé le"]
                                            new_cols_avis = ["Avis", "Désignation", "Poste technique", "Poste de travail principal", "Statut", "Date de création"]
                                            subset_avis = rename_safe(subset_avis, old_cols_avis, new_cols_avis)
                                            subset_avis["KPI impacté"] = kpi
                                            subset_avis["Action recommandée"] = "Créer un OT pour cet Avis ou clarifier son statut."
                                            sheet_data = pd.concat([sheet_data, subset_avis])

                                    if not sheet_data.empty:
                                        base_name = poste_export.replace(" ", "_").replace("/", "_")[:20]
                                        kpi_name = kpi.replace("/", "_").replace(" ", "_")[:10]
                                        sheet_name = f"{base_name}_{kpi_name}"[:31]
                                        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

                        output.seek(0)
                        nom_fichier = "Plan_Action_Tous_Postes.xlsx" if selected_poste_export == "All" else f"Plan_Action_{selected_poste_export.replace(' ', '_')}.xlsx"
                        st.download_button(
                            label="✅ Télécharger le fichier",
                            data=output.getvalue(),
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            else:
                st.info("✅ Aucune anomalie détectée. Export désactivé.")

        except FileNotFoundError:
            st.error("❌ Fichier(s) introuvable(s). Activez 'Charger de nouveaux fichiers' ou placez ot.xlsx et avis.xlsx dans le répertoire.")
        except pd.errors.EmptyDataError:
            st.error("❌ Un des fichiers est vide.")
        except KeyError as e:
            st.error(f"❌ Colonne manquante : {e}")
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
            import traceback
            st.code(traceback.format_exc())

# ==================================================
# EXECUTION
# ==================================================
if __name__ == "__main__":
    try:
        locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_ALL, 'fr_FR')
        except:
            pass
    main()
