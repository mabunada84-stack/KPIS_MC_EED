# -*- coding: utf-8 -*-
"""app.py"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import locale
import altair as alt
import random
import time
from datetime import datetime
import os

# ==================================================
# CSS PERSONNALISÉE
# ==================================================
def inject_custom_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
        
        :root {
            --primary: #1e3a5f;
            --primary-light: #2c5282;
            --accent: #ed8936;
            --success: #38a169;
            --danger: #e53e3e;
            --warning: #d69e2e;
            --bg: #f7fafc;
            --card: #ffffff;
            --text: #1a202c;
            --text-secondary: #718096;
            --border: #e2e8f0;
            --shadow: 0 4px 20px rgba(0,0,0,0.08);
            --radius: 16px;
        }
        
        .stApp {
            background: linear-gradient(135deg, #f0f4f8 0%, #e2e8f0 50%, #f0f4f8 100%);
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        
        .main-header {
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
            padding: 28px 40px;
            border-radius: var(--radius);
            margin-bottom: 20px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.12);
            position: relative;
            overflow: hidden;
        }
        
        .main-header::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -10%;
            width: 300px;
            height: 300px;
            background: rgba(255,255,255,0.05);
            border-radius: 50%;
        }
        
        .main-header h1 {
            color: white;
            font-size: 30px;
            font-weight: 800;
            margin: 0;
            letter-spacing: -0.5px;
        }
        
        .main-header .subtitle {
            color: rgba(255,255,255,0.7);
            font-size: 14px;
            font-weight: 400;
            margin-top: 4px;
        }
        
        .main-header .date-badge {
            position: absolute;
            top: 28px;
            right: 40px;
            background: rgba(255,255,255,0.15);
            backdrop-filter: blur(10px);
            padding: 8px 18px;
            border-radius: 30px;
            color: white;
            font-size: 13px;
            font-weight: 500;
            border: 1px solid rgba(255,255,255,0.2);
        }
        
        .section-title {
            font-size: 18px;
            font-weight: 700;
            color: var(--primary);
            margin-bottom: 12px;
            padding-left: 14px;
            border-left: 4px solid var(--accent);
        }
        
        .chart-container {
            background: var(--card);
            border-radius: var(--radius);
            padding: 20px;
            box-shadow: var(--shadow);
            border: 1px solid var(--border);
        }
        
        .chart-title {
            font-size: 14px;
            font-weight: 700;
            color: var(--primary);
            margin-bottom: 12px;
            text-align: center;
        }
        
        .synthese-row {
            display: flex;
            align-items: center;
            padding: 12px 18px;
            background: var(--card);
            border-radius: 10px;
            margin-bottom: 6px;
            border: 1px solid var(--border);
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
            transition: all 0.2s ease;
        }
        
        .synthese-row:hover {
            transform: translateX(4px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.08);
            border-left: 4px solid var(--accent);
        }
        
        .synthese-kpi-name {
            font-weight: 700;
            color: var(--primary);
            font-size: 13px;
            min-width: 260px;
        }
        
        .synthese-count {
            background: var(--primary);
            color: white;
            padding: 5px 14px;
            border-radius: 20px;
            font-weight: 800;
            font-size: 14px;
            min-width: 50px;
            text-align: center;
            margin: 0 14px;
        }
        
        .synthese-action {
            color: var(--text-secondary);
            font-size: 12px;
            flex: 1;
        }
        
        .top-badge {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 6px;
            font-size: 10px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .top-badge.red { background: #fed7d7; color: #c53030; }
        .top-badge.orange { background: #feebc8; color: #c05621; }
        .top-badge.yellow { background: #fefcbf; color: #975a16; }
        
        .empty-state {
            text-align: center;
            padding: 50px 20px;
            color: var(--text-secondary);
        }
        
        .empty-state .icon { font-size: 44px; margin-bottom: 14px; }
        .empty-state h3 { color: var(--text); font-size: 17px; font-weight: 600; margin-bottom: 6px; }
        
        /* SIDEBAR */
        div[data-testid="stSidebar"] {
            background: linear-gradient(180deg, var(--primary) 0%, #0f2744 100%);
        }
        div[data-testid="stSidebar"] * { color: rgba(255,255,255,0.9) !important; }
        div[data-testid="stSidebar"] .stSelectbox label,
        div[data-testid="stSidebar"] .stMultiSelect label,
        div[data-testid="stSidebar"] .stDateInput label,
        div[data-testid="stSidebar"] .stCheckbox label {
            color: rgba(255,255,255,0.8) !important;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        div[data-testid="stSidebar"] div[data-testid="stWidget"] {
            background: rgba(255,255,255,0.08);
            border-radius: 10px;
            padding: 4px 12px;
            margin-bottom: 6px;
            border: 1px solid rgba(255,255,255,0.1);
        }
        div[data-testid="stSidebar"] .stSelectbox > div > div,
        div[data-testid="stSidebar"] .stMultiSelect > div > div,
        div[data-testid="stSidebar"] .stDateInput > div > div {
            background: rgba(255,255,255,0.95) !important;
            border-radius: 8px;
        }
        
        .stTableContainer { width: 100% !important; }
        
        .dataframe th {
            background: var(--primary) !important;
            color: white !important;
            font-weight: 700 !important;
            font-size: 11px !important;
            text-transform: uppercase !important;
            letter-spacing: 0.3px !important;
            padding: 10px 8px !important;
            border: none !important;
            white-space: nowrap;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        
        .dataframe td {
            padding: 8px !important;
            font-size: 11px !important;
            border-bottom: 1px solid var(--border) !important;
        }
        
        .dataframe tr:hover td { background: #f7fafc !important; }
        .dataframe tr:nth-child(even) td { background: #f9fbfd !important; }
        .dataframe tr:nth-child(even):hover td { background: #edf2f7 !important; }
        
        .stTabs [data-baseweb="tab-list"] {
            gap: 4px;
            background: #edf2f7;
            padding: 4px;
            border-radius: 10px;
            margin-bottom: 16px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: 600;
            font-size: 14px;
        }
        .stTabs [aria-selected="true"] {
            background: white !important;
            color: var(--primary) !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
            border: none;
            border-radius: 10px;
            padding: 12px 28px;
            font-weight: 700;
            font-size: 14px;
            box-shadow: 0 4px 12px rgba(30, 58, 95, 0.3);
            width: 100%;
        }
        .stButton > button[kind="primary"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(30, 58, 95, 0.4);
        }
        .stDownloadButton > button {
            background: linear-gradient(135deg, var(--success) 0%, #276749 100%);
            border: none;
            border-radius: 10px;
            padding: 14px 28px;
            font-weight: 700;
            font-size: 15px;
            color: white;
            width: 100%;
            box-shadow: 0 4px 12px rgba(56, 161, 105, 0.3);
        }
        
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #a1a1a1; }

        @media (max-width: 768px) {
            .main-header { padding: 18px 20px; }
            .main-header h1 { font-size: 20px; }
            .main-header .date-badge { position: static; margin-top: 8px; display: inline-block; }
            .synthese-kpi-name { min-width: 140px; font-size: 11px; }
            .synthese-action { font-size: 11px; }
        }
    </style>
    """, unsafe_allow_html=True)


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

    if "hse_affiche" not in st.session_state:
        st.session_state.hse_affiche = False

    if not st.session_state.hse_affiche:
        consigne = random.choice(consignes)
        st.markdown("""
        <div style="min-height:100vh; display:flex; flex-direction:column; align-items:center; justify-content:center; background:linear-gradient(135deg, #1a365d 0%, #2d3748 50%, #1a365d 100%); padding:40px;">
            <div style="font-size:64px; margin-bottom:20px;">🦺</div>
            <h1 style="text-align:center; font-size:42px; color:#fff; font-weight:900; margin:0; letter-spacing:-1px;">
                HSE - CONSIGNE DE SÉCURITÉ
            </h1>
            <p style="text-align:center; color:rgba(255,255,255,0.6); font-size:18px; margin-top:8px; font-weight:400; letter-spacing:3px; text-transform:uppercase;">
                Sécurité &bull; Santé &bull; Environnement
            </p>
            <div style="background:linear-gradient(135deg, #f6e05e 0%, #ed8936 100%); padding:36px 48px; border-radius:20px; font-size:28px; font-weight:700; text-align:center; margin-top:40px; margin-bottom:40px; color:#1a202c; max-width:800px; box-shadow:0 20px 60px rgba(0,0,0,0.3);">
                ⚠️ {consigne}
            </div>
            <h2 style="text-align:center; color:#48bb78; font-size:32px; font-weight:900; letter-spacing:-0.5px;">
                Aucun travail n'est plus urgent que la sécurité
            </h2>
            <div style="margin-top:40px; width:200px; height:4px; background:rgba(255,255,255,0.1); border-radius:2px; overflow:hidden;">
                <div style="width:100%; height:100%; background:linear-gradient(90deg, #48bb78, #38a169); border-radius:2px; animation:loading 5.5s ease-in-out forwards;"></div>
            </div>
            <style>@keyframes loading {{ from {{ width: 0%; }} to {{ width: 100%; }} }}</style>
        </div>
        """.format(consigne=consigne), unsafe_allow_html=True)

        time.sleep(6)
        st.session_state.hse_affiche = True
        st.rerun()
        st.stop()

    # ==================================================
    # FONCTIONS UTILITAIRES
    # ==================================================
    def rename_safe(df, old_names, new_names):
        mapping = {old: new for old, new in zip(old_names, new_names) if old in df.columns}
        return df.rename(columns=mapping)

    def contient_mot(texte, liste_mots):
        texte = str(texte)
        return any(mot in texte for liste_mots_element in liste_mots for mot in liste_mots_element.split())

    def categorie_age(age):
        if age <= 1: return "<1 mois"
        elif age >= 3: return ">3 mois"
        else: return "1 mois < <3 mois"

    def calcul_kpi(numerateur, denominateur, si_zero=100):
        return np.where(denominateur == 0, si_zero, (numerateur / denominateur) * 100)

    def creer_pivot(dataframe, filtre, colonne, postes_to_reindex):
        pivot = pd.pivot_table(dataframe[filtre], index="Poste travail princ.", columns=colonne, values="Ordre", aggfunc="count", fill_value=0)
        return pivot.reindex(postes_to_reindex, fill_value=0)

    def exclure_cresseurs(df):
        if "Poste travail princ." in df.columns:
            return df[~df["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)].copy()
        return df

    def calculate_all_kpis_and_intermediate_dfs(df_input, avis_df_input, now_timestamp, all_postes_list):
        results = {}
        df_temp = df_input.copy()
        avis_df_temp = avis_df_input.copy()

        mots_prep = ["CRPR ATPD", "CRPR ATMR", "CRPR ATER", "CRPR ATRS", "CRPR ATMO", "ATPD", "ATMR", "ATER", "ATRS", "ATMO"]
        df_temp["Backlog préparation"] = np.where(df_temp["Statut utilisateur"].apply(lambda x: contient_mot(x, mots_prep)), "CARACTERISE", "NON CARACTERISE")
        mots_planif = ["ATPL ATEI", "ATPL ATAL", "ATPL ATER", "ATPL AGAR", "ATPL ATHS", "ATEI", "ATAL", "ATAS", "AGAR", "ATHS"]
        df_temp["Backlog planification"] = np.where(df_temp["Statut utilisateur"].apply(lambda x: contient_mot(x, mots_planif)), "CARACTERISE", "NON CARACTERISE")

        for date_col, age_mois_col, age_col in [('Créé le', "Age mois préparation", "Age préparation"), ('Date de début planifiée', "Age mois planification", "Age planification"), ('Date de début planifiée', "Age mois exécution", "Age exécution")]:
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
        results['df_processed'] = df_temp

        analyse = creer_pivot(df_temp, df_temp["Nº appel pl.entret."].fillna(0) == 0, "Statut OT", all_postes_list)
        for col in ["CLOT", "CRÉÉ", "LANC", "TCLO"]: analyse[col] = analyse.get(col, 0)
        analyse["Total"] = analyse[["CLOT", "CRÉÉ", "LANC", "TCLO"]].sum(axis=1)
        analyse["TAUX_REALISATION_CORRECTIF/PT"] = calcul_kpi(analyse["TCLO"], analyse["Total"])

        prep = creer_pivot(df_temp, df_temp["Statut OT"] == "CRÉÉ", "Age préparation", all_postes_list)
        for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: prep[col] = prep.get(col, 0)
        prep["Total"] = prep[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
        prep["OT préparation <1 mois"] = calcul_kpi(prep["<1 mois"], prep["Total"])
        prep["OT préparation >3 mois"] = calcul_kpi(prep[">3 mois"], prep["Total"], 0)
        prep["OT préparation 1mois< <3mois"] = calcul_kpi(prep["1 mois < <3 mois"], prep["Total"], 0)

        planif = creer_pivot(df_temp, (df_temp["Statut OT"] == "LANC") & (df_temp["Contient SOPL"] == 0), "Age planification", all_postes_list)
        for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: planif[col] = planif.get(col, 0)
        planif["Total"] = planif[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
        planif["OT planification <1 mois"] = calcul_kpi(planif["<1 mois"], planif["Total"])
        planif["OT planification >3 mois"] = calcul_kpi(planif[">3 mois"], planif["Total"], 0)
        planif["OT planification 1mois< <3mois"] = calcul_kpi(planif["1 mois < <3 mois"], planif["Total"], 0)

        execu = creer_pivot(df_temp, (df_temp["Statut OT"] == "LANC") & (df_temp["Contient SOPL"] == 1), "Age exécution", all_postes_list)
        for col in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: execu[col] = execu.get(col, 0)
        execu["Total"] = execu[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
        execu["OT exécution <1 mois"] = calcul_kpi(execu["<1 mois"], execu["Total"])
        execu["OT exécution >3 mois"] = calcul_kpi(execu[">3 mois"], execu["Total"], 0)
        execu["OT exécution 1mois< <3mois"] = calcul_kpi(execu["1 mois < <3 mois"], execu["Total"], 0)

        lanc = pd.pivot_table(df_temp[df_temp["Statut OT"] == "LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
        for col in ["OUI", "NON"]: lanc[col] = lanc.get(col, 0)
        lanc["Total"] = lanc["OUI"] + lanc["NON"]
        lanc["OT LANC ESTIME"] = calcul_kpi(lanc["OUI"], lanc["Total"])

        prep_carac = pd.pivot_table(df_temp[df_temp["Statut OT"] == "CRÉÉ"], index="Poste travail princ.", columns="Backlog préparation", values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
        for col in ["CARACTERISE", "NON CARACTERISE"]: prep_carac[col] = prep_carac.get(col, 0)
        prep_carac["Total"] = prep_carac["CARACTERISE"] + prep_carac["NON CARACTERISE"]
        prep_carac["Backlog préparation caractérisé"] = calcul_kpi(prep_carac["CARACTERISE"], prep_carac["Total"])

        planif_carac = pd.pivot_table(df_temp[df_temp["Statut OT"] == "LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
        for col in ["CARACTERISE", "NON CARACTERISE"]: planif_carac[col] = planif_carac.get(col, 0)
        planif_carac["Total"] = planif_carac["CARACTERISE"] + planif_carac["NON CARACTERISE"]
        planif_carac["Backlog planification caractérisé"] = calcul_kpi(planif_carac["CARACTERISE"], planif_carac["Total"])

        for kpi_name, col_name in [("OT CONFIME", "OT CONFIME"), ("OT_COR_EGAL", "OT_COR_EGAL")]:
            df_pivot = pd.pivot_table(df_temp, index="Poste travail princ.", columns=col_name, values="Ordre", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
            for c in ["OUI", "NON"]: df_pivot[c] = df_pivot.get(c, 0)
            df_pivot["Total"] = df_pivot["OUI"] + df_pivot["NON"]
            df_pivot[col_name] = calcul_kpi(df_pivot["OUI"], df_pivot["Total"])
            results[kpi_name.lower().replace(" ", "_")] = df_pivot

        avis_df_filtered_local = avis_df_temp[(avis_df_temp["Ordre"].isna()) | (avis_df_temp["Ordre"].astype(str).str.strip() == "")].copy()
        results['avis_df_filtered'] = avis_df_filtered_local

        tableau_croise_avis = pd.pivot_table(avis_df_filtered_local, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(all_postes_list, fill_value=0)
        for col in ["APRQ", "APRV", "APRV AVAU", "REJT"]: tableau_croise_avis[col] = tableau_croise_avis.get(col, 0)
        tableau_croise_avis["Total"] = tableau_croise_avis[["APRQ", "APRV", "APRV AVAU", "REJT"]].sum(axis=1)
        tableau_croise_avis["appel avis approuvé"] = calcul_kpi(tableau_croise_avis["APRV"], tableau_croise_avis["Total"])

        calculated_kpis_df = pd.concat([
            analyse[["TAUX_REALISATION_CORRECTIF/PT"]], prep[["OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois"]],
            planif[["OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois"]], execu[["OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]],
            tableau_croise_avis[["appel avis approuvé"]], lanc[["OT LANC ESTIME"]], prep_carac[["Backlog préparation caractérisé"]],
            planif_carac[["Backlog planification caractérisé"]], results['ot_confime'][["OT CONFIME"]], results['ot_cor_egal'][["OT_COR_EGAL"]]
        ], axis=1)

        results['calculated_kpis_df'] = calculated_kpis_df
        return results

    def highlight_kpis(row):
        styles = []
        if row.name == 'CIBLE': return ['padding: 8px 6px; background-color: #1e3a5f; color: #ffffff; font-weight: 700; font-size: 11px;'] * len(row)
        if row.name == 'Total général': return ['padding: 8px 6px; background-color: #e2e8f0; color: #1a202c; font-weight: 800; font-size: 11px; border-top: 3px solid #1e3a5f;'] * len(row)
        for col in row.index:
            s = 'padding: 8px 6px; font-size: 11px;'
            try: v = float(row[col])
            except: styles.append(s); continue
            if col in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
                s += 'background-color:#c6efce; color:#006100; font-weight:600;' if v >= 80 else ('background-color:#ffeb9c; color:#9c6500; font-weight:600;' if v >= 75 else 'background-color:#ffc7ce; color:#9c0006; font-weight:600;')
            elif col in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
                s += 'background-color:#c6efce; color:#006100; font-weight:600;' if v <= 15 else 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
            elif col in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
                s += 'background-color:#c6efce; color:#006100; font-weight:600;' if v <= 5 else 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
            elif col == "TAUX_REALISATION_CORRECTIF/PT":
                s += 'background-color:#c6efce; color:#006100; font-weight:600;' if v >= 85 else ('background-color:#ffeb9c; color:#9c6500; font-weight:600;' if v >= 80 else 'background-color:#ffc7ce; color:#9c0006; font-weight:600;')
            elif col == "appel avis approuvé":
                s += 'background-color:#c6efce; color:#006100; font-weight:600;' if v >= 95 else ('background-color:#ffeb9c; color:#9c6500; font-weight:600;' if v >= 90 else 'background-color:#ffc7ce; color:#9c0006; font-weight:600;')
            elif col in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
                s += 'background-color:#c6efce; color:#006100; font-weight:600;' if v >= 100 else ('background-color:#ffeb9c; color:#9c6500; font-weight:600;' if v >= 95 else 'background-color:#ffc7ce; color:#9c0006; font-weight:600;')
            styles.append(s)
        return styles

    def highlight_anomalies(row):
        styles = []
        if row.name == 'Total général': return ['padding: 8px 6px; background-color: #1e3a5f; color: #ffffff; font-weight: 800; font-size: 11px;'] * len(row)
        for val in row:
            try: v = int(val)
            except: styles.append('padding: 8px 6px; font-size: 11px;'); continue
            if v == 0: styles.append('padding: 8px 6px; font-size: 11px; color: #a0aec0;')
            elif v <= 3: styles.append('padding: 8px 6px; font-size: 11px; background-color: #ffeb9c; color: #9c6500; font-weight: 600;')
            elif v <= 10: styles.append('padding: 8px 6px; font-size: 11px; background-color: #fed7d7; color: #c53030; font-weight: 600;')
            else: styles.append('padding: 8px 6px; font-size: 11px; background-color: #fc8181; color: #742a2a; font-weight: 800;')
        return styles

    def highlight_classification_table_kpis(row):
        styles = ['padding: 8px 6px; font-size: 11px;'] * len(row)
        for i, col in enumerate(row.index):
            if col in ["Score KPIs Quantité", "Score KPIs Qualité", "Total performance "]:
                try:
                    s = str(row[col]); v = float(s.replace(' %', '').strip()) if ' %' in s else float(s)
                    if v >= 90: styles[i] = 'padding: 8px 6px; font-size: 11px; background-color: #c6efce; color: #006100; font-weight: 700;'
                    elif v >= 80: styles[i] = 'padding: 8px 6px; font-size: 11px; background-color: #ffeb9c; color: #9c6500; font-weight: 700;'
                    else: styles[i] = 'padding: 8px 6px; font-size: 11px; background-color: #ffc7ce; color: #9c0006; font-weight: 700;'
                except: pass
            if row.name == 'Total général':
                styles[i] = 'padding: 8px 6px; font-size: 11px; background-color: #e2e8f0; color: #1a202c; font-weight: 800; border-top: 3px solid #1e3a5f;'
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

    # ==================================================
    # SIDEBAR
    # ==================================================
    with st.sidebar:
        st.markdown("""
        <div style="padding: 16px 0 8px 0;">
            <div style="font-size:24px; margin-bottom:4px;">⚙️</div>
            <div style="font-size:15px; font-weight:800; color:white;">Filtres & Paramètres</div>
            <div style="font-size:10px; color:rgba(255,255,255,0.5); text-transform:uppercase; letter-spacing:1px;">Configuration</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        use_new_files = st.toggle("📁 Charger de nouveaux fichiers", value=False, key="toggle_files")

        ot_file = None
        avis_file = None

        if use_new_files:
            ot_file = st.file_uploader("Fichier OT", type=["xlsx"], key="up_ot")
            avis_file = st.file_uploader("Fichier AVIS", type=["xlsx"], key="up_avis")
        else:
            date_fichier = datetime.now().strftime("%d/%m/%Y")
            if os.path.exists("ot.xlsx"):
                timestamp = os.path.getmtime("ot.xlsx")
                date_fichier = datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y")
            st.markdown(f"""
            <div style="background:rgba(255,255,255,0.1); padding:10px 14px; border-radius:8px; border:1px solid rgba(255,255,255,0.15);">
                <div style="font-size:10px; color:rgba(255,255,255,0.5); text-transform:uppercase; letter-spacing:1px;">Données</div>
                <div style="font-size:13px; color:white; font-weight:600; margin-top:2px;">📅 {date_fichier}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("**🎯 Postes**")
        selected_postes = st.multiselect("Poste de travail", ["All"] + all_postes_master_list if 'all_postes_master_list' in dir() else ["All"], ["All"], key="sel_postes")
        st.markdown("**🏭 Atelier**")
        selected_ateliers = st.multiselect("Atelier", ["All", "Sulfurique (PS)", "Phosphorique (PP)", "Engrais (TSP/REX)", "Feed (MCP/DCP)"], ["All"], key="sel_ateliers")
        st.markdown("**🏢 Division**")
        selected_divisions = st.multiselect("Division", ["All", "SF1", "SF2"], ["All"], key="sel_div")
        st.markdown("---")
        st.markdown("**📅 Période**")
        date_range = st.date_input("Date début planifiée", value=(datetime(2025, 1, 1).date(), datetime.today().date()), format="DD/MM/YYYY", key="date_range")

    if not use_new_files or (ot_file is not None and avis_file is not None):
        try:
            if use_new_files:
                df_ot_raw = pd.read_excel(ot_file)
                avis_df_raw = pd.read_excel(avis_file)
                date_fichier = datetime.now().strftime("%d/%m/%Y")
            else:
                df_ot_raw = pd.read_excel("ot.xlsx")
                avis_df_raw = pd.read_excel("avis.xlsx")

            # Exclure cresseurs sur les fichiers bruts
            df_ot_raw = exclure_cresseurs(df_ot_raw)
            avis_df_raw = exclure_cresseurs(avis_df_raw)

            for col in ["Créé le", "Date de début planifiée", "Date de clôture", "Début réel", "Fin réelle"]:
                if col in df_ot_raw.columns: df_ot_raw[col] = pd.to_datetime(df_ot_raw[col], errors="coerce")
            for col in ["Créé le", "Début souhaité", "Date de la clôture"]:
                if col in avis_df_raw.columns: avis_df_raw[col] = pd.to_datetime(avis_df_raw[col], errors="coerce")

            # Liste des postes sans cresseurs
            all_postes_master_list = sorted(
                df_ot_raw[
                    df_ot_raw["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)
                ]["Poste travail princ."].dropna().unique().tolist()
            )

            if "All" in selected_postes or len(selected_postes) == 0: selected_postes = all_postes_master_list
            if "All" in selected_ateliers or len(selected_ateliers) == 0: selected_ateliers = ["All"]
            if "All" in selected_divisions or len(selected_divisions) == 0: selected_divisions = ["All"]

            if len(date_range) == 2:
                start_date = pd.to_datetime(date_range[0])
                end_date = pd.to_datetime(date_range[1])
            else:
                start_date = pd.to_datetime(datetime(2025, 1, 1))
                end_date = pd.to_datetime(datetime.today())

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

            df = df_ot_raw[(df_ot_raw["Poste travail princ."].isin(valid_postes)) & (df_ot_raw["Date de début planifiée"].between(start_date, end_date))].copy()
            avis_df = avis_df_raw[avis_df_raw["Poste travail princ."].isin(valid_postes)].copy()

            df = df[df["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)].drop_duplicates()
            avis_df = avis_df[(avis_df["Ordre"].isna()) | (avis_df["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates()

            # Exclusion cresseurs finale
            df = exclure_cresseurs(df)
            avis_df = exclure_cresseurs(avis_df)

            if "Statut système" in df.columns: df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            now = pd.Timestamp.now()
            results = calculate_all_kpis_and_intermediate_dfs(df, avis_df, now, valid_postes)
            calculated_kpis_df = results['calculated_kpis_df']
            df_processed = results['df_processed']

            cible = pd.DataFrame([{
                "TAUX_REALISATION_CORRECTIF/PT": 85, "OT préparation <1 mois": 80, "OT préparation >3 mois": 5, "OT préparation 1mois< <3mois": 15,
                "OT planification <1 mois": 80, "OT planification >3 mois": 5, "OT planification 1mois< <3mois": 15,
                "OT exécution <1 mois": 80, "OT exécution >3 mois": 5, "OT exécution 1mois< <3mois": 15,
                "appel avis approuvé": 95, "OT LANC ESTIME": 100, "Backlog préparation caractérisé": 100,
                "Backlog planification caractérisé": 100, "OT CONFIME": 100, "OT_COR_EGAL": 100,
            }], index=["CIBLE"])

            # ==================================================
            # ANOMALIES
            # ==================================================
            anomalies_ot_records = []
            anomalies_avis_records = []

            for poste in valid_postes:
                if poste not in df_processed["Poste travail princ."].values: continue
                df_poste = df_processed[df_processed["Poste travail princ."] == poste]
                avis_poste = results['avis_df_filtered'][results['avis_df_filtered']["Poste travail princ."] == poste]

                checks = [
                    ("TAUX_REALISATION_CORRECTIF/PT", "TAUX_REALISATION_CORRECTIF/PT",
                     df_poste[(df_poste["Nº appel pl.entret."].fillna(0) == 0) & (~df_poste["Statut OT"].isin(["CLOT", "TCLO"]))],
                     "Améliorer le taux de réalisation des OT."),
                    ("OT préparation <1 mois", "OT préparation <1 mois",
                     df_poste[(df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Age préparation"] != "<1 mois")],
                     "Réduire l'âge de préparation des OT (< 1 mois)."),
                    ("OT préparation >3 mois", "OT préparation >3 mois",
                     df_poste[(df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Age préparation"] == ">3 mois")],
                     "Traiter les OT avec préparation > 3 mois."),
                    ("OT planification <1 mois", "OT planification <1 mois",
                     df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 0) & (df_poste["Age planification"] != "<1 mois")],
                     "Réduire l'âge de planification des OT (< 1 mois)."),
                    ("OT planification >3 mois", "OT planification >3 mois",
                     df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 0) & (df_poste["Age planification"] == ">3 mois")],
                     "Traiter les OT avec planification > 3 mois."),
                    ("OT exécution <1 mois", "OT exécution <1 mois",
                     df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 1) & (df_poste["Age exécution"] != "<1 mois")],
                     "Réduire l'âge d'exécution des OT (< 1 mois)."),
                    ("OT exécution >3 mois", "OT exécution >3 mois",
                     df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 1) & (df_poste["Age exécution"] == ">3 mois")],
                     "Traiter les OT avec exécution > 3 mois."),
                    ("OT LANC ESTIME", "OT LANC ESTIME",
                     df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["OT LANC ESTIME"] == "NON")],
                     "Estimer les coûts des OT lancés."),
                    ("Backlog préparation caractérisé", "Backlog préparation caractérisé",
                     df_poste[(df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Backlog préparation"] == "NON CARACTERISE")],
                     "Caractériser le backlog de préparation."),
                    ("Backlog planification caractérisé", "Backlog planification caractérisé",
                     df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Backlog planification"] == "NON CARACTERISE")],
                     "Caractériser le backlog de planification."),
                    ("OT CONFIME", "OT CONFIME",
                     df_poste[df_poste["OT CONFIME"] == "NON"],
                     "Confirmer les OT terminés."),
                    ("OT_COR_EGAL", "OT_COR_EGAL",
                     df_poste[df_poste["OT_COR_EGAL"] == "NON"],
                     "Rapprocher les coûts réels et budgétés."),
                ]

                for kpi_name, cible_col, subset, action in checks:
                    val_kpi = calculated_kpis_df.loc[poste, kpi_name] if poste in calculated_kpis_df.index else 100
                    if pd.notna(val_kpi) and val_kpi < cible.loc['CIBLE', cible_col]:
                        cnt = len(subset)
                        if cnt > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": kpi_name, "Nb OT impactés": cnt, "Action Suggérée": action})

                val_avis = calculated_kpis_df.loc[poste, "appel avis approuvé"] if poste in calculated_kpis_df.index else 100
                if pd.notna(val_avis) and val_avis < cible.loc['CIBLE', "appel avis approuvé"]:
                    cnt = len(avis_poste)
                    if cnt > 0: anomalies_avis_records.append({"Poste travail princ.": poste, "KPI": "appel avis approuvé", "Nb OT impactés": cnt, "Action Suggérée": "Créer un OT pour les avis sans ordre."})

            df_anomalies_ot = pd.DataFrame(anomalies_ot_records)
            df_anomalies_avis = pd.DataFrame(anomalies_avis_records)

            if not df_anomalies_ot.empty:
                pivot_ot = df_anomalies_ot.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb OT impactés", aggfunc="sum", fill_value=0)
            else: pivot_ot = pd.DataFrame()

            if not df_anomalies_avis.empty:
                pivot_avis = df_anomalies_avis.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb OT impactés", aggfunc="sum", fill_value=0)
                if "appel avis approuvé" in pivot_avis.columns: pivot_avis = pivot_avis.rename(columns={"appel avis approuvé": "Nb Avis sans ordre"})
            else: pivot_avis = pd.DataFrame()

            anomalies_dashboard = pivot_ot.join(pivot_avis, how='outer').fillna(0).astype(int)
            if not anomalies_dashboard.empty:
                anomalies_dashboard["Total éléments impactés"] = anomalies_dashboard.sum(axis=1)
                total_row = pd.DataFrame(anomalies_dashboard.sum()).T; total_row.index = ["Total général"]
                anomalies_dashboard = pd.concat([anomalies_dashboard, total_row])
            else: anomalies_dashboard = pd.DataFrame()

            # Synthèse par KPI
            all_anomalies = pd.concat([df_anomalies_ot, df_anomalies_avis], ignore_index=True) if not df_anomalies_ot.empty or not df_anomalies_avis.empty else pd.DataFrame()
            if not all_anomalies.empty:
                synthese_par_kpi = all_anomalies.groupby(["KPI", "Action Suggérée"])["Nb OT impactés"].sum().reset_index().sort_values("Nb OT impactés", ascending=False).reset_index(drop=True)
            else: synthese_par_kpi = pd.DataFrame()

            # Classification (sans les colonnes Métier/Atelier/Division dans le tableau affiché)
            qty_kpis = ["TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois", "OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois", "OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]
            qual_kpis = ["appel avis approuvé", "OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]

            class_results = []
            for poste in calculated_kpis_df.index:
                row = calculated_kpis_df.loc[poste]
                s_qty = sum(get_kpi_score(k, row[k], cible.loc['CIBLE', k]) for k in qty_kpis if k in row.index)
                s_qual = sum(get_kpi_score(k, row[k], cible.loc['CIBLE', k]) for k in qual_kpis if k in row.index)
                p_qty = (s_qty / len(qty_kpis) * 100) if qty_kpis else 0
                p_qual = (s_qual / len(qual_kpis) * 100) if qual_kpis else 0
                class_results.append({"Poste travail princ.": poste, "Score KPIs Quantité": p_qty, "Score KPIs Qualité": p_qual, "Total performance ": (p_qty+p_qual)/2})
            df_class = pd.DataFrame(class_results)

            # Colonnes Métier/Atelier/Division : uniquement pour les charts, PAS dans le tableau
            df_class["Métier"] = df_class["Poste travail princ."].apply(get_groupe_metier)
            df_class["Atelier"] = df_class["Poste travail princ."].apply(get_groupe_atelier)
            df_class["Division"] = df_class["Poste travail princ."].apply(get_groupe_division)

            # Tableau KPIs final
            total_general_kpi = pd.DataFrame(calculated_kpis_df.mean()).T; total_general_kpi.index = ["Total général"]
            final_kpi = pd.concat([cible, calculated_kpis_df, total_general_kpi]).round(2)

            # Tableau Classement SANS Métier/Atelier/Division
            df_class_display = df_class[["Poste travail princ.", "Score KPIs Quantité", "Score KPIs Qualité", "Total performance "]].copy()
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

            # ==================================================
            # HEADER
            # ==================================================
            st.markdown(f"""
            <div class="main-header">
                <h1>📊 KPI Dashboard MC & FEED</h1>
                <div class="subtitle">Maintenance Conditionnelle • Suivi des indicateurs de performance</div>
                <div class="date-badge">📅 {date_fichier}</div>
            </div>
            """, unsafe_allow_html=True)

            # ==================================================
            # ONGLETS
            # ==================================================
            tab_kpi, tab_anomalies = st.tabs(["📊 TABLEAU DE BORD DES KPIs", "🚨 TABLEAU DE BORD DES ANOMALIES"])

            with tab_kpi:
                st.markdown('<p class="section-title">Indicateurs par Poste de Travail</p>', unsafe_allow_html=True)
                st.table(final_kpi.style.apply(highlight_kpis, axis=1).format("{:.2f}"))

                st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
                st.markdown('<p class="section-title">🏆 Classement des Postes par Performance</p>', unsafe_allow_html=True)
                st.table(df_class_display.style.apply(highlight_classification_table_kpis, axis=1))

            with tab_anomalies:
                if not anomalies_dashboard.empty:
                    st.markdown('<p class="section-title">Anomalies par Poste de Travail</p>', unsafe_allow_html=True)
                    st.table(anomalies_dashboard.style.apply(highlight_anomalies, axis=1))
                else:
                    st.markdown("""
                    <div class="empty-state">
                        <div class="icon">✅</div>
                        <h3>Aucune anomalie détectée</h3>
                        <p>Tous les KPIs atteignent leurs cibles.</p>
                    </div>
                    """, unsafe_allow_html=True)

            # ==================================================
            # SYNTHÈSE PAR KPI
            # ==================================================
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">📋 Synthèse des Actions par KPI</p>', unsafe_allow_html=True)

            if not synthese_par_kpi.empty:
                for idx, row in synthese_par_kpi.iterrows():
                    nb = int(row["Nb OT impactés"])
                    badge = "red" if nb >= 20 else ("orange" if nb >= 5 else "yellow")
                    label = "Critique" if nb >= 20 else ("Moyen" if nb >= 5 else "Mineur")
                    st.markdown(f"""
                    <div class="synthese-row">
                        <div class="synthese-kpi-name">{row['KPI']}</div>
                        <div class="synthese-count">{nb}</div>
                        <div class="synthese-action">{row['Action Suggérée']}</div>
                        <span class="top-badge {badge}">{label}</span>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="empty-state" style="padding:30px;">
                    <div class="icon">🎉</div>
                    <h3>Aucune action requise</h3>
                    <p>Tous les indicateurs sont dans le vert.</p>
                </div>
                """, unsafe_allow_html=True)

            # ==================================================
            # TOP 5
            # ==================================================
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">🔻 Top 5 Postes à Améliorer</p>', unsafe_allow_html=True)

            col_top1, col_top2, col_top3 = st.columns(3)
            with col_top1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">📦 Impact Quantité</div>', unsafe_allow_html=True)
                t5 = df_class.nsmallest(5, "Score KPIs Quantité")[["Poste travail princ.", "Score KPIs Quantité"]].reset_index(drop=True)
                t5.columns = ["Poste", "Score (%)"]
                t5["Score (%)"] = t5["Score (%)"].round(1)
                st.dataframe(t5, use_container_width=True, hide_index=True, height=230)
                st.markdown('</div>', unsafe_allow_html=True)

            with col_top2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">🎯 Impact Qualité</div>', unsafe_allow_html=True)
                t5 = df_class.nsmallest(5, "Score KPIs Qualité")[["Poste travail princ.", "Score KPIs Qualité"]].reset_index(drop=True)
                t5.columns = ["Poste", "Score (%)"]
                t5["Score (%)"] = t5["Score (%)"].round(1)
                st.dataframe(t5, use_container_width=True, hide_index=True, height=230)
                st.markdown('</div>', unsafe_allow_html=True)

            with col_top3:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">⭐ Performance Globale</div>', unsafe_allow_html=True)
                t5 = df_class.nsmallest(5, "Total performance ")[["Poste travail princ.", "Total performance "]].reset_index(drop=True)
                t5.columns = ["Poste", "Score (%)"]
                t5["Score (%)"] = t5["Score (%)"].round(1)
                st.dataframe(t5, use_container_width=True, hide_index=True, height=230)
                st.markdown('</div>', unsafe_allow_html=True)

            # ==================================================
            # GRAPHIQUES
            # ==================================================
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">📈 Analyse par Catégorie</p>', unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)

            with c1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">🔧 Par Métier</div>', unsafe_allow_html=True)
                df_m = df_class.groupby("Métier")["Total performance "].mean().reset_index()
                df_m.columns = ["Métier", "Performance"]
                df_m["Performance"] = df_m["Performance"].round(1)

                bars = alt.Chart(df_m).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, height=20).encode(
                    x=alt.X('Performance:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title=''),
                    y=alt.Y('Métier:O', sort='-x', title='', axis=alt.Axis(labelFontSize=12, labelFontWeight='600')),
                    color=alt.Color('Performance:Q', scale=alt.Scale(domain=[0, 50, 80, 100], range=['#e53e3e', '#ed8936', '#ecc94b', '#38a169']), legend=None),
                    tooltip=['Métier', 'Performance']
                )
                txt = alt.Chart(df_m).mark_text(align='left', baseline='middle', dx=6, fontSize=13, fontWeight='700').encode(
                    y=alt.Y('Métier:O', sort='-x'), text=alt.Text('Performance:Q', format='.1f')
                )
                st.altair_chart((bars + txt).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent').properties(height=max(120, 40 * len(df_m) + 20)), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with c2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">🏭 Par Atelier</div>', unsafe_allow_html=True)
                df_a = df_class.groupby("Atelier")["Total performance "].mean().reset_index()
                df_a.columns = ["Atelier", "Performance"]
                df_a["Performance"] = df_a["Performance"].round(1)

                bars = alt.Chart(df_a).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, height=20).encode(
                    x=alt.X('Performance:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title=''),
                    y=alt.Y('Atelier:O', sort='-x', title='', axis=alt.Axis(labelFontSize=12, labelFontWeight='600')),
                    color=alt.Color('Performance:Q', scale=alt.Scale(domain=[0, 50, 80, 100], range=['#e53e3e', '#ed8936', '#ecc94b', '#38a169']), legend=None),
                    tooltip=['Atelier', 'Performance']
                )
                txt = alt.Chart(df_a).mark_text(align='left', baseline='middle', dx=6, fontSize=13, fontWeight='700').encode(
                    y=alt.Y('Atelier:O', sort='-x'), text=alt.Text('Performance:Q', format='.1f')
                )
                st.altair_chart((bars + txt).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent').properties(height=max(120, 40 * len(df_a) + 20)), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with c3:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">🏢 Par Division</div>', unsafe_allow_html=True)
                df_d = df_class.groupby("Division")["Total performance "].mean().reset_index()
                df_d.columns = ["Division", "Performance"]
                df_d["Performance"] = df_d["Performance"].round(1)

                bars = alt.Chart(df_d).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, height=20).encode(
                    x=alt.X('Performance:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title=''),
                    y=alt.Y('Division:O', sort='-x', title='', axis=alt.Axis(labelFontSize=12, labelFontWeight='600')),
                    color=alt.Color('Performance:Q', scale=alt.Scale(domain=[0, 50, 80, 100], range=['#e53e3e', '#ed8936', '#ecc94b', '#38a169']), legend=None),
                    tooltip=['Division', 'Performance']
                )
                txt = alt.Chart(df_d).mark_text(align='left', baseline='middle', dx=6, fontSize=13, fontWeight='700').encode(
                    y=alt.Y('Division:O', sort='-x'), text=alt.Text('Performance:Q', format='.1f')
                )
                st.altair_chart((bars + txt).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent').properties(height=max(120, 40 * len(df_d) + 20)), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            # Graphiques complémentaires
            st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)
            col_dist, col_qq = st.columns(2)

            with col_dist:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">📊 Distribution des Scores</div>', unsafe_allow_html=True)
                df_dist = df_class[["Poste travail princ.", "Total performance "]].copy()
                df_dist.columns = ["Poste", "Score"]
                bins = [0, 50, 60, 70, 80, 90, 100.1]
                labels = ["0-50%", "50-60%", "60-70%", "70-80%", "80-90%", "90-100%"]
                df_dist["Tranche"] = pd.cut(df_dist["Score"], bins=bins, labels=labels, right=False)
                df_hist = df_dist.groupby("Tranche", observed=False).size().reset_index(name="Nb")
                color_map = {"0-50%": "#e53e3e", "50-60%": "#e53e3e", "60-70%": "#ed8936", "70-80%": "#d69e2e", "80-90%": "#ecc94b", "90-100%": "#38a169"}

                bars_h = alt.Chart(df_hist).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                    x=alt.X('Tranche:O', sort=labels, title='', axis=alt.Axis(labelAngle=0, labelFontSize=11)),
                    y=alt.Y('Nb:Q', title='Nb postes', axis=alt.Axis(grid=False, ticks=False)),
                    color=alt.Color('Tranche:O', scale=alt.Scale(domain=labels, range=[color_map[l] for l in labels]), legend=None),
                    tooltip=['Tranche', 'Nb']
                )
                txt_h = alt.Chart(df_hist).mark_text(align='center', baseline='bottom', dy=-4, fontSize=13, fontWeight='800').encode(
                    x=alt.X('Tranche:O', sort=labels), y='Nb:Q', text=alt.Text('Nb:Q')
                )
                st.altair_chart((bars_h + txt_h).configure_axisX(grid=False, ticks=False, domain=False).configure_view(stroke='transparent').properties(height=200), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with col_qq:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">⚖️ Quantité vs Qualité</div>', unsafe_allow_html=True)
                df_qq = pd.DataFrame({
                    "Type": ["Quantité", "Qualité"],
                    "Score": [df_class["Score KPIs Quantité"].mean(), df_class["Score KPIs Qualité"].mean()]
                })
                df_qq["Score"] = df_qq["Score"].round(1)
                df_qq["Couleur"] = ["#4299e1", "#48bb78"]

                bars_q = alt.Chart(df_qq).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X('Type:O', title='', axis=alt.Axis(labelFontSize=13, labelFontWeight='600')),
                    y=alt.Y('Score:Q', scale=alt.Scale(domain=[0, 100]), title='Score moyen (%)', axis=alt.Axis(grid=False, ticks=False)),
                    color=alt.Color('Couleur:N', scale=None, legend=None),
                    tooltip=['Type', 'Score']
                )
                txt_q = alt.Chart(df_qq).mark_text(align='center', baseline='bottom', dy=-5, fontSize=16, fontWeight='800').encode(
                    x='Type:O', y='Score:Q', text=alt.Text('Score:Q', format='.1f')
                )
                st.altair_chart((bars_q + txt_q).configure_axisX(grid=False, ticks=False, domain=False).configure_view(stroke='transparent').properties(height=200, width=280), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            # ==================================================
            # EXPORT
            # ==================================================
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">💾 Export des Plans d\'Action</p>', unsafe_allow_html=True)

            postes_avec_anomalies = anomalies_dashboard[anomalies_dashboard.index != "Total général"].index.tolist() if not anomalies_dashboard.empty else []

            if postes_avec_anomalies:
                col_exp1, col_exp2 = st.columns([1, 1])
                with col_exp1:
                    selected_poste_export = st.selectbox(
                        "Poste de travail :",
                        options=["📌 Tous les postes"] + postes_avec_anomalies,
                        key="sel_export"
                    )
                with col_exp2:
                    st.markdown("<div style='height:32px;'></div>", unsafe_allow_html=True)
                    generate_btn = st.button("📥 Générer le fichier Excel", type="primary", key="btn_export", use_container_width=True)

                if generate_btn:
                    with st.spinner("Génération en cours..."):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            postes_a_traiter = postes_avec_anomalies if selected_poste_export == "📌 Tous les postes" else [selected_poste_export]

                            for poste_export in postes_a_traiter:
                                kpis_en_defaut = df_anomalies_ot[df_anomalies_ot["Poste travail princ."] == poste_export]["KPI"].unique().tolist()
                                if not df_anomalies_avis.empty and "appel avis approuvé" in df_anomalies_avis[df_anomalies_avis["Poste travail princ."] == poste_export]["KPI"].values:
                                    kpis_en_defaut.append("appel avis approuvé")

                                for kpi in kpis_en_defaut:
                                    sheet_data = pd.DataFrame()

                                    if kpi != "appel avis approuvé":
                                        df_poste_filtered = df_processed[df_processed["Poste travail princ."] == poste_export].copy()
                                        conditions_map = {
                                            "TAUX_REALISATION_CORRECTIF/PT": (df_poste_filtered["Nº appel pl.entret."].fillna(0) == 0) & (~df_poste_filtered["Statut OT"].isin(["CLOT", "TCLO"])),
                                            "OT préparation <1 mois": (df_poste_filtered["Statut OT"] == "CRÉÉ") & (df_poste_filtered["Age préparation"] != "<1 mois"),
                                            "OT préparation >3 mois": (df_poste_filtered["Statut OT"] == "CRÉÉ") & (df_poste_filtered["Age préparation"] == ">3 mois"),
                                            "OT planification <1 mois": (df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 0) & (df_poste_filtered["Age planification"] != "<1 mois"),
                                            "OT planification >3 mois": (df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 0) & (df_poste_filtered["Age planification"] == ">3 mois"),
                                            "OT exécution <1 mois": (df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 1) & (df_poste_filtered["Age exécution"] != "<1 mois"),
                                            "OT exécution >3 mois": (df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 1) & (df_poste_filtered["Age exécution"] == ">3 mois"),
                                            "OT LANC ESTIME": (df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["OT LANC ESTIME"] == "NON"),
                                            "Backlog préparation caractérisé": (df_poste_filtered["Statut OT"] == "CRÉÉ") & (df_poste_filtered["Backlog préparation"] == "NON CARACTERISE"),
                                            "Backlog planification caractérisé": (df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Backlog planification"] == "NON CARACTERISE"),
                                            "OT CONFIME": df_poste_filtered["OT CONFIME"] == "NON",
                                            "OT_COR_EGAL": df_poste_filtered["OT_COR_EGAL"] == "NON",
                                        }
                                        condition = conditions_map.get(kpi, pd.Series(False, index=df_poste_filtered.index))
                                        subset_ot = df_poste_filtered[condition]

                                        if not subset_ot.empty:
                                            old_cols = ["Ordre", "Désignation", "Emplacement technique", "Poste travail princ.", "Statut système", "Statut utilisateur", "Date de début planifiée", "Type d'ordre", "Backlog préparation", "Backlog planification"]
                                            new_cols = ["Ordre de travail", "Désignation", "Poste technique", "Poste de travail principal", "Statut système", "Statut utilisateur", "Date de début planifiée", "Type d'ordre", "Caractérisation backlog Préparation", "Caractérisation backlog Planification"]
                                            subset_ot = rename_safe(subset_ot, old_cols, new_cols)
                                            subset_ot["KPI impacté"] = kpi
                                            subset_ot["Action recommandée"] = f"Corriger l'indicateur {kpi}."
                                            sheet_data = pd.concat([sheet_data, subset_ot])

                                    if kpi == "appel avis approuvé":
                                        subset_avis = results['avis_df_filtered'][results['avis_df_filtered']["Poste travail princ."] == poste_export].copy()
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
                        nom_fichier = "Plan_Action_Tous_Postes.xlsx" if selected_poste_export == "📌 Tous les postes" else f"Plan_Action_{selected_poste_export.replace(' ', '_')}.xlsx"
                        st.download_button(
                            label="✅ Télécharger le fichier Excel",
                            data=output.getvalue(),
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            else:
                st.markdown("""
                <div class="empty-state" style="padding:30px;">
                    <div class="icon">🎉</div>
                    <h3>Aucun export nécessaire</h3>
                    <p>Toutes les anomalies ont été résolues.</p>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("""
            <div style="text-align:center; padding:24px 0 8px 0; color:#a0aec0; font-size:11px;">
                KPI Dashboard MC & FEED • Maintenance Conditionnelle • {date}
            </div>
            """.format(date=date_fichier), unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    try:
        locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_ALL, 'fr_FR')
        except:
            pass
    main()
