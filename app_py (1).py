# -*- coding: utf-8 -*-
"""app.py - KPI Dashboard MC et FEED (version améliorée)"""

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
# FORCAGE DE LA LANGUE FRANCAISE POUR LE CALENDRIER
# ==================================================
def main():
    try:
        locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_ALL, 'fr_FR')
        except:
            pass

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

    # ==================================================
    # ECRAN HSE
    # ==================================================
    if "hse_affiche" not in st.session_state:
        st.session_state.hse_affiche = False

    if not st.session_state.hse_affiche:
        consigne = random.choice(consignes)
        st.markdown("""
        <style>
            .hse-container {
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                min-height: 80vh;
                font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            }
            .hse-icon {
                width: 100px;
                height: 100px;
                border-radius: 50%;
                background: linear-gradient(135deg, #f59e0b, #d97706);
                display: flex;
                align-items: center;
                justify-content: center;
                margin-bottom: 30px;
                box-shadow: 0 20px 40px rgba(245, 158, 11, 0.3);
                animation: pulse-glow 2s ease-in-out infinite;
            }
            @keyframes pulse-glow {
                0%, 100% { box-shadow: 0 20px 40px rgba(245, 158, 11, 0.3); }
                50% { box-shadow: 0 20px 60px rgba(245, 158, 11, 0.5); }
            }
            .hse-icon svg { width: 50px; height: 50px; fill: white; }
            .hse-title {
                font-size: 48px;
                font-weight: 800;
                color: #0f172a;
                margin-bottom: 8px;
                letter-spacing: -1px;
            }
            .hse-subtitle {
                font-size: 18px;
                color: #64748b;
                font-weight: 500;
                letter-spacing: 3px;
                margin-bottom: 40px;
            }
            .hse-consigne-box {
                background: linear-gradient(135deg, #fffbeb, #fef3c7);
                border-left: 6px solid #f59e0b;
                padding: 35px 45px;
                border-radius: 16px;
                font-size: 28px;
                font-weight: 700;
                text-align: center;
                color: #92400e;
                max-width: 800px;
                box-shadow: 0 10px 30px rgba(245, 158, 11, 0.15);
                animation: slide-up 0.8s ease-out;
            }
            @keyframes slide-up {
                from { opacity: 0; transform: translateY(30px); }
                to { opacity: 1; transform: translateY(0); }
            }
            .hse-motto {
                font-size: 38px;
                font-weight: 900;
                color: #059669;
                margin-top: 45px;
                letter-spacing: -0.5px;
            }
            .hse-progress {
                margin-top: 30px;
                display: flex;
                align-items: center;
                gap: 8px;
                color: #94a3b8;
                font-size: 14px;
            }
            .hse-progress-dot {
                width: 8px;
                height: 8px;
                border-radius: 50%;
                background: #f59e0b;
                animation: blink 1s ease-in-out infinite;
            }
            @keyframes blink {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.3; }
            }
        </style>
        <div class="hse-container">
            <div class="hse-icon">
                <svg viewBox="0 0 24 24"><path d="M12 2L3 7v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V7l-9-5zm0 10.99h7c-.53 4.12-3.28 7.79-7 8.94V12H5V6.3l7-3.89v10.58z"/></svg>
            </div>
            <div class="hse-title">HSE - CONSIGNE DE SÉCURITÉ</div>
            <div class="hse-subtitle">SÉCURITÉ &bull; SANTÉ &bull; ENVIRONNEMENT</div>
            <div class="hse-consigne-box">⚠️ {consigne}</div>
            <div class="hse-motto">Aucun travail n'est plus urgent que la sécurité</div>
            <div class="hse-progress">
                <div class="hse-progress-dot"></div>
                Chargement en cours...
                <div class="hse-progress-dot"></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        time.sleep(6)
        st.session_state.hse_affiche = True
        st.rerun()
        st.stop()

    # ==================================================
    # CSS GLOBAL - PRÉSENTATION WEB
    # ==================================================
    st.markdown("""
    <style>
        /* --- Reset & Base --- */
        .block-container {
            max-width: 1600px !important;
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
        h1, h2, h3 { font-family: 'Segoe UI', system-ui, sans-serif; }

        /* --- Header Dashboard --- */
        .dash-header {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            padding: 28px 32px;
            border-radius: 16px;
            margin-bottom: 24px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.15);
        }
        .dash-header h1 {
            color: white !important;
            font-size: 32px !important;
            font-weight: 800 !important;
            margin: 0 !important;
            letter-spacing: -0.5px;
        }
        .dash-header p {
            color: #94a3b8 !important;
            font-size: 14px !important;
            margin: 6px 0 0 0 !important;
        }

        /* --- Section Card --- */
        .section-card {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            overflow: hidden;
            margin-bottom: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }
        .section-header {
            padding: 18px 24px;
            font-size: 18px;
            font-weight: 700;
            color: white;
            letter-spacing: -0.3px;
        }
        .section-header-kpi {
            background: linear-gradient(135deg, #0f172a, #334155);
        }
        .section-header-anomalie {
            background: linear-gradient(135deg, #b91c1c, #dc2626);
        }
        .section-header-action {
            background: linear-gradient(135deg, #b45309, #d97706);
        }
        .section-header-classement {
            background: linear-gradient(135deg, #6d28d9, #7c3aed);
        }
        .section-header-chart {
            background: linear-gradient(135deg, #0f766e, #14b8a6);
        }
        .section-body {
            padding: 20px 24px;
        }

        /* --- Toggle Buttons --- */
        .toggle-container {
            display: flex;
            gap: 8px;
            margin-bottom: 20px;
        }
        .toggle-btn {
            padding: 12px 24px;
            border-radius: 12px;
            font-size: 14px;
            font-weight: 600;
            border: 2px solid transparent;
            cursor: pointer;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .toggle-btn-active-kpi {
            background: #0f172a;
            color: white;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.3);
        }
        .toggle-btn-inactive {
            background: white;
            color: #475569;
            border-color: #e2e8f0;
        }
        .toggle-btn-inactive:hover {
            background: #f8fafc;
            border-color: #cbd5e1;
        }
        .toggle-btn-active-anomalie {
            background: #dc2626;
            color: white;
            box-shadow: 0 4px 12px rgba(220, 38, 38, 0.3);
        }

        /* --- Tables --- */
        .dataframe {
            font-size: 12px !important;
        }
        .dataframe th {
            background: #f8fafc !important;
            color: #1e293b !important;
            font-weight: 600 !important;
            font-size: 11px !important;
            padding: 10px 12px !important;
            border-bottom: 2px solid #e2e8f0 !important;
            white-space: nowrap;
        }
        .dataframe td {
            padding: 8px 12px !important;
            font-size: 11px !important;
            border-bottom: 1px solid #f1f5f9 !important;
        }
        .dataframe tr:hover td {
            background: #f8fafc !important;
        }

        /* --- Action Cards --- */
        .action-card {
            border-left: 5px solid;
            border-radius: 12px;
            padding: 20px 24px;
            margin-bottom: 16px;
            transition: box-shadow 0.2s;
        }
        .action-card:hover {
            box-shadow: 0 4px 16px rgba(0,0,0,0.08);
        }
        .action-card-critical { border-color: #ef4444; background: #fef2f2; }
        .action-card-warning { border-color: #f59e0b; background: #fffbeb; }
        .action-card-info { border-color: #64748b; background: #f8fafc; }
        .action-card-title {
            font-size: 14px;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .action-card-badge {
            display: inline-flex;
            align-items: center;
            padding: 2px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
        }
        .badge-critical { background: #fecaca; color: #991b1b; }
        .badge-warning { background: #fde68a; color: #92400e; }
        .badge-info { background: #e2e8f0; color: #334155; }
        .action-card-action {
            font-size: 13px;
            color: #64748b;
            margin-bottom: 12px;
        }
        .action-card-postes {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            padding-top: 12px;
            border-top: 1px solid rgba(0,0,0,0.06);
        }
        .poste-chip {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 500;
            color: #475569;
        }
        .poste-chip-count {
            background: #e2e8f0;
            color: #334155;
            font-weight: 700;
            padding: 1px 6px;
            border-radius: 6px;
            font-size: 11px;
        }
        .action-summary {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-top: 16px;
            border-top: 1px solid #e2e8f0;
            margin-top: 8px;
            font-size: 13px;
            color: #64748b;
        }
        .action-big-number {
            font-size: 36px;
            font-weight: 800;
            color: #1e293b;
            line-height: 1;
        }
        .action-big-label {
            font-size: 11px;
            color: #94a3b8;
            font-weight: 500;
        }
        .action-header-row {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 16px;
        }

        /* --- Chart container --- */
        .chart-box {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }
        .chart-title {
            font-size: 14px;
            font-weight: 700;
            padding: 14px 20px 0 20px;
            color: #1e293b;
        }
        .chart-subtitle {
            font-size: 11px;
            color: #94a3b8;
            padding: 0 20px 4px 20px;
        }

        /* --- Scrollable table wrapper --- */
        .table-scroll {
            overflow-x: auto;
            border-radius: 12px;
        }

        /* --- Info message --- */
        .stAlert {
            border-radius: 12px !important;
        }
    </style>
    """, unsafe_allow_html=True)

    # ==================================================
    # TON DASHBOARD COMMENCE ICI
    # ==================================================
    st.markdown("""
    <div class="dash-header">
        <h1>KPI Dashboard MC et FEED</h1>
        <p>Tableau de bord des indicateurs de performance maintenance</p>
    </div>
    """, unsafe_allow_html=True)

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
        if row.name == 'CIBLE': return ['padding: 8px; background-color: #dbeafe; color: #1e40af; font-weight: bold; border-bottom: 2px solid #93c5fd;'] * len(row)
        for col in row.index:
            current_style = 'padding: 8px; white-space: nowrap;'
            try: value = float(row[col])
            except (ValueError, TypeError): styles.append(current_style); continue
            if col in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
                if value >= 80: current_style += 'background-color:#d1fae5; color:#065f46; font-weight:600;'
                elif value >= 75: current_style += 'background-color:#fef3c7; color:#92400e; font-weight:600;'
                else: current_style += 'background-color:#fee2e2; color:#991b1b; font-weight:600;'
            elif col in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
                if value <= 15: current_style += 'background-color:#d1fae5; color:#065f46; font-weight:600;'
                else: current_style += 'background-color:#fee2e2; color:#991b1b; font-weight:600;'
            elif col in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
                if value <= 5: current_style += 'background-color:#d1fae5; color:#065f46; font-weight:600;'
                else: current_style += 'background-color:#fee2e2; color:#991b1b; font-weight:600;'
            elif col == "TAUX_REALISATION_CORRECTIF/PT":
                if value >= 85: current_style += 'background-color:#d1fae5; color:#065f46; font-weight:600;'
                elif value >= 80: current_style += 'background-color:#fef3c7; color:#92400e; font-weight:600;'
                else: current_style += 'background-color:#fee2e2; color:#991b1b; font-weight:600;'
            elif col == "appel avis approuvé":
                if value >= 95: current_style += 'background-color:#d1fae5; color:#065f46; font-weight:600;'
                elif value >= 90: current_style += 'background-color:#fef3c7; color:#92400e; font-weight:600;'
                else: current_style += 'background-color:#fee2e2; color:#991b1b; font-weight:600;'
            elif col in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
                if value >= 100: current_style += 'background-color:#d1fae5; color:#065f46; font-weight:600;'
                elif value >= 95: current_style += 'background-color:#fef3c7; color:#92400e; font-weight:600;'
                else: current_style += 'background-color:#fee2e2; color:#991b1b; font-weight:600;'
            styles.append(current_style)
        return styles

    def highlight_anomalies(row):
        styles = []
        if row.name == 'Total général':
            return ['padding: 8px; background-color: #fecaca; color: #7f1d1d; font-weight: bold; border-top: 2px solid #f87171;'] * len(row)
        for col in row.index:
            current_style = 'padding: 8px; white-space: nowrap;'
            try:
                value = int(row[col])
            except (ValueError, TypeError):
                styles.append(current_style)
                continue
            if value > 0:
                current_style += 'background-color:#fee2e2; color:#991b1b; font-weight:700;'
            else:
                current_style += 'color:#e2e8f0;'
            styles.append(current_style)
        return styles

    def highlight_classification_table_kpis(row):
        styles = [''] * len(row)
        for i, col in enumerate(row.index):
            if col in ["Score KPIs Quantité", "Score KPIs Qualité", "Total performance "]:
                try:
                    s = str(row[col]); value = float(s.replace(' %', '').strip()) if ' %' in s else float(s)
                    if value >= 90: styles[i] = 'background-color: #d1fae5; color: #065f46; font-weight: 700;'
                    elif value >= 80: styles[i] = 'background-color: #fef3c7; color: #92400e; font-weight: 700;'
                    else: styles[i] = 'background-color: #fee2e2; color: #991b1b; font-weight: 700;'
                except (ValueError, TypeError): pass
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

    def get_action_severity(kpi_name):
        if "TAUX_REALISATION" in kpi_name: return "critical"
        if any(x in kpi_name for x in ["préparation", "planification", "exécution"]): return "warning"
        return "info"

    use_new_files = st.toggle("Charger de nouveaux fichiers OT et AVIS", value=False)

    ot_file = None
    avis_file = None

    if use_new_files:
        col1, col2 = st.columns(2)
        with col1: ot_file = st.file_uploader("📂 Charger le fichier OT", type=["xlsx"])
        with col2: avis_file = st.file_uploader("📂 Charger le fichier AVIS", type=["xlsx"])
    else:
        date_fichier = datetime.now().strftime("%d/%m/%Y")
        if os.path.exists("ot.xlsx"):
            timestamp = os.path.getmtime("ot.xlsx")
            date_fichier = datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y")
        st.info(f"Data Version : V1.0 – Mise à jour du : {date_fichier}")

    if not use_new_files or (ot_file is not None and avis_file is not None):
        try:
            if use_new_files:
                df_ot_raw = pd.read_excel(ot_file)
                avis_df_raw = pd.read_excel(avis_file)
                date_fichier = datetime.now().strftime("%d/%m/%Y")
            else:
                df_ot_raw = pd.read_excel("ot.xlsx")
                avis_df_raw = pd.read_excel("avis.xlsx")

            for col in ["Créé le", "Date de début planifiée", "Date de clôture", "Début réel", "Fin réelle"]:
                if col in df_ot_raw.columns: df_ot_raw[col] = pd.to_datetime(df_ot_raw[col], errors="coerce")
            for col in ["Créé le", "Début souhaité", "Date de la clôture"]:
                if col in avis_df_raw.columns: avis_df_raw[col] = pd.to_datetime(avis_df_raw[col], errors="coerce")

            # ============================================================
            # EXCLUSION TOTALE DES CRESSEURS SUR TOUT
            # ============================================================
            df_ot_raw = df_ot_raw[~df_ot_raw["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)]
            avis_df_raw = avis_df_raw[~avis_df_raw["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)]

            all_postes_master_list = sorted(df_ot_raw[df_ot_raw["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())

            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                selected_postes = st.multiselect("Poste de travail", ["All"] + all_postes_master_list, ["All"])
                if "All" in selected_postes or len(selected_postes) == 0: selected_postes = all_postes_master_list
            with col_f2:
                ateliers_options = ["All", "Sulfurique (PS)", "Phosphorique (PP)", "Engrais (TSP/REX)", "Feed (MCP/DCP)"]
                selected_ateliers = st.multiselect("Atelier", ateliers_options, ["All"])
                if "All" in selected_ateliers or len(selected_ateliers) == 0: selected_ateliers = ["All"]
            with col_f3:
                divisions_options = ["All", "SF1", "SF2"]
                selected_divisions = st.multiselect("Division", divisions_options, ["All"])
                if "All" in selected_divisions or len(selected_divisions) == 0: selected_divisions = ["All"]

            default_start = datetime(2025, 1, 1).date()
            default_end = datetime.today().date()

            date_range = st.date_input(
                "📅 Filtre Date de début planifiée (Du - Au)",
                value=(default_start, default_end),
                format="DD/MM/YYYY"
            )

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

            # Double exclusion cresseur (sécurité)
            df = df_ot_raw[(df_ot_raw["Poste travail princ."].isin(valid_postes)) & (df_ot_raw["Date de début planifiée"].between(start_date, end_date)) & (~df_ot_raw["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False))].copy()
            avis_df = avis_df_raw[(avis_df_raw["Poste travail princ."].isin(valid_postes)) & (~avis_df_raw["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False))].copy()

            df = df[df["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)].drop_duplicates()
            avis_df = avis_df[(avis_df["Ordre"].isna()) | (avis_df["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates()

            if "Statut système" in df.columns: df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]
            st.success(f"OT après filtre : {len(df)} lignes | AVIS après filtre : {len(avis_df)} lignes")

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

            # ============================================================
            # ANOMALIES DETECTION
            # ============================================================
            anomalies_ot_records = []
            anomalies_avis_records = []

            for poste in valid_postes:
                if poste not in df_processed["Poste travail princ."].values: continue
                df_poste = df_processed[df_processed["Poste travail princ."] == poste]
                avis_poste = results['avis_df_filtered'][results['avis_df_filtered']["Poste travail princ."] == poste]

                val_real = calculated_kpis_df.loc[poste, "TAUX_REALISATION_CORRECTIF/PT"] if poste in calculated_kpis_df.index else 100
                if pd.notna(val_real) and val_real < cible.loc['CIBLE', "TAUX_REALISATION_CORRECTIF/PT"]:
                    count_anom = len(df_poste[(df_poste["Nº appel pl.entret."].fillna(0) == 0) & (~df_poste["Statut OT"].isin(["CLOT", "TCLO"]))])
                    if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "TAUX_REALISATION_CORRECTIF/PT", "Nb OT impactés": count_anom, "Action Suggérée": "Améliorer le taux de réalisation des OT."})

                val_prep_inf1 = calculated_kpis_df.loc[poste, "OT préparation <1 mois"] if poste in calculated_kpis_df.index else 100
                if pd.notna(val_prep_inf1) and val_prep_inf1 < cible.loc['CIBLE', "OT préparation <1 mois"]:
                    count_anom = len(df_poste[(df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Age préparation"] != "<1 mois")])
                    if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "OT préparation <1 mois", "Nb OT impactés": count_anom, "Action Suggérée": "Réduire l'âge de préparation des OT."})

                val_prep_sup3 = calculated_kpis_df.loc[poste, "OT préparation >3 mois"] if poste in calculated_kpis_df.index else 0
                if pd.notna(val_prep_sup3) and val_prep_sup3 > cible.loc['CIBLE', "OT préparation >3 mois"]:
                    count_anom = len(df_poste[(df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Age préparation"] == ">3 mois")])
                    if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "OT préparation >3 mois", "Nb OT impactés": count_anom, "Action Suggérée": "Traiter les OT ayant un âge de préparation > 3 mois."})

                val_planif_inf1 = calculated_kpis_df.loc[poste, "OT planification <1 mois"] if poste in calculated_kpis_df.index else 100
                if pd.notna(val_planif_inf1) and val_planif_inf1 < cible.loc['CIBLE', "OT planification <1 mois"]:
                    count_anom = len(df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 0) & (df_poste["Age planification"] != "<1 mois")])
                    if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "OT planification <1 mois", "Nb OT impactés": count_anom, "Action Suggérée": "Réduire l'âge de planification des OT."})

                val_planif_sup3 = calculated_kpis_df.loc[poste, "OT planification >3 mois"] if poste in calculated_kpis_df.index else 0
                if pd.notna(val_planif_sup3) and val_planif_sup3 > cible.loc['CIBLE', "OT planification >3 mois"]:
                    count_anom = len(df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 0) & (df_poste["Age planification"] == ">3 mois")])
                    if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "OT planification >3 mois", "Nb OT impactés": count_anom, "Action Suggérée": "Traiter les OT ayant un âge de planification > 3 mois."})

                val_exec_inf1 = calculated_kpis_df.loc[poste, "OT exécution <1 mois"] if poste in calculated_kpis_df.index else 100
                if pd.notna(val_exec_inf1) and val_exec_inf1 < cible.loc['CIBLE', "OT exécution <1 mois"]:
                    count_anom = len(df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 1) & (df_poste["Age exécution"] != "<1 mois")])
                    if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "OT exécution <1 mois", "Nb OT impactés": count_anom, "Action Suggérée": "Réduire l'âge d'exécution des OT."})

                val_exec_sup3 = calculated_kpis_df.loc[poste, "OT exécution >3 mois"] if poste in calculated_kpis_df.index else 0
                if pd.notna(val_exec_sup3) and val_exec_sup3 > cible.loc['CIBLE', "OT exécution >3 mois"]:
                    count_anom = len(df_poste[(df_poste["Statut OT"] == "LANC") & (df_poste["Contient SOPL"] == 1) & (df_poste["Age exécution"] == ">3 mois")])
                    if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": "OT exécution >3 mois", "Nb OT impactés": count_anom, "Action Suggérée": "Traiter les OT ayant un âge d'exécution > 3 mois."})

                for kpi_name, condition, action in [
                    ("OT LANC ESTIME", (df_poste["Statut OT"] == "LANC") & (df_poste["OT LANC ESTIME"] == "NON"), "Estimer les coûts des OT lancés."),
                    ("Backlog préparation caractérisé", (df_poste["Statut OT"] == "CRÉÉ") & (df_poste["Backlog préparation"] == "NON CARACTERISE"), "Caractériser le backlog de préparation."),
                    ("Backlog planification caractérisé", (df_poste["Statut OT"] == "LANC") & (df_poste["Backlog planification"] == "NON CARACTERISE"), "Caractériser le backlog de planification."),
                    ("OT CONFIME", df_poste["OT CONFIME"] == "NON", "Confirmer les OT terminés."),
                    ("OT_COR_EGAL", df_poste["OT_COR_EGAL"] == "NON", "Rapprocher les coûts réels et budgétés.")
                ]:
                    val_kpi = calculated_kpis_df.loc[poste, kpi_name] if poste in calculated_kpis_df.index else 100
                    if pd.notna(val_kpi) and val_kpi < cible.loc['CIBLE', kpi_name]:
                        count_anom = len(df_poste[condition])
                        if count_anom > 0: anomalies_ot_records.append({"Poste travail princ.": poste, "KPI": kpi_name, "Nb OT impactés": count_anom, "Action Suggérée": action})

                val_avis = calculated_kpis_df.loc[poste, "appel avis approuvé"] if poste in calculated_kpis_df.index else 100
                if pd.notna(val_avis) and val_avis < cible.loc['CIBLE', "appel avis approuvé"]:
                    count_avis_anom = len(avis_poste)
                    if count_avis_anom > 0: anomalies_avis_records.append({"Poste travail princ.": poste, "KPI": "appel avis approuvé", "Nb Avis impactés": count_avis_anom, "Action Suggérée": "Créer un OT pour les avis sans ordre."})

            df_anomalies_ot = pd.DataFrame(anomalies_ot_records)
            df_anomalies_avis = pd.DataFrame(anomalies_avis_records)

            if not df_anomalies_ot.empty:
                pivot_ot = df_anomalies_ot.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb OT impactés", aggfunc="sum", fill_value=0)
            else: pivot_ot = pd.DataFrame(columns=["Poste travail princ."])

            if not df_anomalies_avis.empty:
                pivot_avis = df_anomalies_avis.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb Avis impactés", aggfunc="sum", fill_value=0)
                if "appel avis approuvé" in pivot_avis.columns: pivot_avis = pivot_avis.rename(columns={"appel avis approuvé": "Nb Avis sans ordre"})
            else: pivot_avis = pd.DataFrame(columns=["Poste travail princ."])

            anomalies_dashboard = pivot_ot.join(pivot_avis, how='outer').fillna(0).astype(int)
            if not anomalies_dashboard.empty:
                anomalies_dashboard["Total éléments impactés"] = anomalies_dashboard.sum(axis=1)
                total_row = pd.DataFrame(anomalies_dashboard.sum()).T; total_row.index = ["Total général"]
                anomalies_dashboard = pd.concat([anomalies_dashboard, total_row])
            else: anomalies_dashboard = pd.DataFrame()

            # ============================================================
            # CLASSIFICATION
            # ============================================================
            qty_kpis = ["TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois", "OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois", "OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]
            qual_kpis = ["appel avis approuvé", "OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]

            class_results = []
            for poste in calculated_kpis_df.index:
                row = calculated_kpis_df.loc[poste]
                s_qty = sum(get_kpi_score(k, row[k], cible.loc['CIBLE', k]) for k in qty_kpis if k in row.index)
                s_qual = sum(get_kpi_score(k, row[k], cible.loc['CIBLE', k]) for k in qual_kpis if k in row.index)
                p_qty = (s_qty / len(qty_kpis) * 100) if qty_kpis else 0
                p_qual = (s_qual / len(qual_kpis) * 100) if qual_kpis else 0
                class_results.append({
                    "Poste travail princ.": poste,
                    "Score KPIs Quantité": p_qty,
                    "Score KPIs Qualité": p_qual,
                    "Total performance ": (p_qty+p_qual)/2
                })
            df_class = pd.DataFrame(class_results)

            # ============================================================
            # 1. BOUTON TOGGLE : KPIs OU ANOMALIES
            # ============================================================
            st.markdown("""
            <div class="toggle-container">
                <div class="toggle-btn toggle-btn-active-kpi" id="btn-kpi" onclick="
                    document.getElementById('btn-kpi').className = 'toggle-btn toggle-btn-active-kpi';
                    document.getElementById('btn-anomalie').className = 'toggle-btn toggle-btn-inactive';
                    document.getElementById('section-kpi').style.display = 'block';
                    document.getElementById('section-anomalie').style.display = 'none';
                ">📊 TABLEAU DE BORD DES KPIs</div>
                <div class="toggle-btn toggle-btn-inactive" id="btn-anomalie" onclick="
                    document.getElementById('btn-anomalie').className = 'toggle-btn toggle-btn-active-anomalie';
                    document.getElementById('btn-kpi').className = 'toggle-btn toggle-btn-inactive';
                    document.getElementById('section-anomalie').style.display = 'block';
                    document.getElementById('section-kpi').style.display = 'none';
                ">⚠️ TABLEAU DE BORD DES ANOMALIES</div>
            </div>
            """, unsafe_allow_html=True)

            # --- SECTION KPIs (visible par défaut) ---
            st.markdown('<div id="section-kpi">', unsafe_allow_html=True)
            total_general_kpi = pd.DataFrame(calculated_kpis_df.mean()).T; total_general_kpi.index = ["Total général"]
            final_kpi = pd.concat([cible, calculated_kpis_df, total_general_kpi]).round(2)

            st.markdown("""
            <div class="section-card">
                <div class="section-header section-header-kpi">
                    📊 TABLEAU DE BORD DES KPIs
                    <span style="font-weight:400; font-size:13px; opacity:0.8; margin-left:8px;">Indicateurs de performance par poste de travail</span>
                </div>
                <div class="section-body">
                    <div class="table-scroll">
            """, unsafe_allow_html=True)
            st.table(final_kpi.style.apply(highlight_kpis, axis=1).format("{:.2f}"))
            st.markdown('</div></div></div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # --- SECTION ANOMALIES (masquée par défaut) ---
            st.markdown('<div id="section-anomalie" style="display:none;">', unsafe_allow_html=True)
            st.markdown("""
            <div class="section-card">
                <div class="section-header section-header-anomalie">
                    ⚠️ TABLEAU DE BORD DES ANOMALIES
                    <span style="font-weight:400; font-size:13px; opacity:0.8; margin-left:8px;">Éléments impactés par KPI et par poste</span>
                </div>
                <div class="section-body">
                    <div class="table-scroll">
            """, unsafe_allow_html=True)
            if not anomalies_dashboard.empty:
                st.dataframe(anomalies_dashboard.style.apply(highlight_anomalies, axis=1), use_container_width=True)
            else:
                st.info("✅ Aucune anomalie détectée. Tous les KPIs atteignent leurs cibles.")
            st.markdown('</div></div></div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # ============================================================
            # 2. SYNTHESE DES ACTIONS KPI - REGROUPEE PAR KPI
            # ============================================================
            st.markdown("""
            <div class="section-card">
                <div class="section-header section-header-action">
                    🎯 Synthèse des Actions KPI par Poste de Travail
                    <span style="font-weight:400; font-size:13px; opacity:0.8; margin-left:8px;">KPIs regroupés par nombre d'OT impactés et actions</span>
                </div>
                <div class="section-body">
            """, unsafe_allow_html=True)

            if not df_anomalies_ot.empty:
                # Regroupement par KPI
                grouped = df_anomalies_ot.groupby("KPI").agg(
                    nb_total_ot=("Nb OT impactés", "sum"),
                    action=("Action Suggérée", "first"),
                    postes_list=("Poste travail princ.", lambda x: list(zip(x, df_anomalies_ot.loc[x.index, "Nb OT impactés"])))
                ).sort_values("nb_total_ot", ascending=False).reset_index()

                total_all_ot = grouped["nb_total_ot"].sum()

                for idx, row in grouped.iterrows():
                    severity = get_action_severity(row["KPI"])
                    badge_class = "badge-critical" if severity == "critical" else ("badge-warning" if severity == "warning" else "badge-info")
                    card_class = "action-card-critical" if severity == "critical" else ("action-card-warning" if severity == "warning" else "action-card-info")

                    postes_html = ""
                    for poste, nb in sorted(row["postes_list"], key=lambda x: x[1], reverse=True):
                        postes_html += f'<div class="poste-chip">{poste} <span class="poste-chip-count">{nb}</span></div>'

                    st.markdown(f"""
                    <div class="action-card {card_class}">
                        <div class="action-header-row">
                            <div style="flex:1;">
                                <div class="action-card-title">
                                    {row["KPI"]}
                                    <span class="action-card-badge {badge_class}">{row["nb_total_ot"]} OT</span>
                                </div>
                                <div class="action-card-action">{row["action"]}</div>
                            </div>
                            <div style="text-align:right; flex-shrink:0;">
                                <div class="action-big-number">{row["nb_total_ot"]}</div>
                                <div class="action-big-label">OT impactés</div>
                            </div>
                        </div>
                        <div class="action-card-postes">{postes_html}</div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown(f"""
                <div class="action-summary">
                    <span>⚠️ <strong>{len(grouped)}</strong> KPI(s) en dessous de la cible</span>
                    <span>Total : <strong>{total_all_ot}</strong> OT impactés</span>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("✅ Tous les KPIs atteignent leurs cibles. Aucune action immédiate requise.")

            st.markdown('</div></div>', unsafe_allow_html=True)

            # ============================================================
            # 3. CLASSEMENT DES POSTES + TOP 5
            # ============================================================
            st.markdown("""
            <div class="section-card">
                <div class="section-header section-header-classement">
                    🏆 Classement des Postes de Travail par Qualité des KPIs
                </div>
                <div class="section-body">
            """, unsafe_allow_html=True)

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
            st.table(df_class_display.style.apply(highlight_classification_table_kpis, axis=1))

            st.markdown('</div></div>', unsafe_allow_html=True)

            # TOP 5 CARDS
            st.markdown('<div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 24px;">', unsafe_allow_html=True)

            for col_idx, (title, sort_col, color, icon) in enumerate([
                ("Impactant la Quantité", "Score KPIs Quantité", "#e11d48", "📉"),
                ("Impactant la Qualité", "Score KPIs Qualité", "#d97706", "🔧"),
                ("Impactant la Performance Globale", "Total performance ", "#0f172a", "📊"),
            ]):
                top5 = df_class.nsmallest(5, sort_col)[["Poste travail princ.", sort_col]].reset_index(drop=True)
                cards_html = ""
                for i, (_, r) in enumerate(top5.iterrows()):
                    val = r[sort_col]
                    bg = "#d1fae5; color:#065f46" if val >= 90 else ("#fef3c7; color:#92400e" if val >= 80 else "#fee2e2; color:#991b1b")
                    medal = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"][i]
                    cards_html += f'''
                    <div style="display:flex; align-items:center; justify-content:space-between; padding:10px 12px; background:#f8fafc; border-radius:10px; margin-bottom:8px; border:1px solid #e2e8f0;">
                        <div style="display:flex; align-items:center; gap:8px; min-width:0;">
                            <span style="font-size:14px;">{medal}</span>
                            <span style="font-size:12px; font-weight:600; color:#334155; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">{r["Poste travail princ."]}</span>
                        </div>
                        <span style="font-size:11px; font-weight:700; padding:3px 8px; border-radius:6px; background:{bg}; flex-shrink:0;">{val:.1f}%</span>
                    </div>'''

                st.markdown(f'''
                <div class="chart-box">
                    <div style="background:{color}; padding:14px 20px;">
                        <div style="color:white; font-size:13px; font-weight:700;">{icon} Top 5 Postes</div>
                        <div style="color:rgba(255,255,255,0.7); font-size:11px;">{title}</div>
                    </div>
                    <div style="padding:12px 16px;">{cards_html}</div>
                </div>
                ''', unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            # ============================================================
            # 4. CHARTS AMELIORES
            # ============================================================
            st.markdown("""
            <div class="section-card">
                <div class="section-header section-header-chart">
                    📈 Performance par Groupe
                    <span style="font-weight:400; font-size:13px; opacity:0.8; margin-left:8px;">Moyenne par Métier, Atelier et Division</span>
                </div>
                <div class="section-body">
            """, unsafe_allow_html=True)

            df_class["Métier"] = df_class["Poste travail princ."].apply(get_groupe_metier)
            df_class["Atelier"] = df_class["Poste travail princ."].apply(get_groupe_atelier)
            df_class["Division"] = df_class["Poste travail princ."].apply(get_groupe_division)

            chart_configs = [
                ("Métier", "métier", "#3b82f6", "Électrique, Mécanique, Instrumentation..."),
                ("Atelier", "atelier", "#ef4444", "Sulfurique, Phosphorique, Engrais, Feed"),
                ("Division", "division", "#10b981", "SF1 vs SF2"),
            ]

            cols = st.columns(3)
            for i, (title, group_col, base_color, subtitle) in enumerate(chart_configs):
                with cols[i]:
                    df_group = df_class.groupby(group_col.capitalize())["Total performance "].mean().reset_index()
                    df_group.columns = ["Groupe", "Performance"]
                    df_group = df_group.sort_values("Performance", ascending=True)

                    # Couleur par performance
                    color_scale = alt.Color(
                        "Performance:Q",
                        scale=alt.Scale(
                            domain=[0, 80, 90, 100],
                            range=["#fee2e2", "#fef3c7", "#d1fae5"],
                            type="threshold"
                        ),
                        legend=None
                    )

                    chart = alt.Chart(df_group).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
                        x=alt.X("Performance:Q", scale=alt.Scale(domain=[0, 100]), axis=alt.Axis(grid=False, ticks=False, title=None, labels=False)),
                        y=alt.Y("Groupe:N", sort=None, axis=alt.Axis(labelFontSize=12, labelFontWeight=600, title=None, ticks=False)),
                        color=color_scale,
                    ).properties(height=max(120, len(df_group) * 50), width=350)

                    text = chart.mark_text(
                        align="left",
                        baseline="middle",
                        dx=8,
                        fontSize=13,
                        fontWeight="bold",
                        color="#334155"
                    ).encode(
                        text=alt.Text("Performance:Q", format=".1f %")
                    )

                    st.markdown(f'''
                    <div class="chart-box" style="margin-bottom:0;">
                        <div class="chart-title">{title}</div>
                        <div class="chart-subtitle">{subtitle}</div>
                        <div style="padding: 12px 8px 16px 8px;">
                    ''', unsafe_allow_html=True)
                    st.altair_chart(chart + text, use_container_width=True)
                    st.markdown('</div></div>', unsafe_allow_html=True)

            st.markdown('</div></div>', unsafe_allow_html=True)

            # ============================================================
            # 5. EXPORT
            # ============================================================
            st.markdown("""
            <div class="section-card">
                <div class="section-header" style="background: linear-gradient(135deg, #065f46, #059669);">
                    📥 Exporter les plans d'action détaillés
                </div>
                <div class="section-body">
            """, unsafe_allow_html=True)

            postes_avec_anomalies = anomalies_dashboard[anomalies_dashboard.index != "Total général"].index.tolist() if not anomalies_dashboard.empty else []

            if postes_avec_anomalies:
                selected_poste_export = st.selectbox("Sélectionnez le poste de travail pour générer le fichier Excel :", options=["All"] + postes_avec_anomalies)

                if st.button("📥 Générer et télécharger le fichier Excel", type="primary"):
                    with st.spinner("Génération du fichier en cours..."):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            postes_a_traiter = postes_avec_anomalies if selected_poste_export == "All" else [selected_poste_export]

                            for poste_export in postes_a_traiter:
                                kpis_en_defaut = df_anomalies_ot[df_anomalies_ot["Poste travail princ."] == poste_export]["KPI"].unique().tolist()
                                if "appel avis approuvé" in pivot_avis.columns and poste_export in pivot_avis.index and pivot_avis.loc[poste_export, "Nb Avis sans ordre"] > 0:
                                    kpis_en_defaut.append("appel avis approuvé")

                                for kpi in kpis_en_defaut:
                                    sheet_data = pd.DataFrame()

                                    if kpi != "appel avis approuvé":
                                        df_poste_filtered = df_processed[df_processed["Poste travail princ."] == poste_export].copy()
                                        if kpi == "TAUX_REALISATION_CORRECTIF/PT": subset_ot = df_poste_filtered[(df_poste_filtered["Nº appel pl.entret."].fillna(0) == 0) & (~df_poste_filtered["Statut OT"].isin(["CLOT", "TCLO"]))]
                                        elif kpi == "OT préparation <1 mois": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "CRÉÉ") & (df_poste_filtered["Age préparation"] != "<1 mois")]
                                        elif kpi == "OT préparation >3 mois": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "CRÉÉ") & (df_poste_filtered["Age préparation"] == ">3 mois")]
                                        elif kpi == "OT planification <1 mois": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 0) & (df_poste_filtered["Age planification"] != "<1 mois")]
                                        elif kpi == "OT planification >3 mois": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 0) & (df_poste_filtered["Age planification"] == ">3 mois")]
                                        elif kpi == "OT exécution <1 mois": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 1) & (df_poste_filtered["Age exécution"] != "<1 mois")]
                                        elif kpi == "OT exécution >3 mois": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Contient SOPL"] == 1) & (df_poste_filtered["Age exécution"] == ">3 mois")]
                                        elif kpi == "OT LANC ESTIME": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["OT LANC ESTIME"] == "NON")]
                                        elif kpi == "Backlog préparation caractérisé": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "CRÉÉ") & (df_poste_filtered["Backlog préparation"] == "NON CARACTERISE")]
                                        elif kpi == "Backlog planification caractérisé": subset_ot = df_poste_filtered[(df_poste_filtered["Statut OT"] == "LANC") & (df_poste_filtered["Backlog planification"] == "NON CARACTERISE")]
                                        elif kpi == "OT CONFIME": subset_ot = df_poste_filtered[df_poste_filtered["OT CONFIME"] == "NON"]
                                        elif kpi == "OT_COR_EGAL": subset_ot = df_poste_filtered[df_poste_filtered["OT_COR_EGAL"] == "NON"]
                                        else: subset_ot = pd.DataFrame()

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
                                            subset_avis["Action recommandée"] = "Créer un Ordre de Travail pour cet Avis ou clarifier son statut."
                                            sheet_data = pd.concat([sheet_data, subset_avis])

                                    if not sheet_data.empty:
                                        base_name = poste_export.replace(" ", "_").replace("/", "_")[:20]
                                        kpi_name = kpi.replace("/", "_").replace(" ", "_")[:10]
                                        sheet_name = f"{base_name}_{kpi_name}"[:31]
                                        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

                        output.seek(0)
                        nom_fichier = "Plan_Action_Tous_Postes.xlsx" if selected_poste_export == "All" else f"Plan_Action_{selected_poste_export.replace(' ', '_')}.xlsx"
                        st.download_button(
                            label="✅ Cliquez ici pour télécharger le fichier",
                            data=output.getvalue(),
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.info("✅ Aucune anomalie n'a été détectée. L'export est désactivé.")

            st.markdown('</div></div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Une erreur est survenue lors du traitement des fichiers : {e}")

# ==================================================
# EXECUTION DU SCRIPT
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
