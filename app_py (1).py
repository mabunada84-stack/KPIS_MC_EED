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

def inject_custom_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
        :root {
            --primary: #1e3a5f; --primary-light: #2c5282; --accent: #ed8936;
            --success: #38a169; --danger: #e53e3e; --warning: #d69e2e;
            --bg: #f7fafc; --card: #ffffff; --text: #1a202c;
            --text-secondary: #718096; --border: #e2e8f0;
            --shadow: 0 4px 20px rgba(0,0,0,0.08); --radius: 16px;
        }
        .stApp {
            background: linear-gradient(135deg, #f0f4f8 0%, #e2e8f0 50%, #f0f4f8 100%);
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        .main-header {
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
            padding: 28px 40px; border-radius: var(--radius); margin-bottom: 24px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.12); position: relative; overflow: hidden;
        }
        .main-header::before {
            content: ''; position: absolute; top: -50%; right: -10%; width: 300px; height: 300px;
            background: rgba(255,255,255,0.05); border-radius: 50%;
        }
        .main-header h1 { color: white; font-size: 30px; font-weight: 800; margin: 0; letter-spacing: -0.5px; }
        .main-header .subtitle { color: rgba(255,255,255,0.7); font-size: 14px; font-weight: 400; margin-top: 4px; }
        .main-header .date-badge {
            position: absolute; top: 28px; right: 40px; background: rgba(255,255,255,0.15);
            backdrop-filter: blur(10px); padding: 8px 18px; border-radius: 30px;
            color: white; font-size: 13px; font-weight: 500; border: 1px solid rgba(255,255,255,0.2);
        }
        .section-title {
            font-size: 18px; font-weight: 700; color: var(--primary); margin-bottom: 12px;
            padding-left: 14px; border-left: 4px solid var(--accent);
        }
        .chart-container {
            background: var(--card); border-radius: var(--radius); padding: 20px;
            box-shadow: var(--shadow); border: 1px solid var(--border);
        }
        .chart-title { font-size: 14px; font-weight: 700; color: var(--primary); margin-bottom: 12px; text-align: center; }
        .synthese-row {
            display: flex; align-items: center; padding: 12px 18px; background: var(--card);
            border-radius: 10px; margin-bottom: 6px; border: 1px solid var(--border);
            box-shadow: 0 2px 6px rgba(0,0,0,0.03); transition: all 0.2s ease;
        }
        .synthese-row:hover { transform: translateX(4px); box-shadow: 0 4px 16px rgba(0,0,0,0.08); border-left: 4px solid var(--accent); }
        .synthese-kpi-name { font-weight: 700; color: var(--primary); font-size: 13px; min-width: 260px; }
        .synthese-count { background: var(--primary); color: white; padding: 5px 14px; border-radius: 20px; font-weight: 800; font-size: 14px; min-width: 50px; text-align: center; margin: 0 14px; }
        .synthese-action { color: var(--text-secondary); font-size: 12px; flex: 1; }
        .top-badge { display: inline-block; padding: 3px 10px; border-radius: 6px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
        .top-badge.red { background: #fed7d7; color: #c53030; }
        .top-badge.orange { background: #feebc8; color: #c05621; }
        .top-badge.yellow { background: #fefcbf; color: #975a16; }
        .empty-state { text-align: center; padding: 50px 20px; color: var(--text-secondary); }
        .empty-state .icon { font-size: 44px; margin-bottom: 14px; }
        .empty-state h3 { color: var(--text); font-size: 17px; font-weight: 600; margin-bottom: 6px; }
        
        div[data-testid="stSidebar"] { background: linear-gradient(180deg, var(--primary) 0%, #0f2744 100%); }
        div[data-testid="stSidebar"] * { color: rgba(255,255,255,0.9) !important; }
        div[data-testid="stSidebar"] .stSelectbox label,
        div[data-testid="stSidebar"] .stMultiSelect label,
        div[data-testid="stSidebar"] .stDateInput label { color: rgba(255,255,255,0.8) !important; font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
        div[data-testid="stSidebar"] div[data-testid="stWidget"] { background: rgba(255,255,255,0.08); border-radius: 10px; padding: 4px 12px; margin-bottom: 6px; border: 1px solid rgba(255,255,255,0.1); }
        div[data-testid="stSidebar"] .stSelectbox > div > div,
        div[data-testid="stSidebar"] .stMultiSelect > div > div,
        div[data-testid="stSidebar"] .stDateInput > div > div { background: rgba(255,255,255,0.95) !important; border-radius: 8px; }
        
        .table-full-width {
            width: 100%; border-collapse: collapse; font-family: 'Inter', sans-serif; font-size: 10px;
            display: block; overflow-x: auto; -webkit-overflow-scrolling: touch;
        }
        .table-full-width thead th {
            background: var(--primary); color: white; font-weight: 700; font-size: 9px;
            text-transform: uppercase; letter-spacing: 0.3px; padding: 8px 6px; border: none;
            white-space: nowrap; position: sticky; top: 0; z-index: 10;
        }
        .table-full-width tbody td { padding: 6px 6px; border-bottom: 1px solid var(--border); white-space: nowrap; }
        .table-full-width tbody tr:nth-child(even) td { background: #f9fbfd; }
        .table-full-width tbody tr:hover td { background: #edf2f7 !important; }
        
        .stTabs [data-baseweb="tab-list"] { gap: 4px; background: #edf2f7; padding: 4px; border-radius: 10px; margin-bottom: 16px; }
        .stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 10px 20px; font-weight: 600; font-size: 14px; }
        .stTabs [aria-selected="true"] { background: white !important; color: var(--primary) !important; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        
        .stButton > button[kind="primary"] { background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%); border: none; border-radius: 10px; padding: 12px 28px; font-weight: 700; font-size: 14px; box-shadow: 0 4px 12px rgba(30, 58, 95, 0.3); width: 100%; }
        .stDownloadButton > button { background: linear-gradient(135deg, var(--success) 0%, #276749 100%); border: none; border-radius: 10px; padding: 14px 28px; font-weight: 700; font-size: 15px; color: white; width: 100%; box-shadow: 0 4px 12px rgba(56, 161, 105, 0.3); }
        
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }
        @media (max-width: 768px) {
            .main-header { padding: 18px 20px; }
            .main-header h1 { font-size: 20px; }
            .main-header .date-badge { position: static; margin-top: 8px; display: inline-block; }
            .synthese-kpi-name { min-width: 140px; font-size: 11px; }
        }
    </style>
    """, unsafe_allow_html=True)

def main():
    try: locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try: locale.setlocale(locale.LC_ALL, 'fr_FR')
        except: pass

    inject_custom_css()

    consignes = [ "Port obligatoire des EPI avant toute intervention.", "Port obligatoire du casque de sécurité.", "Port obligatoire des lunettes de protection.", "Port obligatoire des gants adaptés au travail.", "Utiliser les protections auditives dans les zones bruyantes.", "Vérifier l'absence de tension avant toute intervention électrique.", "Respecter la procédure de consignation et déconsignation.", "Ne jamais intervenir sur un équipement en marche.", "Baliser et sécuriser la zone de travail.", "Maintenir le poste de travail propre et ordonné.", "Vérifier l'état des outils avant utilisation.", "Utiliser uniquement du matériel homologué.", "Respecter les permis de travail en vigueur.", "Identifier les risques avant de commencer une tâche.", "Signaler immédiatement toute situation dangereuse.", "Signaler tout incident ou presque accident.", "Ne jamais neutraliser un dispositif de sécurité.", "Vérifier les détecteurs de gaz avant utilisation.", "Vérifier la bonne ventilation des zones de travail.", "Respecter les règles des espaces confinés.", "Contrôler l'atmosphère avant d'entrer dans un espace confiné.", "Utiliser les points d'ancrage pour les travaux en hauteur.", "Vérifier l'état des échafaudages avant utilisation.", "Sécuriser les outils lors des travaux en hauteur.", "Ne pas travailler seul lors d'opérations à risque.", "Contrôler les élingues avant chaque levage.", "Respecter les limites de charge des équipements.", "Vérifier l'état des appareils de levage.", "Maintenir les voies de circulation dégagées.", "Respecter la signalisation de sécurité.", "Vérifier les extincteurs à proximité du chantier.", "Connaître les issues de secours les plus proches.", "Respecter les procédures d'arrêt d'urgence.", "Vérifier les flexibles et raccords avant mise en service.", "Contrôler les fuites avant démarrage d'un équipement.", "Respecter les distances de sécurité.", "Ne jamais contourner une procédure HSE.", "Porter les EPI adaptés au risque identifié.", "Prévenir son responsable avant toute intervention particulière.", "Analyser les risques avant chaque démarrage de chantier.", "Vérifier la stabilité des équipements.", "Utiliser les bons outils pour la bonne tâche.", "Respecter les consignes spécifiques du chantier.", "Ne jamais prendre de raccourci au détriment de la sécurité.", "Arrêter immédiatement les travaux en cas de danger.", "Protéger l'environnement lors des interventions.", "Collecter et trier correctement les déchets.", "Éviter toute pollution accidentelle.", "Respecter les consignes de stockage des produits dangereux.", "Lire les fiches de sécurité avant manipulation.", "Vérifier les équipements avant chaque prise de poste.", "S'assurer de la disponibilité des moyens de secours.", "Communiquer clairement avec l'équipe avant intervention.", "Respecter les règles de circulation des engins.", "Garder une vigilance permanente sur son environnement.", "Prendre le temps d'effectuer le travail en sécurité.", "La sécurité est l'affaire de tous.", "Chaque incident peut être évité par la prévention.", "Aucun travail n'est plus urgent que la sécurité.", "Zéro accident commence par un comportement sûr." ]

    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        consigne = random.choice(consignes)
        st.markdown("""<div style="min-height:100vh; display:flex; flex-direction:column; align-items:center; justify-content:center; background:linear-gradient(135deg, #1a365d 0%, #2d3748 50%, #1a365d 100%); padding:40px;">
            <div style="font-size:64px; margin-bottom:20px;">🦺</div>
            <h1 style="text-align:center; font-size:42px; color:#fff; font-weight:900; margin:0; letter-spacing:-1px;">HSE - CONSIGNE DE SÉCURITÉ</h1>
            <p style="text-align:center; color:rgba(255,255,255,0.6); font-size:18px; margin-top:8px; font-weight:400; letter-spacing:3px; text-transform:uppercase;">Sécurité &bull; Santé &bull; Environnement</p>
            <div style="background:linear-gradient(135deg, #f6e05e 0%, #ed8936 100%); padding:36px 48px; border-radius:20px; font-size:28px; font-weight:700; text-align:center; margin-top:40px; margin-bottom:40px; color:#1a202c; max-width:800px; box-shadow:0 20px 60px rgba(0,0,0,0.3);">⚠️ {consigne}</div>
            <h2 style="text-align:center; color:#48bb78; font-size:32px; font-weight:900; letter-spacing:-0.5px;">Aucun travail n'est plus urgent que la sécurité</h2>
            <div style="margin-top:40px; width:200px; height:4px; background:rgba(255,255,255,0.1); border-radius:2px; overflow:hidden;"><div style="width:100%; height:100%; background:linear-gradient(90deg, #48bb78, #38a169); border-radius:2px; animation:loading 5.5s ease-in-out forwards;"></div></div>
            <style>@keyframes loading {{ from {{ width: 0%; }} to {{ width: 100%; }} }}</style>
        </div>""".format(consigne=consigne), unsafe_allow_html=True)
        time.sleep(6); st.session_state.hse_affiche = True; st.rerun(); st.stop()

    def rename_safe(df, old_names, new_names):
        return df.rename(columns={old: new for old, new in zip(old_names, new_names) if old in df.columns})
    def contient_mot(texte, liste_mots):
        texte = str(texte); return any(mot in texte for l in liste_mots for mot in l.split())
    def categorie_age(age):
        if age <= 1: return "<1 mois"
        elif age >= 3: return ">3 mois"
        return "1 mois < <3 mois"
    def calcul_kpi(n, d, sz=100): return np.where(d == 0, sz, (n / d) * 100)
    def creer_pivot(df, f, c, p):
        return pd.pivot_table(df[f], index="Poste travail princ.", columns=c, values="Ordre", aggfunc="count", fill_value=0).reindex(p, fill_value=0)
    def exclure_cresseurs(df):
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)].copy() if "Poste travail princ." in df.columns else df

    def calculate_all_kpis(df_input, avis_input, now, postes):
        res = {}; df = df_input.copy(); av = avis_input.copy()
        mp = ["CRPR ATPD", "CRPR ATMR", "CRPR ATER", "CRPR ATRS", "CRPR ATMO", "ATPD", "ATMR", "ATER", "ATRS", "ATMO"]
        df["Backlog préparation"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mp)), "CARACTERISE", "NON CARACTERISE")
        mplan = ["ATPL ATEI", "ATPL ATAL", "ATPL ATER", "ATPL AGAR", "ATPL ATHS", "ATEI", "ATAL", "ATAS", "AGAR", "ATHS"]
        df["Backlog planification"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mplan)), "CARACTERISE", "NON CARACTERISE")
        for dc, am, ac in [('Créé le', "Age mois préparation", "Age préparation"), ('Date de début planifiée', "Age mois planification", "Age planification"), ('Date de début planifiée', "Age mois exécution", "Age exécution")]:
            if dc in df.columns:
                df[dc] = pd.to_datetime(df[dc], errors='coerce')
                df[am] = ((now.year - df[dc].dt.year) * 12 + (now.month - df[dc].dt.month)).round(2)
                df[ac] = df[am].apply(categorie_age)
            else: df[am] = np.nan; df[ac] = "Inconnu"
        df["OT CONFIME"] = np.where(df["Statut système"].str.contains("CLO", na=False) & df["Statut système"].str.contains("CONF", na=False), "OUI", "NON")
        df["Contient SOPL"] = df["Statut utilisateur"].str.contains("SOPL", na=False).map({True: 1, False: 0})
        df["OT LANC ESTIME"] = np.where(df["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
        df["OT_COR_EGAL"] = np.where((df["Total coûts budgétés"].fillna(0) - df["Total coûts réels"].fillna(0)) == 0, "OUI", "NON")
        res['df_processed'] = df
        an = creer_pivot(df, df["Nº appel pl.entret."].fillna(0) == 0, "Statut OT", postes)
        for c in ["CLOT", "CRÉÉ", "LANC", "TCLO"]: an[c] = an.get(c, 0)
        an["Total"] = an[["CLOT", "CRÉÉ", "LANC", "TCLO"]].sum(axis=1); an["TAUX_REALISATION_CORRECTIF/PT"] = calcul_kpi(an["TCLO"], an["Total"])
        pr = creer_pivot(df, df["Statut OT"] == "CRÉÉ", "Age préparation", postes)
        for c in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: pr[c] = pr.get(c, 0)
        pr["Total"] = pr[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
        pr["OT préparation <1 mois"] = calcul_kpi(pr["<1 mois"], pr["Total"]); pr["OT préparation >3 mois"] = calcul_kpi(pr[">3 mois"], pr["Total"], 0); pr["OT préparation 1mois< <3mois"] = calcul_kpi(pr["1 mois < <3 mois"], pr["Total"], 0)
        pl = creer_pivot(df, (df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 0), "Age planification", postes)
        for c in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: pl[c] = pl.get(c, 0)
        pl["Total"] = pl[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
        pl["OT planification <1 mois"] = calcul_kpi(pl["<1 mois"], pl["Total"]); pl["OT planification >3 mois"] = calcul_kpi(pl[">3 mois"], pl["Total"], 0); pl["OT planification 1mois< <3mois"] = calcul_kpi(pl["1 mois < <3 mois"], pl["Total"], 0)
        ex = creer_pivot(df, (df["Statut OT"] == "LANC") & (df["Contient SOPL"] == 1), "Age exécution", postes)
        for c in ["<1 mois", ">3 mois", "1 mois < <3 mois"]: ex[c] = ex.get(c, 0)
        ex["Total"] = ex[["<1 mois", "1 mois < <3 mois", ">3 mois"]].sum(axis=1)
        ex["OT exécution <1 mois"] = calcul_kpi(ex["<1 mois"], ex["Total"]); ex["OT exécution >3 mois"] = calcul_kpi(ex[">3 mois"], ex["Total"], 0); ex["OT exécution 1mois< <3mois"] = calcul_kpi(ex["1 mois < <3 mois"], ex["Total"], 0)
        la = pd.pivot_table(df[df["Statut OT"] == "LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["OUI", "NON"]: la[c] = la.get(c, 0)
        la["Total"] = la["OUI"] + la["NON"]; la["OT LANC ESTIME"] = calcul_kpi(la["OUI"], la["Total"])
        pc = pd.pivot_table(df[df["Statut OT"] == "CRÉÉ"], index="Poste travail princ.", columns="Backlog préparation", values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["CARACTERISE", "NON CARACTERISE"]: pc[c] = pc.get(c, 0)
        pc["Total"] = pc["CARACTERISE"] + pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"] = calcul_kpi(pc["CARACTERISE"], pc["Total"])
        plc = pd.pivot_table(df[df["Statut OT"] == "LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["CARACTERISE", "NON CARACTERISE"]: plc[c] = plc.get(c, 0)
        plc["Total"] = plc["CARACTERISE"] + plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"] = calcul_kpi(plc["CARACTERISE"], plc["Total"])
        for kn, cn in [("OT CONFIME", "OT CONFIME"), ("OT_COR_EGAL", "OT_COR_EGAL")]:
            pv = pd.pivot_table(df, index="Poste travail princ.", columns=cn, values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
            for c in ["OUI", "NON"]: pv[c] = pv.get(c, 0)
            pv["Total"] = pv["OUI"] + pv["NON"]; pv[cn] = calcul_kpi(pv["OUI"], pv["Total"]); res[kn.lower().replace(" ", "_")] = pv
        avf = av[(av["Ordre"].isna()) | (av["Ordre"].astype(str).str.strip() == "")].copy(); res['avis_df_filtered'] = avf
        tca = pd.pivot_table(avf, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["APRQ", "APRV", "APRV AVAU", "REJT"]: tca[c] = tca.get(c, 0)
        tca["Total"] = tca[["APRQ", "APRV", "APRV AVAU", "REJT"]].sum(axis=1); tca["appel avis approuvé"] = calcul_kpi(tca["APRV"], tca["Total"])
        res['calculated_kpis_df'] = pd.concat([an[["TAUX_REALISATION_CORRECTIF/PT"]], pr[["OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois"]], pl[["OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois"]], ex[["OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]], tca[["appel avis approuvé"]], la[["OT LANC ESTIME"]], pc[["Backlog préparation caractérisé"]], plc[["Backlog planification caractérisé"]], res['ot_confime'][["OT CONFIME"]], res['ot_cor_egal'][["OT_COR_EGAL"]]], axis=1)
        return res

    def get_kpi_style(v, c):
        try: val = float(v)
        except: return ""
        if c in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]:
            return "background-color:#c6efce; color:#006100; font-weight:600;" if val >= 80 else ("background-color:#ffeb9c; color:#9c6500; font-weight:600;" if val >= 75 else "background-color:#ffc7ce; color:#9c0006; font-weight:600;")
        if c in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]:
            return "background-color:#c6efce; color:#006100; font-weight:600;" if val <= 15 else "background-color:#ffc7ce; color:#9c0006; font-weight:600;"
        if c in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]:
            return "background-color:#c6efce; color:#006100; font-weight:600;" if val <= 5 else "background-color:#ffc7ce; color:#9c0006; font-weight:600;"
        if c == "TAUX_REALISATION_CORRECTIF/PT":
            return "background-color:#c6efce; color:#006100; font-weight:600;" if val >= 85 else ("background-color:#ffeb9c; color:#9c6500; font-weight:600;" if val >= 80 else "background-color:#ffc7ce; color:#9c0006; font-weight:600;")
        if c == "appel avis approuvé":
            return "background-color:#c6efce; color:#006100; font-weight:600;" if val >= 95 else ("background-color:#ffeb9c; color:#9c6500; font-weight:600;" if val >= 90 else "background-color:#ffc7ce; color:#9c0006; font-weight:600;")
        if c in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]:
            return "background-color:#c6efce; color:#006100; font-weight:600;" if val >= 100 else ("background-color:#ffeb9c; color:#9c6500; font-weight:600;" if val >= 95 else "background-color:#ffc7ce; color:#9c0006; font-weight:600;")
        return ""

    def get_anomalie_style(v):
        try: val = int(v)
        except: return ""
        if val == 0: return "color:#a0aec0;"
        if val <= 3: return "background-color:#ffeb9c; color:#9c6500; font-weight:600;"
        if val <= 10: return "background-color:#fed7d7; color:#c53030; font-weight:600;"
        return "background-color:#fc8181; color:#742a2a; font-weight:800;"

    def get_class_style(v):
        try: val = float(str(v).replace(' %', '').strip()); 
        except: return ""
        return "background-color:#c6efce; color:#006100; font-weight:700;" if val >= 90 else ("background-color:#ffeb9c; color:#9c6500; font-weight:700;" if val >= 80 else "background-color:#ffc7ce; color:#9c0006; font-weight:700;")

    def df_to_html_kpi(df):
        df = df.rename_axis("Poste de travail").reset_index()
        h = '<table class="table-full-width"><thead><tr>' + ''.join(f'<th>{c}</th>' for c in df.columns) + '</tr></thead><tbody>'
        for i, r in df.iterrows():
            is_c = r["Poste de travail"] == "CIBLE"; is_t = r["Poste de travail"] == "Total général"
            h += f'<tr style="{"border-top:3px solid #1e3a5f;" if is_t else ""}">'
            for c in df.columns:
                v = r[c]
                if is_c: h += f'<td style="padding:8px 6px; background-color:#1e3a5f; color:#ffffff; font-weight:700; font-size:10px;">{v}</td>'
                elif is_t:
                    s = get_kpi_style(v, c)
                    s = s if s else "background-color:#e2e8f0; color:#1a202c;"
                    h += f'<td style="padding:8px 6px; {s} font-weight:800; font-size:10px;">{v}</td>'
                else: h += f'<td style="padding:6px 6px; font-size:10px; {get_kpi_style(v, c)}">{v}</td>'
            h += '</tr>'
        return h + '</tbody></table>'

    def df_to_html_anomalies(df):
        df = df.rename_axis("Poste de travail").reset_index()
        h = '<table class="table-full-width"><thead><tr>' + ''.join(f'<th>{c}</th>' for c in df.columns) + '</tr></thead><tbody>'
        for i, r in df.iterrows():
            is_t = r["Poste de travail"] == "Total général"
            h += '<tr>'
            for c in df.columns:
                v = r[c]
                if is_t:
                    s = get_anomalie_style(v)
                    s = s if s else "color:#1a202c;"
                    h += f'<td style="padding:8px 6px; {s} font-weight:800; font-size:10px;">{v}</td>'
                else: h += f'<td style="padding:6px 6px; font-size:10px; {get_anomalie_style(v)}">{v}</td>'
            h += '</tr>'
        return h + '</tbody></table>'

    def df_to_html_class(df):
        h = '<table class="table-full-width"><thead><tr>' + ''.join(f'<th>{c}</th>' for c in df.columns) + '</tr></thead><tbody>'
        for i, r in df.iterrows():
            is_t = r.get("Poste travail princ.") == "Total général"
            h += f'<tr style="{"border-top:3px solid #1e3a5f;" if is_t else ""}">'
            for c in df.columns:
                v = r[c]
                if is_t:
                    s = get_class_style(v) if c in ["Score KPIs Quantité", "Score KPIs Qualité"] else "background-color:#e2e8f0; color:#1a202c;"
                    h += f'<td style="padding:8px 6px; {s} font-weight:800; font-size:10px;">{v}</td>'
                else:
                    s = get_class_style(v) if c in ["Score KPIs Quantité", "Score KPIs Qualité"] else ""
                    h += f'<td style="padding:6px 6px; font-size:10px; {s}">{v}</td>'
            h += '</tr>'
        return h + '</tbody></table>'

    def get_kpi_score(k, a, t):
        if pd.isna(a) or pd.isna(t): return 0
        if k in ["OT préparation <1 mois", "OT planification <1 mois", "OT exécution <1 mois"]: return 1 if a >= 75 else 0
        if k in ["OT préparation 1mois< <3mois", "OT planification 1mois< <3mois", "OT exécution 1mois< <3mois"]: return 1 if a <= 15 else 0
        if k in ["OT préparation >3 mois", "OT planification >3 mois", "OT exécution >3 mois"]: return 1 if a <= 5 else 0
        if k == "TAUX_REALISATION_CORRECTIF/PT": return 1 if a >= 80 else 0
        if k == "appel avis approuvé": return 1 if a >= 90 else 0
        if k in ["OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]: return 1 if a >= 95 else 0
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
        if "PS" in p: return "Sulfurique"
        if "PP" in p: return "Phosphorique"
        if "TSP" in p or "REX" in p: return "Engrais"
        if "MCP" in p or "DCP" in p: return "Feed"
        return "Autre"
    def get_division(p):
        p = str(p).upper()
        if "SF1" in p: return "SF1"
        if "SF2" in p: return "SF2"
        return "Autre"

    # ==================================================
    # SIDEBAR
    # ==================================================
    with st.sidebar:
        st.markdown("""<div style="padding: 16px 0 8px 0;"><div style="font-size:24px; margin-bottom:4px;">⚙️</div><div style="font-size:15px; font-weight:800; color:white;">Filtres & Paramètres</div><div style="font-size:10px; color:rgba(255,255,255,0.5); text-transform:uppercase; letter-spacing:1px;">Configuration</div></div>""", unsafe_allow_html=True)
        st.markdown("---")
        use_new_files = st.toggle("📁 Charger de nouveaux fichiers", value=False, key="toggle_files")
        ot_file = avis_file = None
        all_postes_master_list = []
        
        if use_new_files:
            ot_file = st.file_uploader("Fichier OT", type=["xlsx"], key="up_ot")
            avis_file = st.file_uploader("Fichier AVIS", type=["xlsx"], key="up_avis")
        else:
            date_fichier = datetime.now().strftime("%d/%m/%Y")
            if os.path.exists("ot.xlsx"):
                try:
                    date_fichier = datetime.fromtimestamp(os.path.getmtime("ot.xlsx")).strftime("%d/%m/%Y")
                    _tmp = exclure_cresseurs(pd.read_excel("ot.xlsx"))
                    all_postes_master_list = sorted(_tmp[_tmp["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
                except: pass
            st.markdown(f"""<div style="background:rgba(255,255,255,0.1); padding:10px 14px; border-radius:8px; border:1px solid rgba(255,255,255,0.15);"><div style="font-size:10px; color:rgba(255,255,255,0.5); text-transform:uppercase; letter-spacing:1px;">Données</div><div style="font-size:13px; color:white; font-weight:600; margin-top:2px;">📅 {date_fichier}</div></div>""", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("**🎯 Postes**")
        selected_postes = st.multiselect("Poste de travail", ["All"] + all_postes_master_list, ["All"], key="sel_postes")
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
                df_ot_raw = pd.read_excel(ot_file); avis_df_raw = pd.read_excel(avis_file); date_fichier = datetime.now().strftime("%d/%m/%Y")
            else:
                df_ot_raw = pd.read_excel("ot.xlsx"); avis_df_raw = pd.read_excel("avis.xlsx")
            
            df_ot_raw = exclure_cresseurs(df_ot_raw); avis_df_raw = exclure_cresseurs(avis_df_raw)
            for c in ["Créé le", "Date de début planifiée", "Date de clôture", "Début réel", "Fin réelle"]:
                if c in df_ot_raw.columns: df_ot_raw[c] = pd.to_datetime(df_ot_raw[c], errors="coerce")
            for c in ["Créé le", "Début souhaité", "Date de la clôture"]:
                if c in avis_df_raw.columns: avis_df_raw[c] = pd.to_datetime(avis_df_raw[c], errors="coerce")

            if not all_postes_master_list: all_postes_master_list = sorted(df_ot_raw[df_ot_raw["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
            if "All" in selected_postes or not selected_postes: selected_postes = all_postes_master_list
            if "All" in selected_ateliers or not selected_ateliers: selected_ateliers = ["All"]
            if "All" in selected_divisions or not selected_divisions: selected_divisions = ["All"]
            
            start_date = pd.to_datetime(date_range[0]) if len(date_range) == 2 else pd.to_datetime(datetime(2025, 1, 1))
            end_date = pd.to_datetime(date_range[1]) if len(date_range) == 2 else pd.to_datetime(datetime.today())

            def match_filters(poste):
                p = str(poste).upper()
                if "All" not in selected_ateliers:
                    m = False
                    if "Sulfurique (PS)" in selected_ateliers and "PS" in p: m = True
                    if "Phosphorique (PP)" in selected_ateliers and "PP" in p: m = True
                    if "Engrais (TSP/REX)" in selected_ateliers and ("TSP" in p or "REX" in p): m = True
                    if "Feed (MCP/DCP)" in selected_ateliers and ("MCP" in p or "DCP" in p): m = True
                    if not m: return False
                if "All" not in selected_divisions:
                    m = False
                    if "SF1" in selected_divisions and "SF1" in p: m = True
                    if "SF2" in selected_divisions and "SF2" in p: m = True
                    if not m: return False
                return True

            valid_postes = [p for p in all_postes_master_list if match_filters(p) and p in selected_postes]
            df = df_ot_raw[(df_ot_raw["Poste travail princ."].isin(valid_postes)) & (df_ot_raw["Date de début planifiée"].between(start_date, end_date))].copy()
            avis_df = avis_df_raw[avis_df_raw["Poste travail princ."].isin(valid_postes)].copy()
            df = exclure_cresseurs(df[df["Poste travail princ."].astype(str).str.startswith(("SF1", "SF2"), na=False)].drop_duplicates())
            avis_df = exclure_cresseurs(avis_df[(avis_df["Ordre"].isna()) | (avis_df["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates())
            if "Statut système" in df.columns: df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            now = pd.Timestamp.now()
            results = calculate_all_kpis(df, avis_df, now, valid_postes)
            calculated_kpis_df = results['calculated_kpis_df']; df_processed = results['df_processed']

            cible = pd.DataFrame([{"TAUX_REALISATION_CORRECTIF/PT": 85, "OT préparation <1 mois": 80, "OT préparation >3 mois": 5, "OT préparation 1mois< <3mois": 15, "OT planification <1 mois": 80, "OT planification >3 mois": 5, "OT planification 1mois< <3mois": 15, "OT exécution <1 mois": 80, "OT exécution >3 mois": 5, "OT exécution 1mois< <3mois": 15, "appel avis approuvé": 95, "OT LANC ESTIME": 100, "Backlog préparation caractérisé": 100, "Backlog planification caractérisé": 100, "OT CONFIME": 100, "OT_COR_EGAL": 100}], index=["CIBLE"])

            # ANOMALIES
            ano_ot, ano_av = [], []
            for poste in valid_postes:
                if poste not in df_processed["Poste travail princ."].values: continue
                dp = df_processed[df_processed["Poste travail princ."] == poste]
                ap = results['avis_df_filtered'][results['avis_df_filtered']["Poste travail princ."] == poste]
                for kn, cc, sub, act in [
                    ("TAUX_REALISATION_CORRECTIF/PT", "TAUX_REALISATION_CORRECTIF/PT", dp[(dp["Nº appel pl.entret."].fillna(0) == 0) & (~dp["Statut OT"].isin(["CLOT", "TCLO"]))], "Améliorer le taux de réalisation des OT."),
                    ("OT préparation <1 mois", "OT préparation <1 mois", dp[(dp["Statut OT"] == "CRÉÉ") & (dp["Age préparation"] != "<1 mois")], "Réduire l'âge de préparation des OT (< 1 mois)."),
                    ("OT préparation >3 mois", "OT préparation >3 mois", dp[(dp["Statut OT"] == "CRÉÉ") & (dp["Age préparation"] == ">3 mois")], "Traiter les OT avec préparation > 3 mois."),
                    ("OT planification <1 mois", "OT planification <1 mois", dp[(dp["Statut OT"] == "LANC") & (dp["Contient SOPL"] == 0) & (dp["Age planification"] != "<1 mois")], "Réduire l'âge de planification des OT (< 1 mois)."),
                    ("OT planification >3 mois", "OT planification >3 mois", dp[(dp["Statut OT"] == "LANC") & (dp["Contient SOPL"] == 0) & (dp["Age planification"] == ">3 mois")], "Traiter les OT avec planification > 3 mois."),
                    ("OT exécution <1 mois", "OT exécution <1 mois", dp[(dp["Statut OT"] == "LANC") & (dp["Contient SOPL"] == 1) & (dp["Age exécution"] != "<1 mois")], "Réduire l'âge d'exécution des OT (< 1 mois)."),
                    ("OT exécution >3 mois", "OT exécution >3 mois", dp[(dp["Statut OT"] == "LANC") & (dp["Contient SOPL"] == 1) & (dp["Age exécution"] == ">3 mois")], "Traiter les OT avec exécution > 3 mois."),
                    ("OT LANC ESTIME", "OT LANC ESTIME", dp[(dp["Statut OT"] == "LANC") & (dp["OT LANC ESTIME"] == "NON")], "Estimer les coûts des OT lancés."),
                    ("Backlog préparation caractérisé", "Backlog préparation caractérisé", dp[(dp["Statut OT"] == "CRÉÉ") & (dp["Backlog préparation"] == "NON CARACTERISE")], "Caractériser le backlog de préparation."),
                    ("Backlog planification caractérisé", "Backlog planification caractérisé", dp[(dp["Statut OT"] == "LANC") & (dp["Backlog planification"] == "NON CARACTERISE")], "Caractériser le backlog de planification."),
                    ("OT CONFIME", "OT CONFIME", dp[dp["OT CONFIME"] == "NON"], "Confirmer les OT terminés."),
                    ("OT_COR_EGAL", "OT_COR_EGAL", dp[dp["OT_COR_EGAL"] == "NON"], "Rapprocher les coûts réels et budgétés.")]:
                    vk = calculated_kpis_df.loc[poste, kn] if poste in calculated_kpis_df.index else 100
                    if pd.notna(vk) and vk < cible.loc['CIBLE', cc]:
                        cnt = len(sub)
                        if cnt > 0: ano_ot.append({"Poste travail princ.": poste, "KPI": kn, "Nb OT impactés": cnt, "Action Suggérée": act})
                va = calculated_kpis_df.loc[poste, "appel avis approuvé"] if poste in calculated_kpis_df.index else 100
                if pd.notna(va) and va < cible.loc['CIBLE', "appel avis approuvé"]:
                    cnt = len(ap)
                    if cnt > 0: ano_av.append({"Poste travail princ.": poste, "KPI": "appel avis approuvé", "Nb OT impactés": cnt, "Action Suggérée": "Créer un OT pour les avis sans ordre."})

            df_ano_ot = pd.DataFrame(ano_ot); df_ano_av = pd.DataFrame(ano_av)
            pot = df_ano_ot.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb OT impactés", aggfunc="sum", fill_value=0) if not df_ano_ot.empty else pd.DataFrame()
            pav = df_ano_av.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb OT impactés", aggfunc="sum", fill_value=0).rename(columns={"appel avis approuvé": "Nb Avis sans ordre"}) if not df_ano_av.empty else pd.DataFrame()
            ano_dash = pot.join(pav, how='outer').fillna(0).astype(int)
            if not ano_dash.empty:
                ano_dash["Total éléments impactés"] = ano_dash.sum(axis=1)
                ano_dash.loc["Total général"] = ano_dash.sum()
            
            all_ano = pd.concat([df_ano_ot, df_ano_av], ignore_index=True) if not df_ano_ot.empty or not df_ano_av.empty else pd.DataFrame()
            synthese = all_ano.groupby(["KPI", "Action Suggérée"])["Nb OT impactés"].sum().reset_index().sort_values("Nb OT impactés", ascending=False).reset_index(drop=True) if not all_ano.empty else pd.DataFrame()

            qty_k = ["TAUX_REALISATION_CORRECTIF/PT", "OT préparation <1 mois", "OT préparation >3 mois", "OT préparation 1mois< <3mois", "OT planification <1 mois", "OT planification >3 mois", "OT planification 1mois< <3mois", "OT exécution <1 mois", "OT exécution >3 mois", "OT exécution 1mois< <3mois"]
            qual_k = ["appel avis approuvé", "OT LANC ESTIME", "Backlog préparation caractérisé", "Backlog planification caractérisé", "OT CONFIME", "OT_COR_EGAL"]
            
            class_r = []
            for poste in calculated_kpis_df.index:
                r = calculated_kpis_df.loc[poste]
                sq = sum(get_kpi_score(k, r[k], cible.loc['CIBLE', k]) for k in qty_k if k in r.index)
                sl = sum(get_kpi_score(k, r[k], cible.loc['CIBLE', k]) for k in qual_k if k in r.index)
                pq = (sq / len(qty_k) * 100) if qty_k else 0
                pl = (sl / len(qual_k) * 100) if qual_k else 0
                class_r.append({"Poste travail princ.": poste, "Score KPIs Quantité": pq, "Score KPIs Qualité": pl})
            df_class = pd.DataFrame(class_r)
            df_class["Métier"] = df_class["Poste travail princ."].apply(get_metier)
            df_class["Atelier"] = df_class["Poste travail princ."].apply(get_atelier)
            df_class["Division"] = df_class["Poste travail princ."].apply(get_division)

            tg = pd.DataFrame(calculated_kpis_df.mean()).T; tg.index = ["Total général"]
            final_kpi = pd.concat([cible, calculated_kpis_df, tg]).round(2)

            # Classement sans Total performance
            df_cl_disp = df_class[["Poste travail princ.", "Score KPIs Quantité", "Score KPIs Qualité"]].copy()
            df_cl_disp["Score KPIs Quantité"] = df_cl_disp["Score KPIs Quantité"].apply(lambda x: f"{x:.2f} %")
            df_cl_disp["Score KPIs Qualité"] = df_cl_disp["Score KPIs Qualité"].apply(lambda x: f"{x:.2f} %")
            tg_cl = pd.DataFrame([{"Poste travail princ.": "Total général", "Score KPIs Quantité": f"{df_class['Score KPIs Quantité'].mean():.2f} %", "Score KPIs Qualité": f"{df_class['Score KPIs Qualité'].mean():.2f} %"}])
            df_cl_disp = pd.concat([df_cl_disp, tg_cl], ignore_index=True)

            # HEADER
            st.markdown(f"""<div class="main-header"><h1>📊 KPI Dashboard MC & FEED</h1><div class="subtitle">Maintenance Conditionnelle • Suivi des indicateurs de performance</div><div class="date-badge">📅 {date_fichier}</div></div>""", unsafe_allow_html=True)

            # ONGLETS
            tab_kpi, tab_ano = st.tabs(["📊 TABLEAU DE BORD DES KPIs", "🚨 TABLEAU DE BORD DES ANOMALIES"])
            with tab_kpi:
                st.markdown('<p class="section-title">Indicateurs par Poste de Travail</p>', unsafe_allow_html=True)
                st.markdown(df_to_html_kpi(final_kpi), unsafe_allow_html=True)
                st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
                st.markdown('<p class="section-title">🏆 Classement des Postes par Performance</p>', unsafe_allow_html=True)
                st.markdown(df_to_html_class(df_cl_disp), unsafe_allow_html=True)
            with tab_ano:
                if not ano_dash.empty:
                    st.markdown('<p class="section-title">Anomalies par Poste de Travail</p>', unsafe_allow_html=True)
                    st.markdown(df_to_html_anomalies(ano_dash), unsafe_allow_html=True)
                else:
                    st.markdown("""<div class="empty-state"><div class="icon">✅</div><h3>Aucune anomalie détectée</h3><p>Tous les KPIs atteignent leurs cibles.</p></div>""", unsafe_allow_html=True)

            # SYNTHÈSE
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">📋 Synthèse des Actions par KPI</p>', unsafe_allow_html=True)
            if not synthese.empty:
                for i, r in synthese.iterrows():
                    nb = int(r["Nb OT impactés"])
                    bg = "red" if nb >= 20 else ("orange" if nb >= 5 else "yellow")
                    lb = "Critique" if nb >= 20 else ("Moyen" if nb >= 5 else "Mineur")
                    st.markdown(f"""<div class="synthese-row"><div class="synthese-kpi-name">{r['KPI']}</div><div class="synthese-count">{nb}</div><div class="synthese-action">{r['Action Suggérée']}</div><span class="top-badge {bg}">{lb}</span></div>""", unsafe_allow_html=True)
            else:
                st.markdown("""<div class="empty-state" style="padding:30px;"><div class="icon">🎉</div><h3>Aucune action requise</h3><p>Tous les indicateurs sont dans le vert.</p></div>""", unsafe_allow_html=True)

            # CHARTS
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">📈 Analyse par Catégorie</p>', unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            
            with c1:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">🔧 Par Métier</div>', unsafe_allow_html=True)
                df_m = df_class.groupby("Métier")["Score KPIs Quantité"].mean().reset_index().round(1)
                df_m.columns = ["Catégorie", "Score"]
                if not df_m.empty:
                    b = alt.Chart(df_m).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(x=alt.X('Score:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title=''), y=alt.Y('Catégorie:O', sort='-x', title='', axis=alt.Axis(labelFontSize=12)), color=alt.Color('Score:Q', scale=alt.Scale(domain=[0, 50, 80, 100], range=['#e53e3e', '#ed8936', '#ecc94b', '#38a169']), legend=None), tooltip=['Catégorie', 'Score']).properties(height=max(120, 40 * len(df_m)))
                    t = alt.Chart(df_m).mark_text(align='left', baseline='middle', dx=6, fontSize=13, fontWeight=700).encode(y=alt.Y('Catégorie:O', sort='-x'), text=alt.Text('Score:Q', format='.1f'))
                    st.altair_chart((b+t).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent'), use_container_width=True)
                else: st.info("Aucune donnée")
                st.markdown('</div>', unsafe_allow_html=True)

            with c2:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">🏭 Par Atelier</div>', unsafe_allow_html=True)
                df_a = df_class.groupby("Atelier")["Score KPIs Quantité"].mean().reset_index().round(1)
                df_a.columns = ["Catégorie", "Score"]
                if not df_a.empty:
                    b = alt.Chart(df_a).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(x=alt.X('Score:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title=''), y=alt.Y('Catégorie:O', sort='-x', title='', axis=alt.Axis(labelFontSize=12)), color=alt.Color('Score:Q', scale=alt.Scale(domain=[0, 50, 80, 100], range=['#e53e3e', '#ed8936', '#ecc94b', '#38a169']), legend=None), tooltip=['Catégorie', 'Score']).properties(height=max(120, 40 * len(df_a)))
                    t = alt.Chart(df_a).mark_text(align='left', baseline='middle', dx=6, fontSize=13, fontWeight=700).encode(y=alt.Y('Catégorie:O', sort='-x'), text=alt.Text('Score:Q', format='.1f'))
                    st.altair_chart((b+t).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent'), use_container_width=True)
                else: st.info("Aucune donnée")
                st.markdown('</div>', unsafe_allow_html=True)

            with c3:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.markdown('<div class="chart-title">🏢 Par Division</div>', unsafe_allow_html=True)
                df_d = df_class.groupby("Division")["Score KPIs Quantité"].mean().reset_index().round(1)
                df_d.columns = ["Catégorie", "Score"]
                if not df_d.empty:
                    b = alt.Chart(df_d).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(x=alt.X('Score:Q', scale=alt.Scale(domain=[0, 100], clamp=True), title=''), y=alt.Y('Catégorie:O', sort='-x', title='', axis=alt.Axis(labelFontSize=12)), color=alt.Color('Score:Q', scale=alt.Scale(domain=[0, 50, 80, 100], range=['#e53e3e', '#ed8936', '#ecc94b', '#38a169']), legend=None), tooltip=['Catégorie', 'Score']).properties(height=max(120, 40 * len(df_d)))
                    t = alt.Chart(df_d).mark_text(align='left', baseline='middle', dx=6, fontSize=13, fontWeight=700).encode(y=alt.Y('Catégorie:O', sort='-x'), text=alt.Text('Score:Q', format='.1f'))
                    st.altair_chart((b+t).configure_axis(grid=False, ticks=False, domain=False).configure_view(stroke='transparent'), use_container_width=True)
                else: st.info("Aucune donnée")
                st.markdown('</div>', unsafe_allow_html=True)

            # EXPORT
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            st.markdown('<p class="section-title">💾 Export des Plans d\'Action</p>', unsafe_allow_html=True)
            postes_ano = ano_dash[ano_dash.index != "Total général"].index.tolist() if not ano_dash.empty else []
            if postes_ano:
                ce1, ce2 = st.columns([1, 1])
                with ce1: sel_exp = st.selectbox("Poste de travail :", options=["📌 Tous les postes"] + postes_ano, key="sel_exp")
                with ce2: 
                    st.markdown("<div style='height:32px;'></div>", unsafe_allow_html=True)
                    gen_btn = st.button("📥 Générer le fichier Excel", type="primary", key="btn_exp", use_container_width=True)
                if gen_btn:
                    with st.spinner("Génération en cours..."):
                        out = io.BytesIO()
                        with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                            for pe in (postes_ano if sel_exp == "📌 Tous les postes" else [sel_exp]):
                                kds = df_ano_ot[df_ano_ot["Poste travail princ."] == pe]["KPI"].unique().tolist()
                                if not df_ano_av.empty and "appel avis approuvé" in df_ano_av[df_ano_av["Poste travail princ."] == pe]["KPI"].values: kds.append("appel avis approuvé")
                                for kpi in kds:
                                    sd = pd.DataFrame()
                                    if kpi != "appel avis approuvé":
                                        dpf = df_processed[df_processed["Poste travail princ."] == pe].copy()
                                        cmap = {"TAUX_REALISATION_CORRECTIF/PT": (dpf["Nº appel pl.entret."].fillna(0) == 0) & (~dpf["Statut OT"].isin(["CLOT", "TCLO"])), "OT préparation <1 mois": (dpf["Statut OT"] == "CRÉÉ") & (dpf["Age préparation"] != "<1 mois"), "OT préparation >3 mois": (dpf["Statut OT"] == "CRÉÉ") & (dpf["Age préparation"] == ">3 mois"), "OT planification <1 mois": (dpf["Statut OT"] == "LANC") & (dpf["Contient SOPL"] == 0) & (dpf["Age planification"] != "<1 mois"), "OT planification >3 mois": (dpf["Statut OT"] == "LANC") & (dpf["Contient SOPL"] == 0) & (dpf["Age planification"] == ">3 mois"), "OT exécution <1 mois": (dpf["Statut OT"] == "LANC") & (dpf["Contient SOPL"] == 1) & (dpf["Age exécution"] != "<1 mois"), "OT exécution >3 mois": (dpf["Statut OT"] == "LANC") & (dpf["Contient SOPL"] == 1) & (dpf["Age exécution"] == ">3 mois"), "OT LANC ESTIME": (dpf["Statut OT"] == "LANC") & (dpf["OT LANC ESTIME"] == "NON"), "Backlog préparation caractérisé": (dpf["Statut OT"] == "CRÉÉ") & (dpf["Backlog préparation"] == "NON CARACTERISE"), "Backlog planification caractérisé": (dpf["Statut OT"] == "LANC") & (dpf["Backlog planification"] == "NON CARACTERISE"), "OT CONFIME": dpf["OT CONFIME"] == "NON", "OT_COR_EGAL": dpf["OT_COR_EGAL"] == "NON"}
                                        sub = dpf[cmap.get(kpi, pd.Series(False, index=dpf.index))]
                                        if not sub.empty:
                                            sub = rename_safe(sub, ["Ordre", "Désignation", "Emplacement technique", "Poste travail princ.", "Statut système", "Statut utilisateur", "Date de début planifiée", "Type d'ordre", "Backlog préparation", "Backlog planification"], ["Ordre de travail", "Désignation", "Poste technique", "Poste de travail principal", "Statut système", "Statut utilisateur", "Date de début planifiée", "Type d'ordre", "Caractérisation backlog Préparation", "Caractérisation backlog Planification"])
                                            sub["KPI impacté"] = kpi; sub["Action recommandée"] = f"Corriger l'indicateur {kpi}."; sd = pd.concat([sd, sub])
                                    if kpi == "appel avis approuvé":
                                        sa = results['avis_df_filtered'][results['avis_df_filtered']["Poste travail princ."] == pe].copy()
                                        if not sa.empty:
                                            sa = rename_safe(sa, ["Avis", "Désignation texte", "Emplacement technique", "Poste travail princ.", "Statut utilisateur", "Créé le"], ["Avis", "Désignation", "Poste technique", "Poste de travail principal", "Statut", "Date de création"])
                                            sa["KPI impacté"] = kpi; sa["Action recommandée"] = "Créer un OT pour cet Avis."; sd = pd.concat([sd, sa])
                                    if not sd.empty: sd.to_excel(w, sheet_name=f"{pe.replace(' ', '_').replace('/', '_')[:20]}_{kpi.replace('/', '_').replace(' ', '_')[:10]}"[:31], index=False)
                        out.seek(0)
                        nf = "Plan_Action_Tous_Postes.xlsx" if sel_exp == "📌 Tous les postes" else f"Plan_Action_{sel_exp.replace(' ', '_')}.xlsx"
                        st.download_button(label="✅ Télécharger le fichier Excel", data=out.getvalue(), file_name=nf, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else: st.markdown("""<div class="empty-state" style="padding:30px;"><div class="icon">🎉</div><h3>Aucun export nécessaire</h3><p>Toutes les anomalies ont été résolues.</p></div>""", unsafe_allow_html=True)
            st.markdown(f"""<div style="text-align:center; padding:24px 0 8px 0; color:#a0aec0; font-size:11px;">KPI Dashboard MC & FEED • Maintenance Conditionnelle • {date_fichier}</div>""", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")

if __name__ == "__main__":
    try: locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try: locale.setlocale(locale.LC_ALL, 'fr_FR')
        except: pass
    main()
