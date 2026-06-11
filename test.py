# -*- coding: utf-8 -*-
"""app.py"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import locale
import random
import time
from datetime import datetime
import os

def inject_custom_css():
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--p:#1e3a5f;--pl:#2c5282;--b:#e2e8f0;--r:12px}
    .stApp{background:linear-gradient(135deg,#f0f4f8,#e2e8f0,#f0f4f8);font-family:'Inter',sans-serif}
    .mh{background:linear-gradient(135deg,var(--p),var(--pl));padding:18px 30px;border-radius:var(--r);margin-bottom:14px;box-shadow:0 8px 30px rgba(0,0,0,.1);position:relative;overflow:hidden}
    .mh h1{color:#fff;font-size:22px;font-weight:800;margin:0}
    .mh .sub{color:rgba(255,255,255,.7);font-size:11px;margin-top:2px}
    .mh .db{position:absolute;top:18px;right:30px;background:rgba(255,255,255,.15);backdrop-filter:blur(10px);padding:5px 12px;border-radius:20px;color:#fff;font-size:11px;font-weight:500;border:1px solid rgba(255,255,255,.2)}
    .stl{font-size:14px;font-weight:700;color:var(--p);margin-bottom:6px;padding-left:10px;border-left:3px solid var(--pl)}
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:9px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:8px;text-transform:uppercase;letter-spacing:.3px;padding:5px 4px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw tbody td{padding:3px 4px;border-bottom:1px solid var(--b);white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f9fbfd}
    .tw tbody tr:hover td{background:#edf2f7!important}
    .stTabs [data-baseweb="tab-list"]{gap:3px;background:#edf2f7;padding:3px;border-radius:8px;margin-bottom:10px}
    .stTabs [data-baseweb="tab"]{border-radius:6px;padding:7px 16px;font-weight:600;font-size:13px}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--p)!important;box-shadow:0 2px 6px rgba(0,0,0,.08)}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:8px;padding:9px 20px;font-weight:700;font-size:12px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1;border-radius:3px}::-webkit-scrollbar-thumb{background:#c1c1c1;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:8px;padding:3px 10px;margin-bottom:4px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:6px}
    .es{text-align:center;padding:24px;color:#718096}.es .ei{font-size:28px;margin-bottom:6px}
    @media(max-width:768px){.mh{padding:12px 14px}.mh h1{font-size:16px}.mh .db{position:static;margin-top:4px;display:inline-block}}
    </style>""", unsafe_allow_html=True)

def main():
    try: locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try: locale.setlocale(locale.LC_ALL, 'fr_FR')
        except: pass

    inject_custom_css()

    consignes = ["Port obligatoire des EPI avant toute intervention.","Port obligatoire du casque de sécurité.","Port obligatoire des lunettes de protection.","Port obligatoire des gants adaptés au travail.","Utiliser les protections auditives dans les zones bruyantes.","Vérifier l'absence de tension avant toute intervention électrique.","Respecter la procédure de consignation et déconsignation.","Ne jamais intervenir sur un équipement en marche.","Baliser et sécuriser la zone de travail.","Maintenir le poste de travail propre et ordonné.","Vérifier l'état des outils avant utilisation.","Utiliser uniquement du matériel homologué.","Respecter les permis de travail en vigueur.","Identifier les risques avant de commencer une tâche.","Signaler immédiatement toute situation dangereuse.","Signaler tout incident ou presque accident.","Ne jamais neutraliser un dispositif de sécurité.","Vérifier les détecteurs de gaz avant utilisation.","Vérifier la bonne ventilation des zones de travail.","Respecter les règles des espaces confinés.","Contrôler l'atmosphère avant d'entrer dans un espace confiné.","Utiliser les points d'ancrage pour les travaux en hauteur.","Vérifier l'état des échafaudages avant utilisation.","Sécuriser les outils lors des travaux en hauteur.","Ne pas travailler seul lors d'opérations à risque.","Contrôler les élingues avant chaque levage.","Respecter les limites de charge des équipements.","Vérifier l'état des appareils de levage.","Maintenir les voies de circulation dégagées.","Respecter la signalisation de sécurité.","Vérifier les extincteurs à proximité du chantier.","Connaître les issues de secours les plus proches.","Respecter les procédures d'arrêt d'urgence.","Vérifier les flexibles et raccords avant mise en service.","Contrôler les fuites avant démarrage d'un équipement.","Respecter les distances de sécurité.","Ne jamais contourner une procédure HSE.","Porter les EPI adaptés au risque identifié.","Prévenir son responsable avant toute intervention particulière.","Analyser les risques avant chaque démarrage de chantier.","Vérifier la stabilité des équipements.","Utiliser les bons outils pour la bonne tâche.","Respecter les consignes spécifiques du chantier.","Ne jamais prendre de raccourci au détriment de la sécurité.","Arrêter immédiatement les travaux en cas de danger.","Protéger l'environnement lors des interventions.","Collecter et trier correctement les déchets.","Éviter toute pollution accidentelle.","Respecter les consignes de stockage des produits dangereux.","Lire les fiches de sécurité avant manipulation.","Vérifier les équipements avant chaque prise de poste.","S'assurer de la disponibilité des moyens de secours.","Communiquer clairement avec l'équipe avant intervention.","Respecter les règles de circulation des engins.","Garder une vigilance permanente sur son environnement.","Prendre le temps d'effectuer le travail en sécurité.","La sécurité est l'affaire de tous.","Chaque incident peut être évité par la prévention.","Aucun travail n'est plus urgent que la sécurité.","Zéro accident commence par un comportement sûr."]

    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        consigne = random.choice(consignes)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:42px;color:#fff;font-weight:900;margin:0;letter-spacing:-1px">HSE - CONSIGNE DE SÉCURITÉ</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:18px;margin-top:8px;font-weight:400;letter-spacing:3px;text-transform:uppercase">Sécurité &bull; Santé &bull; Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:28px;font-weight:700;text-align:center;margin-top:40px;margin-bottom:40px;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:32px;font-weight:900;letter-spacing:-.5px">Aucun travail n'est plus urgent que la sécurité</h2>
        <div style="margin-top:40px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden"><div style="width:100%%;height:100%%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 5.5s ease-in-out forwards"></div></div>
        <style>@keyframes ld{from{width:0}to{width:100%%}}</style>
        </div>""" % consigne, unsafe_allow_html=True)
        time.sleep(6); st.session_state.hse_affiche = True; st.rerun(); st.stop()

    def contient_mot(t, lm):
        t = str(t); return any(m in t for l in lm for m in l.split())
    def categorie_age(a):
        if a <= 1: return "<1 mois"
        elif a >= 3: return ">3 mois"
        return "1 mois < <3 mois"
    def calcul_kpi(n, d, sz=100): return np.where(d == 0, sz, (n / d) * 100)
    def creer_pivot(df, f, c, p):
        return pd.pivot_table(df[f], index="Poste travail princ.", columns=c, values="Ordre", aggfunc="count", fill_value=0).reindex(p, fill_value=0)
    def exclure_cresseurs(df):
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)].copy() if "Poste travail princ." in df.columns else df

    def calculate_all_kpis(df_input, avis_input, now, postes):
        res = {}; df = df_input.copy(); av = avis_input.copy()
        mp = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
        df["Backlog préparation"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mp)), "CARACTERISE", "NON CARACTERISE")
        mplan = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
        df["Backlog planification"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mplan)), "CARACTERISE", "NON CARACTERISE")
        for dc, am, ac in [('Créé le',"Age mois préparation","Age préparation"),('Date de début planifiée',"Age mois planification","Age planification"),('Date de début planifiée',"Age mois exécution","Age exécution")]:
            if dc in df.columns:
                df[dc] = pd.to_datetime(df[dc], errors='coerce')
                df[am] = ((now.year - df[dc].dt.year)*12 + (now.month - df[dc].dt.month)).round(2)
                df[ac] = df[am].apply(categorie_age)
            else: df[am] = np.nan; df[ac] = "Inconnu"
        df["OT CONFIME"] = np.where(df["Statut système"].str.contains("CLO", na=False) & df["Statut système"].str.contains("CONF", na=False), "OUI", "NON")
        df["Contient SOPL"] = df["Statut utilisateur"].str.contains("SOPL", na=False).map({True:1, False:0})
        df["OT LANC ESTIME"] = np.where(df["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
        df["OT_COR_EGAL"] = np.where((df["Total coûts budgétés"].fillna(0) - df["Total coûts réels"].fillna(0)) == 0, "OUI", "NON")
        res['df_processed'] = df
        an = creer_pivot(df, df["Nº appel pl.entret."].fillna(0)==0, "Statut OT", postes)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c] = an.get(c, 0)
        an["Total"] = an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1); an["TAUX_REALISATION_CORRECTIF/PT"] = calcul_kpi(an["TCLO"], an["Total"])
        pr = creer_pivot(df, df["Statut OT"]=="CRÉÉ", "Age préparation", postes)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pr[c] = pr.get(c, 0)
        pr["Total"] = pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pr["OT préparation <1 mois"] = calcul_kpi(pr["<1 mois"], pr["Total"]); pr["OT préparation >3 mois"] = calcul_kpi(pr[">3 mois"], pr["Total"], 0); pr["OT préparation 1mois< <3mois"] = calcul_kpi(pr["1 mois < <3 mois"], pr["Total"], 0)
        pl = creer_pivot(df, (df["Statut OT"]=="LANC") & (df["Contient SOPL"]==0), "Age planification", postes)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pl[c] = pl.get(c, 0)
        pl["Total"] = pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pl["OT planification <1 mois"] = calcul_kpi(pl["<1 mois"], pl["Total"]); pl["OT planification >3 mois"] = calcul_kpi(pl[">3 mois"], pl["Total"], 0); pl["OT planification 1mois< <3mois"] = calcul_kpi(pl["1 mois < <3 mois"], pl["Total"], 0)
        ex = creer_pivot(df, (df["Statut OT"]=="LANC") & (df["Contient SOPL"]==1), "Age exécution", postes)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: ex[c] = ex.get(c, 0)
        ex["Total"] = ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        ex["OT exécution <1 mois"] = calcul_kpi(ex["<1 mois"], ex["Total"]); ex["OT exécution >3 mois"] = calcul_kpi(ex[">3 mois"], ex["Total"], 0); ex["OT exécution 1mois< <3mois"] = calcul_kpi(ex["1 mois < <3 mois"], ex["Total"], 0)
        la = pd.pivot_table(df[df["Statut OT"]=="LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["OUI","NON"]: la[c] = la.get(c, 0)
        la["Total"] = la["OUI"]+la["NON"]; la["OT LANC ESTIME"] = calcul_kpi(la["OUI"], la["Total"])
        pc = pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"], index="Poste travail princ.", columns="Backlog préparation", values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: pc[c] = pc.get(c, 0)
        pc["Total"] = pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"] = calcul_kpi(pc["CARACTERISE"], pc["Total"])
        plc = pd.pivot_table(df[df["Statut OT"]=="LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: plc[c] = plc.get(c, 0)
        plc["Total"] = plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"] = calcul_kpi(plc["CARACTERISE"], plc["Total"])
        for kn, cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
            pv = pd.pivot_table(df, index="Poste travail princ.", columns=cn, values="Ordre", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
            for c in ["OUI","NON"]: pv[c] = pv.get(c, 0)
            pv["Total"] = pv["OUI"]+pv["NON"]; pv[cn] = calcul_kpi(pv["OUI"], pv["Total"]); res[kn.lower().replace(" ","_")] = pv
        avf = av[(av["Ordre"].isna())|(av["Ordre"].astype(str).str.strip()=="")].copy(); res['avis_df_filtered'] = avf
        tca = pd.pivot_table(avf, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(postes, fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c] = tca.get(c, 0)
        tca["Total"] = tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["appel avis approuvé"] = calcul_kpi(tca["APRV"], tca["Total"])
        res['calculated_kpis_df'] = pd.concat([an[["TAUX_REALISATION_CORRECTIF/PT"]],pr[["OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois"]],pl[["OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois"]],ex[["OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois"]],tca[["appel avis approuvé"]],la[["OT LANC ESTIME"]],pc[["Backlog préparation caractérisé"]],plc[["Backlog planification caractérisé"]],res['ot_confime'][["OT CONFIME"]],res['ot_cor_egal'][["OT_COR_EGAL"]]], axis=1)
        return res

    def get_kpi_style(v, c):
        try: val = float(v)
        except: return ""
        if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val>=80 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=75 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val<=15 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val<=5 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c == "TAUX_REALISATION_CORRECTIF/PT":
            return "background:#c6efce;color:#006100;font-weight:600" if val>=85 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c == "appel avis approuvé":
            return "background:#c6efce;color:#006100;font-weight:600" if val>=95 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=90 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val>=100 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=95 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        return ""

    def get_anomalie_style(v):
        try: val = int(v)
        except: return ""
        if val == 0: return "color:#a0aec0"
        if val <= 3: return "background:#ffeb9c;color:#9c6500;font-weight:600"
        if val <= 10: return "background:#fed7d7;color:#c53030;font-weight:600"
        return "background:#fc8181;color:#742a2a;font-weight:800"

    def get_class_style(v):
        try: val = float(str(v).replace(' %','').strip())
        except: return ""
        return "background:#c6efce;color:#006100;font-weight:700" if val>=90 else ("background:#ffeb9c;color:#9c6500;font-weight:700" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:700")

    def df_to_html_kpi(df, tc="", cc=None):
        if cc is None: cc = []
        df = df.rename_axis("Poste de travail").reset_index()
        h = '<table class="tw %s"><thead><tr>' % tc + ''.join('<th>%s</th>' % c for c in df.columns) + '</tr></thead><tbody>'
        for _, r in df.iterrows():
            ic = r["Poste de travail"]=="CIBLE"; it = r["Poste de travail"]=="Total général"
            h += '<tr style="%s">' % ("border-top:2px solid #1e3a5f" if it else "")
            for c in df.columns:
                v = r[c]
                if ic:
                    h += '<td style="padding:5px 4px;background:#1e3a5f;color:#fff;font-weight:700;font-size:9px">%s</td>' % v
                elif it:
                    s = get_class_style(v) if c in cc else get_kpi_style(v, c)
                    s = s or "background:#e2e8f0;color:#1a202c"
                    h += '<td style="padding:5px 4px;%s;font-weight:800;font-size:9px">%s</td>' % (s, v)
                else:
                    s = get_class_style(v) if c in cc else get_kpi_style(v, c)
                    h += '<td style="padding:3px 4px;font-size:9px;%s">%s</td>' % (s, v)
            h += '</tr>'
        return h + '</tbody></table>'

    def df_to_html_anomalies(df):
        df = df.rename_axis("Poste de travail").reset_index()
        h = '<table class="tw at"><thead><tr>' + ''.join('<th>%s</th>' % c for c in df.columns) + '</tr></thead><tbody>'
        for _, r in df.iterrows():
            it = r["Poste de travail"]=="Total général"
            h += '<tr>'
            for c in df.columns:
                v = r[c]; s = get_anomalie_style(v) or ("" if not it else "font-weight:800")
                h += '<td style="padding:%spx 4px;font-size:9px;%s">%s</td>' % ("5" if it else "3", s, v)
            h += '</tr>'
        return h + '</tbody></table>'

    def get_kpi_score(k, a, t):
        if pd.isna(a) or pd.isna(t): return 0
        if k in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return 1 if a>=75 else 0
        if k in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return 1 if a<=15 else 0
        if k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return 1 if a<=5 else 0
        if k == "TAUX_REALISATION_CORRECTIF/PT": return 1 if a>=80 else 0
        if k == "appel avis approuvé": return 1 if a>=90 else 0
        if k in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]: return 1 if a>=95 else 0
        return 0

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""<div style="padding:14px 0 6px 0"><div style="font-size:22px;margin-bottom:2px">⚙️</div><div style="font-size:14px;font-weight:800;color:white">Filtres & Paramètres</div><div style="font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Configuration</div></div>""", unsafe_allow_html=True)
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
                    all_postes_master_list = sorted(_tmp[_tmp["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
                except: pass
            st.markdown("""<div style="background:rgba(255,255,255,.1);padding:8px 12px;border-radius:8px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Données</div><div style="font-size:12px;color:white;font-weight:600;margin-top:2px">📅 %s</div></div>""" % date_fichier, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("**🎯 Postes**")
        selected_postes = st.multiselect("Poste de travail", ["All"]+all_postes_master_list, ["All"], key="sel_postes")
        st.markdown("**🏭 Atelier**")
        selected_ateliers = st.multiselect("Atelier", ["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)"], ["All"], key="sel_ateliers")
        st.markdown("**🏢 Division**")
        selected_divisions = st.multiselect("Division", ["All","SF1","SF2"], ["All"], key="sel_div")
        st.markdown("---")
        st.markdown("**📅 Période**")
        date_range = st.date_input("Date début planifiée", value=(datetime(2025,1,1).date(), datetime.today().date()), format="DD/MM/YYYY", key="date_range")

    if not use_new_files or (ot_file is not None and avis_file is not None):
        try:
            if use_new_files:
                df_ot_raw = pd.read_excel(ot_file); avis_df_raw = pd.read_excel(avis_file); date_fichier = datetime.now().strftime("%d/%m/%Y")
            else:
                df_ot_raw = pd.read_excel("ot.xlsx"); avis_df_raw = pd.read_excel("avis.xlsx")
            df_ot_raw = exclure_cresseurs(df_ot_raw); avis_df_raw = exclure_cresseurs(avis_df_raw)
            for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
                if c in df_ot_raw.columns: df_ot_raw[c] = pd.to_datetime(df_ot_raw[c], errors="coerce")
            for c in ["Créé le","Début souhaité","Date de la clôture"]:
                if c in avis_df_raw.columns: avis_df_raw[c] = pd.to_datetime(avis_df_raw[c], errors="coerce")
            if not all_postes_master_list: all_postes_master_list = sorted(df_ot_raw[df_ot_raw["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
            if "All" in selected_postes or not selected_postes: selected_postes = all_postes_master_list
            if "All" in selected_ateliers or not selected_ateliers: selected_ateliers = ["All"]
            if "All" in selected_divisions or not selected_divisions: selected_divisions = ["All"]
            start_date = pd.to_datetime(date_range[0]) if len(date_range)==2 else pd.to_datetime(datetime(2025,1,1))
            end_date = pd.to_datetime(date_range[1]) if len(date_range)==2 else pd.to_datetime(datetime.today())

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
            df = exclure_cresseurs(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)].drop_duplicates())
            avis_df = exclure_cresseurs(avis_df[(avis_df["Ordre"].isna())|(avis_df["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates())
            if "Statut système" in df.columns: df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            now = pd.Timestamp.now()
            results = calculate_all_kpis(df, avis_df, now, valid_postes)
            calculated_kpis_df = results['calculated_kpis_df']; df_processed = results['df_processed']

            # ============ SÉPARATION ============
            qty_k = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois"]
            perf_k = ["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]

            cible = pd.DataFrame([{"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,"OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,"OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,"OT exécution 1mois< <3mois":15,"appel avis approuvé":95,"OT LANC ESTIME":100,"Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,"OT CONFIME":100,"OT_COR_EGAL":100}], index=["CIBLE"])

            calc_qty = calculated_kpis_df[qty_k].copy()
            calc_perf = calculated_kpis_df[perf_k].copy()
            tg_qty = pd.DataFrame(calc_qty.mean()).T; tg_qty.index = ["Total général"]
            tg_perf = pd.DataFrame(calc_perf.mean()).T; tg_perf.index = ["Total général"]
            final_kpi_qty = pd.concat([cible[qty_k], calc_qty, tg_qty]).round(2)
            final_kpi_perf = pd.concat([cible[perf_k], calc_perf, tg_perf]).round(2)

            # ============ SCORES POUR FUSION ============
            score_map = {}
            for poste in calculated_kpis_df.index:
                r = calculated_kpis_df.loc[poste]
                sl = sum(get_kpi_score(k, r[k], cible.loc['CIBLE', k]) for k in perf_k if k in r.index)
                score_map[poste] = f"{(sl/len(perf_k)*100) if perf_k else 0:.2f} %"
            tg_score = f"{np.mean([float(v.replace(' %','')) for v in score_map.values()]):.2f} %"

            merged_perf = final_kpi_perf.copy()
            merged_perf["Score Qualité"] = merged_perf.index.map(lambda x: "100.00 %" if x=="CIBLE" else (tg_score if x=="Total général" else score_map.get(x, "")))

            # ============ ANOMALIES ============
            ano_ot = []
            for poste in valid_postes:
                if poste not in df_processed["Poste travail princ."].values: continue
                dp = df_processed[df_processed["Poste travail princ."]==poste]
                for kn, cc, sub in [
                    ("TAUX_REALISATION_CORRECTIF/PT","TAUX_REALISATION_CORRECTIF/PT",dp[(dp["Nº appel pl.entret."].fillna(0)==0)&(~dp["Statut OT"].isin(["CLOT","TCLO"]))]),
                    ("OT préparation <1 mois","OT préparation <1 mois",dp[(dp["Statut OT"]=="CRÉÉ")&(dp["Age préparation"]!="<1 mois")]),
                    ("OT préparation >3 mois","OT préparation >3 mois",dp[(dp["Statut OT"]=="CRÉÉ")&(dp["Age préparation"]==">3 mois")]),
                    ("OT planification <1 mois","OT planification <1 mois",dp[(dp["Statut OT"]=="LANC")&(dp["Contient SOPL"]==0)&(dp["Age planification"]!="<1 mois")]),
                    ("OT planification >3 mois","OT planification >3 mois",dp[(dp["Statut OT"]=="LANC")&(dp["Contient SOPL"]==0)&(dp["Age planification"]==">3 mois")]),
                    ("OT exécution <1 mois","OT exécution <1 mois",dp[(dp["Statut OT"]=="LANC")&(dp["Contient SOPL"]==1)&(dp["Age exécution"]!="<1 mois")]),
                    ("OT exécution >3 mois","OT exécution >3 mois",dp[(dp["Statut OT"]=="LANC")&(dp["Contient SOPL"]==1)&(dp["Age exécution"]==">3 mois")])]:
                    vk = calculated_kpis_df.loc[poste, kn] if poste in calculated_kpis_df.index else 100
                    if pd.notna(vk) and vk < cible.loc['CIBLE', cc]:
                        cnt = len(sub)
                        if cnt > 0: ano_ot.append({"Poste travail princ.": poste, "KPI": kn, "Nb OT impactés": cnt})

            df_ano_ot = pd.DataFrame(ano_ot)
            ano_dash = pd.DataFrame()
            if not df_ano_ot.empty:
                ano_dash = df_ano_ot.pivot_table(index="Poste travail princ.", columns="KPI", values="Nb OT impactés", aggfunc="sum", fill_value=0).astype(int)
                ano_dash["Total éléments impactés"] = ano_dash.sum(axis=1)
                ano_dash.loc["Total général"] = ano_dash.sum()

            # ============ HEADER ============
            st.markdown('<div class="mh"><h1>📊 KPI Dashboard MC & FEED</h1><div class="sub">Maintenance Conditionnelle • Suivi des indicateurs</div><div class="db">📅 %s</div></div>' % date_fichier, unsafe_allow_html=True)

            # ============ ONGLETS ============
            tab1, tab2 = st.tabs(["📈 INDICATEURS DE PERFORMANCE", "✅ INDICATEUR QUALITÉ"])

            # --- ONGLET 1 : PERFORMANCE (anciennement Quantité) ---
            with tab1:
                c_t, c_b = st.columns([4, 1])
                with c_t:
                    st.markdown('<p class="stl q" style="margin-bottom:0">Indicateurs de Performance par Poste de Travail</p>', unsafe_allow_html=True)
                with c_b:
                    view = st.radio("", ["Tableau KPI", "Anomalies"], horizontal=True, key="vp", label_visibility="collapsed")
                st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
                if view == "Tableau KPI":
                    st.markdown(df_to_html_kpi(final_kpi_qty, "qt"), unsafe_allow_html=True)
                else:
                    if not ano_dash.empty and ano_dash.drop(columns="Total éléments impactés", errors='ignore').sum().sum() > 0:
                        st.markdown(df_to_html_anomalies(ano_dash), unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="es"><div class="ei">✅</div><b>Aucune anomalie détectée</b><br><span style="font-size:11px">Tous les KPIs atteignent leurs cibles.</span></div>', unsafe_allow_html=True)

            # --- ONGLET 2 : QUALITÉ (anciennement Performance) ---
            with tab2:
                st.markdown('<p class="stl p">Indicateur Qualité par Poste de Travail</p>', unsafe_allow_html=True)
                st.markdown(df_to_html_kpi(merged_perf, "pt", cc=["Score Qualité"]), unsafe_allow_html=True)

            # ============ EXPORT ============
            st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
            st.markdown('<p class="stl">💾 Export des Plans d\'Action</p>', unsafe_allow_html=True)
            postes_ano = ano_dash[ano_dash.index!="Total général"].index.tolist() if not ano_dash.empty else []
            if postes_ano:
                ce1, ce2 = st.columns([1, 1])
                with ce1: sel_exp = st.selectbox("Poste de travail :", options=["📌 Tous les postes"]+postes_ano, key="sel_exp")
                with ce2:
                    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                    gen_btn = st.button("📥 Générer le fichier Excel", type="primary", key="btn_exp", use_container_width=True)
                if gen_btn:
                    with st.spinner("Génération en cours..."):
                        out = io.BytesIO()
                        with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                            for pe in (postes_ano if sel_exp=="📌 Tous les postes" else [sel_exp]):
                                kds = df_ano_ot[df_ano_ot["Poste travail princ."]==pe]["KPI"].unique().tolist()
                                for kpi in kds:
                                    dpf = df_processed[df_processed["Poste travail princ."]==pe].copy()
                                    cmap = {"TAUX_REALISATION_CORRECTIF/PT":(dpf["Nº appel pl.entret."].fillna(0)==0)&(~dpf["Statut OT"].isin(["CLOT","TCLO"])),"OT préparation <1 mois":(dpf["Statut OT"]=="CRÉÉ")&(dpf["Age préparation"]!="<1 mois"),"OT préparation >3 mois":(dpf["Statut OT"]=="CRÉÉ")&(dpf["Age préparation"]==">3 mois"),"OT planification <1 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==0)&(dpf["Age planification"]!="<1 mois"),"OT planification >3 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==0)&(dpf["Age planification"]==">3 mois"),"OT exécution <1 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==1)&(dpf["Age exécution"]!="<1 mois"),"OT exécution >3 mois":(dpf["Statut OT"]=="LANC")&(dpf["Contient SOPL"]==1)&(dpf["Age exécution"]==">3 mois")}
                                    if kpi in cmap:
                                        sd = dpf[cmap[kpi]][["Ordre","Poste travail princ.","Statut OT","Statut utilisateur","Créé le","Date de début planifiée","Total coûts budgétés","Total coûts réels"]].copy()
                                        if not sd.empty: sd.to_excel(w, sheet_name=("%s_%s" % (pe[:15], kpi[:15]))[:31], index=False)
                            final_kpi_qty.rename_axis("Poste de travail").reset_index().to_excel(w, sheet_name="KPIs Performance", index=False)
                            merged_perf.rename_axis("Poste de travail").reset_index().to_excel(w, sheet_name="KPIs Qualité", index=False)
                        out.seek(0)
                        st.download_button(label="⬇️ Télécharger le fichier Excel", data=out, file_name="Plan_Action_%s.xlsx" % datetime.now().strftime('%Y%m%d_%H%M'), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        st.success("✅ Fichier généré avec succès !")
            else:
                st.markdown('<div class="es" style="padding:20px"><div class="ei">🎉</div><b>Aucun plan d\'action à exporter</b><br><span style="font-size:11px">Tous les indicateurs sont conformes.</span></div>', unsafe_allow_html=True)

        except Exception as e:
            st.error("❌ Erreur : %s" % str(e))

if __name__ == "__main__":
    main()
