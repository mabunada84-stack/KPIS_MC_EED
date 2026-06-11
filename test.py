# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os
from datetime import datetime

def inject_custom_css():
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--p:#1e3a5f;--pl:#2c5282;--b:#e2e8f0;--r:10px}
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#edf2f7;font-family:'Inter',sans-serif}
    .main .block-container{padding-top:.8rem;padding-bottom:.8rem;max-width:100%!important;padding-left:1rem;padding-right:1rem}
    .mh{background:linear-gradient(135deg,var(--p),var(--pl));padding:12px 20px;border-radius:var(--r);margin-bottom:6px;box-shadow:0 6px 20px rgba(0,0,0,.1);overflow:hidden}
    .mh h1{color:#fff;font-size:18px;font-weight:800;margin:0;display:inline}
    .mh .sub{color:rgba(255,255,255,.7);font-size:10px;margin-top:1px}
    .mh .db{float:right;background:rgba(255,255,255,.15);padding:3px 10px;border-radius:16px;color:#fff;font-size:10px;font-weight:500;border:1px solid rgba(255,255,255,.2);margin-top:2px}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:6px}
    .cc{background:#fff;border-radius:var(--r);padding:10px 14px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid var(--b);text-align:center}
    .cc .cv{font-size:24px;font-weight:900;line-height:1}
    .cc .cl{font-size:8px;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
    .cc.c1{border-top:3px solid #3182ce}.cc.c1 .cv{color:#2b6cb0}
    .cc.c2{border-top:3px solid #38a169}.cc.c2 .cv{color:#276749}
    .cc.c3{border-top:3px solid #805ad5}.cc.c3 .cv{color:#6b46c1}
    .cc.c4{border-top:3px solid #e53e3e}.cc.c4 .cv{color:#c53030}
    .stl{font-size:12px;font-weight:700;color:var(--p);margin:5px 0 2px 0;padding-left:8px;border-left:3px solid var(--pl)}
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:8px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:7px;text-transform:uppercase;letter-spacing:.3px;padding:4px 3px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw tbody td{padding:2px 3px;border-bottom:1px solid #edf2f7;white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important;font-size:8px!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important;font-size:8px!important}
    .stTabs [data-baseweb="tab-list"]{gap:2px;background:#e2e8f0;padding:2px;border-radius:6px;margin-bottom:5px}
    .stTabs [data-baseweb="tab"]{border-radius:5px;padding:5px 12px;font-weight:600;font-size:11px}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--p)!important;box-shadow:0 2px 5px rgba(0,0,0,.07)}
    .sr{display:flex;align-items:center;padding:5px 10px;background:#fff;border-radius:6px;margin-bottom:2px;border:1px solid var(--b);font-size:10px;transition:all .15s}
    .sr:hover{transform:translateX(2px);box-shadow:0 2px 8px rgba(0,0,0,.05);border-left:3px solid var(--pl)}
    .sr .sn{font-weight:700;color:var(--p);min-width:200px;font-size:9px}
    .sr .sc{background:var(--p);color:#fff;padding:2px 8px;border-radius:14px;font-weight:800;font-size:11px;min-width:36px;text-align:center;margin:0 8px}
    .sr .sa{color:#718096;font-size:9px;flex:1}
    .sr .sb{display:inline-block;padding:1px 6px;border-radius:3px;font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.3px}
    .sb.r{background:#fed7d7;color:#c53030}.sb.o{background:#feebc8;color:#c05621}.sb.y{background:#fefcbf;color:#975a16}
    .ca{background:#fff;border-radius:var(--r);padding:10px;margin-top:4px;border:1px solid var(--b);box-shadow:0 2px 6px rgba(0,0,0,.03)}
    .car{display:flex;align-items:center;margin-bottom:4px;font-size:9px}
    .car .cal{width:180px;font-weight:600;color:var(--p);text-align:right;padding-right:8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .car .cab{flex:1;height:16px;background:#edf2f7;border-radius:3px;overflow:hidden;position:relative}
    .car .caf{height:100%;border-radius:3px;display:flex;align-items:center;justify-content:flex-end;padding-right:4px}
    .car .cav{font-size:8px;font-weight:800;color:#fff;text-shadow:0 1px 1px rgba(0,0,0,.2)}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:7px 14px;font-weight:700;font-size:11px;width:100%}
    ::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:2px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:9px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:2px 6px;margin-bottom:2px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:12px;color:#718096;font-size:11px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:14px}.mh .db{float:none;display:block;margin-top:3px}.sr .sn{min-width:120px}}
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
        c = random.choice(consignes)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:42px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SÉCURITÉ</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:18px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Sécurité • Santé • Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:28px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:32px;font-weight:900">Aucun travail n'est plus urgent que la sécurité</h2>
        <div style="margin-top:40px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden"><div style="width:100%%;height:100%%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 5.5s ease-in-out forwards"></div></div>
        <style>@keyframes ld{from{width:0}to{width:100%%}}</style></div>""" % c, unsafe_allow_html=True)
        time.sleep(6); st.session_state.hse_affiche = True; st.rerun(); st.stop()

    def contient_mot(t, lm):
        t = str(t); return any(m in t for l in lm for m in l.split())
    def cat_age(a):
        if a <= 1: return "<1 mois"
        elif a >= 3: return ">3 mois"
        return "1 mois < <3 mois"
    def ckpi(n, d, sz=100): return np.where(d == 0, sz, (n / d) * 100)
    def cpiv(df, f, c, p):
        return pd.pivot_table(df[f], index="Poste travail princ.", columns=c, values="Ordre", aggfunc="count", fill_value=0).reindex(p, fill_value=0)
    def excr(df):
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)].copy() if "Poste travail princ." in df.columns else df

    def calc_kpis(df_i, av_i, now, posts):
        res = {}; df = df_i.copy(); av = av_i.copy()
        mp = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
        df["Backlog préparation"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mp)), "CARACTERISE", "NON CARACTERISE")
        mplan = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
        df["Backlog planification"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mplan)), "CARACTERISE", "NON CARACTERISE")
        for dc, am, ac in [('Créé le',"amp","ap"),('Date de début planifiée',"amlp","alp"),('Date de début planifiée',"amex","aex")]:
            if dc in df.columns:
                df[dc] = pd.to_datetime(df[dc], errors='coerce')
                df[am] = ((now.year - df[dc].dt.year)*12 + (now.month - df[dc].dt.month)).round(2)
                df[ac] = df[am].apply(cat_age)
            else: df[am] = np.nan; df[ac] = "Inconnu"
        df["OT CONFIME"] = np.where(df["Statut système"].str.contains("CLO", na=False) & df["Statut système"].str.contains("CONF", na=False), "OUI", "NON")
        df["Contient SOPL"] = df["Statut utilisateur"].str.contains("SOPL", na=False).map({True:1, False:0})
        df["OT LANC ESTIME"] = np.where(df["Total coûts budgétés"].fillna(0) == 0, "NON", "OUI")
        df["OT_COR_EGAL"] = np.where((df["Total coûts budgétés"].fillna(0) - df["Total coûts réels"].fillna(0)) == 0, "OUI", "NON")
        res['dfp'] = df
        an = cpiv(df, df["Nº appel pl.entret."].fillna(0)==0, "Statut OT", posts)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c] = an.get(c, 0)
        an["Total"] = an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1); an["TAUX_REALISATION_CORRECTIF/PT"] = ckpi(an["TCLO"], an["Total"])
        pr = cpiv(df, df["Statut OT"]=="CRÉÉ", "ap", posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pr[c] = pr.get(c, 0)
        pr["Total"] = pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pr["OT préparation <1 mois"] = ckpi(pr["<1 mois"], pr["Total"]); pr["OT préparation >3 mois"] = ckpi(pr[">3 mois"], pr["Total"], 0); pr["OT préparation 1mois< <3mois"] = ckpi(pr["1 mois < <3 mois"], pr["Total"], 0)
        pl = cpiv(df, (df["Statut OT"]=="LANC") & (df["Contient SOPL"]==0), "alp", posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pl[c] = pl.get(c, 0)
        pl["Total"] = pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pl["OT planification <1 mois"] = ckpi(pl["<1 mois"], pl["Total"]); pl["OT planification >3 mois"] = ckpi(pl[">3 mois"], pl["Total"], 0); pl["OT planification 1mois< <3mois"] = ckpi(pl["1 mois < <3 mois"], pl["Total"], 0)
        ex = cpiv(df, (df["Statut OT"]=="LANC") & (df["Contient SOPL"]==1), "aex", posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: ex[c] = ex.get(c, 0)
        ex["Total"] = ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        ex["OT exécution <1 mois"] = ckpi(ex["<1 mois"], ex["Total"]); ex["OT exécution >3 mois"] = ckpi(ex[">3 mois"], ex["Total"], 0); ex["OT exécution 1mois< <3mois"] = ckpi(ex["1 mois < <3 mois"], ex["Total"], 0)
        la = pd.pivot_table(df[df["Statut OT"]=="LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["OUI","NON"]: la[c] = la.get(c, 0)
        la["Total"] = la["OUI"]+la["NON"]; la["OT LANC ESTIME"] = ckpi(la["OUI"], la["Total"])
        pc = pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"], index="Poste travail princ.", columns="Backlog préparation", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: pc[c] = pc.get(c, 0)
        pc["Total"] = pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"] = ckpi(pc["CARACTERISE"], pc["Total"])
        plc = pd.pivot_table(df[df["Statut OT"]=="LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: plc[c] = plc.get(c, 0)
        plc["Total"] = plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"] = ckpi(plc["CARACTERISE"], plc["Total"])
        for kn, cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
            pv = pd.pivot_table(df, index="Poste travail princ.", columns=cn, values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
            for c in ["OUI","NON"]: pv[c] = pv.get(c, 0)
            pv["Total"] = pv["OUI"]+pv["NON"]; pv[cn] = ckpi(pv["OUI"], pv["Total"]); res[kn.lower().replace(" ","_")] = pv
        avf = av[(av["Ordre"].isna())|(av["Ordre"].astype(str).str.strip()=="")].copy(); res['avf'] = avf
        tca = pd.pivot_table(avf, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c] = tca.get(c, 0)
        tca["Total"] = tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["appel avis approuvé"] = ckpi(tca["APRV"], tca["Total"])
        res['ckdf'] = pd.concat([an[["TAUX_REALISATION_CORRECTIF/PT"]],pr[["OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois"]],pl[["OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois"]],ex[["OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois"]],tca[["appel avis approuvé"]],la[["OT LANC ESTIME"]],pc[["Backlog préparation caractérisé"]],plc[["Backlog planification caractérisé"]],res['ot_confime'][["OT CONFIME"]],res['ot_cor_egal'][["OT_COR_EGAL"]]], axis=1)
        return res

    def ks(v, c):
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

    def cs(v):
        try: val = float(str(v).replace(' %','').strip())
        except: return ""
        return "background:#c6efce;color:#006100;font-weight:700" if val>=90 else ("background:#ffeb9c;color:#9c6500;font-weight:700" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:700")

    def kas(v):
        try: val = int(v)
        except: return ""
        if val == 0: return "color:#cbd5e0"
        if val <= 3: return "background:#ffeb9c;color:#9c6500;font-weight:600"
        if val <= 10: return "background:#fed7d7;color:#c53030;font-weight:600"
        return "background:#fc8181;color:#742a2a;font-weight:800"

    def gscore(k, a, t):
        if pd.isna(a) or pd.isna(t): return 0
        if k in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return 1 if a>=75 else 0
        if k in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return 1 if a<=15 else 0
        if k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return 1 if a<=5 else 0
        if k == "TAUX_REALISATION_CORRECTIF/PT": return 1 if a>=80 else 0
        if k == "appel avis approuvé": return 1 if a>=90 else 0
        if k in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]: return 1 if a>=95 else 0
        return 0

    def html_table(rows, cols, tc, sc_col=None):
        h = '<table class="tw %s"><thead><tr>' % tc + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for r in rows:
            rc = "cb" if r.get("_t") == "cible" else ("tr" if r.get("_t") == "total" else "")
            h += '<tr class="%s">' % rc
            for c in cols:
                v = r.get(c, "")
                if r.get("_t") == "cible":
                    h += '<td>%s</td>' % v
                elif r.get("_t") == "total":
                    s = cs(v) if sc_col and c in sc_col else ks(v, c)
                    s = s or ""
                    h += '<td style="%s">%s</td>' % (s, v)
                else:
                    s = cs(v) if sc_col and c in sc_col else ks(v, c)
                    h += '<td style="%s">%s</td>' % (s, v)
            h += '</tr>'
        return h + '</tbody></table>'

    def html_ano(rows, cols):
        h = '<table class="tw at"><thead><tr>' + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for r in rows:
            rc = "tr" if r.get("_t") == "total" else ""
            h += '<tr class="%s">' % rc
            for c in cols:
                v = r.get(c, ""); s = kas(v) or ""
                h += '<td style="%s">%s</td>' % (s, v)
            h += '</tr>'
        return h + '</tbody></table>'

    def html_synth(data, count_col="Nb"):
        if not data: return '<div class="es"><b>Aucune action requise</b><br>Tous les indicateurs sont conformes.</div>'
        h = ''
        for d in data:
            nb = d[count_col]; bg = "r" if nb>=20 else ("o" if nb>=5 else "y"); lb = "CRITIQUE" if nb>=20 else ("MOYEN" if nb>=5 else "MINEUR")
            h += '<div class="sr"><div class="sn">%s</div><div class="sc">%s</div><div class="sa">%s</div><span class="sb %s">%s</span></div>' % (d["KPI"], nb, d["Action"], bg, lb)
        return h

    def html_cat(data, max_val, color):
        if not data: return ''
        h = '<div class="ca">'
        for label, val in data:
            pct = max(5, (val / max_val * 100)) if max_val > 0 else 0
            h += '<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"><span class="cav">%s</span></div></div></div>' % (label, min(pct, 100), color, val)
        h += '</div>'
        return h

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""<div style="padding:12px 0 4px 0"><div style="font-size:20px;margin-bottom:2px">⚙️</div><div style="font-size:13px;font-weight:800;color:white">Filtres & Paramètres</div><div style="font-size:8px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Configuration</div></div>""", unsafe_allow_html=True)
        st.markdown("---")
        unf = st.toggle("📁 Charger nouveaux fichiers", value=False, key="tf")
        ot_f = av_f = None; apm = []
        if unf:
            ot_f = st.file_uploader("Fichier OT", type=["xlsx"], key="uot")
            av_f = st.file_uploader("Fichier AVIS", type=["xlsx"], key="uav")
        else:
            df_dt = datetime.now().strftime("%d/%m/%Y")
            if os.path.exists("ot.xlsx"):
                try:
                    df_dt = datetime.fromtimestamp(os.path.getmtime("ot.xlsx")).strftime("%d/%m/%Y")
                    _t = excr(pd.read_excel("ot.xlsx"))
                    apm = sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
                except: pass
            st.markdown("""<div style="background:rgba(255,255,255,.1);padding:6px 10px;border-radius:6px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:8px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Données</div><div style="font-size:11px;color:white;font-weight:600;margin-top:1px">📅 %s</div></div>""" % df_dt, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("**🎯 Postes**")
        sp = st.multiselect("Poste", ["All"]+apm, ["All"], key="sp")
        st.markdown("**🏭 Atelier**")
        sa = st.multiselect("Atelier", ["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)"], ["All"], key="sa")
        st.markdown("**🏢 Division**")
        sd = st.multiselect("Division", ["All","SF1","SF2"], ["All"], key="sd")
        st.markdown("---")
        st.markdown("**📅 Période**")
        dr = st.date_input("Date début planifiée", value=(datetime(2025,1,1).date(), datetime.today().date()), format="DD/MM/YYYY", key="dr")

    if not unf or (ot_f is not None and av_f is not None):
        try:
            if unf:
                raw_ot = pd.read_excel(ot_f); raw_av = pd.read_excel(av_f); df_dt = datetime.now().strftime("%d/%m/%Y")
            else:
                raw_ot = pd.read_excel("ot.xlsx"); raw_av = pd.read_excel("avis.xlsx")
            raw_ot = excr(raw_ot); raw_av = excr(raw_av)
            for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
                if c in raw_ot.columns: raw_ot[c] = pd.to_datetime(raw_ot[c], errors="coerce")
            for c in ["Créé le","Début souhaité","Date de la clôture"]:
                if c in raw_av.columns: raw_av[c] = pd.to_datetime(raw_av[c], errors="coerce")
            if not apm: apm = sorted(raw_ot[raw_ot["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)]["Poste travail princ."].dropna().unique().tolist())
            if "All" in sp or not sp: sp = apm
            if "All" in sa or not sa: sa = ["All"]
            if "All" in sd or not sd: sd = ["All"]
            sdt = pd.to_datetime(dr[0]) if len(dr)==2 else pd.to_datetime(datetime(2025,1,1))
            edt = pd.to_datetime(dr[1]) if len(dr)==2 else pd.to_datetime(datetime.today())

            def mf(poste):
                p = str(poste).upper()
                if "All" not in sa:
                    m = False
                    if "Sulfurique (PS)" in sa and "PS" in p: m = True
                    if "Phosphorique (PP)" in sa and "PP" in p: m = True
                    if "Engrais (TSP/REX)" in sa and ("TSP" in p or "REX" in p): m = True
                    if "Feed (MCP/DCP)" in sa and ("MCP" in p or "DCP" in p): m = True
                    if not m: return False
                if "All" not in sd:
                    m = False
                    if "SF1" in sd and "SF1" in p: m = True
                    if "SF2" in sd and "SF2" in p: m = True
                    if not m: return False
                return True

            vp = [p for p in apm if mf(p) and p in sp]
            df = raw_ot[(raw_ot["Poste travail princ."].isin(vp)) & (raw_ot["Date de début planifiée"].between(sdt, edt))].copy()
            avdf = raw_av[raw_av["Poste travail princ."].isin(vp)].copy()
            df = excr(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)].drop_duplicates())
            avdf = excr(avdf[(avdf["Ordre"].isna())|(avdf["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates())
            if "Statut système" in df.columns: df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            now = pd.Timestamp.now()
            res = calc_kpis(df, avdf, now, vp)
            ckdf = res['ckdf']; dfp = res['dfp']

            # ============ SÉPARATION ============
            qk = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois"]
            pk = ["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]
            cible = {"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,"OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,"OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,"OT exécution 1mois< <3mois":15,"appel avis approuvé":95,"OT LANC ESTIME":100,"Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,"OT CONFIME":100,"OT_COR_EGAL":100}

            # ============ SCORES ============
            pscores = {}; qscores = {}
            for poste in ckdf.index:
                r = ckdf.loc[poste]
                sq = sum(gscore(k, r[k], cible[k]) for k in qk if k in r.index)
                sl = sum(gscore(k, r[k], cible[k]) for k in pk if k in r.index)
                pscores[poste] = (sq/len(qk)*100) if qk else 0
                qscores[poste] = (sl/len(pk)*100) if pk else 0

            # ============ ANOMALIES ============
            all_ano = []
            act_map = {
                "TAUX_REALISATION_CORRECTIF/PT": ("Améliorer le taux de réalisation des OT.", "P"),
                "OT préparation <1 mois": ("Réduire l'âge de préparation des OT (< 1 mois).", "P"),
                "OT préparation >3 mois": ("Traiter les OT avec préparation > 3 mois.", "P"),
                "OT planification <1 mois": ("Réduire l'âge de planification des OT (< 1 mois).", "P"),
                "OT planification >3 mois": ("Traiter les OT avec planification > 3 mois.", "P"),
                "OT exécution <1 mois": ("Réduire l'âge d'exécution des OT (< 1 mois).", "P"),
                "OT exécution >3 mois": ("Traiter les OT avec exécution > 3 mois.", "P"),
                "OT LANC ESTIME": ("Estimer les coûts des OT lancés.", "Q"),
                "Backlog préparation caractérisé": ("Caractériser le backlog de préparation.", "Q"),
                "Backlog planification caractérisé": ("Caractériser le backlog de planification.", "Q"),
                "OT CONFIME": ("Confirmer les OT terminés.", "Q"),
                "OT_COR_EGAL": ("Rapprocher les coûts réels et budgétés.", "Q"),
                "appel avis approuvé": ("Créer un OT pour les avis sans ordre.", "Q")
            }
            sub_map_p = {
                "TAUX_REALISATION_CORRECTIF/PT": lambda d: d[(d["Nº appel pl.entret."].fillna(0)==0)&(~d["Statut OT"].isin(["CLOT","TCLO"]))],
                "OT préparation <1 mois": lambda d: d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]!="<1 mois")],
                "OT préparation >3 mois": lambda d: d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]==">3 mois")],
                "OT planification <1 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]!="<1 mois")],
                "OT planification >3 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]==">3 mois")],
                "OT exécution <1 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]!="<1 mois")],
                "OT exécution >3 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]==">3 mois")],
            }
            sub_map_q = {
                "OT LANC ESTIME": lambda d: d[(d["Statut OT"]=="LANC")&(d["OT LANC ESTIME"]=="NON")],
                "Backlog préparation caractérisé": lambda d: d[(d["Statut OT"]=="CRÉÉ")&(d["Backlog préparation"]=="NON CARACTERISE")],
                "Backlog planification caractérisé": lambda d: d[(d["Statut OT"]=="LANC")&(d["Backlog planification"]=="NON CARACTERISE")],
                "OT CONFIME": lambda d: d[d["OT CONFIME"]=="NON"],
                "OT_COR_EGAL": lambda d: d[d["OT_COR_EGAL"]=="NON"],
            }

            for poste in vp:
                if poste not in dfp["Poste travail princ."].values: continue
                dp = dfp[dfp["Poste travail princ."]==poste]
                ap = res['avf'][res['avf']["Poste travail princ."]==poste]
                for kn, sf in sub_map_p.items():
                    vk = ckdf.loc[poste, kn] if poste in ckdf.index else 100
                    if pd.notna(vk) and vk < cible[kn]:
                        cnt = len(sf(dp))
                        if cnt > 0: all_ano.append({"Poste": poste, "KPI": kn, "Nb": cnt, "Action": act_map[kn][0], "Type": "P"})
                for kn, sf in sub_map_q.items():
                    vk = ckdf.loc[poste, kn] if poste in ckdf.index else 100
                    if pd.notna(vk) and vk < cible[kn]:
                        cnt = len(sf(dp))
                        if cnt > 0: all_ano.append({"Poste": poste, "KPI": kn, "Nb": cnt, "Action": act_map[kn][0], "Type": "Q"})
                vk_av = ckdf.loc[poste, "appel avis approuvé"] if poste in ckdf.index else 100
                if pd.notna(vk_av) and vk_av < cible["appel avis approuvé"]:
                    cnt = len(ap)
                    if cnt > 0: all_ano.append({"Poste": poste, "KPI": "appel avis approuvé", "Nb": cnt, "Action": act_map["appel avis approuvé"][0], "Type": "Q"})

            # Synthèse par type
            synth_p = sorted([{"KPI": k, "Nb": sum(a["Nb"] for a in all_ano if a["KPI"]==k), "Action": act_map[k][0]} for k in qk if any(a["KPI"]==k for a in all_ano)], key=lambda x: x["Nb"], reverse=True)
            synth_q = sorted([{"KPI": k, "Nb": sum(a["Nb"] for a in all_ano if a["KPI"]==k), "Action": act_map[k][0]} for k in pk if any(a["KPI"]==k for a in all_ano)], key=lambda x: x["Nb"], reverse=True)

            # Anomaly dashboards par type
            ano_p_df = pd.DataFrame([a for a in all_ano if a["Type"]=="P"])
            ano_q_df = pd.DataFrame([a for a in all_ano if a["Type"]=="Q"])

            def build_ano_dash(ano_df, kpi_list):
                if ano_df.empty: return [], [], []
                pv = ano_df.pivot_table(index="Poste", columns="KPI", values="Nb", aggfunc="sum", fill_value=0).astype(int)
                pv["Total"] = pv.sum(axis=1)
                total_row = pv.sum()
                cols = [c for c in kpi_list if c in pv.columns] + (["Total"] if "Total" in pv.columns else [])
                rows = []
                for idx in pv.index:
                    r = {"_t": "n", "Poste de travail": idx}
                    for c in cols: r[c] = pv.loc[idx, c]
                    rows.append(r)
                tr = {"_t": "total", "Poste de travail": "Total général"}
                for c in cols: tr[c] = int(total_row[c])
                rows.append(tr)
                return cols, rows, [(c, int(total_row[c])) for c in cols if c != "Total"]

            ano_p_cols, ano_p_rows, cat_p = build_ano_dash(ano_p_df, qk)
            ano_q_cols, ano_q_rows, cat_q = build_ano_dash(ano_q_df, pk)

            # ============ TABLES KPI AVEC SCORE + CLASSEMENT ============
            def build_kpi_table(kpi_list, scores, score_name):
                sorted_posts = sorted(scores.keys(), key=lambda x: scores[x], reverse=True)
                cols = ["Poste de travail"] + kpi_list + [score_name]
                rows = []
                # CIBLE
                cr = {"_t": "cible", "Poste de travail": "CIBLE"}
                for k in kpi_list: cr[k] = cible[k]
                cr[score_name] = "100.00 %"
                rows.append(cr)
                # Posts triés
                for p in sorted_posts:
                    r = {"_t": "n", "Poste de travail": p}
                    for k in kpi_list: r[k] = round(ckdf.loc[p, k], 2) if p in ckdf.index else ""
                    r[score_name] = "%.2f %%" % scores[p]
                    rows.append(r)
                # Total
                tr = {"_t": "total", "Poste de travail": "Total général"}
                for k in kpi_list: tr[k] = round(ckdf[k].mean(), 2)
                tr[score_name] = "%.2f %%" % (np.mean(list(scores.values())) if scores else 0)
                rows.append(tr)
                return cols, rows

            perf_cols, perf_rows = build_kpi_table(qk, pscores, "Score Performance")
            qual_cols, qual_rows = build_kpi_table(pk, qscores, "Score Qualité")

            # ============ COMPTEURS ============
            total_ot = len(df)
            avg_perf = np.mean(list(pscores.values())) if pscores else 0
            avg_qual = np.mean(list(qscores.values())) if qscores else 0
            total_ano = sum(a["Nb"] for a in all_ano)

            # ============ RENDER ============
            st.markdown('<div class="mh"><h1>📊 KPI Dashboard MC & FEED</h1><div class="db">📅 %s</div></div>' % df_dt, unsafe_allow_html=True)

            # Compteurs
            st.markdown("""<div class="cr">
            <div class="cc c1"><div class="cv">%s</div><div class="cl">Total OT Analysés</div></div>
            <div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Performance Moyen</div></div>
            <div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Qualité Moyen</div></div>
            <div class="cc c4"><div class="cv">%s</div><div class="cl">Total Anomalies</div></div>
            </div>""" % (total_ot, avg_perf, avg_qual, total_ano), unsafe_allow_html=True)

            tab1, tab2 = st.tabs(["📈 INDICATEURS DE PERFORMANCE", "✅ INDICATEUR QUALITÉ"])

            # --- TAB PERFORMANCE ---
            with tab1:
                c_t, c_b = st.columns([5, 1])
                with c_t:
                    st.markdown('<p class="stl q" style="margin-bottom:0">Indicateurs de Performance par Poste de Travail</p>', unsafe_allow_html=True)
                with c_b:
                    vw = st.radio("", ["Tableau KPI", "Anomalies"], horizontal=True, key="vp", label_visibility="collapsed")
                st.markdown("<div style='height:3px'></div>", unsafe_allow_html=True)

                if vw == "Tableau KPI":
                    st.markdown(html_table(perf_rows, perf_cols, "qt", sc_col=["Score Performance"]), unsafe_allow_html=True)
                else:
                    if ano_p_rows:
                        st.markdown(html_ano(ano_p_rows, ["Poste de travail"] + ano_p_cols), unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="es">✅ <b>Aucune anomalie détectée</b></div>', unsafe_allow_html=True)

                # Synthèse
                st.markdown('<p class="stl q">Synthèse des Actions par KPI</p>', unsafe_allow_html=True)
                st.markdown(html_synth(synth_p), unsafe_allow_html=True)

                # Analyse par Catégorie
                if cat_p:
                    mx = max(v for _, v in cat_p)
                    st.markdown('<p class="stl c">Analyse par Catégorie</p>', unsafe_allow_html=True)
                    st.markdown(html_cat(cat_p, mx, "linear-gradient(90deg,#2b6cb0,#3182ce)"), unsafe_allow_html=True)

            # --- TAB QUALITÉ ---
            with tab2:
                st.markdown('<p class="stl p">Indicateur Qualité par Poste de Travail</p>', unsafe_allow_html=True)
                st.markdown(html_table(qual_rows, qual_cols, "pt", sc_col=["Score Qualité"]), unsafe_allow_html=True)

                # Synthèse
                st.markdown('<p class="stl p">Synthèse des Actions par KPI</p>', unsafe_allow_html=True)
                st.markdown(html_synth(synth_q), unsafe_allow_html=True)

                # Analyse par Catégorie
                if cat_q:
                    mx = max(v for _, v in cat_q)
                    st.markdown('<p class="stl c">Analyse par Catégorie</p>', unsafe_allow_html=True)
                    st.markdown(html_cat(cat_q, mx, "linear-gradient(90deg,#276749,#38a169)"), unsafe_allow_html=True)

            # ============ EXPORT ============
            st.markdown('<p class="stl" style="margin-top:8px">💾 Export des Plans d\'Action</p>', unsafe_allow_html=True)
            pa = list(set(a["Poste"] for a in all_ano))
            if pa:
                ce1, ce2 = st.columns([1, 1])
                with ce1: se = st.selectbox("Poste :", options=["📌 Tous"]+pa, key="se")
                with ce2:
                    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
                    gb = st.button("📥 Générer Excel", type="primary", key="gb", use_container_width=True)
                if gb:
                    with st.spinner("Génération..."):
                        out = io.BytesIO()
                        with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                            for pe in (pa if se=="📌 Tous" else [se]):
                                for a in all_ano:
                                    if a["Poste"] != pe: continue
                                    kn = a["KPI"]
                                    sd = pd.DataFrame()
                                    if kn in sub_map_p:
                                        dpf = dfp[dfp["Poste travail princ."]==pe].copy()
                                        sd = sub_map_p[kn](dpf)[["Ordre","Poste travail princ.","Statut OT","Statut utilisateur","Créé le","Date de début planifiée","Total coûts budgétés","Total coûts réels"]].copy()
                                    elif kn in sub_map_q:
                                        dpf = dfp[dfp["Poste travail princ."]==pe].copy()
                                        sd = sub_map_q[kn](dpf)[["Ordre","Poste travail princ.","Statut OT","Statut utilisateur","Créé le","Date de début planifiée","Total coûts budgétés","Total coûts réels"]].copy()
                                    elif kn == "appel avis approuvé":
                                        sd = res['avf'][res['avf']["Poste travail princ."]==pe][["Avis","Poste travail princ.","Statut utilisateur","Créé le","Début souhaité"]].copy()
                                    if not sd.empty: sd.to_excel(w, sheet_name=("%s_%s" % (pe[:15], kn[:15]))[:31], index=False)
                            pd.DataFrame(perf_rows).to_excel(w, sheet_name="KPIs Performance", index=False)
                            pd.DataFrame(qual_rows).to_excel(w, sheet_name="KPIs Qualité", index=False)
                        out.seek(0)
                        st.download_button(label="⬇️ Télécharger Excel", data=out, file_name="Plan_Action_%s.xlsx" % datetime.now().strftime('%Y%m%d_%H%M'), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        st.success("✅ Fichier généré !")
            else:
                st.markdown('<div class="es">🎉 <b>Aucun plan d\'action à exporter</b></div>', unsafe_allow_html=True)

        except Exception as e:
            st.error("❌ Erreur : %s" % str(e))

if __name__ == "__main__":
    main()
