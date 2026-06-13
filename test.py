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
    .main .block-container{padding-top:.6rem;padding-bottom:.6rem;max-width:100%!important;padding-left:.2rem;padding-right:.2rem}
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
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:8px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:7px;text-transform:uppercase;letter-spacing:.3px;padding:3px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw tbody td{padding:2px 3px;border-bottom:1px solid #edf2f7;white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important;font-size:8px!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important;font-size:8px!important}
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
    .dgrid{display:grid;grid-template-columns:1fr 1fr;gap:4px}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:6px 12px;font-weight:700;font-size:11px;width:100%}
    ::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:2px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:9px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:2px 6px;margin-bottom:2px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:10px;color:#718096;font-size:10px}
    .rh{display:flex;align-items:center;justify-content:space-between;margin-bottom:0}
    .rh .stl{margin:0}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:13px}.mh .db{float:none;display:block;margin-top:2px}.cg,.dgrid{grid-template-columns:1fr}.car .cal{width:100px}.gbr-l{width:90px}}
    </style>""", unsafe_allow_html=True)

def main():
    try: locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try: locale.setlocale(locale.LC_ALL, 'fr_FR')
        except: pass
    inject_custom_css()

    consignes = ["Port obligatoire des EPI avant toute intervention.","Port obligatoire du casque de securite.","Port obligatoire des lunettes de protection.","Port obligatoire des gants adaptes au travail.","Utiliser les protections auditives dans les zones bruyantes.","Verifier l'absence de tension avant toute intervention electrique.","Respecter la procedure de consignation et deconsignation.","Ne jamais intervenir sur un equipement en marche.","Baliser et securiser la zone de travail.","Maintenir le poste de travail propre et ordonne.","Verifier l'etat des outils avant utilisation.","Utiliser uniquement du materiel homologue.","Respecter les permis de travail en vigueur.","Identifier les risques avant de commencer une tache.","Signaler immediatement toute situation dangereuse.","Signaler tout incident ou presque accident.","Ne jamais neutraliser un dispositif de securite.","Verifier les detecteurs de gaz avant utilisation.","Verifier la bonne ventilation des zones de travail.","Respecter les regles des espaces confines.","Controler l'atmosphere avant d'entrer dans un espace confine.","Utiliser les points d'ancrage pour les travaux en hauteur.","Verifier l'etat des echafaudages avant utilisation.","Securiser les outils lors des travaux en hauteur.","Ne pas travailler seul lors d'operations a risque.","Controler les elingues avant chaque levage.","Respecter les limites de charge des equipements.","Verifier l'etat des appareils de levage.","Maintenir les voies de circulation degagees.","Respecter la signalisation de securite.","Verifier les extincteurs a proximite du chantier.","Connaitre les issues de secours les plus proches.","Respecter les procedures d'arret d'urgence.","Verifier les flexibles et raccords avant mise en service.","Controler les fuites avant demarrage d'un equipement.","Respecter les distances de securite.","Ne jamais contourner une procedure HSE.","Porter les EPI adaptes au risque identifie.","Prevenir son responsable avant toute intervention particuliere.","Analyser les risques avant chaque demarrage de chantier.","Verifier la stabilite des equipements.","Utiliser les bons outils pour la bonne tache.","Respecter les consignes specifiques du chantier.","Ne jamais prendre de raccourci au detriment de la securite.","Arreter immediatement les travaux en cas de danger.","Proteger l'environnement lors des interventions.","Collecter et trier correctement les dechets.","Eviter toute pollution accidentelle.","Respecter les consignes de stockage des produits dangereux.","Lire les fiches de securite avant manipulation.","Verifier les equipements avant chaque prise de poste.","S'assurer de la disponibilite des moyens de secours.","Communiquer clairement avec l'equipe avant intervention.","Respecter les regles de circulation des engins.","Garder une vigilance permanente sur son environnement.","Prendre le temps d'effectuer le travail en securite.","La securite est l'affaire de tous.","Chaque incident peut etre evite par la prevention.","Aucun travail n'est plus urgent que la securite.","Zero accident commence par un comportement sur."]

    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        c = random.choice(consignes)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:42px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SECURITE</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:18px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Securite - Sante - Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:28px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:32px;font-weight:900">Aucun travail n'est plus urgent que la securite</h2>
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
    def get_metier(p):
        p = str(p).upper()
        if "E" in p: return "Electrique"
        if "M" in p: return "Mecanique"
        if "R" in p: return "Instrumentation"
        if "G" in p: return "Genie Civil"
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

    def calc_kpis(df_i, av_i, now, posts):
        res = {}; df = df_i.copy(); av = av_i.copy()
        mp = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
        df["Backlog preparation"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, mp)), "CARACTERISE", "NON CARACTERISE")
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
        pc = pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"], index="Poste travail princ.", columns="Backlog preparation", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
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
        res['ckdf'] = pd.DataFrame({
            "TAUX_REALISATION_CORRECTIF/PT": an["TAUX_REALISATION_CORRECTIF/PT"],
            "OT préparation <1 mois": pr["OT préparation <1 mois"],"OT préparation >3 mois": pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois": pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois": pl["OT planification <1 mois"],"OT planification >3 mois": pl["OT planification >3 mois"],"OT planification 1mois< <3mois": pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois": ex["OT exécution <1 mois"],"OT exécution >3 mois": ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois": ex["OT exécution 1mois< <3mois"],
            "appel avis approuvé": tca["appel avis approuvé"],"OT LANC ESTIME": la["OT LANC ESTIME"],
            "Backlog préparation caractérisé": pc["Backlog préparation caractérisé"],"Backlog planification caractérisé": plc["Backlog planification caractérisé"],
            "OT CONFIME": res['ot_confime']["OT CONFIME"],"OT_COR_EGAL": res['ot_cor_egal']["OT_COR_EGAL"]
        })
        return res

    def ks(v, c):
        try: val = float(v)
        except: return ""
        if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val>=80 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=75 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val<=15 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val<=5 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c == "TAUX_REALISATION_CORRECTIF/PT": return "background:#c6efce;color:#006100;font-weight:600" if val>=85 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c == "appel avis approuvé": return "background:#c6efce;color:#006100;font-weight:600" if val>=95 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=90 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]: return "background:#c6efce;color:#006100;font-weight:600" if val>=100 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=95 else "background:#ffc7ce;color:#9c0006;font-weight:600")
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
    def is_lb(k): return k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois","OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]

    def html_table(rows, cols, tc, sc_col=None):
        h = '<table class="tw %s"><thead><tr>' % tc + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for r in rows:
            rc = "cb" if r.get("_t")=="cible" else ("tr" if r.get("_t")=="total" else "")
            h += '<tr class="%s">' % rc
            for c in cols:
                v = r.get(c, "")
                if r.get("_t")=="cible":
                    h += '<td>%s</td>' % v
                else:
                    s = cs(v) if sc_col and c in sc_col else ks(v, c)
                    h += '<td style="%s">%s</td>' % (s or "", v)
            h += '</tr>'
        return h + '</tbody></table>'

    def html_ano(rows, cols):
        h = '<table class="tw at"><thead><tr>' + ''.join('<th>%s</th>' % c for c in cols) + '</tr></thead><tbody>'
        for r in rows:
            h += '<tr class="%s">' % ("tr" if r.get("_t")=="total" else "")
            for c in cols: v = r.get(c,""); h += '<td style="%s">%s</td>' % (kas(v) or "", v)
            h += '</tr>'
        return h + '</tbody></table>'

    def html_synth(kpi_list, actuals, targets, act_map, accent):
        h = ''
        for k in kpi_list:
            av, tv = actuals.get(k,0), targets.get(k,100)
            met = av <= tv if is_lb(k) else av >= tv
            sbg, sclr = ("#c6efce","#006100") if met else ("#ffc7ce","#9c0006")
            scbg = accent if met else "#e53e3e"
            act = "Objectif atteint" if met else act_map.get(k,"")
            h += '<div class="sr"><div class="sn">%s</div><div class="sc" style="background:%s">%.1f%%</div><div class="stg">Cible: %s%%</div><div class="sb" style="color:%s;background:%s">%s</div><div class="sa">%s</div></div>' % (k, scbg, av, tv, sclr, sbg, "ATTEINT" if met else "NON ATTEINT", act)
        return h

    def html_cat(kpi_list, actuals, targets, col_ok, col_fail):
        h = '<div class="ca">'
        for k in kpi_list:
            av, tv = actuals.get(k,0), targets.get(k,100)
            met = av <= tv if is_lb(k) else av >= tv
            bw = min(max(av,3),100); bg = col_ok if met else col_fail
            h += '<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%% / %s%%</div></div>' % (k, bw, bg, av, tv)
        h += '</div>'; return h

    def html_classement(scores, accent):
        sp = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        met_p, not_p = [(p,s) for p,s in sp if s>=80], [(p,s) for p,s in sp if s<80]
        t5, b5 = met_p[:5], not_p[-5:] if len(not_p)>5 else not_p
        h = '<div class="cg"><div><div class="ct" style="color:#38a169">Top 5 — Objectif Atteint</div>'
        if t5:
            for i,(p,s) in enumerate(t5): h += '<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>' % (accent,i+1,p,cs("%.2f"%s),s)
        else: h += '<div style="padding:4px;font-size:8px;color:#718096">Aucun poste</div>'
        h += '</div><div><div class="ct" style="color:#e53e3e">Bottom 5 — Non Atteint</div>'
        if b5:
            for i,(p,s) in enumerate(reversed(b5)): h += '<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>' % (len(b5)-i,p,cs("%.2f"%s),s)
        else: h += '<div style="padding:4px;font-size:8px;color:#38a169">Tous atteints</div>'
        h += '</div></div>'; return h

    def html_kpi_bars(kpi_list, actuals, targets, title, color_ok, color_fail):
        h = '<div class="ca"><div class="ct" style="color:%s">%s</div>' % (color_ok, title)
        for k in kpi_list:
            av, tv = actuals.get(k,0), targets.get(k,100)
            met = av <= tv if is_lb(k) else av >= tv
            bw = min(max(av,0),100); bg = color_ok if met else color_fail
            h += '<div class="car"><div class="cal" style="width:240px">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>' % (k, bw, bg, av)
        h += '</div>'; return h

    def html_bars(data, title, color):
        h = '<div class="ca"><div class="ct" style="color:%s">%s</div>' % (color, title)
        for label, val in sorted(data, key=lambda x: x[1], reverse=True):
            bw = min(max(val,0),100)
            h += '<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>' % (label, bw, color, val)
        h += '</div>'; return h

    def html_grouped_bars(posts, pscores, qscores, title):
        h = '<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>' % title
        h += '<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span></div>'
        sp2 = sorted(posts, key=lambda x: (pscores.get(x,0)+qscores.get(x,0))/2, reverse=True)
        for p in sp2:
            pv, qv = pscores.get(p,0), qscores.get(p,0)
            pw, qw = min(max(pv,0),100), min(max(qv,0),100)
            h += '<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f gb-p" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w"><div class="gbr-f gb-q" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div></div></div>' % (p, pw, pv, qw, qv)
        h += '</div>'; return h

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""<div style="padding:10px 0 4px 0"><div style="font-size:18px;margin-bottom:2px">⚙️</div><div style="font-size:12px;font-weight:800;color:white">Filtres & Parametres</div><div style="font-size:8px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Configuration</div></div>""", unsafe_allow_html=True)
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
            st.markdown("""<div style="background:rgba(255,255,255,.1);padding:5px 8px;border-radius:6px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:7px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Donnees</div><div style="font-size:10px;color:white;font-weight:600;margin-top:1px">📅 %s</div></div>""" % df_dt, unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("**🎯 Postes**")
        sp = st.multiselect("Poste", ["All"]+apm, ["All"], key="sp")
        st.markdown("**🏭 Atelier**")
        sa = st.multiselect("Atelier", ["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)"], ["All"], key="sa")
        st.markdown("**🏢 Division**")
        sd = st.multiselect("Division", ["All","SF1","SF2"], ["All"], key="sd")
        st.markdown("---")
        st.markdown("**📅 Periode**")
        dr = st.date_input("Date debut planifiee", value=(datetime(2025,1,1).date(), datetime.today().date()), format="DD/MM/YYYY", key="dr")

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
            
            # === Donnees AVEC filtre date (pour tableaux detailles) ===
            df = raw_ot[(raw_ot["Poste travail princ."].isin(vp)) & (raw_ot["Date de début planifiée"].between(sdt, edt))].copy()
            avdf = raw_av[raw_av["Poste travail princ."].isin(vp)].copy()
            df = excr(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)].drop_duplicates())
            avdf = excr(avdf[(avdf["Ordre"].isna())|(avdf["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates())
            if "Statut système" in df.columns: df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            # === Donnees SANS filtre date (pour chart tableau de bord) ===
            df_dash = raw_ot[raw_ot["Poste travail princ."].isin(vp)].copy()
            df_dash = excr(df_dash[df_dash["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"), na=False)].drop_duplicates())
            if "Statut système" in df_dash.columns: df_dash["Statut OT"] = df_dash["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            now = pd.Timestamp.now()

            # KPIs avec date (detail)
            res = calc_kpis(df, avdf, now, vp)
            ckdf = res['ckdf']; dfp = res['dfp']

            # KPIs sans date (dashboard)
            res_d = calc_kpis(df_dash, avdf, now, vp)
            ckdf_d = res_d['ckdf']

            qk = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois"]
            pk = ["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]
            cible = {"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,"OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,"OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,"OT exécution 1mois< <3mois":15,"appel avis approuvé":95,"OT LANC ESTIME":100,"Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,"OT CONFIME":100,"OT_COR_EGAL":100}
            act_map = {"TAUX_REALISATION_CORRECTIF/PT":"Ameliorer le taux de realisation des OT.","OT préparation <1 mois":"Reduire l'age de preparation des OT (< 1 mois).","OT préparation >3 mois":"Traiter les OT avec preparation > 3 mois.","OT planification <1 mois":"Reduire l'age de planification des OT (< 1 mois).","OT planification >3 mois":"Traiter les OT avec planification > 3 mois.","OT exécution <1 mois":"Reduire l'age d'execution des OT (< 1 mois).","OT exécution >3 mois":"Traiter les OT avec execution > 3 mois.","OT LANC ESTIME":"Estimer les couts des OT lances.","Backlog préparation caractérisé":"Caracteriser le backlog de preparation.","Backlog planification caractérisé":"Caracteriser le backlog de planification.","OT CONFIME":"Confirmer les OT termines.","OT_COR_EGAL":"Rapprocher les couts reels et budgetes.","appel avis approuvé":"Creer un OT pour les avis sans ordre."}

            # Scores avec date
            pscores = {}; qscores = {}
            for poste in ckdf.index:
                r = ckdf.loc[poste]
                pscores[poste] = (sum(gscore(k,r[k],cible[k]) for k in qk if k in r.index)/len(qk)*100) if qk else 0
                qscores[poste] = (sum(gscore(k,r[k],cible[k]) for k in pk if k in r.index)/len(pk)*100) if pk else 0
            pa = {k: round(ckdf[k].mean(),2) for k in qk}
            qa = {k: round(ckdf[k].mean(),2) for k in pk}

            # Scores sans date (dashboard)
            pscores_d = {}; qscores_d = {}
            for poste in ckdf_d.index:
                r = ckdf_d.loc[poste]
                pscores_d[poste] = (sum(gscore(k,r[k],cible[k]) for k in qk if k in r.index)/len(qk)*100) if qk else 0
                qscores_d[poste] = (sum(gscore(k,r[k],cible[k]) for k in pk if k in r.index)/len(pk)*100) if pk else 0
            pa_d = {k: round(ckdf_d[k].mean(),2) for k in qk}
            qa_d = {k: round(ckdf_d[k].mean(),2) for k in pk}

            # ANOMALIES (avec date)
            all_ano = []
            sub_p = {"TAUX_REALISATION_CORRECTIF/PT":lambda d:d[(d["Nº appel pl.entret."].fillna(0)==0)&(~d["Statut OT"].isin(["CLOT","TCLO"]))],"OT préparation <1 mois":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]!="<1 mois")],"OT préparation >3 mois":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]==">3 mois")],"OT planification <1 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]!="<1 mois")],"OT planification >3 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]==">3 mois")],"OT exécution <1 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]!="<1 mois")],"OT exécution >3 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]==">3 mois")]}
            sub_q = {"OT LANC ESTIME":lambda d:d[(d["Statut OT"]=="LANC")&(d["OT LANC ESTIME"]=="NON")],"Backlog préparation caractérisé":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["Backlog preparation"]=="NON CARACTERISE")],"Backlog planification caractérisé":lambda d:d[(d["Statut OT"]=="LANC")&(d["Backlog planification"]=="NON CARACTERISE")],"OT CONFIME":lambda d:d[d["OT CONFIME"]=="NON"],"OT_COR_EGAL":lambda d:d[d["OT_COR_EGAL"]=="NON"]}
            for poste in vp:
                if poste not in dfp["Poste travail princ."].values: continue
                dp = dfp[dfp["Poste travail princ."]==poste]
                for kn, sf in sub_p.items():
                    vk = ckdf.loc[poste, kn] if poste in ckdf.index else 100
                    if pd.notna(vk) and vk < cible[kn]:
                        cnt = len(sf(dp))
                        if cnt > 0: all_ano.append({"Poste":poste,"KPI":kn,"Nb":cnt,"Type":"P"})
                for kn, sf in sub_q.items():
                    vk = ckdf.loc[poste, kn] if poste in ckdf.index else 100
                    if pd.notna(vk) and vk < cible[kn]:
                        cnt = len(sf(dp))
                        if cnt > 0: all_ano.append({"Poste":poste,"KPI":kn,"Nb":cnt,"Type":"Q"})
                vk_av = ckdf.loc[poste, "appel avis approuvé"] if poste in ckdf.index else 100
                if pd.notna(vk_av) and vk_av < cible["appel avis approuvé"]:
                    cnt = len(res['avf'][res['avf']["Poste travail princ."]==poste])
                    if cnt > 0: all_ano.append({"Poste":poste,"KPI":"appel avis approuvé","Nb":cnt,"Type":"Q"})

            def build_ano(ano_list, kpi_list):
                if not ano_list: return [], []
                adf = pd.DataFrame(ano_list)
                pv = adf.pivot_table(index="Poste", columns="KPI", values="Nb", aggfunc="sum", fill_value=0).astype(int)
                pv["Total"] = pv.sum(axis=1); tot = pv.sum()
                cols = [c for c in kpi_list if c in pv.columns] + ["Total"]; rows = []
                for idx in pv.index:
                    r = {"_t":"n","Poste de travail":idx}
                    for c in cols: r[c] = pv.loc[idx, c]
                    rows.append(r)
                tr = {"_t":"total","Poste de travail":"Total general"}
                for c in cols: tr[c] = int(tot[c])
                rows.append(tr); return ["Poste de travail"]+cols, rows

            ano_p_c, ano_p_r = build_ano([a for a in all_ano if a["Type"]=="P"], qk)
            ano_q_c, ano_q_r = build_ano([a for a in all_ano if a["Type"]=="Q"], pk)

            def build_kpi(kpi_list, scores, sname):
                sp2 = sorted(scores.keys(), key=lambda x: scores[x], reverse=True)
                cols = ["Poste de travail"]+kpi_list+[sname]; rows = []
                cr = {"_t":"cible","Poste de travail":"CIBLE"}
                for k in kpi_list: cr[k] = cible[k]
                cr[sname] = "100.00 %"; rows.append(cr)
                for p in sp2:
                    r = {"_t":"n","Poste de travail":p}
                    for k in kpi_list: r[k] = round(ckdf.loc[p, k], 2) if p in ckdf.index else ""
                    r[sname] = "%.2f %%" % scores[p]; rows.append(r)
                tr = {"_t":"total","Poste de travail":"Total general"}
                for k in kpi_list: tr[k] = round(ckdf[k].mean(), 2)
                tr[sname] = "%.2f %%" % (np.mean(list(scores.values())) if scores else 0)
                rows.append(tr); return cols, rows

            pcols, prows = build_kpi(qk, pscores, "Score Performance")
            qcols, qrows = build_kpi(pk, qscores, "Score Qualite")

            # Donnees dashboard
            df_sc_d = pd.DataFrame([{"Poste":p,"Perf":pscores_d[p],"Qual":qscores_d[p],"Metier":get_metier(p),"Atelier":get_atelier(p),"Division":get_division(p)} for p in vp if p in pscores_d])
            by_at = df_sc_d.groupby("Atelier")[["Perf","Qual"]].mean().round(1) if not df_sc_d.empty else pd.DataFrame()
            by_mt = df_sc_d.groupby("Metier")[["Perf","Qual"]].mean().round(1) if not df_sc_d.empty else pd.DataFrame()
            by_div = df_sc_d.groupby("Division")[["Perf","Qual"]].mean().round(1) if not df_sc_d.empty else pd.DataFrame()
            sf1_p = [p for p in vp if str(p).upper().startswith("SF1") and p in pscores_d]
            sf2_p = [p for p in vp if str(p).upper().startswith("SF2") and p in pscores_d]

            total_ot = len(df); avg_p = np.mean(list(pscores.values())) if pscores else 0
            avg_q = np.mean(list(qscores.values())) if qscores else 0; total_ano = sum(a["Nb"] for a in all_ano)

            # RENDER
            st.markdown('<div class="mh"><h1>📊 KPI Dashboard MC & FEED</h1><div class="db">📅 %s</div></div>' % df_dt, unsafe_allow_html=True)
            st.markdown("""<div class="cr">
            <div class="cc c1"><div class="cv">%s</div><div class="cl">Total OT Analyses</div></div>
            <div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Performance</div></div>
            <div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Qualite</div></div>
            <div class="cc c4"><div class="cv">%s</div><div class="cl">Total Anomalies</div></div>
            </div>""" % (total_ot, avg_p, avg_q, total_ano), unsafe_allow_html=True)

            tab0, tab1, tab2 = st.tabs(["📊 TABLEAU DE BORD", "📈 INDICATEURS DE PERFORMANCE", "✅ INDICATEUR QUALITE"])

            # ==================== DASHBOARD ====================
            with tab0:
                st.markdown('<div class="stl p">📊 Vue d\'ensemble par poste</div>', unsafe_allow_html=True)
                st.markdown(html_grouped_bars(vp, pscores_d, qscores_d, "Performance & Qualite par Poste de Travail"), unsafe_allow_html=True)

                st.markdown('<div class="stl p">🏭 Par Atelier</div>', unsafe_allow_html=True)
                if not by_at.empty:
                    at_data = [(idx, row["Perf"]) for idx, row in by_at.iterrows()]
                    st.markdown(html_bars(at_data, "Score Performance par Atelier", "#2b6cb0"), unsafe_allow_html=True)
                    at_data_q = [(idx, row["Qual"]) for idx, row in by_at.iterrows()]
                    st.markdown(html_bars(at_data_q, "Score Qualite par Atelier", "#276749"), unsafe_allow_html=True)

                st.markdown('<div class="stl p">🔧 Par Metier</div>', unsafe_allow_html=True)
                if not by_mt.empty:
                    mt_data = [(idx, row["Perf"]) for idx, row in by_mt.iterrows()]
                    st.markdown(html_bars(mt_data, "Score Performance par Metier", "#2b6cb0"), unsafe_allow_html=True)
                    mt_data_q = [(idx, row["Qual"]) for idx, row in by_mt.iterrows()]
                    st.markdown(html_bars(mt_data_q, "Score Qualite par Metier", "#276749"), unsafe_allow_html=True)

                st.markdown('<div class="stl p">🏢 Par Division</div>', unsafe_allow_html=True)
                if not by_div.empty:
                    dv_data = [(idx, row["Perf"]) for idx, row in by_div.iterrows()]
                    st.markdown(html_bars(dv_data, "Score Performance par Division", "#2b6cb0"), unsafe_allow_html=True)
                    dv_data_q = [(idx, row["Qual"]) for idx, row in by_div.iterrows()]
                    st.markdown(html_bars(dv_data_q, "Score Qualite par Division", "#276749"), unsafe_allow_html=True)

                st.markdown('<div class="stl p">📊 Synthese globale (sans filtre date)</div>', unsafe_allow_html=True)
                st.markdown(html_synth(qk, pa_d, cible, act_map, "#2b6cb0"), unsafe_allow_html=True)
                st.markdown(html_synth(pk, qa_d, cible, act_map, "#276749"), unsafe_allow_html=True)

                st.markdown('<div class="stl p">🏆 Classement des postes</div>', unsafe_allow_html=True)
                st.markdown(html_classement(pscores_d, "#2b6cb0"), unsafe_allow_html=True)
                st.markdown(html_classement(qscores_d, "#276749"), unsafe_allow_html=True)

            # ==================== PERFORMANCE ====================
            with tab1:
                st.markdown('<div class="rh"><div class="stl p">📈 Indicateurs de Performance par Poste</div></div>', unsafe_allow_html=True)
                st.markdown(html_table(prows, pcols, "pt", ["Score Performance"]), unsafe_allow_html=True)

                if ano_p_c:
                    st.markdown('<div class="rh"><div class="stl a">⚠️ Anomalies Performance</div></div>', unsafe_allow_html=True)
                    st.markdown(html_ano(ano_p_r, ano_p_c), unsafe_allow_html=True)

                st.markdown('<div class="stl p">📊 Synthese Performance</div>', unsafe_allow_html=True)
                st.markdown(html_synth(qk, pa, cible, act_map, "#2b6cb0"), unsafe_allow_html=True)

                st.markdown('<div class="stl p">📈 Barres de progression Performance</div>', unsafe_allow_html=True)
                st.markdown(html_kpi_bars(qk, pa, cible, "Performance Globale", "#2b6cb0", "#e53e3e"), unsafe_allow_html=True)

                st.markdown('<div class="stl p">🏆 Classement Performance</div>', unsafe_allow_html=True)
                st.markdown(html_classement(pscores, "#2b6cb0"), unsafe_allow_html=True)

            # ==================== QUALITE ====================
            with tab2:
                st.markdown('<div class="rh"><div class="stl q">✅ Indicateurs de Qualite par Poste</div></div>', unsafe_allow_html=True)
                st.markdown(html_table(qrows, qcols, "qt", ["Score Qualite"]), unsafe_allow_html=True)

                if ano_q_c:
                    st.markdown('<div class="rh"><div class="stl a">⚠️ Anomalies Qualite</div></div>', unsafe_allow_html=True)
                    st.markdown(html_ano(ano_q_r, ano_q_c), unsafe_allow_html=True)

                st.markdown('<div class="stl q">📊 Synthese Qualite</div>', unsafe_allow_html=True)
                st.markdown(html_synth(pk, qa, cible, act_map, "#276749"), unsafe_allow_html=True)

                st.markdown('<div class="stl q">📈 Barres de progression Qualite</div>', unsafe_allow_html=True)
                st.markdown(html_kpi_bars(pk, qa, cible, "Qualite Globale", "#276749", "#e53e3e"), unsafe_allow_html=True)

                st.markdown('<div class="stl q">🏆 Classement Qualite</div>', unsafe_allow_html=True)
                st.markdown(html_classement(qscores, "#276749"), unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Erreur de chargement: {str(e)}")
    else:
        st.markdown('<div class="es">📁 Veuillez charger les fichiers OT et AVIS pour commencer.</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
