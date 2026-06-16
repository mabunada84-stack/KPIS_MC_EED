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

QK = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois",
      "OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois",
      "OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois",
      "OT exécution 1mois< <3mois"]
PK = ["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé",
      "Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]
ALL_KPI = QK + PK
CIBLE = {"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,
         "OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,
         "OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,
         "OT exécution 1mois< <3mois":15,"appel avis approuvé":95,"OT LANC ESTIME":100,
         "Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,
         "OT CONFIME":100,"OT_COR_EGAL":100}
ACT_MAP = {"TAUX_REALISATION_CORRECTIF/PT":"Ameliorer le taux de realisation des OT d'execution.",
           "OT préparation <1 mois":"Reduire l'age de preparation des OT (< 1 mois).",
           "OT préparation >3 mois":"Traiter les OT avec preparation > 3 mois.",
           "OT planification <1 mois":"Reduire l'age de planification des OT (< 1 mois).",
           "OT planification >3 mois":"Traiter les OT avec planification > 3 mois.",
           "OT exécution <1 mois":"Reduire l'age d'execution des OT (< 1 mois).",
           "OT exécution >3 mois":"Traiter les OT avec execution > 3 mois.",
           "OT LANC ESTIME":"Estimer les couts des OT lances.",
           "Backlog préparation caractérisé":"Caracteriser le backlog de preparation.",
           "Backlog planification caractérisé":"Caracteriser le backlog de planification.",
           "OT CONFIME":"Confirmer les OT termines.",
           "OT_COR_EGAL":"Rapprocher les couts reels et budgetes.",
           "appel avis approuvé":"Creer un OT pour les avis sans ordre.",
           "OT préparation 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT planification 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT exécution 1mois< <3mois":"Reduire les OT entre 1 et 3 mois."}
LOWER_BETTER = ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois",
                "OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]
MP_KW = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
MPLAN_KW = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
CONSIGNES_HSE = [
    "Port obligatoire des EPI avant toute intervention.","Port obligatoire du casque de securite.",
    "Port obligatoire des lunettes de protection.","Port obligatoire des gants adaptes au travail.",
    "Utiliser les protections auditives dans les zones bruyantes.","Verifier l'absence de tension avant toute intervention electrique.",
    "Respecter la procedure de consignation et deconsignation.","Ne jamais intervenir sur un equipement en marche.",
    "Baliser et securiser la zone de travail.","Maintenir le poste de travail propre et ordonne.",
    "Verifier l'etat des outils avant utilisation.","Utiliser uniquement du materiel homologue.",
    "Respecter les permis de travail en vigueur.","Identifier les risques avant de commencer une tache.",
    "Signaler immediatement toute situation dangereuse.","Signaler tout incident ou presque accident.",
    "Ne jamais neutraliser un dispositif de securite.","Verifier les detecteurs de gaz avant utilisation.",
    "Verifier la bonne ventilation des zones de travail.","Respecter les regles des espaces confines.",
    "Controler l'atmosphere avant d'entrer dans un espace confine.","Utiliser les points d'ancrage pour les travaux en hauteur.",
    "Verifier l'etat des echafaudages avant utilisation.","Securiser les outils lors des travaux en hauteur.",
    "Ne pas travailler seul lors d'operations a risque.","Controler les elingues avant chaque levage.",
    "Respecter les limites de charge des equipements.","Verifier l'etat des appareils de levage.",
    "Maintenir les voies de circulation degagees.","Respecter la signalisation de securite.",
    "Verifier les extincteurs a proximite du chantier.","Connaitre les issues de secours les plus proches.",
    "Respecter les procedures d'arret d'urgence.","Verifier les flexibles et raccords avant mise en service.",
    "Controler les fuites avant demarrage d'un equipement.","Respecter les distances de securite.",
    "Ne jamais contourner une procedure HSE.","Porter les EPI adaptes au risque identifie.",
    "Prevenir son responsable avant toute intervention particuliere.","Analyser les risques avant chaque demarrage de chantier.",
    "Verifier la stabilite des equipements.","Utiliser les bons outils pour la bonne tache.",
    "Respecter les consignes specifiques du chantier.","Ne jamais prendre de raccourci au detriment de la securite.",
    "Arreter immediatement les travaux en cas de danger.","Proteger l'environnement lors des interventions.",
    "Collecter et trier correctement les dechets.","Eviter toute pollution accidentelle.",
    "Respecter les consignes de stockage des produits dangereux.","Lire les fiches de securite avant manipulation.",
    "Verifier les equipements avant chaque prise de poste.","S'assurer de la disponibilite des moyens de secours.",
    "Communiquer clairement avec l'equipe avant intervention.","Respecter les regles de circulation des engins.",
    "Garder une vigilance permanente sur son environnement.","Prendre le temps d'effectuer le travail en securite.",
    "La securite est l'affaire de tous.","Chaque incident peut etre evite par la prevention.",
    "Aucun travail n'est plus urgent que la securite.","Zero accident commence par un comportement sur."]

# ============================================================
# FONCTIONS UTILITAIRES NIVEAU MODULE
# ============================================================
def excr(df):
    if "Poste travail princ." in df.columns:
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)].copy()
    return df

def contient_mot(t, lm):
    t = str(t)
    return any(m in t for l in lm for m in l.split())

def cat_age(a):
    if a <= 1: return "<1 mois"
    elif a >= 3: return ">3 mois"
    return "1 mois < <3 mois"

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

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt","r",encoding="utf-8") as f: return f.read().strip()
        except Exception: pass
    return datetime.now().strftime("%d/%m/%Y")

def get_poste_name(row):
    for col in ["Poste travail princ.", "Poste de travail", "Poste travail", "Poste"]:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
            return str(row[col]).strip()
    return "NON ATTRIBUE"

# ============================================================
# CACHE : Chargement des fichiers Excel
# ============================================================
@st.cache_data(ttl=3600, show_spinner="Chargement des fichiers Excel...")
def _cached_load_excel(path_or_bytes, is_bytes):
    if is_bytes:
        return pd.read_excel(io.BytesIO(path_or_bytes))
    return pd.read_excel(path_or_bytes)

def load_raw_data(unf, ot_f, av_f):
    if unf:
        raw_ot = _cached_load_excel(ot_f.read(), is_bytes=True)
        raw_av = _cached_load_excel(av_f.read(), is_bytes=True)
    else:
        raw_ot = _cached_load_excel("ot.xlsx", is_bytes=False)
        raw_av = _cached_load_excel("avis.xlsx", is_bytes=False)
    raw_ot = excr(raw_ot)
    raw_av = excr(raw_av)
    for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
        if c in raw_ot.columns: raw_ot[c] = pd.to_datetime(raw_ot[c],errors="coerce")
    for c in ["Créé le","Début souhaité","Date de la clôture"]:
        if c in raw_av.columns: raw_av[c] = pd.to_datetime(raw_av[c],errors="coerce")
    return raw_ot, raw_av

# ============================================================
# CACHE : Calcul des KPIs
# ============================================================
def _df_fingerprint(df, av, posts, now_ts):
    s = f"{len(df)}|{df['Ordre'].nunique() if 'Ordre' in df.columns else 0}|{len(av)}|{sorted(posts)}|{now_ts}"
    return hashlib.md5(s.encode()).hexdigest()

@st.cache_data(show_spinner="Calcul des KPIs en cours...")
def _cached_calc_kpis(_fp, df, av, now_ts, posts):
    return _calc_kpis_impl(df, av, pd.Timestamp.fromtimestamp(now_ts), list(posts))

def _calc_kpis_impl(df_i, av_i, now, posts):
    res = {}; df = df_i.copy(); av = av_i.copy()
    df["Backlog preparation"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, MP_KW)), "CARACTERISE", "NON CARACTERISE")
    df["Backlog planification"] = np.where(df["Statut utilisateur"].apply(lambda x: contient_mot(x, MPLAN_KW)), "CARACTERISE", "NON CARACTERISE")
    for dc, am, ac in [('Créé le',"amp","ap"), ('Date de début planifiée',"amlp","alp"), ('Date de début planifiée',"amex","aex")]:
        if dc in df.columns:
            df[dc] = pd.to_datetime(df[dc], errors='coerce')
            df[am] = ((now.year - df[dc].dt.year)*12 + (now.month - df[dc].dt.month)).round(2)
            df[ac] = df[am].apply(cat_age)
        else:
            df[am] = np.nan; df[ac] = "Inconnu"
    df["OT CONFIME"] = np.where(df["Statut système"].str.contains("CLO",na=False)&df["Statut système"].str.contains("CONF",na=False),"OUI","NON")
    df["Contient SOPL"] = df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
    df["OT LANC ESTIME"] = np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
    df["OT_COR_EGAL"] = np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
    res['dfp'] = df

    def ckpi(n, d, sz=100):
        return np.where(d==0, sz, (n/d)*100)
    def cpiv(f, col_name, p):
        return pd.pivot_table(f, index="Poste travail princ.", columns=col_name, values="Ordre", aggfunc="count", fill_value=0).reindex(p, fill_value=0)

    # NOUVEAU CALCUL TAUX_REALISATION_CORRECTIF/PT
    df_total = df[
        (df["Nº appel pl.entret."].fillna(0) == 0)
        & (df["Contient SOPL"] == 1)
    ]
    total_ot = df_total.groupby("Poste travail princ.")["Ordre"].count()
    df_cloture = df_total[df_total["Statut OT"].isin(["TCLO", "CLOT"])]
    ot_clotures = df_cloture.groupby("Poste travail princ.")["Ordre"].count()
    an = pd.DataFrame(index=posts)
    an["OT_CLOTURES"] = ot_clotures
    an["Total"] = total_ot
    an = an.fillna(0)
    an["TAUX_REALISATION_CORRECTIF/PT"] = np.where(
        an["Total"] == 0,
        100,
        (an["OT_CLOTURES"] / an["Total"]) * 100
    )

    # Préparation
    pr = cpiv(df[df["Statut OT"]=="CRÉÉ"], "ap", posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pr[c] = pr.get(c, 0)
    pr["Total"] = pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    pr["OT préparation <1 mois"] = ckpi(pr["<1 mois"], pr["Total"])
    pr["OT préparation >3 mois"] = ckpi(pr[">3 mois"], pr["Total"], 0)
    pr["OT préparation 1mois< <3mois"] = ckpi(pr["1 mois < <3 mois"], pr["Total"], 0)

    # Planification
    pl = cpiv(df[(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==0)], "alp", posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pl[c] = pl.get(c, 0)
    pl["Total"] = pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    pl["OT planification <1 mois"] = ckpi(pl["<1 mois"], pl["Total"])
    pl["OT planification >3 mois"] = ckpi(pl[">3 mois"], pl["Total"], 0)
    pl["OT planification 1mois< <3mois"] = ckpi(pl["1 mois < <3 mois"], pl["Total"], 0)

    # Exécution
    ex = cpiv(df[(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1)], "aex", posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: ex[c] = ex.get(c, 0)
    ex["Total"] = ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    ex["OT exécution <1 mois"] = ckpi(ex["<1 mois"], ex["Total"])
    ex["OT exécution >3 mois"] = ckpi(ex[">3 mois"], ex["Total"], 0)
    ex["OT exécution 1mois< <3mois"] = ckpi(ex["1 mois < <3 mois"], ex["Total"], 0)

    # OT LANC ESTIME
    la = pd.pivot_table(df[df["Statut OT"]=="LANC"], index="Poste travail princ.", columns="OT LANC ESTIME", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["OUI","NON"]: la[c] = la.get(c, 0)
    la["Total"] = la["OUI"]+la["NON"]; la["OT LANC ESTIME"] = ckpi(la["OUI"], la["Total"])

    # Backlog préparation
    pc = pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"], index="Poste travail princ.", columns="Backlog preparation", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: pc[c] = pc.get(c, 0)
    pc["Total"] = pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"] = ckpi(pc["CARACTERISE"], pc["Total"])

    # Backlog planification
    plc = pd.pivot_table(df[df["Statut OT"]=="LANC"], index="Poste travail princ.", columns="Backlog planification", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: plc[c] = plc.get(c, 0)
    plc["Total"] = plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"] = ckpi(plc["CARACTERISE"], plc["Total"])

    # OT CONFIME et OT_COR_EGAL
    for kn, cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
        pv = pd.pivot_table(df, index="Poste travail princ.", columns=cn, values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
        for c in ["OUI","NON"]: pv[c] = pv.get(c, 0)
        pv["Total"] = pv["OUI"]+pv["NON"]; pv[cn] = ckpi(pv["OUI"], pv["Total"])
        res[kn.lower().replace(" ","_")] = pv

    # Avis
    avf = av[(av["Ordre"].isna())|(av["Ordre"].astype(str).str.strip()=="")].copy()
    res['avf'] = avf
    tca = pd.pivot_table(avf, index="Poste travail princ.", columns="Statut utilisateur", values="Avis", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
    for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c] = tca.get(c, 0)
    tca["Total"] = tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1)
    tca["appel avis approuvé"] = ckpi(tca["APRV"], tca["Total"])

    # Assemblage final
    res['ckdf'] = pd.DataFrame({
        "TAUX_REALISATION_CORRECTIF/PT": an["TAUX_REALISATION_CORRECTIF/PT"],
        "OT préparation <1 mois":pr["OT préparation <1 mois"],"OT préparation >3 mois":pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
        "OT planification <1 mois":pl["OT planification <1 mois"],"OT planification >3 mois":pl["OT planification >3 mois"],"OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
        "OT exécution <1 mois":ex["OT exécution <1 mois"],"OT exécution >3 mois":ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
        "appel avis approuvé":tca["appel avis approuvé"],"OT LANC ESTIME":la["OT LANC ESTIME"],
        "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],"Backlog planification caractérisé":plc["Backlog planification caractérisé"],
        "OT CONFIME":res['ot_confime']["OT CONFIME"],"OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"]
    })
    return res

# ============================================================
# CACHE : Historique et variations
# ============================================================
@st.cache_data(ttl=1800, show_spinner="Chargement historique...")
def _cached_load_historical(filepath):
    return _load_historical_impl(filepath)

def _load_historical_impl(filepath):
    if not os.path.exists(filepath): return pd.DataFrame()
    try: wb = load_workbook(filepath, read_only=True, data_only=True)
    except Exception: return pd.DataFrame()
    records = []; section = None; headers = None
    for sheet_name in wb.sheetnames:
        try:
            ws = wb[sheet_name]; rows_data = list(ws.iter_rows(values_only=True))
            for row in rows_data:
                cell0 = str(row[0]).strip() if row[0] else ""
                if "INDICATEURS DE PERFORMANCE" in cell0.upper(): section="perf"; headers=None; continue
                elif "INDICATEURS DE QUALITE" in cell0.upper(): section="qual"; headers=None; continue
                elif "ANOMALIES" in cell0.upper(): section=None; continue
                if section and headers is None and cell0:
                    headers = [str(c).strip() if c else "" for c in row]; continue
                if section and headers and cell0 and cell0 not in ("CIBLE","Total general",""):
                    entry = {"Date":sheet_name}
                    for j, h in enumerate(headers):
                        if j < len(row): entry[h] = row[j]
                    entry["_section"] = section; records.append(entry)
        except Exception: continue
    wb.close()
    if not records: return pd.DataFrame()
    df = pd.DataFrame(records)
    df["Date_parsed"] = pd.to_datetime(df["Date"].str.replace("-","/"), format="%d/%m/%Y", errors="coerce")
    return df.sort_values("Date_parsed").reset_index(drop=True)

@st.cache_data(ttl=1800)
def _cached_calc_variations(_hist_fp):
    hist_df = st.session_state.get("__hist_df_raw")
    if hist_df is None or hist_df.empty: return pd.DataFrame()
    return _calc_variations_impl(hist_df)

def _calc_variations_impl(hist_df):
    if hist_df.empty or "Date" not in hist_df.columns: return pd.DataFrame()
    dates = sorted(hist_df["Date"].unique())
    if len(dates) < 2: return pd.DataFrame()
    perf_df = hist_df[hist_df["_section"]=="perf"].copy()
    qual_df = hist_df[hist_df["_section"]=="qual"].copy()
    variations = []
    for i in range(1, len(dates)):
        prev_date, curr_date = dates[i-1], dates[i]
        idx_col = "Poste de travail" if "Poste de travail" in perf_df.columns else None
        if idx_col is None: continue
        for sec_name, sdf, kpi_list in [("Performance", perf_df, QK+["Score Performance"]),("Qualite", qual_df, PK+["Score Qualite"])]:
            prev_d = sdf[sdf["Date"]==prev_date].set_index(idx_col)
            curr_d = sdf[sdf["Date"]==curr_date].set_index(idx_col)
            for poste in set(prev_d.index) & set(curr_d.index):
                for kpi in kpi_list:
                    if kpi not in prev_d.columns or kpi not in curr_d.columns: continue
                    try: pv = float(prev_d.loc[poste, kpi])
                    except Exception: continue
                    try: cv = float(curr_d.loc[poste, kpi])
                    except Exception: continue
                    diff = cv - pv
                    pct = (diff/pv*100) if pv != 0 else (100 if cv != 0 else 0)
                    if abs(diff) <= 0.5: trend = "stabilite"
                    elif diff > 0.5: trend = "hausse"
                    else: trend = "baisse"
                    variations.append({"Date precedente":prev_date,"Date actuelle":curr_date,"Poste":poste,
                        "Type":sec_name,"KPI":kpi,"Valeur precedente":round(pv,2),"Valeur actuelle":round(cv,2),
                        "Ecart":round(diff,2),"Ecart %":round(pct,2),"Tendance":trend})
    return pd.DataFrame(variations)

# ============================================================
def save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_r,ano_p_c,ano_q_r,ano_q_c,sheet_name):
    kpis_dir="kpis"; os.makedirs(kpis_dir,exist_ok=True)
    filepath=os.path.join(kpis_dir,"indicateurs_kpis.xlsx")
    sn=str(sheet_name).replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","")[:31]
    hf=Font(bold=True,color="FFFFFF",size=10); hfl=PatternFill(start_color="1E3A5F",end_color="1E3A5F",fill_type="solid")
    tf=Font(bold=True,size=12,color="1E3A5F")
    tb=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    try: wb=load_workbook(filepath)
    except Exception: wb=Workbook()
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    if sn in wb.sheetnames: del wb[sn]
    ws=wb.create_sheet(sn); rn=1
    def ws_sec(title,cols,rows,sr):
        ws.cell(row=sr,column=1,value=title).font=tf; sr+=1
        for j,c in enumerate(cols,1):
            cl=ws.cell(row=sr,column=j,value=c); cl.font=hf; cl.fill=hfl; cl.alignment=Alignment(horizontal='center'); cl.border=tb
        sr+=1
        for r in rows:
            for j,c in enumerate(cols,1):
                cl=ws.cell(row=sr,column=j,value=r.get(c,"")); cl.border=tb; cl.alignment=Alignment(horizontal='center')
            sr+=1
        return sr+1
    rn=ws_sec("INDICATEURS DE PERFORMANCE",pcols,prows,rn)
    if ano_p_c and ano_p_r: rn=ws_sec("ANOMALIES PERFORMANCE",ano_p_c,ano_p_r,rn)
    rn=ws_sec("INDICATEURS DE QUALITE",qcols,qrows,rn)
    if ano_q_c and ano_q_r: rn=ws_sec("ANOMALIES QUALITE",ano_q_c,ano_q_r,rn)
    try: wb.save(filepath)
    except Exception: pass

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
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}.stl.s{border-left-color:#d69e2e}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.3px;padding:5px 6px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw.st thead th{background:linear-gradient(135deg,#975a16,#d69e2e)}
    .tw tbody td{padding:4px 6px;border-bottom:1px solid #edf2f7;white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important;font-size:12px!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important;font-size:12px!important}
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
    .cg{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .cg>div{background:#fff;border-radius:var(--r);padding:8px 10px;border:1px solid var(--b)}
    .cg .ct{font-size:13px;font-weight:700;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid var(--b)}
    .cgr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f7fafc}
    .cgr:last-child{border:none}
    .cgr .rk{width:18px;font-weight:800;text-align:center}
    .cgr .pn{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cgr .ps{font-weight:800;min-width:55px;text-align:right}
    .rank-card{background:#fff;border-radius:var(--r);padding:12px 16px;border:1px solid var(--b);box-shadow:0 2px 8px rgba(0,0,0,.04)}
    .rank-card .rank-title{font-size:15px;font-weight:800;margin-bottom:8px;padding-bottom:5px;border-bottom:2px solid var(--b)}
    .rank-row{display:flex;align-items:center;padding:5px 0;font-size:13px;border-bottom:1px solid #f7fafc}
    .rank-row:last-child{border:none}
    .rank-row .rank-num{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:13px;color:#fff;margin-right:10px;flex-shrink:0}
    .rank-row .rank-name{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .rank-row .rank-score{font-weight:900;min-width:70px;text-align:right}
    .dgrid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:8px 14px;font-weight:700;font-size:15px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label,div[data-testid="stSidebar"] .stCheckbox label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:3px 8px;margin-bottom:3px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:14px;color:#718096;font-size:14px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:17px}.cg,.dgrid{grid-template-columns:1fr}.car .cal{width:120px}.gbr-l{width:100px}}
    </style>""",unsafe_allow_html=True)

# ============================================================
def main():
    try: locale.setlocale(locale.LC_ALL,'fr_FR.UTF-8')
    except Exception:
        try: locale.setlocale(locale.LC_ALL,'fr_FR')
        except Exception: pass
    inject_custom_css()
    fichier_date = get_date_from_file()

    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        c = random.choice(CONSIGNES_HSE)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:46px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SECURITE</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:22px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Securite - Sante - Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:32px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:36px;font-weight:900">Aucun travail n'est plus urgent que la securite</h2>
        <div style="margin-top:40px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden"><div style="width:100%%;height:100%%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 5.5s ease-in-out forwards"></div></div>
        <style>@keyframes ld{from{width:0}to{width:100%%}}</style></div>"""%c,unsafe_allow_html=True)
        time.sleep(6); st.session_state.hse_affiche = True; st.rerun(); st.stop()

    def ks(v, c):
        try: val = float(v)
        except Exception: return ""
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
        except Exception: return ""
        return "background:#c6efce;color:#006100;font-weight:700" if val>=90 else ("background:#ffeb9c;color:#9c6500;font-weight:700" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:700")
    def kas(v):
        try: val = int(v)
        except Exception: return ""
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
    def is_lb(k): return k in LOWER_BETTER

    def html_table(rows, cols, tc, sc_col=None):
        h = '<table class="tw %s"><thead><tr>'%tc + ''.join('<th>%s</th>'%c for c in cols) + '</tr></thead><tbody>'
        for r in rows:
            rc = "cb" if r.get("_t")=="cible" else ("tr" if r.get("_t")=="total" else "")
            h += '<tr class="%s">'%rc
            for c in cols:
                v = r.get(c, "")
                if r.get("_t") == "cible": h += '<td>%s</td>'%v
                else:
                    s = cs(v) if sc_col and c in sc_col else ks(v, c)
                    h += '<td style="%s">%s</td>'%(s or "", v)
            h += '</tr>'
        return h + '</tbody></table>'
    def html_ano(rows, cols):
        h = '<table class="tw at"><thead><tr>' + ''.join('<th>%s</th>'%c for c in cols) + '</tr></thead><tbody>'
        for r in rows:
            h += '<tr class="%s">'%("tr" if r.get("_t")=="total" else "")
            for c in cols:
                v = r.get(c, ""); h += '<td style="%s">%s</td>'%(kas(v) or "", v)
            h += '</tr>'
        return h + '</tbody></table>'
    def html_actions_table(kpi_list, actuals, targets, act_map):
        h = '<table class="tw st"><thead><tr><th>KPI</th><th>Valeur</th><th>Cible</th><th>Ecart</th><th>Statut</th><th>Action Recommandée</th></tr></thead><tbody>'
        for k in kpi_list:
            av = actuals.get(k, 0); tv = targets.get(k, 100); diff = av - tv
            met = av <= tv if is_lb(k) else av >= tv
            status = "ATTEINT" if met else "NON ATTEINT"
            st_s = "background:#c6efce;color:#006100;font-weight:700" if met else "background:#ffc7ce;color:#9c0006;font-weight:700"
            ec_clr = "#276749" if met else "#c53030"
            action = "Objectif atteint ✓" if met else act_map.get(k, "")
            h += '<tr><td style="font-weight:600">%s</td><td>%.1f%%</td><td>%.0f%%</td><td style="color:%s;font-weight:700">%+.1f%%</td><td style="%s">%s</td><td style="color:#4a5568">%s</td></tr>'%(k, av, tv, ec_clr, diff, st_s, status, action)
        return h + '</tbody></table>'
    def html_classement(scores, accent):
        sp = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        met_p = [(p, s) for p, s in sp if s >= 80]; not_p = [(p, s) for p, s in sp if s < 80]
        t5 = met_p[:5]; b5 = not_p[-5:] if len(not_p) > 5 else not_p
        h = '<div class="cg"><div><div class="ct" style="color:#38a169">Top 5 — Objectif Atteint</div>'
        if t5:
            for i, (p, s) in enumerate(t5):
                h += '<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(accent, i+1, p, cs("%.2f"%s), s)
        else: h += '<div style="padding:6px;font-size:12px;color:#718096">Aucun poste</div>'
        h += '</div><div><div class="ct" style="color:#e53e3e">Bottom 5 — Non Atteint</div>'
        if b5:
            for i, (p, s) in enumerate(reversed(b5)):
                h += '<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(len(b5)-i, p, cs("%.2f"%s), s)
        else: h += '<div style="padding:6px;font-size:12px;color:#38a169">Tous atteints</div>'
        h += '</div></div>'; return h
    def html_kpi_bars(kpi_list, actuals, targets, title, color_ok, color_fail):
        h = '<div class="ca"><div class="ct" style="color:%s">%s</div>'%(color_ok, title)
        for k in kpi_list:
            av = actuals.get(k, 0); tv = targets.get(k, 100); met = av <= tv if is_lb(k) else av >= tv
            bw = min(max(av, 0), 100); bg = color_ok if met else color_fail
            h += '<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>'%(k, bw, bg, av)
        return h + '</div>'
    def html_grouped_bars(posts, pscores, qscores, title):
        h = '<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>'%title
        h += '<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span></div>'
        for p in sorted(posts, key=lambda x: (pscores.get(x,0)+qscores.get(x,0))/2, reverse=True):
            pv, qv = pscores.get(p, 0), qscores.get(p, 0)
            h += '<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f gb-p" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w"><div class="gbr-f gb-q" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div></div></div>'%(p, min(max(pv,0),100), pv, min(max(qv,0),100), qv)
        return h + '</div>'

    def anl_pie_chart(data, names_col, values_col, title, colors=None, threshold=5.0):
        if data.empty or data[values_col].sum() == 0: return None
        df = data.copy()
        total = df[values_col].sum()
        df['_pct'] = df[values_col] / total * 100
        small_mask = df['_pct'] < threshold
        has_small = small_mask.any()
        if not has_small:
            fig = px.pie(df, names=names_col, values=values_col, title=title,
                         color_discrete_sequence=colors or px.colors.qualitative.Set2)
            fig.update_traces(textposition='inside', textinfo='percent+label', textfont_size=12)
            fig.update_layout(margin=dict(t=50,b=20,l=20,r=20), height=450, autosize=True,
                             title_font_size=15, legend=dict(font_size=11, orientation="h", yanchor="bottom", y=-0.12))
            return fig
        large_df = df[~small_mask].copy()
        small_df = df[small_mask].copy()
        others_row = pd.DataFrame({
            names_col: ['Autres (' + str(len(small_df)) + ' secteurs)'],
            values_col: [small_df[values_col].sum()], '_pct': [small_df['_pct'].sum()]
        })
        outer_df = pd.concat([large_df, others_row], ignore_index=True)
        base_colors = colors or px.colors.qualitative.Set2
        outer_colors = list(base_colors[:len(large_df)]) + ['#a0aec0']
        inner_colors = list(px.colors.qualitative.Pastel[:len(small_df)])
        if len(inner_colors) < len(small_df):
            inner_colors = (inner_colors * ((len(small_df)//len(inner_colors))+1))[:len(small_df)]
        fig = go.Figure()
        fig.add_trace(go.Pie(
            labels=outer_df[names_col].tolist(), values=outer_df[values_col].tolist(),
            name="Principal", marker_colors=outer_colors,
            domain={'x': [0.0, 1.0], 'y': [0.22, 1.0]},
            textinfo='percent+label', textposition='inside', textfont_size=11, hole=0.42,
            pull=[0.04 if 'Autres' in str(lbl) else 0 for lbl in outer_df[names_col]]
        ))
        fig.add_trace(go.Pie(
            labels=small_df[names_col].tolist(), values=small_df[values_col].tolist(),
            name="Détail Autres", marker_colors=inner_colors,
            domain={'x': [0.18, 0.82], 'y': [0.0, 0.22]},
            textinfo='percent+label', textposition='inside', textfont_size=9, hole=0.3
        ))
        fig.update_layout(
            title=dict(text=title + ' <span style="font-size:11px;color:#a0aec0">(pie of 2 pie — secteurs < ' + str(threshold) + '%)</span>', font=dict(size=14)),
            height=580, autosize=True, margin=dict(t=60, b=10, l=10, r=10),
            legend=dict(font_size=10, orientation="h", yanchor="bottom", y=-0.02),
            annotations=[dict(text="Détail<br>secteurs minces", x=0.5, y=0.11, font=dict(size=9, color="#718096"), showarrow=False)],
            showlegend=True
        )
        return fig

    def export_btn(df, filename):
        buf = io.BytesIO(); df.to_excel(buf, index=False, engine='openpyxl'); buf.seek(0)
        st.download_button("📥 Exporter Excel", data=buf, file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""<div style="padding:10px 0 4px 0"><div style="font-size:22px;margin-bottom:2px">⚙️</div><div style="font-size:14px;font-weight:800;color:white">Filtres & Parametres</div><div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Configuration</div></div>""",unsafe_allow_html=True)
        st.markdown("---")
        show_filters = st.checkbox("Afficher les filtres", value=True, key="show_filters")
        if show_filters:
            unf = st.toggle("📁 Charger nouveaux fichiers", value=False, key="tf")
            ot_f = av_f = None; apm = []
            if unf:
                ot_f = st.file_uploader("Fichier OT", type=["xlsx"], key="uot")
                av_f = st.file_uploader("Fichier AVIS", type=["xlsx"], key="uav")
            else:
                if os.path.exists("ot.xlsx"):
                    try:
                        _t = excr(pd.read_excel("ot.xlsx"))
                        apm = sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                    except Exception: pass
                st.markdown("""<div style="background:rgba(255,255,255,.1);padding:6px 10px;border-radius:6px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Donnees</div><div style="font-size:14px;color:white;font-weight:600;margin-top:2px">📅 %s</div></div>"""%fichier_date,unsafe_allow_html=True)
            st.markdown("---"); st.markdown("**🎯 Postes**")
            sp = st.multiselect("Poste", ["All"]+apm, ["All"], key="sp")
            st.markdown("**🏭 Atelier**")
            sa = st.multiselect("Atelier", ["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)"], ["All"], key="sa")
            st.markdown("**🏢 Division**")
            sd = st.multiselect("Division", ["All","SF1","SF2"], ["All"], key="sd")
            st.markdown("---"); st.markdown("**📅 Periode**")
            dr = st.date_input("Date debut planifiee", value=(datetime(2025,1,1).date(), datetime.today().date()), format="DD/MM/YYYY", key="dr")
        else:
            unf = False; ot_f = av_f = None; apm = []; sp = ["All"]; sa = ["All"]; sd = ["All"]
            dr = (datetime(2025,1,1).date(), datetime.today().date())
            if os.path.exists("ot.xlsx"):
                try:
                    _t = excr(pd.read_excel("ot.xlsx"))
                    apm = sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                except Exception: pass

    # ===================== DATA LOADING =====================
    data_ready = not unf or (ot_f is not None and av_f is not None)
    if data_ready:
        with st.spinner("Chargement des données..."):
            raw_ot, raw_av = load_raw_data(unf, ot_f, av_f)

            if not apm:
                apm = sorted(raw_ot[raw_ot["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
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

            df = raw_ot[(raw_ot["Poste travail princ."].isin(vp))&(raw_ot["Date de début planifiée"].between(sdt, edt))].copy()
            df = excr(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)].drop_duplicates())
            if "Statut système" in df.columns:
                df["Statut OT"] = df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            avdf = raw_av[raw_av["Poste travail princ."].isin(vp)].copy()
            av_date_col = None
            for col in ["Créé le", "Début souhaité", "Date de la clôture"]:
                if col in avdf.columns and avdf[col].notna().any():
                    av_date_col = col; break
            if av_date_col:
                avdf = avdf[avdf[av_date_col].between(sdt, edt)]
            avdf = excr(avdf[(avdf["Ordre"].isna())|(avdf["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates())

            df_dash = raw_ot[raw_ot["Poste travail princ."].isin(vp)].copy()
            df_dash = excr(df_dash[df_dash["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)].drop_duplicates())
            if "Statut système" in df_dash.columns:
                df_dash["Statut OT"] = df_dash["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]

            now = pd.Timestamp.now()

            fp = _df_fingerprint(df, avdf, vp, now.timestamp())
            res = _cached_calc_kpis(fp, df, avdf, now.timestamp(), tuple(sorted(vp)))

            fp_d = _df_fingerprint(df_dash, avdf, vp, now.timestamp())
            res_d = _cached_calc_kpis(fp_d, df_dash, avdf, now.timestamp(), tuple(sorted(vp)))

            ckdf = res['ckdf']; dfp = res['dfp']
            ckdf_d = res_d['ckdf']

            pa = {k: round(ckdf[k].mean(), 2) for k in QK}
            qa = {k: round(ckdf[k].mean(), 2) for k in PK}
            pa_d = {k: round(ckdf_d[k].mean(), 2) for k in QK}
            qa_d = {k: round(ckdf_d[k].mean(), 2) for k in PK}

            pscores = {}; qscores = {}
            for poste in ckdf.index:
                r = ckdf.loc[poste]
                pscores[poste] = (sum(gscore(k, r[k], CIBLE[k]) for k in QK if k in r.index)/len(QK)*100) if QK else 0
                qscores[poste] = (sum(gscore(k, r[k], CIBLE[k]) for k in PK if k in r.index)/len(PK)*100) if PK else 0
            pscores_d = {}; qscores_d = {}
            for poste in ckdf_d.index:
                r = ckdf_d.loc[poste]
                pscores_d[poste] = (sum(gscore(k, r[k], CIBLE[k]) for k in QK if k in r.index)/len(QK)*100) if QK else 0
                qscores_d[poste] = (sum(gscore(k, r[k], CIBLE[k]) for k in PK if k in r.index)/len(PK)*100) if PK else 0

            sub_p = {
                "TAUX_REALISATION_CORRECTIF/PT": lambda d: d[(d["Nº appel pl.entret."].fillna(0)==0)&(d["Contient SOPL"]==1)&(~d["Statut OT"].isin(["CLOT","TCLO"]))],
                "OT préparation <1 mois": lambda d: d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]!="<1 mois")],
                "OT préparation >3 mois": lambda d: d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]==">3 mois")],
                "OT planification <1 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]!="<1 mois")],
                "OT planification >3 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]==">3 mois")],
                "OT exécution <1 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]!="<1 mois")],
                "OT exécution >3 mois": lambda d: d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]==">3 mois")]
            }
            sub_q = {
                "OT LANC ESTIME": lambda d: d[(d["Statut OT"]=="LANC")&(d["OT LANC ESTIME"]=="NON")],
                "Backlog préparation caractérisé": lambda d: d[(d["Statut OT"]=="CRÉÉ")&(d["Backlog preparation"]=="NON CARACTERISE")],
                "Backlog planification caractérisé": lambda d: d[(d["Statut OT"]=="LANC")&(d["Backlog planification"]=="NON CARACTERISE")],
                "OT CONFIME": lambda d: d[(d["Statut OT"].str.contains("CLO",na=False))&(d["OT CONFIME"]=="NON")],
                "OT_COR_EGAL": lambda d: d[(d["Statut OT"].str.contains("CLO",na=False))&(d["OT_COR_EGAL"]=="NON")],
                "appel avis approuvé": lambda d: pd.DataFrame()
            }

            ano_p_rows = []; ano_q_rows = []
            for k, fn in sub_p.items():
                try:
                    ad = fn(dfp)
                    if not ad.empty:
                        for _, row in ad.iterrows():
                            ano_p_rows.append({"Poste de travail": get_poste_name(row), "KPI": k, "Ordre": row.get("Ordre",""), "Description": row.get("Description","")})
                except Exception: pass
            for k, fn in sub_q.items():
                try:
                    ad = fn(dfp)
                    if not ad.empty:
                        for _, row in ad.iterrows():
                            ano_q_rows.append({"Poste de travail": get_poste_name(row), "KPI": k, "Ordre": row.get("Ordre",""), "Description": row.get("Description","")})
                except Exception: pass
            try:
                avf = res['avf']
                if not avf.empty:
                    for _, row in avf.iterrows():
                        ano_q_rows.append({"Poste de travail": get_poste_name(row), "KPI": "appel avis approuvé", "Ordre": "", "Avis": row.get("Avis","")})
            except Exception: pass

            ano_p_counts = pd.DataFrame(ano_p_rows).groupby(["Poste de travail","KPI"]).size().reset_index(name="Nombre") if ano_p_rows else pd.DataFrame(columns=["Poste de travail","KPI","Nombre"])
            ano_q_counts = pd.DataFrame(ano_q_rows).groupby(["Poste de travail","KPI"]).size().reset_index(name="Nombre") if ano_q_rows else pd.DataFrame(columns=["Poste de travail","KPI","Nombre"])

            hist_path = os.path.join("kpis", "indicateurs_kpis.xlsx")
            hist_df = _cached_load_historical(hist_path) if os.path.exists(hist_path) else pd.DataFrame()
            st.session_state["__hist_df_raw"] = hist_df
            hist_fp = hashlib.md5((str(len(hist_df)) + str(hist_df.columns.tolist())).encode()).hexdigest() if not hist_df.empty else "empty"
            var_df = _cached_calc_variations(hist_fp)

            if not var_df.empty:
                journal = var_df.copy()
                journal["Significatif"] = journal["Ecart %"].abs() >= 5
                journal = journal[journal["Significatif"]].copy()
                journal["Sens"] = journal.apply(lambda r: "Amelioration" if ((r["Tendance"]=="hausse" and r["KPI"] not in LOWER_BETTER) or (r["Tendance"]=="baisse" and r["KPI"] in LOWER_BETTER)) else "Degradation", axis=1)
                journal = journal.sort_values(["Date actuelle","Sens","Ecart %"], ascending=[True, False, False])
            else:
                journal = pd.DataFrame()

            if not var_df.empty:
                scores_var = {}
                for poste in var_df["Poste"].unique():
                    pv = var_df[var_df["Poste"]==poste]
                    scores_var[poste] = sum((-r["Ecart %"] if r["KPI"] in LOWER_BETTER else r["Ecart %"]) for _, r in pv.iterrows())
                ranked = sorted(scores_var.items(), key=lambda x: x[1], reverse=True)
                top5 = pd.DataFrame(ranked[:5], columns=["Poste","Score variation"]) if ranked else pd.DataFrame()
                bot5 = pd.DataFrame(ranked[-5:][::-1], columns=["Poste","Score variation"]) if ranked else pd.DataFrame()
            else:
                top5 = pd.DataFrame(columns=["Poste","Score variation"])
                bot5 = pd.DataFrame(columns=["Poste","Score variation"])

            dist_atelier = df_dash.groupby(df_dash["Poste travail princ."].apply(get_atelier))["Ordre"].count().reset_index()
            dist_atelier.columns = ["Atelier", "Nombre"]
            dist_metier = df_dash.groupby(df_dash["Poste travail princ."].apply(get_metier))["Ordre"].count().reset_index()
            dist_metier.columns = ["Metier", "Nombre"]
            dist_division = df_dash.groupby(df_dash["Poste travail princ."].apply(get_division))["Ordre"].count().reset_index()
            dist_division.columns = ["Division", "Nombre"]
            dist_statut = df_dash.groupby("Statut OT")["Ordre"].count().reset_index()
            dist_statut.columns = ["Statut", "Nombre"]

            pcols = ["Poste de travail"] + QK + ["Score Performance"]
            prows = []
            for poste in ckdf.index:
                r = ckdf.loc[poste]; row = {"Poste de travail": poste}
                for k in QK: row[k] = round(r[k], 2) if k in r.index else 0
                row["Score Performance"] = round(pscores.get(poste, 0), 2)
                prows.append(row)
            cible_row = {"Poste de travail": "CIBLE"}
            for k in QK: cible_row[k] = CIBLE[k]
            cible_row["Score Performance"] = 80; prows.append(cible_row)
            tot_row = {"Poste de travail": "Moyenne"}
            for k in QK: tot_row[k] = round(pa.get(k, 0), 2)
            tot_row["Score Performance"] = round(np.mean(list(pscores.values())), 2) if pscores else 0
            prows.append(tot_row)

            qcols = ["Poste de travail"] + PK + ["Score Qualite"]
            qrows = []
            for poste in ckdf.index:
                r = ckdf.loc[poste]; row = {"Poste de travail": poste}
                for k in PK: row[k] = round(r[k], 2) if k in r.index else 0
                row["Score Qualite"] = round(qscores.get(poste, 0), 2)
                qrows.append(row)
            cible_row_q = {"Poste de travail": "CIBLE"}
            for k in PK: cible_row_q[k] = CIBLE[k]
            cible_row_q["Score Qualite"] = 80; qrows.append(cible_row_q)
            tot_row_q = {"Poste de travail": "Moyenne"}
            for k in PK: tot_row_q[k] = round(qa.get(k, 0), 2)
            tot_row_q["Score Qualite"] = round(np.mean(list(qscores.values())), 2) if qscores else 0
            qrows.append(tot_row_q)

            ano_p_exp = ano_p_counts.pivot_table(index="Poste de travail", columns="KPI", values="Nombre", aggfunc="sum", fill_value=0).reset_index() if not ano_p_counts.empty else pd.DataFrame()
            ano_q_exp = ano_q_counts.pivot_table(index="Poste de travail", columns="KPI", values="Nombre", aggfunc="sum", fill_value=0).reset_index() if not ano_q_counts.empty else pd.DataFrame()

            save_kpis_to_excel(prows, pcols, qrows, qcols,
                               ano_p_exp.to_dict('records') if not ano_p_exp.empty else [],
                               list(ano_p_exp.columns) if not ano_p_exp.empty else [],
                               ano_q_exp.to_dict('records') if not ano_q_exp.empty else [],
                               list(ano_q_exp.columns) if not ano_q_exp.empty else [],
                               fichier_date)

            # ===================== RENDU =====================
            avg_p_score = round(np.mean(list(pscores.values())), 1) if pscores else 0
            avg_q_score = round(np.mean(list(qscores.values())), 1) if qscores else 0
            avg_g_score = round((avg_p_score + avg_q_score) / 2, 1)

            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📊 Dashboard", "⚡ Performance", "✅ Qualité", "⚠️ Anomalies", "📈 Suivi Amélioration", "📥 Export"
            ])

            with tab1:
                st.markdown('<div class="mh"><h1>📊 Dashboard KPI — Maintenance</h1><span class="db">📅 %s</span></div>'%fichier_date, unsafe_allow_html=True)
                st.markdown('<div class="cr"><div class="cc c1"><div class="cv">%.1f%%</div><div class="cl">Score Performance</div></div><div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Qualite</div></div><div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Global</div></div><div class="cc c4"><div class="cv">%d</div><div class="cl">Postes Suivis</div></div></div>'%(avg_p_score, avg_q_score, avg_g_score, len(vp)), unsafe_allow_html=True)
                st.markdown('<div class="stl p">Scores par Poste</div>', unsafe_allow_html=True)
                st.markdown(html_grouped_bars(vp, pscores, qscores, "Performance vs Qualité par Poste"), unsafe_allow_html=True)
                st.markdown('<div class="stl c">Classement</div>', unsafe_allow_html=True)
                st.markdown(html_classement(pscores, "#276749"), unsafe_allow_html=True)
                st.markdown('<div class="stl q">Distributions (Pie of 2 Pie si secteurs minces)</div>', unsafe_allow_html=True)
                pie_cols = st.columns(4)
                with pie_cols[0]:
                    fig1 = anl_pie_chart(dist_atelier, "Atelier", "Nombre", "Par Atelier")
                    if fig1: st.plotly_chart(fig1, use_container_width=True)
                with pie_cols[1]:
                    fig2 = anl_pie_chart(dist_metier, "Metier", "Nombre", "Par Métier")
                    if fig2: st.plotly_chart(fig2, use_container_width=True)
                with pie_cols[2]:
                    fig3 = anl_pie_chart(dist_division, "Division", "Nombre", "Par Division")
                    if fig3: st.plotly_chart(fig3, use_container_width=True)
                with pie_cols[3]:
                    fig4 = anl_pie_chart(dist_statut, "Statut", "Nombre", "Par Statut OT")
                    if fig4: st.plotly_chart(fig4, use_container_width=True)

            with tab2:
                st.markdown('<div class="stl p">Indicateurs de Performance — Barres</div>', unsafe_allow_html=True)
                st.markdown(html_kpi_bars(QK, pa, CIBLE, "Performance Globale (moyenne)", "#38a169", "#e53e3e"), unsafe_allow_html=True)
                st.markdown('<div class="stl p">Detail par Poste</div>', unsafe_allow_html=True)
                st.markdown(html_table(prows, pcols, "pt", {"Score Performance"}), unsafe_allow_html=True)

            with tab3:
                st.markdown('<div class="stl q">Indicateurs de Qualité — Barres</div>', unsafe_allow_html=True)
                st.markdown(html_kpi_bars(PK, qa, CIBLE, "Qualité Globale (moyenne)", "#3182ce", "#e53e3e"), unsafe_allow_html=True)
                st.markdown('<div class="stl q">Detail par Poste</div>', unsafe_allow_html=True)
                st.markdown(html_table(qrows, qcols, "qt", {"Score Qualite"}), unsafe_allow_html=True)

            with tab4:
                st.markdown('<div class="stl a">Anomalies Performance</div>', unsafe_allow_html=True)
                if not ano_p_counts.empty:
                    ano_p_pivot = ano_p_counts.pivot_table(index="Poste de travail", columns="KPI", values="Nombre", aggfunc="sum", fill_value=0).reset_index()
                    ano_p_pivot["Total"] = ano_p_pivot.iloc[:, 1:].sum(axis=1)
                    ano_p_pivot = ano_p_pivot.sort_values("Total", ascending=False)
                    tot_p = {"Poste de travail": "Total general", "Total": int(ano_p_pivot["Total"].sum())}
                    for c in ano_p_pivot.columns:
                        if c not in ("Poste de travail","Total"): tot_p[c] = int(ano_p_pivot[c].sum())
                    tot_p["_t"] = "total"
                    ano_p_html_rows = [dict(r, _t="") for _, r in ano_p_pivot.iterrows()] + [tot_p]
                    st.markdown(html_ano(ano_p_html_rows, list(ano_p_pivot.columns)), unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Aucune anomalie performance detectee</div>', unsafe_allow_html=True)
                st.markdown('<div class="stl a">Anomalies Qualité</div>', unsafe_allow_html=True)
                if not ano_q_counts.empty:
                    ano_q_pivot = ano_q_counts.pivot_table(index="Poste de travail", columns="KPI", values="Nombre", aggfunc="sum", fill_value=0).reset_index()
                    ano_q_pivot["Total"] = ano_q_pivot.iloc[:, 1:].sum(axis=1)
                    ano_q_pivot = ano_q_pivot.sort_values("Total", ascending=False)
                    tot_q = {"Poste de travail": "Total general", "Total": int(ano_q_pivot["Total"].sum())}
                    for c in ano_q_pivot.columns:
                        if c not in ("Poste de travail","Total"): tot_q[c] = int(ano_q_pivot[c].sum())
                    tot_q["_t"] = "total"
                    ano_q_html_rows = [dict(r, _t="") for _, r in ano_q_pivot.iterrows()] + [tot_q]
                    st.markdown(html_ano(ano_q_html_rows, list(ano_q_pivot.columns)), unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Aucune anomalie qualite detectee</div>', unsafe_allow_html=True)

            with tab5:
                st.markdown('<div class="mh"><h1>📈 Suivi d\'Amélioration</h1><span class="db">Historique & Tendances</span></div>', unsafe_allow_html=True)
                if hist_df.empty:
                    st.markdown('<div class="es">Aucune donnée historique disponible.<br>Les données seront accumulées à chaque sauvegarde Excel dans le dossier <code>kpis/</code>.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="stl s">Journal des Variations Significatives (|écart| ≥ 5%)</div>', unsafe_allow_html=True)
                    if not journal.empty:
                        jh = '<table class="tw st"><thead><tr><th>Date</th><th>Poste</th><th>Type</th><th>KPI</th><th>Préc.</th><th>Act.</th><th>Ecart %%</th><th>Tendance</th><th>Sens</th></tr></thead><tbody>'
                        for _, r in journal.iterrows():
                            trend_icon = "🔺" if r["Tendance"]=="hausse" else ("🔻" if r["Tendance"]=="baisse" else "➖")
                            sens_s = "background:#c6efce;color:#006100;font-weight:700" if r["Sens"]=="Amelioration" else "background:#ffc7ce;color:#9c0006;font-weight:700"
                            jh += '<tr><td>%s</td><td>%s</td><td>%s</td><td style="font-weight:600">%s</td><td>%.1f</td><td>%.1f</td><td style="font-weight:700">%.1f%%</td><td>%s %s</td><td style="%s">%s</td></tr>'%(
                                r["Date actuelle"], r["Poste"], r["Type"], r["KPI"],
                                r["Valeur precedente"], r["Valeur actuelle"], r["Ecart %"],
                                trend_icon, r["Tendance"], sens_s, r["Sens"])
                        jh += '</tbody></table>'
                        st.markdown(jh, unsafe_allow_html=True)
                        if st.button("📥 Exporter Journal", key="exp_journal"):
                            export_btn(journal, "journal_variations.xlsx")
                    else:
                        st.markdown('<div class="es">Aucune variation significative detectee</div>', unsafe_allow_html=True)
                    st.markdown('<div class="stl c">Classement par Progression</div>', unsafe_allow_html=True)
                    cls_cols = st.columns(2)
                    with cls_cols[0]:
                        st.markdown('<div class="rank-card"><div class="rank-title" style="color:#38a169">🏆 Top 5 Progression</div>', unsafe_allow_html=True)
                        if not top5.empty:
                            for i, (_, r) in enumerate(top5.iterrows()):
                                clr = "#38a169" if i==0 else ("#48bb78" if i==1 else ("#68d391" if i==2 else "#9ae6b4"))
                                st.markdown('<div class="rank-row"><div class="rank-num" style="background:%s">%s</div><div class="rank-name">%s</div><div class="rank-score" style="color:#276749">%+.1f</div></div>'%(clr, i+1, r["Poste"], r["Score variation"]), unsafe_allow_html=True)
                        else: st.markdown('<div style="padding:8px;color:#718096;font-size:12px">Aucune donnée</div>', unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)
                    with cls_cols[1]:
                        st.markdown('<div class="rank-card"><div class="rank-title" style="color:#e53e3e">⚠️ Top 5 Régression</div>', unsafe_allow_html=True)
                        if not bot5.empty:
                            for i, (_, r) in enumerate(bot5.iterrows()):
                                clr = "#e53e3e" if i==0 else ("#fc8181" if i==1 else ("#feb2b2" if i==2 else "#fed7d7"))
                                st.markdown('<div class="rank-row"><div class="rank-num" style="background:%s">%s</div><div class="rank-name">%s</div><div class="rank-score" style="color:#c53030">%+.1f</div></div>'%(clr, i+1, r["Poste"], r["Score variation"]), unsafe_allow_html=True)
                        else: st.markdown('<div style="padding:8px;color:#718096;font-size:12px">Aucune donnée</div>', unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)
                    st.markdown('<div class="stl q">Tendances Historiques par KPI</div>', unsafe_allow_html=True)
                    kpi_sel = st.multiselect("Sélectionner les KPIs", ALL_KPI, default=QK[:3], key="kpi_trend_sel")
                    if kpi_sel:
                        for kpi in kpi_sel:
                            kpi_hist = var_df[var_df["KPI"]==kpi].copy()
                            if kpi_hist.empty:
                                st.markdown('<div class="es">Aucun historique pour %s</div>'%kpi, unsafe_allow_html=True); continue
                            fig_t = go.Figure()
                            for poste in kpi_hist["Poste"].unique()[:10]:
                                pd_kpi = kpi_hist[kpi_hist["Poste"]==poste].sort_values("Date actuelle")
                                clr = "#276749" if "Amelioration" in pd_kpi["Sens"].values[-1:] else "#c53030"
                                fig_t.add_trace(go.Scatter(x=pd_kpi["Date actuelle"], y=pd_kpi["Valeur actuelle"],
                                    mode='lines+markers+text', name=poste, line=dict(color=clr, width=2),
                                    marker=dict(size=6), text=pd_kpi["Valeur actuelle"].apply(lambda x: "%.1f"%x),
                                    textposition="top center", textfont_size=9))
                            fig_t.update_layout(height=350, margin=dict(t=40,b=20,l=40,r=20),
                                title=dict(text=kpi, font_size=14, font_color="#1e3a5f"),
                                xaxis_title="Date", yaxis_title="Valeur",
                                legend=dict(orientation="h", yanchor="bottom", y=-0.3, font_size=10))
                            st.plotly_chart(fig_t, use_container_width=True)
                    st.markdown('<div class="stl a">Plan d\'Actions Recommandées</div>', unsafe_allow_html=True)
                    st.markdown(html_actions_table(ALL_KPI, {**pa, **qa}, CIBLE, ACT_MAP), unsafe_allow_html=True)
                    with st.expander("📋 Voir l'historique complet des variations"):
                        if not var_df.empty:
                            st.dataframe(var_df.sort_values(["Date actuelle","Poste","KPI"]), use_container_width=True, height=400)
                            export_btn(var_df, "variations_completes.xlsx")
                        else: st.markdown('<div class="es">Aucune variation</div>', unsafe_allow_html=True)

            with tab6:
                st.markdown('<div class="mh"><h1>📥 Export des Données</h1></div>', unsafe_allow_html=True)
                st.markdown('<div class="stl p">Performance</div>', unsafe_allow_html=True)
                export_btn(pd.DataFrame(prows), "performance_kpis.xlsx")
                st.markdown('<div class="stl q">Qualité</div>', unsafe_allow_html=True)
                export_btn(pd.DataFrame(qrows), "qualite_kpis.xlsx")
                st.markdown('<div class="stl a">Anomalies Performance</div>', unsafe_allow_html=True)
                if not ano_p_counts.empty: export_btn(ano_p_counts, "anomalies_performance.xlsx")
                else: st.markdown('<div class="es">Aucune anomalie</div>', unsafe_allow_html=True)
                st.markdown('<div class="stl a">Anomalies Qualité</div>', unsafe_allow_html=True)
                if not ano_q_counts.empty: export_btn(ano_q_counts, "anomalies_qualite.xlsx")
                else: st.markdown('<div class="es">Aucune anomalie</div>', unsafe_allow_html=True)
                if not var_df.empty:
                    st.markdown('<div class="stl s">Variations Historiques</div>', unsafe_allow_html=True)
                    export_btn(var_df, "variations_historiques.xlsx")
                if not journal.empty:
                    st.markdown('<div class="stl s">Journal</div>', unsafe_allow_html=True)
                    export_btn(journal, "journal_amelioration.xlsx")
                st.markdown('<div class="stl c">Données Brutes (filtrées)</div>', unsafe_allow_html=True)
                export_btn(df, "donnees_brutes_filtrees.xlsx")

    else:
        st.markdown("""<div style="text-align:center;padding:80px;color:#718096">
        <div style="font-size:64px;margin-bottom:20px">📁</div>
        <h2 style="font-size:24px;font-weight:700;color:#1e3a5f">Veuillez charger les fichiers OT et AVIS</h2>
        <p style="font-size:16px;margin-top:10px">Utilisez le panneau de gauche pour activer le chargement de nouveaux fichiers.</p>
        </div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
