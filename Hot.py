# -*- coding: utf-8 -*-
# VERSION V3.0 - Admin KPIs & Page Code de Calcul Sécurisée
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os, json
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(layout="wide", page_title="Dashboard KPI V3")

# ==========================================
# 1. SYSTÈME DE CONFIGURATION DYNAMIQUE (JSON)
# ==========================================
CONFIG_FILE = "config_kpis.json"
ACCESS_CODE = "96800221"

def get_default_config():
    return {
        "access_code": ACCESS_CODE,
        "kpis": {
            "TAUX_REALISATION_CORRECTIF/PT": {
                "type": "statut_count", 
                "numerateurs": ["CLOT", "TCLO"], # CLOT AJOUTÉ ICI
                "denominateurs": ["CLOT", "CRÉÉ", "LANC", "TCLO"],
                "cible": 85.0, "seuil_score": 80.0, "lower_better": False, "categorie": "Performance",
                "formule_texte": "((CLOT + TCLO) / (CLOT + CRÉÉ + LANC + TCLO)) * 100"
            },
            "OT préparation <1 mois": {"type": "hardcoded", "cible": 80.0, "seuil_score": 75.0, "lower_better": False, "categorie": "Performance", "formule_texte": "(OT CRÉÉ âge ≤ 1 mois / Total CRÉÉ) * 100"},
            "OT préparation >3 mois": {"type": "hardcoded", "cible": 5.0, "seuil_score": 5.0, "lower_better": True, "categorie": "Performance", "formule_texte": "(OT CRÉÉ âge ≥ 3 mois / Total CRÉÉ) * 100"},
            "OT préparation 1mois< <3mois": {"type": "hardcoded", "cible": 15.0, "seuil_score": 15.0, "lower_better": True, "categorie": "Performance", "formule_texte": "(OT CRÉÉ 1 < âge < 3 / Total CRÉÉ) * 100"},
            "OT planification <1 mois": {"type": "hardcoded", "cible": 80.0, "seuil_score": 75.0, "lower_better": False, "categorie": "Performance", "formule_texte": "(OT LANC sans SOPL âge ≤ 1 / Total LANC sans SOPL) * 100"},
            "OT planification >3 mois": {"type": "hardcoded", "cible": 5.0, "seuil_score": 5.0, "lower_better": True, "categorie": "Performance", "formule_texte": "(OT LANC sans SOPL âge ≥ 3 / Total LANC sans SOPL) * 100"},
            "OT planification 1mois< <3mois": {"type": "hardcoded", "cible": 15.0, "seuil_score": 15.0, "lower_better": True, "categorie": "Performance", "formule_texte": "(OT LANC sans SOPL 1 < âge < 3 / Total LANC sans SOPL) * 100"},
            "OT exécution <1 mois": {"type": "hardcoded", "cible": 80.0, "seuil_score": 75.0, "lower_better": False, "categorie": "Performance", "formule_texte": "(OT LANC avec SOPL âge ≤ 1 / Total LANC avec SOPL) * 100"},
            "OT exécution >3 mois": {"type": "hardcoded", "cible": 5.0, "seuil_score": 5.0, "lower_better": True, "categorie": "Performance", "formule_texte": "(OT LANC avec SOPL âge ≥ 3 / Total LANC avec SOPL) * 100"},
            "OT exécution 1mois< <3mois": {"type": "hardcoded", "cible": 15.0, "seuil_score": 15.0, "lower_better": True, "categorie": "Performance", "formule_texte": "(OT LANC avec SOPL 1 < âge < 3 / Total LANC avec SOPL) * 100"},
            "appel avis approuvé": {"type": "hardcoded", "cible": 95.0, "seuil_score": 90.0, "lower_better": False, "categorie": "Qualite", "formule_texte": "(Avis APRV / Total Avis sans ordre) * 100"},
            "OT LANC ESTIME": {"type": "hardcoded", "cible": 100.0, "seuil_score": 95.0, "lower_better": False, "categorie": "Qualite", "formule_texte": "(OT LANC Coûts budgétés > 0 / Total LANC) * 100"},
            "Backlog préparation caractérisé": {"type": "hardcoded", "cible": 100.0, "seuil_score": 95.0, "lower_better": False, "categorie": "Qualite", "formule_texte": "(OT CRÉÉ CARACTERISE / Total CRÉÉ) * 100"},
            "Backlog planification caractérisé": {"type": "hardcoded", "cible": 100.0, "seuil_score": 95.0, "lower_better": False, "categorie": "Qualite", "formule_texte": "(OT LANC CARACTERISE / Total LANC) * 100"},
            "OT CONFIME": {"type": "hardcoded", "cible": 100.0, "seuil_score": 95.0, "lower_better": False, "categorie": "Qualite", "formule_texte": "(OT Statut CLO+CONF / Total OT) * 100"},
            "OT_COR_EGAL": {"type": "hardcoded", "cible": 100.0, "seuil_score": 95.0, "lower_better": False, "categorie": "Qualite", "formule_texte": "(OT Coûts réels = budgétés / Total OT) * 100"}
        }
    }

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            saved = json.load(f)
            default = get_default_config()
            default["access_code"] = saved.get("access_code", ACCESS_CODE)
            for k in default["kpis"]:
                if k in saved["kpis"]:
                    default["kpis"][k].update(saved["kpis"][k])
            for k in saved["kpis"]:
                if k not in default["kpis"]:
                    default["kpis"][k] = saved["kpis"][k]
            return default
    return get_default_config()

def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

APP_CONFIG = load_config()

# ==========================================
# 2. VARIABLES GLOBALES
# ==========================================
QK = ["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois",
      "OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois",
      "OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois",
      "OT exécution 1mois< <3mois"]
PK = ["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé",
      "Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]
ALL_KPI = QK + PK

ACT_MAP = {"TAUX_REALISATION_CORRECTIF/PT":"Ameliorer le taux de realisation des OT.","OT préparation <1 mois":"Reduire l'age de preparation (< 1 mois).","OT préparation >3 mois":"Traiter les OT avec preparation > 3 mois.","OT planification <1 mois":"Reduire l'age de planification (< 1 mois).","OT planification >3 mois":"Traiter les OT avec planification > 3 mois.","OT exécution <1 mois":"Reduire l'age d'execution (< 1 mois).","OT exécution >3 mois":"Traiter les OT avec execution > 3 mois.","OT LANC ESTIME":"Estimer les couts des OT lances.","Backlog préparation caractérisé":"Caracteriser le backlog de preparation.","Backlog planification caractérisé":"Caracteriser le backlog de planification.","OT CONFIME":"Confirmer les OT termines.","OT_COR_EGAL":"Rapprocher les couts reels et budgetes.","appel avis approuvé":"Creer un OT pour les avis sans ordre.","OT préparation 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.","OT planification 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.","OT exécution 1mois< <3mois":"Reduire les OT entre 1 et 3 mois."}
LOWER_BETTER = ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois","OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]
MP_KW = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
MPLAN_KW = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
CONSIGNES_HSE = ["Port obligatoire des EPI avant toute intervention.","Port obligatoire du casque de securite.","Port obligatoire des lunettes de protection.","Verifier l'absence de tension avant toute intervention electrique.","Respecter la procedure de consignation et deconsignation.","Ne jamais intervenir sur un equipement en marche.","Baliser et securiser la zone de travail.","Aucun travail n'est plus urgent que la securite."]

# ==========================================
# 3. PAGES SÉCURISÉES (ADMIN & CODE CALCUL)
# ==========================================
def render_admin_page():
    st.title("🛠️ Administration des KPIs & Formules")
    tab_edit, tab_add = st.tabs(["✏️ Modifier un KPI", "➕ Ajouter un KPI"])
    with tab_edit:
        kpi_list = list(APP_CONFIG["kpis"].keys())
        selected_kpi = st.selectbox("Choisir le KPI", kpi_list)
        if selected_kpi:
            kpi_data = APP_CONFIG["kpis"][selected_kpi]
            with st.form(f"edit_{selected_kpi}"):
                if kpi_data.get("type") == "statut_count":
                    all_possible_status = ["CLOT", "CRÉÉ", "LANC", "TCLO", "APRQ", "APRV", "REJT"]
                    new_num = st.multiselect("Colonnes Numérateur", all_possible_status, default=kpi_data.get("numerateurs", []))
                    new_den = st.multiselect("Colonnes Dénominateur", all_possible_status, default=kpi_data.get("denominateurs", []))
                    kpi_data["numerateurs"] = new_num
                    kpi_data["denominateurs"] = new_den
                    kpi_data["formule_texte"] = f"(({ ' + '.join(new_num) if new_num else '0'}) / ({' + '.join(new_den) if new_den else '0'})) * 100"
                new_cible = st.number_input("Cible (%)", 0.0, 100.0, float(kpi_data.get("cible", 100)))
                new_seuil = st.number_input("Seuil Score (%)", 0.0, 100.0, float(kpi_data.get("seuil_score", 100)))
                new_lb = st.checkbox("Moins c'est mieux", value=kpi_data.get("lower_better", False))
                if st.form_submit_button("💾 Sauvegarder"):
                    kpi_data["cible"], kpi_data["seuil_score"], kpi_data["lower_better"] = new_cible, new_seuil, new_lb
                    save_config(APP_CONFIG); st.success("Mis à jour !")
    with tab_add:
        with st.form("add_new_kpi"):
            new_name = st.text_input("Nom du nouveau KPI")
            new_cat = st.selectbox("Catégorie", ["Performance", "Qualite"])
            all_possible_status = ["CLOT", "CRÉÉ", "LANC", "TCLO", "APRQ", "APRV", "REJT"]
            add_num = st.multiselect("Statuts Numérateur", all_possible_status)
            add_den = st.multiselect("Statuts Dénominateur", all_possible_status)
            add_cible = st.number_input("Cible (%)", 0.0, 100.0, 100.0)
            add_seuil = st.number_input("Seuil Score (%)", 0.0, 100.0, 95.0)
            add_lb = st.checkbox("Moins c'est mieux ?", value=False)
            if st.form_submit_button("🚀 Ajouter"):
                if new_name and add_den:
                    APP_CONFIG["kpis"][new_name] = {"type": "statut_count", "numerateurs": add_num, "denominateurs": add_den, "cible": add_cible, "seuil_score": add_seuil, "lower_better": add_lb, "categorie": new_cat, "formule_texte": f"(({ ' + '.join(add_num)}) / ({' + '.join(add_den)})) * 100"}
                    save_config(APP_CONFIG); st.success(f"KPI '{new_name}' ajouté !"); st.rerun()
                else: st.error("Nom et Dénominateur requis.")

def render_code_calcul_page():
    st.markdown("<h1 style='text-align:center; color:#1e3a5f;'>📄 Code de Calcul & Formules</h1>", unsafe_allow_html=True)
    st.caption("Ce document se met à jour automatiquement selon les modifications de la page Administration.")
    st.markdown("---")
    
    st.subheader("1. Dictionnaire des Codes (Colonnes ajoutées)")
    all_codes = set()
    for kpi_data in APP_CONFIG["kpis"].values():
        if kpi_data.get("type") == "statut_count":
            all_codes.update(kpi_data.get("numerateurs", []))
            all_codes.update(kpi_data.get("denominateurs", []))
    if all_codes: st.dataframe(pd.DataFrame({"Code Colonne": sorted(list(all_codes))}), use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("2. Formules de Calcul des KPIs")
    for kpi_name, kpi_data in APP_CONFIG["kpis"].items():
        with st.expander(f"📊 {kpi_name} (Cible: {kpi_data.get('cible')}%)"):
            c1, c2 = st.columns(2)
            c1.markdown("**Formule :**"); c1.code(kpi_data.get("formule_texte", "N/A"), language="text")
            c2.markdown("**Logique Algo :**")
            if kpi_data.get("type") == "statut_count":
                py_code = f"numerateur = df['Statut OT'].isin({kpi_data.get('numerateurs', [])})\ndenominateur = df['Statut OT'].isin({kpi_data.get('denominateurs', [])})\nreturn ckpi(numerateur, denominateur)"
                c2.code(py_code, language="python")
            else: c2.code("# Logique métier complexe (Hardcoded)", language="python")
            
    st.markdown("---")
    st.subheader("3. Méthode de Calcul du Score")
    st.markdown("Le score est binaire : **1 point** si le seuil est atteint, **0 sinon**.")
    dynamic_gscore_code = "def gscore(kpi_name, valeur):\n    seuils = {\n"
    for k, v in APP_CONFIG["kpis"].items():
        symbole = "≤" if v.get("lower_better") else "≥"
        st.markdown(f"- **{k}** : `Valeur {symbole} {v.get('seuil_score')}%` ➡️ **1 point**")
        dynamic_gscore_code += f'        "{k}": {{"seuil": {v.get("seuil_score", 100)}, "lower_better": {str(v.get("lower_better", False))}}},\n'
    dynamic_gscore_code += "    }\n    cfg = seuils.get(kpi_name)\n    if cfg['lower_better']: return 1 if valeur <= cfg['seuil'] else 0\n    else: return 1 if valeur >= cfg['seuil'] else 0"
    st.code(dynamic_gscore_code, language="python")

# ==========================================
# 4. FONCTIONS UTILITAIRES & DATA
# ==========================================
def get_date_from_file():
    if os.path.exists("date.txt"):
        try: 
            with open("date.txt","r",encoding="utf-8") as f: return f.read().strip()
        except Exception: pass
    return datetime.now().strftime("%d/%m/%Y")

def inject_custom_css():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{width:250px!important}
    .main .block-container{max-width:100%!important;padding-left:0.5rem!important;padding-right:0.5rem!important}
    .mh{background:linear-gradient(135deg,#1e3a5f,#2c5282);padding:12px 20px;border-radius:10px;margin-bottom:6px;box-shadow:0 6px 20px rgba(0,0,0,.1)}
    .mh h1{color:#fff;font-size:20px;font-weight:800;margin:0;display:inline}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:6px}
    .cc{background:#fff;border-radius:10px;padding:10px 12px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid #e2e8f0;text-align:center}
    .cc .cv{font-size:26px;font-weight:900;line-height:1}.cc .cl{font-size:11px;color:#718096;font-weight:700;text-transform:uppercase;margin-top:2px}
    .cc.c1{border-top:3px solid #3182ce}.cc.c1 .cv{color:#2b6cb0}
    .cc.c2{border-top:3px solid #38a169}.cc.c2 .cv{color:#276749}
    .cc.c3{border-top:3px solid #805ad5}.cc.c3 .cv{color:#6b46c1}
    .cc.c4{border-top:3px solid #e53e3e}.cc.c4 .cv{color:#c53030}
    .tw{width:100%;border-collapse:collapse;font-size:12px;display:block;overflow-x:auto;margin:0}
    .tw thead th{background:#1e3a5f;color:#fff;font-weight:700;font-size:11px;padding:5px 6px;position:sticky;top:0;z-index:10}
    .tw tbody td{padding:4px 6px;border-bottom:1px solid #edf2f7}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}.tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important}
    </style>""", unsafe_allow_html=True)

def contient_mot(t, lm):
    t=str(t)
    for l in lm:
        for m in l.split():
            if m in t: return True
    return False

def cat_age(a):
    if a<=1: return "<1 mois"
    elif a>=3: return ">3 mois"
    return "1 mois < <3 mois"

def ckpi(n,d,sz=100):
    return np.where(d==0,sz,(n/d)*100)

def cpiv(df,f,c,p):
    return pd.pivot_table(df[f],index="Poste travail princ.",columns=c,values="Ordre",aggfunc="count",fill_value=0).reindex(p,fill_value=0)

def excr(df):
    if "Poste travail princ." in df.columns:
        return df[~df["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)].copy()
    return df

# ==========================================
# 5. FONCTION DE CALCUL DES KPIs
# ==========================================
def calc_kpis(df_i, av_i, now, posts, all_kpi_list, cible_dict, lower_better_list, act_map_dict):
    res={}
    df=df_i.copy()
    av=av_i.copy()
    
    df["Backlog preparation"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MP_KW)),"CARACTERISE","NON CARACTERISE")
    df["Backlog planification"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MPLAN_KW)),"CARACTERISE","NON CARACTERISE")
    
    for dc,am,ac in [('Créé le',"amp","ap"),('Date de début planifiée',"amlp","alp"),('Date de début planifiée',"amex","aex")]:
        if dc in df.columns:
            df[dc]=pd.to_datetime(df[dc],errors='coerce')
            df[am]=((now.year-df[dc].dt.year)*12+(now.month-df[dc].dt.month)).round(2)
            df[ac]=df[am].apply(cat_age)
        else:
            df[am]=np.nan; df[ac]="Inconnu"
            
    df["OT CONFIME"]=np.where(df["Statut système"].str.contains("CLO",na=False)&df["Statut système"].str.contains("CONF",na=False),"OUI","NON")
    df["Contient SOPL"]=df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
    df["OT LANC ESTIME"]=np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
    df["OT_COR_EGAL"]=np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
    res['dfp']=df

    # --- Calculs Hardcodés (Complexes) ---
    an=cpiv(df,df["Nº appel pl.entret."].fillna(0)==0,"Statut OT",posts)
    for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
    an["Total"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
    
    pr=cpiv(df,df["Statut OT"]=="CRÉÉ","ap",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pr[c]=pr.get(c,0)
    pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    
    pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==0),"alp",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pl[c]=pl.get(c,0)
    pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    
    ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: ex[c]=ex.get(c,0)
    ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)

    la=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: la[c]=la.get(c,0)
    la["Total"]=la["OUI"]+la["NON"]

    pc=pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
    pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]

    plc=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
    plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]

    conf_pv=pd.pivot_table(df,index="Poste travail princ.",columns="OT CONFIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: conf_pv[c]=conf_pv.get(c,0)
    conf_pv["Total"]=conf_pv["OUI"]+conf_pv["NON"]

    cor_pv=pd.pivot_table(df,index="Poste travail princ.",columns="OT_COR_EGAL",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: cor_pv[c]=cor_pv.get(c,0)
    cor_pv["Total"]=cor_pv["OUI"]+cor_pv["NON"]

    avf=av[(av["Ordre"].isna())|(av["Ordre"].astype(str).str.strip()=="")].copy()
    tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c]=tca.get(c,0)
    tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1)

    # Assemblage du DataFrame Final
    ckdf = pd.DataFrame(index=posts)
    ckdf["TAUX_REALISATION_CORRECTIF/PT"] = ckpi(an["CLOT"] + an["TCLO"], an["Total"]) # CLOT AJOUTÉ
    ckdf["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"])
    ckdf["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0)
    ckdf["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)
    ckdf["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"])
    ckdf["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0)
    ckdf["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)
    ckdf["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"])
    ckdf["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0)
    ckdf["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)
    ckdf["appel avis approuvé"]=ckpi(tca["APRV"],tca["Total"])
    ckdf["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"])
    ckdf["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"])
    ckdf["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])
    ckdf["OT CONFIME"]=ckpi(conf_pv["OUI"],conf_pv["Total"])
    ckdf["OT_COR_EGAL"]=ckpi(cor_pv["OUI"],cor_pv["Total"])

    # --- Calculs Dynamiques (Depuis le JSON) ---
    for kpi_name, kpi_conf in APP_CONFIG["kpis"].items():
        if kpi_conf.get("type") == "statut_count" and kpi_name not in ckdf.columns:
            num_statuses = kpi_conf.get("numerateurs", [])
            den_statuses = kpi_conf.get("denominateurs", [])
            df_num = df[df["Statut OT"].isin(num_statuses)]
            df_den = df[df["Statut OT"].isin(den_statuses)]
            pivot_num = pd.pivot_table(df_num, index="Poste travail princ.", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
            pivot_den = pd.pivot_table(df_den, index="Poste travail princ.", values="Ordre", aggfunc="count", fill_value=0).reindex(posts, fill_value=0)
            ckdf[kpi_name] = ckpi(pivot_num, pivot_den)
            
            if kpi_name not in all_kpi_list: all_kpi_list.append(kpi_name)
            cible_dict[kpi_name] = kpi_conf.get("cible", 100)
            act_map_dict[kpi_name] = f"Suivre le KPI {kpi_name}."
            if kpi_conf.get("lower_better") and kpi_name not in lower_better_list: lower_better_list.append(kpi_name)

    res['ckdf'] = ckdf
    return res

# ==========================================
# 6. MAIN APPLICATION
# ==========================================
def main():
    try: locale.setlocale(locale.LC_ALL,'fr_FR.UTF-8')
    except Exception: pass

    inject_custom_css()
    fichier_date=get_date_from_file()

    # Gestion des variables modifiables par le JSON
    all_kpi_list = ALL_KPI.copy()
    cible_dict = {k: v["cible"] for k, v in APP_CONFIG["kpis"].items()}
    lower_better_list = LOWER_BETTER.copy()
    act_map_dict = ACT_MAP.copy()

    # --- ROUTING PAGES SÉCURISÉES ---
    query_params = st.query_params
    page = query_params.get("page", ["dashboard"])[0]

    if page in ["admin", "code_calcul"]:
        if "admin_auth" not in st.session_state: st.session_state.admin_auth = False
        if not st.session_state.admin_auth:
            st.markdown("<div style='display:flex; justify-content:center; margin-top:15%;'><div style='background:white; padding:40px; border-radius:15px; box-shadow:0 10px 30px rgba(0,0,0,0.1); text-align:center; width:400px;'>", unsafe_allow_html=True)
            st.subheader("🔐 Accès Restreint")
            code_saisi = st.text_input("Code d'accès", type="password", key="code_auth_secure")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("⚙️ Administration"):
                    if code_saisi == APP_CONFIG.get("access_code", ACCESS_CODE):
                        st.session_state.admin_auth = True; st.query_params["page"] = "admin"; st.rerun()
                    else: st.error("Code incorrect !")
            with col2:
                if st.button("📄 Code de calcul"):
                    if code_saisi == APP_CONFIG.get("access_code", ACCESS_CODE):
                        st.session_state.admin_auth = True; st.query_params["page"] = "code_calcul"; st.rerun()
                    else: st.error("Code incorrect !")
            st.markdown("</div></div>", unsafe_allow_html=True)
            st.stop()
        else:
            if page == "admin": render_admin_page()
            elif page == "code_calcul": render_code_calcul_page()
            if st.button("🚪 Retourner au Dashboard"): del st.query_params["page"]; st.rerun()
            st.stop()

    # --- ÉCRAN HSE ---
    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche = False
    if not st.session_state.hse_affiche:
        c=random.choice(CONSIGNES_HSE)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:46px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SECURITE</h1>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:32px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:36px;font-weight:900">Aucun travail n'est plus urgent que la securite</h2>
        </div>"""%c,unsafe_allow_html=True)
        time.sleep(5); st.session_state.hse_affiche=True; st.rerun(); st.stop()

    # --- SIDEBAR ---
    with st.sidebar:
        st.markdown("**⚙️ Filtres**")
        unf=st.toggle("📁 Charger nouveaux fichiers",value=False,key="tf")
        ot_f=None; av_f=None; apm=[]
        if unf:
            ot_f=st.file_uploader("Fichier OT",type=["xlsx"],key="uot")
            av_f=st.file_uploader("Fichier AVIS",type=["xlsx"],key="uav")
        else:
            st.markdown(f"📅 Données: {fichier_date}")
          
             if os.path.exists("ot.xlsx"):
                try:
                    _t=excr(pd.read_excel("ot.xlsx", engine="openpyxl"))
                    apm=sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                except Exception: pass
        sp=st.multiselect("Poste",["All"]+apm,["All"],key="sp")
        sa=st.multiselect("Atelier",["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)"],["All"],key="sa")
        sd=st.multiselect("Division",["All","SF1","SF2"],["All"],key="sd")
        dr=st.date_input("Periode",value=(datetime(2025,1,1).date(),datetime.today().date()),format="DD/MM/YYYY",key="dr")

    # --- DATA LOADING ---
    if not unf or (ot_f is not None and av_f is not None):
        try:
            raw_ot = pd.read_excel(ot_f, engine="openpyxl") if unf else pd.read_excel("ot.xlsx", engine="openpyxl")
            raw_av = pd.read_excel(av_f, engine="openpyxl") if unf else pd.read_excel("avis.xlsx", engine="openpyxl")
            raw_ot = excr(raw_ot); raw_av = excr(raw_av)
            
            for c in ["Créé le","Date de début planifiée"]: 
                if c in raw_ot.columns: raw_ot[c]=pd.to_datetime(raw_ot[c],errors="coerce")
                
            if not apm: apm = sorted(raw_ot[raw_ot["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
            sp_sel = apm if "All" in sp or not sp else [p for p in sp if p in apm]
            sdt = pd.to_datetime(dr[0]); edt = pd.to_datetime(dr[1])

            # Reconstruction du code coupé (Filtres ateliers)
            def mf(poste):
                p=str(poste).upper()
                if "All" not in sa:
                    m=False
                    for a in sa:
                        if a=="Sulfurique (PS)" and "PS" in p: m=True
                        elif a=="Phosphorique (PP)" and "PP" in p: m=True
                        elif a=="Engrais (TSP/REX)" and ("TSP" in p or "REX" in p): m=True
                        elif a=="Feed (MCP/DCP)" and ("MCP" in p or "DCP" in p): m=True
                    if not m: return False
                if "All" not in sd:
                    d=False
                    for d_sel in sd:
                        if d_sel in p: d=True
                    if not d: return False
                return True

            df_i = raw_ot[(raw_ot["Poste travail princ."].isin(sp_sel)) & (raw_ot["Poste travail princ."].apply(mf))].copy()
            av_i = raw_av[(raw_av["Poste travail princ."].isin(sp_sel)) & (raw_av["Poste travail princ."].apply(mf))].copy()
            posts = sorted(df_i["Poste travail princ."].dropna().unique().tolist())
            if not posts: posts = ["Aucun"]
            
            res = calc_kpis(df_i, av_i, datetime.now(), posts, all_kpi_list, cible_dict, lower_better_list, act_map_dict)
            ckdf = res['ckdf']

            # --- AFFICHAGE ---
            st.markdown('<div class="mh"><h1>Dashboard KPI V3.0</h1><div class="db">%s</div></div>'%fichier_date, unsafe_allow_html=True)
            
            # Cartes Résumé
            st.markdown('<div class="cr">', unsafe_allow_html=True)
            st.markdown('<div class="cc c1"><div class="cv">%d</div><div class="cl">Postes</div></div>'%len(posts), unsafe_allow_html=True)
            st.markdown('<div class="cc c2"><div class="cv">%d</div><div class="cl">Total OT</div></div>'%len(df_i), unsafe_allow_html=True)
            st.markdown('<div class="cc c3"><div class="cv">%d</div><div class="cl">Total Avis</div></div>'%len(av_i), unsafe_allow_html=True)
            st.markdown('<div class="cc c4"><div class="cv">%d</div><div class="cl">KPIs Suivis</div></div>'%len(all_kpi_list), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Onglets
            tab1, tab2 = st.tabs(["📊 Tableau de Bord", "📄 Code de calcul (Sécurisé)"])
            
            with tab1:
                st.dataframe(ckdf.style.format("{:.1f}"), use_container_width=True, height=600)
                
            with tab2:
                st.warning("⚠️ Cette page est protégée. Cliquez ci-dessous pour accéder au détail des formules et codes de calcul.")
                if st.button("🔓 Accéder à la page Code de Calcul"):
                    st.query_params["page"] = "code_calcul"
                    st.rerun()

        except Exception as e:
            st.error(f"Erreur lors du chargement des données : {e}")

if __name__ == "__main__":
    main()
