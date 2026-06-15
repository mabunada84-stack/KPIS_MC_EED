# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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
ACT_MAP = {"TAUX_REALISATION_CORRECTIF/PT":"Ameliorer le taux de realisation des OT.",
           "OT préparation <1 mois":"Reduire l'age de preparation des OT (< 1 mois).",
           "OT préparation >3 mois":"Traiter les OT avec preparation > 3 mois.",
           "OT préparation 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT planification <1 mois":"Reduire l'age de planification des OT (< 1 mois).",
           "OT planification >3 mois":"Traiter les OT avec planification > 3 mois.",
           "OT planification 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT exécution <1 mois":"Reduire l'age d'execution des OT (< 1 mois).",
           "OT exécution >3 mois":"Traiter les OT avec execution > 3 mois.",
           "OT exécution 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "OT LANC ESTIME":"Estimer les couts des OT lances.",
           "Backlog préparation caractérisé":"Caracteriser le backlog de preparation.",
           "Backlog planification caractérisé":"Caracteriser le backlog de planification.",
           "OT CONFIME":"Confirmer les OT termines.",
           "OT_COR_EGAL":"Rapprocher les couts reels et budgetes.",
           "appel avis approuvé":"Creer un OT pour les avis sans ordre."}
LOWER_BETTER = ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois",
                "OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]
MP_KW = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
MPLAN_KW = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
EXEC_KW = ["ATEX","EXEC","EXE","PRER","PRÉR"]
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
def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt","r",encoding="utf-8") as f: return f.read().strip()
        except Exception: pass
    return datetime.now().strftime("%d/%m/%Y")

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

def get_caract_type(statut_user,keywords):
    s=str(statut_user).upper(); matched=[kw for kw in keywords if kw in s]
    return max(matched,key=len) if matched else "AUTRE"

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
    .tw.bt thead th{background:linear-gradient(135deg,#553c9a,#805ad5)}
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
    .car .cab-wrap{flex:1;position:relative}
    .car .cab{height:24px;background:#edf2f7;border-radius:4px;overflow:hidden;width:100%}
    .car .caf{height:100%;border-radius:4px;transition:width .3s}
    .car .cav-out{font-size:12px;font-weight:800;color:#1a202c;min-width:55px;text-align:right;padding-left:6px}
    .target-marker{position:absolute;top:-6px;bottom:-6px;width:3px;background:#e53e3e;z-index:2;border-radius:1px;box-shadow:0 0 6px rgba(229,62,62,0.5)}
    .target-label{position:absolute;top:-20px;font-size:10px;color:#e53e3e;font-weight:800;transform:translateX(-50%);white-space:nowrap;z-index:3;background:rgba(255,255,255,0.92);padding:1px 5px;border-radius:3px;border:1px solid #e53e3e}
    .cg{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .cg>div{background:#fff;border-radius:var(--r);padding:8px 10px;border:1px solid var(--b)}
    .cg .ct{font-size:13px;font-weight:700;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid var(--b)}
    .cgr{display:flex;align-items:center;padding:3px 0;font-size:12px;border-bottom:1px solid #f7fafc}
    .cgr:last-child{border:none}
    .cgr .rk{width:18px;font-weight:800;text-align:center}
    .cgr .pn{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cgr .ps{font-weight:800;min-width:55px;text-align:right}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:8px 14px;font-weight:700;font-size:15px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label,div[data-testid="stSidebar"] .stCheckbox label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:3px 8px;margin-bottom:3px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:14px;color:#718096;font-size:14px}
    .g-green{background:#c6efce;color:#006100;font-weight:600}
    .g-yellow{background:#ffeb9c;color:#9c6500;font-weight:600}
    .g-red{background:#ffc7ce;color:#9c0006;font-weight:600}
    .pie-summary-table{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:13px;margin-top:12px}
    .pie-summary-table thead th{background:var(--p);color:#fff;font-weight:700;font-size:12px;padding:8px 12px;text-align:left}
    .pie-summary-table tbody td{padding:6px 12px;border-bottom:1px solid #edf2f7}
    .pie-summary-table tbody tr:nth-child(even) td{background:#f7fafc}
    .pie-summary-table .tot-row td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important}
    .pie-status-dot{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:8px;vertical-align:middle}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:17px}.cg{grid-template-columns:1fr}.car .cal{width:120px}}
    </style>""",unsafe_allow_html=True)

# ============================================================
def main():
    try: locale.setlocale(locale.LC_ALL,'fr_FR.UTF-8')
    except Exception:
        try: locale.setlocale(locale.LC_ALL,'fr_FR')
        except Exception: pass
    inject_custom_css()
    fichier_date=get_date_from_file()

    if "hse_affiche" not in st.session_state: st.session_state.hse_affiche=False
    if not st.session_state.hse_affiche:
        c=random.choice(CONSIGNES_HSE)
        st.markdown("""<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a365d,#2d3748,#1a365d);padding:40px">
        <div style="font-size:64px;margin-bottom:20px">🦺</div>
        <h1 style="text-align:center;font-size:46px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SECURITE</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:22px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Securite - Sante - Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:32px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:36px;font-weight:900">Aucun travail n'est plus urgent que la securite</h2>
        <div style="margin-top:40px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden"><div style="width:100%%;height:100%%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 5.5s ease-in-out forwards"></div></div>
        <style>@keyframes ld{from{width:0}to{width:100%%}}</style></div>"""%c,unsafe_allow_html=True)
        time.sleep(6); st.session_state.hse_affiche=True; st.rerun(); st.stop()

    def contient_mot(t,lm):
        t=str(t); return any(m in t for l in lm for m in l.split())
    def cat_age(a):
        if a<=1: return "<1 mois"
        elif a>=3: return ">3 mois"
        return "1 mois < <3 mois"
    def ckpi(n,d,sz=100): return np.where(d==0,sz,(n/d)*100)
    def cpiv(df,f,c,p):
        return pd.pivot_table(df[f],index="Poste travail princ.",columns=c,values="Ordre",aggfunc="count",fill_value=0).reindex(p,fill_value=0)
    def excr(df):
        if "Poste travail princ." in df.columns:
            return df[~df["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)].copy()
        return df
    def get_metier(p):
        p=str(p).upper()
        if "E" in p: return "Electrique"
        if "M" in p: return "Mecanique"
        if "R" in p: return "Instrumentation"
        if "G" in p: return "Genie Civil"
        return "Autre"
    def get_atelier(p):
        p=str(p).upper()
        if "PS" in p: return "Sulfurique"
        if "PP" in p: return "Phosphorique"
        if "TSP" in p or "REX" in p: return "Engrais"
        if "MCP" in p or "DCP" in p: return "Feed"
        return "Autre"
    def get_division(p):
        p=str(p).upper()
        if "SF1" in p: return "SF1"
        if "SF2" in p: return "SF2"
        return "Autre"

    def find_designation_col(df):
        possible=["Désignation","Designation","Description","Texte","Désignation OT","Texte ordre","Désignation de l'ordre"]
        for col in possible:
            if col in df.columns: return col
        for col in df.columns:
            cl=str(col).lower()
            if "désign" in cl or "design" in cl or "texte" in cl or "descr" in cl: return col
        return None

    def is_mince_data(labels, values):
        return len(labels) <= 2 or sum(values) < 10

    def calc_kpis(df_i,av_i,now,posts):
        res={}; df=df_i.copy(); av=av_i.copy()
        df["Backlog preparation"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MP_KW)),"CARACTERISE","NON CARACTERISE")
        df["Backlog planification"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MPLAN_KW)),"CARACTERISE","NON CARACTERISE")
        df["Backlog execution"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,EXEC_KW)),"CARACTERISE","NON CARACTERISE")
        for dc,am,ac in [('Créé le',"amp","ap"),('Date de début planifiée',"amlp","alp"),('Date de début planifiée',"amex","aex")]:
            if dc in df.columns:
                df[dc]=pd.to_datetime(df[dc],errors='coerce')
                df[am]=((now.year-df[dc].dt.year)*12+(now.month-df[dc].dt.month)).round(2)
                df[ac]=df[am].apply(cat_age)
            else: df[am]=np.nan; df[ac]="Inconnu"
        df["OT CONFIME"]=np.where(df["Statut système"].str.contains("CLO",na=False)&df["Statut système"].str.contains("CONF",na=False),"OUI","NON")
        df["Contient SOPL"]=df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
        df["OT LANC ESTIME"]=np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
        df["OT_COR_EGAL"]=np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
        res['dfp']=df
        an=cpiv(df,df["Nº appel pl.entret."].fillna(0)==0,"Statut OT",posts)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
        an["Total"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1); an["TAUX_REALISATION_CORRECTIF/PT"]=ckpi(an["TCLO"],an["Total"])
        pr=cpiv(df,df["Statut OT"]=="CRÉÉ","ap",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pr[c]=pr.get(c,0)
        pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"]); pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0); pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)
        pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==0),"alp",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pl[c]=pl.get(c,0)
        pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"]); pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0); pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)
        ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: ex[c]=ex.get(c,0)
        ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"]); ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0); ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)
        la=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: la[c]=la.get(c,0)
        la["Total"]=la["OUI"]+la["NON"]; la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"])
        pc=pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
        pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"])
        plc=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
        plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])
        exc_back=pd.pivot_table(df[(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1)],index="Poste travail princ.",columns="Backlog execution",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: exc_back[c]=exc_back.get(c,0)
        exc_back["Total"]=exc_back["CARACTERISE"]+exc_back["NON CARACTERISE"]
        res['exc_back']=exc_back
        for kn,cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
            pv=pd.pivot_table(df,index="Poste travail princ.",columns=cn,values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
            for c in ["OUI","NON"]: pv[c]=pv.get(c,0)
            pv["Total"]=pv["OUI"]+pv["NON"]; pv[cn]=ckpi(pv["OUI"],pv["Total"]); res[kn.lower().replace(" ","_")]=pv
        avf=av[(av["Ordre"].isna())|(av["Ordre"].astype(str).str.strip()=="")].copy(); res['avf']=avf
        tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c]=tca.get(c,0)
        tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["appel avis approuvé"]=ckpi(tca["APRV"],tca["Total"])
        res['ckdf']=pd.DataFrame({
            "TAUX_REALISATION_CORRECTIF/PT":an["TAUX_REALISATION_CORRECTIF/PT"],
            "OT préparation <1 mois":pr["OT préparation <1 mois"],"OT préparation >3 mois":pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois":pl["OT planification <1 mois"],"OT planification >3 mois":pl["OT planification >3 mois"],"OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois":ex["OT exécution <1 mois"],"OT exécution >3 mois":ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
            "appel avis approuvé":tca["appel avis approuvé"],"OT LANC ESTIME":la["OT LANC ESTIME"],
            "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],"Backlog planification caractérisé":plc["Backlog planification caractérisé"],
            "OT CONFIME":res['ot_confime']["OT CONFIME"],"OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"]
        })
        return res

    def ks(v,c):
        try: val=float(v)
        except Exception: return ""
        if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val>=80 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=75 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val<=15 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val<=5 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c=="TAUX_REALISATION_CORRECTIF/PT":
            return "background:#c6efce;color:#006100;font-weight:600" if val>=85 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c=="appel avis approuvé":
            return "background:#c6efce;color:#006100;font-weight:600" if val>=95 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=90 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val>=100 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=95 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        return ""
    def cs(v):
        try: val=float(str(v).replace(' %','').strip())
        except Exception: return ""
        return "background:#c6efce;color:#006100;font-weight:700" if val>=90 else ("background:#ffeb9c;color:#9c6500;font-weight:700" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:700")
    def kas(v):
        try: val=int(v)
        except Exception: return ""
        if val==0: return "color:#cbd5e0"
        if val<=3: return "background:#ffeb9c;color:#9c6500;font-weight:600"
        if val<=10: return "background:#fed7d7;color:#c53030;font-weight:600"
        return "background:#fc8181;color:#742a2a;font-weight:800"
    def gscore(k,a,t):
        if pd.isna(a) or pd.isna(t): return 0
        if k in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return 1 if a>=75 else 0
        if k in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return 1 if a<=15 else 0
        if k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return 1 if a<=5 else 0
        if k=="TAUX_REALISATION_CORRECTIF/PT": return 1 if a>=80 else 0
        if k=="appel avis approuvé": return 1 if a>=90 else 0
        if k in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]: return 1 if a>=95 else 0
        return 0
    def is_lb(k): return k in LOWER_BETTER

    def html_table(rows,cols,tc,sc_col=None):
        h='<table class="tw %s"><thead><tr>'%tc+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            rc="cb" if r.get("_t")=="cible" else ("tr" if r.get("_t")=="total" else "")
            h+='<tr class="%s">'%rc
            for c in cols:
                v=r.get(c,"")
                if r.get("_t")=="cible": h+='<td>%s</td>'%v
                else: s=cs(v) if sc_col and c in sc_col else ks(v,c); h+='<td style="%s">%s</td>'%(s or "",v)
            h+='</tr>'
        return h+'</tbody></table>'
    def html_ano(rows,cols):
        h='<table class="tw at"><thead><tr>'+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            h+='<tr class="%s">'%("tr" if r.get("_t")=="total" else "")
            for c in cols: v=r.get(c,""); h+='<td style="%s">%s</td>'%(kas(v) or "",v)
            h+='</tr>'
        return h+'</tbody></table>'
    def html_actions_table(kpi_list,actuals,targets,act_map):
        h='<table class="tw at"><thead><tr><th>KPI</th><th>Valeur Actuelle</th><th>Cible</th><th>Ecart</th><th>Statut</th><th>Action Recommandée</th></tr></thead><tbody>'
        has_rows=False
        for k in kpi_list:
            av=actuals.get(k,0); tv=targets.get(k,100); diff=av-tv
            met=av<=tv if is_lb(k) else av>=tv
            if met: continue
            has_rows=True; ec_clr="#c53030"
            action=act_map.get(k,"")
            h+='<tr><td style="font-weight:600">%s</td><td>%.1f%%</td><td>%.0f%%</td><td style="color:%s;font-weight:700">%+.1f%%</td><td style="background:#ffc7ce;color:#9c0006;font-weight:700">NON ATTEINT</td><td style="color:#4a5568">%s</td></tr>'%(k,av,tv,ec_clr,diff,action)
        if not has_rows:
            h+='<tr><td colspan="6" style="text-align:center;padding:16px;color:#276749;font-weight:700;font-size:14px">✅ Tous les KPI sont atteints — Aucune action requise</td></tr>'
        return h+'</tbody></table>'
    def html_kpi_bars_with_target(kpi_list, actuals, targets, title, color_ok, color_fail):
        h='<div class="ca"><div class="ct" style="color:%s">%s</div>'%(color_ok,title)
        for k in kpi_list:
            av=actuals.get(k,0); tv=targets.get(k,100)
            met=av<=tv if is_lb(k) else av>=tv
            bw=min(max(av,0),100); bg=color_ok if met else color_fail
            tp=min(max(tv,0),100)
            h+='<div class="car" style="padding-top:18px"><div class="cal">%s</div><div class="cab-wrap"><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="target-marker" style="left:%s%%"></div><div class="target-label" style="left:%s%%">Cible %.0f%%</div></div><div class="cav-out">%.1f%%</div></div>'%(k,bw,bg,tp,tp,tv,av)
        return h+'</div>'

    # --- Nouveau : Chart Scores par Poste avec trait objectif ---
    def create_score_chart(scores_dict, target, title, color_ok, color_nok):
        postes=list(scores_dict.keys()); scores=list(scores_dict.values())
        colors=[color_ok if s>=target else color_nok for s in scores]
        fig=go.Figure()
        fig.add_trace(go.Bar(y=postes,x=scores,orientation='h',marker_color=colors,
            text=[f"{s:.1f}%" for s in scores],textposition='outside',name='Score'))
        fig.add_vline(x=target,line_dash="dash",line_width=2.5,line_color="#E53E3E",
            annotation_text=f"Objectif {target}%",annotation_position="top right",
            annotation_font_color="#E53E3E",annotation_font_size=12,annotation_font_weight="bold")
        xmax=max(max(scores)+10,target+15)
        fig.update_layout(title=dict(text=f"<b>{title}</b>",font_size=15,x=0.5),
            height=max(300,len(postes)*28+80),yaxis=dict(autorange="reversed",tickfont_size=11),
            xaxis=dict(range=[0,xmax],title_text="Score (%)",tickfont_size=11),
            margin=dict(l=160,r=50,t=60,b=40),showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)')
        fig.update_yaxes(gridcolor='#e2e8f0'); fig.update_xaxes(gridcolor='#e2e8f0')
        return fig

    # --- Nouveau : Pie chart avec gestion mince ---
    def make_pie_with_mince(labels, values, title, is_mince, color_map=None):
        total=sum(values)
        if total==0: return None
        dc=["#3182CE","#38A169","#D69E2E","#805AD5","#E53E3E","#DD6B20","#319795","#D53F8C"]
        if color_map:
            colors=[color_map.get(l,dc[i%len(dc)]) for i,l in enumerate(labels)]
        else:
            colors=[dc[i%len(dc)] for i in range(len(labels))]
        if is_mince:
            fig=make_subplots(rows=1,cols=2,specs=[[{'type':'domain'},{'type':'domain'}]],
                subplot_titles=['Répartition','Détail'])
            fig.add_trace(go.Pie(labels=labels,values=values,textinfo='label+percent',
                textfont_size=11,marker_colors=colors,hole=0.3,name="Répartition"),row=1,col=1)
            fig.add_trace(go.Pie(labels=labels,values=values,textinfo='value+percent',
                textfont_size=11,marker_colors=colors,hole=0.55,name="Détail",
                pull=[0.03]*len(labels)),row=1,col=2)
            fig.update_layout(title_text=f"<b>{title}</b>",title_x=0.5,height=400,
                showlegend=True,legend=dict(orientation="h",yanchor="bottom",y=-0.12,
                xanchor="center",x=0.5,font_size=10))
        else:
            fig=go.Figure(go.Pie(labels=labels,values=values,textinfo='label+percent',
                textfont_size=13,marker_colors=colors,hole=0.4,pull=[0.03]*len(labels),
                hovertemplate='<b>%{label}</b><br>Nombre: %{value}<br>Part: %{percent}<extra></extra>'))
            fig.update_layout(title_text=f"<b>{title}</b>",title_x=0.5,height=420,
                showlegend=True,legend=dict(orientation="h",yanchor="bottom",y=-0.08,
                xanchor="center",x=0.5,font_size=12),
                annotations=[dict(text=f"<b>{total}</b>",x=0.5,y=0.5,font_size=20,
                    font_color="#1a202c",showarrow=False)])
        return fig

    # --- Nouveau : Analyse par mot-clé (OMS / Thermographie) ---
    def create_keyword_analysis(df, keyword, desig_col, posts, title):
        if desig_col is None or desig_col not in df.columns:
            return None, None, None, 0
        mask=df[desig_col].astype(str).str.contains(keyword,case=False,na=False)
        filtered=df[mask].copy()
        if filtered.empty:
            return None, None, None, 0
        statut_cols=["CRÉÉ","LANC","TCLO","CLOT"]
        pivot=pd.pivot_table(filtered,index="Poste travail princ.",columns="Statut OT",
            values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in statut_cols: pivot[c]=pivot.get(c,0)
        pivot["Total"]=pivot[statut_cols].sum(axis=1)
        pivot=pivot.sort_values("Total",ascending=False)
        # HTML table
        cols_t=["Poste de travail"]+statut_cols+["Total"]
        rows_t=[]
        for p in pivot.index:
            row={"Poste de travail":p}
            for c in statut_cols+["Total"]: row[c]=int(pivot.loc[p,c])
            row["_t"]=""; rows_t.append(row)
        tot_r={"Poste de travail":"Total général"}
        for c in statut_cols+["Total"]: tot_r[c]=int(pivot[c].sum())
        tot_r["_t"]="total"; rows_t.append(tot_r)
        table_html=html_table(rows_t,cols_t,"bt")
        # Pie par statut
        sc=filtered["Statut OT"].value_counts()
        sl=sc.index.tolist(); sv=sc.values.tolist()
        pie1=make_pie_with_mince(sl,sv,f"{title} — Par Statut OT",is_mince_data(sl,sv))
        # Pie par poste top 5
        pc=pivot["Total"].sort_values(ascending=False).head(5)
        if len(pivot)>5: pc["Autres"]=pivot["Total"].iloc[5:].sum()
        pl2=pc.index.tolist(); pv2=pc.values.tolist()
        pie2=make_pie_with_mince(pl2,pv2,f"{title} — Par Poste (Top)",is_mince_data(pl2,pv2))
        return table_html, pie1, pie2, len(filtered)

    # --- Nouveau : Tableau + Pie Backlog ---
    def create_backlog_section(backlog_pivot, title, tc_class):
        for c in ["CARACTERISE","NON CARACTERISE","Total","% Caractérisé"]:
            if c not in backlog_pivot.columns: backlog_pivot[c]=0
        backlog_pivot["Total"]=backlog_pivot["CARACTERISE"]+backlog_pivot["NON CARACTERISE"]
        backlog_pivot["% Caractérisé"]=np.where(backlog_pivot["Total"]>0,
            (backlog_pivot["CARACTERISE"]/backlog_pivot["Total"]*100).round(1),0)
        cols_b=["Poste de travail","Caractérisé","Non Caractérisé","Total","% Caractérisé"]
        rows_b=[]
        for p in backlog_pivot.index:
            row={"Poste de travail":p,"Caractérisé":int(backlog_pivot.loc[p,"CARACTERISE"]),
                "Non Caractérisé":int(backlog_pivot.loc[p,"NON CARACTERISE"]),
                "Total":int(backlog_pivot.loc[p,"Total"]),
                "% Caractérisé":"%.1f%%"%backlog_pivot.loc[p,"% Caractérisé"]}
            row["_t"]=""; rows_b.append(row)
        tc=int(backlog_pivot["CARACTERISE"].sum()); tnc=int(backlog_pivot["NON CARACTERISE"].sum())
        ta=tc+tnc; moy=(tc/ta*100) if ta>0 else 0
        mr={"Poste de travail":"Moyenne / Total","Caractérisé":tc,"Non Caractérisé":tnc,
            "Total":ta,"% Caractérisé":"%.1f%%"%moy}
        mr["_t"]="total"; rows_b.append(mr)
        table_html=html_table(rows_b,cols_b,tc_class,sc_col={"% Caractérisé"})
        labels_b=["Caractérisé","Non Caractérisé"]; values_b=[tc,tnc]
        pie_fig=make_pie_with_mince(labels_b,values_b,title,
            is_mince_data(labels_b,values_b),{"Caractérisé":"#38A169","Non Caractérisé":"#E53E3E"})
        return table_html, pie_fig

    def create_ot_status_pie(df_ot):
        if df_ot is None or df_ot.empty or "Statut OT" not in df_ot.columns:
            return None, None
        sc=df_ot["Statut OT"].value_counts().reset_index()
        sc.columns=["Statut","Nombre"]; total=sc["Nombre"].sum()
        if total==0: return None, None
        sc["Pourcentage"]=(sc["Nombre"]/total*100).round(1)
        status_colors={"CRÉÉ":"#3182CE","LANC":"#38A169","TCLO":"#D69E2E","CLOT":"#805AD5"}
        colors=[status_colors.get(s,"#A0AEC0") for s in sc["Statut"]]
        fig=go.Figure(go.Pie(labels=sc["Statut"].tolist(),values=sc["Nombre"].tolist(),
            textinfo='label+percent',texttemplate='%{label}<br>%{percent}',textposition='inside',
            textfont_size=13,marker_colors=colors,hole=0.4,pull=[0.04]*len(sc),
            hovertemplate='<b>%{label}</b><br>Nombre: %{value} OTs<br>Part: %{percent}<extra></extra>'))
        fig.update_layout(title=dict(text="<b>Répartition des OT par Statut</b>",font_size=16,x=0.5),
            height=480,margin=dict(t=70,b=50,l=30,r=30),
            legend=dict(font_size=12,orientation="h",yanchor="bottom",y=-0.08,xanchor="center",x=0.5),
            showlegend=True,annotations=[dict(text=f"<b>{total}</b><br>OTs",x=0.5,y=0.5,
                font_size=20,font_color="#1a202c",showarrow=False)])
        return fig, sc

    def html_pie_summary_table(sc_df):
        if sc_df is None or sc_df.empty: return ""
        status_colors={"CRÉÉ":"#3182CE","LANC":"#38A169","TCLO":"#D69E2E","CLOT":"#805AD5"}
        h='<table class="pie-summary-table"><thead><tr><th>Statut</th><th>Nombre d\'OT</th><th>Pourcentage</th></tr></thead><tbody>'
        for _,r in sc_df.iterrows():
            clr=status_colors.get(r["Statut"],"#A0AEC0")
            h+='<tr><td><span class="pie-status-dot" style="background:%s"></span><b>%s</b></td><td style="text-align:center;font-weight:700">%d</td><td style="text-align:center;font-weight:700">%.1f%%</td></tr>'%(clr,r["Statut"],int(r["Nombre"]),r["Pourcentage"])
        tn=int(sc_df["Nombre"].sum()); tp=round(sc_df["Pourcentage"].sum(),1)
        h+='<tr class="tot-row"><td><b>TOTAL</b></td><td style="text-align:center">%d</td><td style="text-align:center">%.1f%%</td></tr>'%(tn,tp)
        return h+'</tbody></table>'

    def export_btn(df,filename):
        buf=io.BytesIO(); df.to_excel(buf,index=False,engine='openpyxl'); buf.seek(0)
        st.download_button("📥 Exporter Excel",data=buf,file_name=filename,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""<div style="padding:10px 0 4px 0"><div style="font-size:22px;margin-bottom:2px">⚙️</div><div style="font-size:14px;font-weight:800;color:white">Filtres & Parametres</div><div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Configuration</div></div>""",unsafe_allow_html=True)
        st.markdown("---")
        show_filters=st.checkbox("Afficher les filtres",value=True,key="show_filters")
        if show_filters:
            unf=st.toggle("📁 Charger nouveaux fichiers",value=False,key="tf")
            ot_f=av_f=None; apm=[]
            if unf:
                ot_f=st.file_uploader("Fichier OT",type=["xlsx"],key="uot")
                av_f=st.file_uploader("Fichier AVIS",type=["xlsx"],key="uav")
            else:
                if os.path.exists("ot.xlsx"):
                    try:
                        _t=excr(pd.read_excel("ot.xlsx"))
                        apm=sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                    except Exception: pass
                st.markdown("""<div style="background:rgba(255,255,255,.1);padding:8px 10px;border-radius:6px;font-size:12px;color:rgba(255,255,255,.7)">📁 Fichiers par defaut : ot.xlsx & avis.xlsx</div>""",unsafe_allow_html=True)
            sel_posts=st.multiselect("Postes de travail",options=apm,default=apm,key="sp") if apm else []
            metier_opt=["Tous","Electrique","Mecanique","Instrumentation","Genie Civil","Autre"]
            sel_metier=st.selectbox("Metier",metier_opt,key="sm")
            atelier_opt=["Tous","Sulfurique","Phosphorique","Engrais","Feed","Autre"]
            sel_atelier=st.selectbox("Atelier",atelier_opt,key="sa")
            div_opt=["Tous","SF1","SF2"]
            sel_div=st.selectbox("Division",div_opt,key="sd")
        else:
            sel_posts=[]; sel_metier="Tous"; sel_atelier="Tous"; sel_div="Tous"; unf=False; ot_f=None; av_f=None
            if os.path.exists("ot.xlsx"):
                try:
                    _t=excr(pd.read_excel("ot.xlsx"))
                    apm=sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                    sel_posts=apm[:]
                except Exception: pass

    # ===================== DATA LOADING =====================
    error_msg=""
    try:
        if unf and ot_f and av_f:
            df_raw=pd.read_excel(ot_f); av_raw=pd.read_excel(av_f)
        elif os.path.exists("ot.xlsx") and os.path.exists("avis.xlsx"):
            df_raw=pd.read_excel("ot.xlsx"); av_raw=pd.read_excel("avis.xlsx")
        else:
            error_msg="Fichiers de donnees introuvables. Veuillez charger ot.xlsx et avis.xlsx."; df_raw=pd.DataFrame(); av_raw=pd.DataFrame()
    except Exception as e:
        error_msg="Erreur de chargement : %s"%str(e); df_raw=pd.DataFrame(); av_raw=pd.DataFrame()
    if error_msg:
        st.markdown('<div class="es">⚠️ %s</div>'%error_msg,unsafe_allow_html=True); st.stop()

    df_raw=excr(df_raw); now=datetime.now()
    all_posts=sorted(df_raw["Poste travail princ."].dropna().unique().tolist()) if "Poste travail princ." in df_raw.columns else []

    if sel_posts: df_filt=df_raw[df_raw["Poste travail princ."].isin(sel_posts)].copy()
    else: df_filt=df_raw.copy()
    if sel_metier!="Tous" and df_filt.shape[0]>0:
        df_filt=df_filt[df_filt["Poste travail princ."].apply(lambda p:get_metier(p)==sel_metier)]
    if sel_atelier!="Tous" and df_filt.shape[0]>0:
        df_filt=df_filt[df_filt["Poste travail princ."].apply(lambda p:get_atelier(p)==sel_atelier)]
    if sel_div!="Tous" and df_filt.shape[0]>0:
        df_filt=df_filt[df_filt["Poste travail princ."].apply(lambda p:get_division(p)==sel_div)]

    active_posts=sorted(df_filt["Poste travail princ."].dropna().unique().tolist()) if "Poste travail princ." in df_filt.columns else []
    if not active_posts:
        st.markdown('<div class="es">⚠️ Aucun poste de travail selectionne ou correspondant aux filtres.</div>',unsafe_allow_html=True); st.stop()

    if "Poste travail princ." in av_raw.columns and active_posts:
        av_filt=av_raw[av_raw["Poste travail princ."].isin(active_posts)].copy()
    else: av_filt=av_raw.copy()

    res=calc_kpis(df_filt,av_filt,now,active_posts)
    ckdf=res['ckdf']; dfp=res['dfp']

    # ===================== CALCULATE SCORES =====================
    perf_scores={}
    for p in active_posts:
        if p in ckdf.index:
            sl=[gscore(k,ckdf.loc[p,k],CIBLE.get(k,100)) for k in QK if k in ckdf.columns]
            perf_scores[p]=(sum(sl)/len(sl)*100) if sl else 0
        else: perf_scores[p]=0
    qual_scores={}
    for p in active_posts:
        if p in ckdf.index:
            sl=[gscore(k,ckdf.loc[p,k],CIBLE.get(k,100)) for k in PK if k in ckdf.columns]
            qual_scores[p]=(sum(sl)/len(sl)*100) if sl else 0
        else: qual_scores[p]=0

    global_perf={}
    for k in QK:
        if k in ckdf.columns:
            vals=pd.to_numeric(ckdf[k],errors='coerce').dropna()
            global_perf[k]=vals.mean() if not vals.empty else 0
        else: global_perf[k]=0
    global_qual={}
    for k in PK:
        if k in ckdf.columns:
            vals=pd.to_numeric(ckdf[k],errors='coerce').dropna()
            global_qual[k]=vals.mean() if not vals.empty else 0
        else: global_qual[k]=0

    global_perf_score=sum(perf_scores.values())/len(perf_scores) if perf_scores else 0
    global_qual_score=sum(qual_scores.values())/len(qual_scores) if qual_scores else 0
    total_ots=len(df_filt)
    taux_real=global_perf.get("TAUX_REALISATION_CORRECTIF/PT",0)

    # ===================== ANOMALIES (calcul) =====================
    def calc_anomalies(kpi_list, ckdf, posts, cible_map, lower_better):
        ano_rows=[]
        for p in posts:
            if p not in ckdf.index: continue
            nb_ano=0; details=[]
            for k in kpi_list:
                if k not in ckdf.columns: continue
                try: val=float(ckdf.loc[p,k])
                except Exception: continue
                tgt=cible_map.get(k,100)
                met=(val<=tgt) if k in lower_better else (val>=tgt)
                if not met:
                    nb_ano+=1; details.append(k)
            if nb_ano>0:
                ano_rows.append({"Poste de travail":p,"Nombre anomalies":nb_ano,"Details":", ".join(details),"_t":""})
        # Tri par nombre d'anomalies décroissant puis inversion
        ano_rows.sort(key=lambda x:x["Nombre anomalies"],reverse=True)
        ano_rows=ano_rows[::-1]  # INVERSION
        tot={"Poste de travail":"Total","Nombre anomalies":sum(r["Nombre anomalies"] for r in ano_rows),"Details":"","_t":"total"}
        ano_rows.append(tot)
        return ano_rows

    ano_perf_rows=calc_anomalies(QK,ckdf,active_posts,CIBLE,LOWER_BETTER)
    ano_qual_rows=calc_anomalies(PK,ckdf,active_posts,CIBLE,LOWER_BETTER)

    # ===================== HEADER =====================
    st.markdown('<div class="mh"><h1>📊 Dashboard KPI Maintenance</h1><span class="db">📅 %s</span></div>'%fichier_date,unsafe_allow_html=True)

    # ===================== SUMMARY CARDS =====================
    card_data=[("c1","%d"%total_ots,"Total OT"),("c2","%.1f%%"%taux_real,"Taux Realisation"),
        ("c3","%.1f%%"%global_perf_score,"Score Performance"),("c4","%.1f%%"%global_qual_score,"Score Qualite")]
    cards_html='<div class="cr">'+''.join('<div class="cc %s"><div class="cv">%s</div><div class="cl">%s</div></div>'%(c,v,l) for c,v,l in card_data)+'</div>'
    st.markdown(cards_html,unsafe_allow_html=True)

    # ===================== TABS =====================
    tab1,tab2,tab3,tab4,tab5,tab6=st.tabs(["⚡ Synthèse Performance","🛡️ Synthèse Qualité",
        "🎯 Actions Recommandées","📊 Répartition OT","🔬 Analyse OMS & Thermographie","📋 Analyse Backlog"])

    # ============ TAB 1 : SYNTHESE PERFORMANCE ============
    with tab1:
        st.markdown(html_kpi_bars_with_target(QK,global_perf,CIBLE,"Indicateurs de Performance — Valeur globale avec cible","#38a169","#e53e3e"),unsafe_allow_html=True)

        st.markdown('<div class="stl p" style="margin-top:10px">Détail par Poste de Travail — Performance</div>',unsafe_allow_html=True)
        pcols=["Poste de travail"]+QK+["Score Performance"]
        prows=[]
        for p in active_posts:
            if p in ckdf.index:
                row={"Poste de travail":p}
                for k in QK: row[k]="%.1f%%"%ckdf.loc[p,k]
                row["Score Performance"]="%.2f%%"%perf_scores[p]
                row["_t"]=""; prows.append(row)
        # Cible row
        crow={"Poste de travail":"CIBLE"}
        for k in QK: crow[k]="%.0f%%"%CIBLE.get(k,100)
        crow["Score Performance"]="≥ 80.00%%"; crow["_t"]="cible"; prows.append(crow)
        st.markdown(html_table(prows,pcols,"pt",sc_col={"Score Performance"}),unsafe_allow_html=True)

        # Scores par Poste avec trait objectif
        st.markdown('<div class="stl p" style="margin-top:10px">Scores par Poste de Travail — Performance</div>',unsafe_allow_html=True)
        st.plotly_chart(create_score_chart(perf_scores,80,"Scores Performance par Poste","#38a169","#e53e3e"),use_container_width=True)

        # Anomalies Performance (INVERSE)
        st.markdown('<div class="stl a" style="margin-top:10px">Anomalies Performance</div>',unsafe_allow_html=True)
        st.markdown(html_ano(ano_perf_rows,["Poste de travail","Nombre anomalies","Details"]),unsafe_allow_html=True)

    # ============ TAB 2 : SYNTHESE QUALITE ============
    with tab2:
        st.markdown(html_kpi_bars_with_target(PK,global_qual,CIBLE,"Indicateurs de Qualité — Valeur globale avec cible","#3182ce","#e53e3e"),unsafe_allow_html=True)

        st.markdown('<div class="stl q" style="margin-top:10px">Détail par Poste de Travail — Qualité</div>',unsafe_allow_html=True)
        qcols=["Poste de travail"]+PK+["Score Qualité"]
        qrows=[]
        for p in active_posts:
            if p in ckdf.index:
                row={"Poste de travail":p}
                for k in PK: row[k]="%.1f%%"%ckdf.loc[p,k]
                row["Score Qualité"]="%.2f%%"%qual_scores[p]
                row["_t"]=""; qrows.append(row)
        crow2={"Poste de travail":"CIBLE"}
        for k in PK: crow2[k]="%.0f%%"%CIBLE.get(k,100)
        crow2["Score Qualité"]="≥ 80.00%%"; crow2["_t"]="cible"; qrows.append(crow2)
        st.markdown(html_table(qrows,qcols,"qt",sc_col={"Score Qualité"}),unsafe_allow_html=True)

        # Scores par Poste avec trait objectif
        st.markdown('<div class="stl q" style="margin-top:10px">Scores par Poste de Travail — Qualité</div>',unsafe_allow_html=True)
        st.plotly_chart(create_score_chart(qual_scores,80,"Scores Qualité par Poste","#3182ce","#e53e3e"),use_container_width=True)

        # Anomalies Qualité (INVERSE)
        st.markdown('<div class="stl a" style="margin-top:10px">Anomalies Qualité</div>',unsafe_allow_html=True)
        st.markdown(html_ano(ano_qual_rows,["Poste de travail","Nombre anomalies","Details"]),unsafe_allow_html=True)

    # ============ TAB 3 : ACTIONS RECOMMANDEES ============
    with tab3:
        st.markdown('<div class="stl a">Actions Recommandées — KPI Non Atteints</div>',unsafe_allow_html=True)
        all_actuals={**global_perf,**global_qual}
        st.markdown(html_actions_table(ALL_KPI,all_actuals,CIBLE,ACT_MAP),unsafe_allow_html=True)

    # ============ TAB 4 : REPARTITION OT ============
    with tab4:
        fig_pie, sc_df = create_ot_status_pie(dfp)
        if fig_pie:
            col_p1,col_p2=st.columns([1,1])
            with col_p1: st.plotly_chart(fig_pie,use_container_width=True)
            with col_p2: st.markdown(html_pie_summary_table(sc_df),unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">⚠️ Aucune donnée OT disponible.</div>',unsafe_allow_html=True)

    # ============ TAB 5 : ANALYSE OMS & THERMOGRAPHIE (FUSIONNE) ============
    with tab5:
        desig_col=find_designation_col(dfp)

        # --- OMS ---
        st.markdown('<div class="stl s">🔬 Analyse OMS</div>',unsafe_allow_html=True)
        if desig_col:
            oms_tbl,oms_pie1,oms_pie2,oms_cnt=create_keyword_analysis(dfp,"OMS",desig_col,active_posts,"OMS")
            if oms_tbl:
                st.markdown(f'<div style="margin-bottom:6px;font-size:13px;color:#4a5568;font-weight:600">🔍 {oms_cnt} OT(s) trouvé(s) contenant <b>"OMS"</b> dans la désignation (colonne : {desig_col})</div>',unsafe_allow_html=True)
                st.markdown(oms_tbl,unsafe_allow_html=True)
                c_o1,c_o2=st.columns(2)
                with c_o1:
                    if oms_pie1: st.plotly_chart(oms_pie1,use_container_width=True)
                with c_o2:
                    if oms_pie2: st.plotly_chart(oms_pie2,use_container_width=True)
            else:
                st.markdown('<div class="es">⚠️ Aucun OT contenant "OMS" trouvé.</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">⚠️ Colonne "Désignation" non trouvée. Vérifiez les colonnes du fichier OT.</div>',unsafe_allow_html=True)

        st.markdown('<hr style="margin:16px 0;border:none;border-top:2px solid #e2e8f0">',unsafe_allow_html=True)

        # --- Thermographie ---
        st.markdown('<div class="stl s">🌡️ Analyse Thermographie</div>',unsafe_allow_html=True)
        if desig_col:
            th_tbl,th_pie1,th_pie2,th_cnt=create_keyword_analysis(dfp,"Thermographie",desig_col,active_posts,"Thermographie")
            if th_tbl:
                st.markdown(f'<div style="margin-bottom:6px;font-size:13px;color:#4a5568;font-weight:600">🔍 {th_cnt} OT(s) trouvé(s) contenant <b>"Thermographie"</b> dans la désignation (colonne : {desig_col})</div>',unsafe_allow_html=True)
                st.markdown(th_tbl,unsafe_allow_html=True)
                c_t1,c_t2=st.columns(2)
                with c_t1:
                    if th_pie1: st.plotly_chart(th_pie1,use_container_width=True)
                with c_t2:
                    if th_pie2: st.plotly_chart(th_pie2,use_container_width=True)
            else:
                st.markdown('<div class="es">⚠️ Aucun OT contenant "Thermographie" trouvé.</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">⚠️ Colonne "Désignation" non trouvée.</div>',unsafe_allow_html=True)

    # ============ TAB 6 : ANALYSE BACKLOG ============
    with tab6:
        # --- Backlog Préparation ---
        st.markdown('<div class="stl c">📋 Backlog Préparation</div>',unsafe_allow_html=True)
        pc_df=dfp[dfp["Statut OT"]=="CRÉÉ"].copy() if dfp is not None else pd.DataFrame()
        if not pc_df.empty and active_posts:
            pc_pivot=pd.pivot_table(pc_df,index="Poste travail princ.",columns="Backlog preparation",
                values="Ordre",aggfunc="count",fill_value=0).reindex(active_posts,fill_value=0)
            pc_pivot.index.name="Poste de travail"
            prep_tbl,prep_pie=create_backlog_section(pc_pivot,"Backlog Préparation","bt")
            c_b1,c_b2=st.columns([2,1])
            with c_b1: st.markdown(prep_tbl,unsafe_allow_html=True)
            with c_b2:
                if prep_pie: st.plotly_chart(prep_pie,use_container_width=True)
                else: st.markdown('<div class="es">Aucun backlog</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Aucun OT en backlog de préparation.</div>',unsafe_allow_html=True)

        st.markdown('<hr style="margin:12px 0;border:none;border-top:1px solid #e2e8f0">',unsafe_allow_html=True)

        # --- Backlog Planification ---
        st.markdown('<div class="stl c">📋 Backlog Planification</div>',unsafe_allow_html=True)
        pl_df=dfp[(dfp["Statut OT"]=="LANC")&(dfp["Contient SOPL"]==0)].copy() if dfp is not None else pd.DataFrame()
        if not pl_df.empty and active_posts:
            pl_pivot=pd.pivot_table(pl_df,index="Poste travail princ.",columns="Backlog planification",
                values="Ordre",aggfunc="count",fill_value=0).reindex(active_posts,fill_value=0)
            pl_pivot.index.name="Poste de travail"
            plan_tbl,plan_pie=create_backlog_section(pl_pivot,"Backlog Planification","bt")
            c_b3,c_b4=st.columns([2,1])
            with c_b3: st.markdown(plan_tbl,unsafe_allow_html=True)
            with c_b4:
                if plan_pie: st.plotly_chart(plan_pie,use_container_width=True)
                else: st.markdown('<div class="es">Aucun backlog</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Aucun OT en backlog de planification.</div>',unsafe_allow_html=True)

        st.markdown('<hr style="margin:12px 0;border:none;border-top:1px solid #e2e8f0">',unsafe_allow_html=True)

        # --- Backlog Exécution ---
        st.markdown('<div class="stl c">📋 Backlog Exécution</div>',unsafe_allow_html=True)
        ex_df=dfp[(dfp["Statut OT"]=="LANC")&(dfp["Contient SOPL"]==1)].copy() if dfp is not None else pd.DataFrame()
        if not ex_df.empty and active_posts:
            ex_pivot=pd.pivot_table(ex_df,index="Poste travail princ.",columns="Backlog execution",
                values="Ordre",aggfunc="count",fill_value=0).reindex(active_posts,fill_value=0)
            ex_pivot.index.name="Poste de travail"
            exec_tbl,exec_pie=create_backlog_section(ex_pivot,"Backlog Exécution","bt")
            c_b5,c_b6=st.columns([2,1])
            with c_b5: st.markdown(exec_tbl,unsafe_allow_html=True)
            with c_b6:
                if exec_pie: st.plotly_chart(exec_pie,use_container_width=True)
                else: st.markdown('<div class="es">Aucun backlog</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Aucun OT en backlog d\'exécution.</div>',unsafe_allow_html=True)

    # ===================== SAVE & EXPORT =====================
    # Préparation données pour export Excel
    pcols_exp=["Poste de travail"]+QK+["Score Performance"]
    prows_exp=[]
    for p in active_posts:
        if p in ckdf.index:
            row={"Poste de travail":p}
            for k in QK: row[k]=round(float(ckdf.loc[p,k]),2)
            row["Score Performance"]=round(perf_scores[p],2)
            prows_exp.append(row)
    qcols_exp=["Poste de travail"]+PK+["Score Qualité"]
    qrows_exp=[]
    for p in active_posts:
        if p in ckdf.index:
            row={"Poste de travail":p}
            for k in PK: row[k]=round(float(ckdf.loc[p,k]),2)
            row["Score Qualité"]=round(qual_scores[p],2)
            qrows_exp.append(row)
    ano_p_cols=["Poste de travail","Nombre anomalies","Details"]
    ano_q_cols=["Poste de travail","Nombre anomalies","Details"]

    save_kpis_to_excel(prows_exp,pcols_exp,qrows_exp,qcols_exp,
        ano_perf_rows,ano_p_cols,ano_qual_rows,ano_q_cols,fichier_date)

    st.markdown('<div style="margin-top:12px;padding:10px;background:#fff;border-radius:10px;border:1px solid #e2e8f0;text-align:center;font-size:12px;color:#718096">✅ KPIs sauvegardés automatiquement dans <b>kpis/indicateurs_kpis.xlsx</b> — Onglet : %s</div>'%fichier_date,unsafe_allow_html=True)

if __name__=="__main__":
    main()
