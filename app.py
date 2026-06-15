# -*- coding: utf-8 -*-
# VERSION V2.0 - Synthese & Actions en premier, Pie charts, Objectifs sur barres
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

st.set_page_config(layout="wide", page_title="Dashboard KPI V2")

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

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt","r",encoding="utf-8") as f: return f.read().strip()
        except Exception: pass
    return datetime.now().strftime("%d/%m/%Y")

def save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_r,ano_p_c,ano_q_r,ano_q_c,sheet_name):
    kpis_dir="kpis"
    os.makedirs(kpis_dir,exist_ok=True)
    filepath=os.path.join(kpis_dir,"indicateurs_kpis.xlsx")
    sn=str(sheet_name).replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","")[:31]
    hf=Font(bold=True,color="FFFFFF",size=10)
    hfl=PatternFill(start_color="1E3A5F",end_color="1E3A5F",fill_type="solid")
    tf=Font(bold=True,size=12,color="1E3A5F")
    tb=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    try:
        wb=load_workbook(filepath)
    except Exception:
        wb=Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    if sn in wb.sheetnames:
        del wb[sn]
    ws=wb.create_sheet(sn)
    rn=1
    def ws_sec(title,cols,rows,sr):
        ws.cell(row=sr,column=1,value=title).font=tf
        sr+=1
        for j,c in enumerate(cols,1):
            cl=ws.cell(row=sr,column=j,value=c)
            cl.font=hf
            cl.fill=hfl
            cl.alignment=Alignment(horizontal='center')
            cl.border=tb
        sr+=1
        for r in rows:
            for j,c in enumerate(cols,1):
                cl=ws.cell(row=sr,column=j,value=r.get(c,""))
                cl.border=tb
                cl.alignment=Alignment(horizontal='center')
            sr+=1
        return sr+1
    rn=ws_sec("INDICATEURS DE PERFORMANCE",pcols,prows,rn)
    if ano_p_c and ano_p_r:
        rn=ws_sec("ANOMALIES PERFORMANCE",ano_p_c,ano_p_r,rn)
    rn=ws_sec("INDICATEURS DE QUALITE",qcols,qrows,rn)
    if ano_q_c and ano_q_r:
        rn=ws_sec("ANOMALIES QUALITE",ano_q_c,ano_q_r,rn)
    try:
        wb.save(filepath)
    except Exception:
        pass

def load_historical_kpis(filepath):
    if not os.path.exists(filepath):
        return pd.DataFrame()
    try:
        wb=load_workbook(filepath,read_only=True,data_only=True)
    except Exception:
        return pd.DataFrame()
    records=[]
    section=None
    headers=None
    for sheet_name in wb.sheetnames:
        try:
            ws=wb[sheet_name]
            rows_data=list(ws.iter_rows(values_only=True))
            for row in rows_data:
                cell0=str(row[0]).strip() if row[0] else ""
                if "INDICATEURS DE PERFORMANCE" in cell0.upper():
                    section="perf"
                    headers=None
                    continue
                elif "INDICATEURS DE QUALITE" in cell0.upper():
                    section="qual"
                    headers=None
                    continue
                elif "ANOMALIES" in cell0.upper():
                    section=None
                    continue
                if section and headers is None and cell0:
                    headers=[str(c).strip() if c else "" for c in row]
                    continue
                if section and headers and cell0 and cell0 not in ("CIBLE","Total general",""):
                    entry={"Date":sheet_name}
                    for j,h in enumerate(headers):
                        if j<len(row):
                            entry[h]=row[j]
                    entry["_section"]=section
                    records.append(entry)
        except Exception:
            continue
    wb.close()
    if not records:
        return pd.DataFrame()
    df=pd.DataFrame(records)
    df["Date_parsed"]=pd.to_datetime(df["Date"].str.replace("-","/"),format="%d/%m/%Y",errors="coerce")
    return df.sort_values("Date_parsed").reset_index(drop=True)

def calculate_variations(hist_df):
    if hist_df.empty or "Date" not in hist_df.columns:
        return pd.DataFrame()
    dates=sorted(hist_df["Date"].unique())
    if len(dates)<2:
        return pd.DataFrame()
    perf_df=hist_df[hist_df["_section"]=="perf"].copy()
    qual_df=hist_df[hist_df["_section"]=="qual"].copy()
    variations=[]
    for i in range(1,len(dates)):
        prev_date=dates[i-1]
        curr_date=dates[i]
        if "Poste de travail" in perf_df.columns:
            prev_perf=perf_df[perf_df["Date"]==prev_date].set_index("Poste de travail")
            curr_perf=perf_df[perf_df["Date"]==curr_date].set_index("Poste de travail")
        else:
            prev_perf=pd.DataFrame()
            curr_perf=pd.DataFrame()
        if "Poste de travail" in qual_df.columns:
            prev_qual=qual_df[qual_df["Date"]==prev_date].set_index("Poste de travail")
            curr_qual=qual_df[qual_df["Date"]==curr_date].set_index("Poste de travail")
        else:
            prev_qual=pd.DataFrame()
            curr_qual=pd.DataFrame()
        for sec_name,prev_d,curr_d,kpi_list in [("Performance",prev_perf,curr_perf,QK+["Score Performance"]),("Qualite",prev_qual,curr_qual,PK+["Score Qualite"])]:
            for poste in set(prev_d.index)&set(curr_d.index):
                for kpi in kpi_list:
                    if kpi not in prev_d.columns or kpi not in curr_d.columns:
                        continue
                    try:
                        pv=float(prev_d.loc[poste,kpi])
                    except Exception:
                        continue
                    try:
                        cv=float(curr_d.loc[poste,kpi])
                    except Exception:
                        continue
                    diff=cv-pv
                    if pv!=0:
                        pct=diff/pv*100
                    else:
                        pct=100 if cv!=0 else 0
                    if abs(diff)<=0.5:
                        trend="stabilite"
                    elif diff>0.5:
                        trend="hausse"
                    else:
                        trend="baisse"
                    variations.append({"Date precedente":prev_date,"Date actuelle":curr_date,"Poste":poste,
                        "Type":sec_name,"KPI":kpi,"Valeur precedente":round(pv,2),"Valeur actuelle":round(cv,2),
                        "Ecart":round(diff,2),"Ecart %":round(pct,2),"Tendance":trend})
    return pd.DataFrame(variations)

def generate_journal(var_df):
    if var_df.empty:
        return pd.DataFrame()
    j=var_df.copy()
    j["Significatif"]=j["Ecart %"].abs()>=5
    j=j[j["Significatif"]].copy()
    def sens_func(r):
        if r["Tendance"]=="hausse" and r["KPI"] not in LOWER_BETTER:
            return "Amelioration"
        elif r["Tendance"]=="baisse" and r["KPI"] in LOWER_BETTER:
            return "Amelioration"
        else:
            return "Degradation"
    j["Sens"]=j.apply(sens_func,axis=1)
    return j.sort_values(["Date actuelle","Sens","Ecart %"],ascending=[True,False,False])

def calculate_rankings(var_df):
    if var_df.empty:
        return pd.DataFrame(),pd.DataFrame()
    scores={}
    for poste in var_df["Poste"].unique():
        pv=var_df[var_df["Poste"]==poste].copy()
        s=0
        for _,r in pv.iterrows():
            if r["KPI"] in LOWER_BETTER:
                s=s+(-r["Ecart %"])
            else:
                s=s+r["Ecart %"]
        scores[poste]=s
    ranked=sorted(scores.items(),key=lambda x:x[1],reverse=True)
    top=pd.DataFrame(ranked[:5],columns=["Poste","Score variation"])
    if len(ranked)>5:
        bot=pd.DataFrame(ranked[-5:][::-1],columns=["Poste","Score variation"])
    else:
        bot=pd.DataFrame(columns=["Poste","Score variation"])
    return top,bot

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
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:8px 14px;font-weight:700;font-size:15px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label,div[data-testid="stSidebar"] .stCheckbox label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:3px 8px;margin-bottom:3px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:14px;color:#718096;font-size:14px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:17px}.cg{grid-template-columns:1fr}.car .cal{width:120px}.gbr-l{width:100px}}
    </style>""",unsafe_allow_html=True)

def main():
    try:
        locale.setlocale(locale.LC_ALL,'fr_FR.UTF-8')
    except Exception:
        try:
            locale.setlocale(locale.LC_ALL,'fr_FR')
        except Exception:
            pass

    inject_custom_css()
    fichier_date=get_date_from_file()

    if "hse_affiche" not in st.session_state:
        st.session_state.hse_affiche=False

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
        time.sleep(6)
        st.session_state.hse_affiche=True
        st.rerun()
        st.stop()

    # ---- Fonctions locales ----
    def contient_mot(t,lm):
        t=str(t)
        for l in lm:
            for m in l.split():
                if m in t:
                    return True
        return False

    def cat_age(a):
        if a<=1:
            return "<1 mois"
        elif a>=3:
            return ">3 mois"
        return "1 mois < <3 mois"

    def ckpi(n,d,sz=100):
        return np.where(d==0,sz,(n/d)*100)

    def cpiv(df,f,c,p):
        return pd.pivot_table(df[f],index="Poste travail princ.",columns=c,values="Ordre",aggfunc="count",fill_value=0).reindex(p,fill_value=0)

    def excr(df):
        if "Poste travail princ." in df.columns:
            return df[~df["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)].copy()
        return df

    def is_lb(k):
        return k in LOWER_BETTER

    def calc_kpis(df_i,av_i,now,posts):
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
                df[am]=np.nan
                df[ac]="Inconnu"
        df["OT CONFIME"]=np.where(df["Statut système"].str.contains("CLO",na=False)&df["Statut système"].str.contains("CONF",na=False),"OUI","NON")
        df["Contient SOPL"]=df["Statut utilisateur"].str.contains("SOPL",na=False).map({True:1,False:0})
        df["OT LANC ESTIME"]=np.where(df["Total coûts budgétés"].fillna(0)==0,"NON","OUI")
        df["OT_COR_EGAL"]=np.where((df["Total coûts budgétés"].fillna(0)-df["Total coûts réels"].fillna(0))==0,"OUI","NON")
        res['dfp']=df

        an=cpiv(df,df["Nº appel pl.entret."].fillna(0)==0,"Statut OT",posts)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]:
            an[c]=an.get(c,0)
        an["Total"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
        an["TAUX_REALISATION_CORRECTIF/PT"]=ckpi(an["TCLO"],an["Total"])

        pr=cpiv(df,df["Statut OT"]=="CRÉÉ","ap",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]:
            pr[c]=pr.get(c,0)
        pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"])
        pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0)
        pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)

        pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==0),"alp",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]:
            pl[c]=pl.get(c,0)
        pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"])
        pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0)
        pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)

        ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois"]:
            ex[c]=ex.get(c,0)
        ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
        ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"])
        ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0)
        ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)

        la=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]:
            la[c]=la.get(c,0)
        la["Total"]=la["OUI"]+la["NON"]
        la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"])

        pc=pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]:
            pc[c]=pc.get(c,0)
        pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]
        pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"])

        plc=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]:
            plc[c]=plc.get(c,0)
        plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]
        plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])

        for kn,cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
            pv=pd.pivot_table(df,index="Poste travail princ.",columns=cn,values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
            for c in ["OUI","NON"]:
                pv[c]=pv.get(c,0)
            pv["Total"]=pv["OUI"]+pv["NON"]
            pv[cn]=ckpi(pv["OUI"],pv["Total"])
            res[kn.lower().replace(" ","_")]=pv

        avf=av[(av["Ordre"].isna())|(av["Ordre"].astype(str).str.strip()=="")].copy()
        res['avf']=avf

        tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]:
            tca[c]=tca.get(c,0)
        tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1)
        tca["appel avis approuvé"]=ckpi(tca["APRV"],tca["Total"])

        res['ckdf']=pd.DataFrame({
            "TAUX_REALISATION_CORRECTIF/PT":an["TAUX_REALISATION_CORRECTIF/PT"],
            "OT préparation <1 mois":pr["OT préparation <1 mois"],
            "OT préparation >3 mois":pr["OT préparation >3 mois"],
            "OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois":pl["OT planification <1 mois"],
            "OT planification >3 mois":pl["OT planification >3 mois"],
            "OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois":ex["OT exécution <1 mois"],
            "OT exécution >3 mois":ex["OT exécution >3 mois"],
            "OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
            "appel avis approuvé":tca["appel avis approuvé"],
            "OT LANC ESTIME":la["OT LANC ESTIME"],
            "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],
            "Backlog planification caractérisé":plc["Backlog planification caractérisé"],
            "OT CONFIME":res['ot_confime']["OT CONFIME"],
            "OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"]
        })
        return res

    # ---- Style cellules ----
    def ks(v,c):
        try:
            val=float(v)
        except Exception:
            return ""
        if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
            if val>=80:
                return "background:#c6efce;color:#006100;font-weight:600"
            elif val>=75:
                return "background:#ffeb9c;color:#9c6500;font-weight:600"
            else:
                return "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
            if val<=15:
                return "background:#c6efce;color:#006100;font-weight:600"
            else:
                return "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
            if val<=5:
                return "background:#c6efce;color:#006100;font-weight:600"
            else:
                return "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c=="TAUX_REALISATION_CORRECTIF/PT":
            if val>=85:
                return "background:#c6efce;color:#006100;font-weight:600"
            elif val>=80:
                return "background:#ffeb9c;color:#9c6500;font-weight:600"
            else:
                return "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c=="appel avis approuvé":
            if val>=95:
                return "background:#c6efce;color:#006100;font-weight:600"
            elif val>=90:
                return "background:#ffeb9c;color:#9c6500;font-weight:600"
            else:
                return "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]:
            if val>=100:
                return "background:#c6efce;color:#006100;font-weight:600"
            elif val>=95:
                return "background:#ffeb9c;color:#9c6500;font-weight:600"
            else:
                return "background:#ffc7ce;color:#9c0006;font-weight:600"
        return ""

    def cs(v):
        try:
            val=float(str(v).replace(' %','').strip())
        except Exception:
            return ""
        if val>=90:
            return "background:#c6efce;color:#006100;font-weight:700"
        elif val>=80:
            return "background:#ffeb9c;color:#9c6500;font-weight:700"
        else:
            return "background:#ffc7ce;color:#9c0006;font-weight:700"

    def kas(v):
        try:
            val=int(v)
        except Exception:
            return ""
        if val==0:
            return "color:#cbd5e0"
        elif val<=3:
            return "background:#ffeb9c;color:#9c6500;font-weight:600"
        elif val<=10:
            return "background:#fed7d7;color:#c53030;font-weight:600"
        else:
            return "background:#fc8181;color:#742a2a;font-weight:800"

    def gscore(k,a,t):
        if pd.isna(a) or pd.isna(t):
            return 0
        if k in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
            return 1 if a>=75 else 0
        if k in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
            return 1 if a<=15 else 0
        if k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
            return 1 if a<=5 else 0
        if k=="TAUX_REALISATION_CORRECTIF/PT":
            return 1 if a>=80 else 0
        if k=="appel avis approuvé":
            return 1 if a>=90 else 0
        if k in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]:
            return 1 if a>=95 else 0
        return 0

    # ---- HTML helpers ----
    def html_table(rows,cols,tc,sc_col=None):
        h='<table class="tw %s"><thead><tr>'%tc+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            rc="cb" if r.get("_t")=="cible" else ("tr" if r.get("_t")=="total" else "")
            h+='<tr class="%s">'%rc
            for c in cols:
                v=r.get(c,"")
                if r.get("_t")=="cible":
                    h+='<td>%s</td>'%v
                else:
                    s=cs(v) if sc_col and c in sc_col else ks(v,c)
                    h+='<td style="%s">%s</td>'%(s or "",v)
            h+='</tr>'
        return h+'</tbody></table>'

    def html_ano(rows,cols):
        h='<table class="tw at"><thead><tr>'+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            h+='<tr class="%s">'%("tr" if r.get("_t")=="total" else "")
            for c in cols:
                v=r.get(c,"")
                h+='<td style="%s">%s</td>'%(kas(v) or "",v)
            h+='</tr>'
        return h+'</tbody></table>'

    def html_actions_table(kpi_list,actuals,targets,act_map):
        h='<table class="tw at"><thead><tr><th>KPI</th><th>Valeur Actuelle</th><th>Cible</th><th>Ecart</th><th>Statut</th><th>Action Recommandée</th></tr></thead><tbody>'
        for k in kpi_list:
            av=actuals.get(k,0)
            tv=targets.get(k,100)
            diff=av-tv
            met=av<=tv if is_lb(k) else av>=tv
            status="ATTEINT" if met else "NON ATTEINT"
            st_s="background:#c6efce;color:#006100;font-weight:700" if met else "background:#ffc7ce;color:#9c0006;font-weight:700"
            ec_clr="#276749" if met else "#c53030"
            action="Objectif atteint" if met else act_map.get(k,"")
            h+='<tr><td style="font-weight:600">%s</td><td>%.1f%%</td><td>%.0f%%</td><td style="color:%s;font-weight:700">%+.1f%%</td><td style="%s">%s</td><td style="color:#4a5568">%s</td></tr>'%(k,av,tv,ec_clr,diff,st_s,status,action)
        return h+'</tbody></table>'

    def html_classement(scores,accent):
        sp=sorted(scores.items(),key=lambda x:x[1],reverse=True)
        met_p=[(p,s) for p,s in sp if s>=80]
        not_p=[(p,s) for p,s in sp if s<80]
        t5=met_p[:5]
        b5=not_p[-5:] if len(not_p)>5 else not_p
        h='<div class="cg"><div><div class="ct" style="color:#38a169">Top 5 — Objectif Atteint</div>'
        if t5:
            for i,(p,s) in enumerate(t5):
                h+='<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(accent,i+1,p,cs("%.2f"%s),s)
        else:
            h+='<div style="padding:6px;font-size:12px;color:#718096">Aucun poste</div>'
        h+='</div><div><div class="ct" style="color:#e53e3e">Bottom 5 — Non Atteint</div>'
        if b5:
            for i,(p,s) in enumerate(reversed(b5)):
                h+='<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(len(b5)-i,p,cs("%.2f"%s),s)
        else:
            h+='<div style="padding:6px;font-size:12px;color:#38a169">Tous atteints</div>'
        h+='</div></div>'
        return h

    def html_grouped_bars(posts,pscores,qscores,title):
        h='<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>'%title
        h+='<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span></div>'
        for p in sorted(posts,key=lambda x:(pscores.get(x,0)+qscores.get(x,0))/2,reverse=True):
            pv=pscores.get(p,0)
            qv=qscores.get(p,0)
            h+='<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f gb-p" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w"><div class="gbr-f gb-q" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div></div></div>'%(p,min(max(pv,0),100),pv,min(max(qv,0),100),qv)
        return h+'</div>'

    # ---- Pie intelligent ----
    def anl_pie_chart(data,names_col,values_col,title,colors=None,min_pct=3.0):
        if data.empty:
            return None
        df=data[[names_col,values_col]].dropna().copy()
        df[values_col]=pd.to_numeric(df[values_col],errors='coerce').fillna(0)
        total=df[values_col].sum()
        if total==0:
            return None
        df["_pct"]=df[values_col]/total*100
        big=df[df["_pct"]>=min_pct].copy()
        small=df[df["_pct"]<min_pct].copy()
        has_small=len(small)>=1 and small[values_col].sum()>0
        if not has_small:
            fig=px.pie(df,names=names_col,values=values_col,title="<b>%s</b>"%title,
                       color_discrete_sequence=colors or px.colors.qualitative.Set2)
            fig.update_traces(textposition='inside',textinfo='percent+label+value',textfont_size=12,pull=[0.02]*len(df))
            fig.update_layout(margin=dict(t=60,b=50,l=20,r=20),height=460,autosize=True,title_font_size=15,
                legend=dict(font_size=11,orientation="h",yanchor="bottom",y=-0.15,title_text="Legende",title_font_size=12))
            return fig
        else:
            others_label="Autres (%d secteurs)"%len(small)
            others_row=pd.DataFrame([{names_col:others_label,values_col:small[values_col].sum(),"_pct":small["_pct"].sum()}])
            main_df=pd.concat([big,others_row],ignore_index=True)
            sub_df=small.sort_values(values_col,ascending=False).copy()
            base_colors=colors or px.colors.qualitative.Set2
            main_colors=[]
            for i in range(len(main_df)):
                if i==len(main_df)-1:
                    main_colors.append("#CBD5E0")
                else:
                    main_colors.append(base_colors[i%len(base_colors)])
            sub_colors=[base_colors[(len(big)+i)%len(base_colors)] for i in range(len(sub_df))]
            fig=make_subplots(rows=1,cols=2,specs=[[{"type":"pie"},{"type":"pie"}]],
                subplot_titles=["<b>%s</b>"%title,"<b>Detail 'Autres' (%d secteurs)</b>"%len(small)],horizontal_spacing=0.08)
            fig.add_trace(go.Pie(labels=main_df[names_col].tolist(),values=main_df[values_col].tolist(),
                textinfo='percent+label+value',textposition='inside',textfont_size=12,marker_colors=main_colors,
                pull=[0.03 if i==len(main_df)-1 else 0.01 for i in range(len(main_df))]),row=1,col=1)
            fig.add_trace(go.Pie(labels=sub_df[names_col].tolist(),values=sub_df[values_col].tolist(),
                textinfo='percent+label+value',textposition='inside',textfont_size=11,marker_colors=sub_colors,hole=0.3),row=1,col=2)
            fig.update_layout(margin=dict(t=60,b=50,l=10,r=10),height=460,autosize=True,title_font_size=15,showlegend=True,
                legend=dict(font_size=10,orientation="h",yanchor="bottom",y=-0.08,title_text="Legende",title_font_size=11))
            return fig

    # ---- Barres Top5/Bottom5 avec ligne objectif ----
    def make_rank_bar_chart(top_df,bottom_df,kpi_name,target_val,color_top,color_bottom):
        all_items=[]
        if not top_df.empty:
            for _,r in top_df.iterrows():
                all_items.append({"Poste":r["Poste"],"Score":r["Score"],"Groupe":"Top 5"})
        if not bottom_df.empty:
            for _,r in bottom_df.iterrows():
                all_items.append({"Poste":r["Poste"],"Score":r["Score"],"Groupe":"Bottom 5"})
        if not all_items:
            return None
        rdf=pd.DataFrame(all_items).sort_values("Score",ascending=True)
        bar_colors=[color_top if g=="Top 5" else color_bottom for g in rdf["Groupe"]]
        fig=go.Figure()
        fig.add_trace(go.Bar(y=rdf["Poste"],x=rdf["Score"],orientation='h',marker_color=bar_colors,
            text=["%.1f%%"%s for s in rdf["Score"]],textposition='outside',
            textfont=dict(size=12,color="#1a202c",family="Inter"),
            hovertemplate="<b>%{y}</b><br>Score: %{x:.1f}%%<extra></extra>"))
        fig.add_vline(x=target_val,line_dash="dash",line_width=2.5,line_color="#e53e3e",
            annotation_text="Objectif: %d%%"%target_val,annotation_position="top right",
            annotation_font=dict(size=12,color="#e53e3e",family="Inter",weight="bold"),
            annotation_bgcolor="rgba(255,255,255,0.85)",annotation_bordercolor="#e53e3e",annotation_borderwidth=1)
        fig.update_layout(title="<b>%s</b>"%kpi_name,height=max(220,len(rdf)*48+100),
            margin=dict(l=180,r=80,t=60,b=30),
            xaxis=dict(range=[0,max(110,rdf["Score"].max()*1.15)],title="Score (%)",gridcolor="#edf2f7",zeroline=False),
            yaxis=dict(tickfont=dict(size=12,family="Inter")),
            plot_bgcolor="white",font=dict(family="Inter"),showlegend=False,bargap=0.35)
        return fig

    # ---- Barres KPI avec lignes cible ----
    def make_kpi_bar_chart_with_target(kpi_list,actuals,targets,title,color_ok,color_fail):
        names=[]
        vals=[]
        colors=[]
        cibles=[]
        for k in kpi_list:
            av=actuals.get(k,0)
            tv=targets.get(k,100)
            met=av<=tv if is_lb(k) else av>=tv
            names.append(k)
            vals.append(round(av,1))
            colors.append(color_ok if met else color_fail)
            cibles.append(tv)
        fig=go.Figure()
        fig.add_trace(go.Bar(x=names,y=vals,marker_color=colors,
            text=["%.1f%%"%v for v in vals],textposition='outside',
            textfont=dict(size=11,color="#1a202c",family="Inter"),
            hovertemplate="<b>%{x}</b><br>Valeur: %{y:.1f}%%<extra></extra>",name="Valeur"))
        for i,(n,tv) in enumerate(zip(names,cibles)):
            fig.add_shape(type="line",x0=i-0.4,x1=i+0.4,y0=tv,y1=tv,
                line=dict(color="#e53e3e",width=2.5,dash="dash"))
            fig.add_annotation(x=i,y=tv,text="Cible %d%%"%tv,showarrow=False,yshift=8,
                font=dict(size=9,color="#e53e3e",family="Inter",weight="bold"),
                bgcolor="rgba(255,255,255,0.8)",bordercolor="#e53e3e",borderwidth=0.5)
        fig.update_layout(title="<b>%s</b>"%title,height=420,margin=dict(l=20,r=20,t=60,b=140),
            xaxis=dict(tickangle=-45,tickfont=dict(size=10,family="Inter")),
            yaxis=dict(range=[0,max(max(vals),max(cibles))*1.2],title="%",gridcolor="#edf2f7"),
            plot_bgcolor="white",font=dict(family="Inter"),showlegend=False,bargap=0.3)
        return fig

    def export_btn(df,filename):
        buf=io.BytesIO()
        df.to_excel(buf,index=False,engine='openpyxl')
        buf.seek(0)
        st.download_button("📥 Exporter Excel",data=buf,file_name=filename,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""<div style="padding:10px 0 4px 0"><div style="font-size:22px;margin-bottom:2px">⚙️</div><div style="font-size:14px;font-weight:800;color:white">Filtres & Parametres</div><div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Configuration</div></div>""",unsafe_allow_html=True)
        st.markdown("---")
        show_filters=st.checkbox("Afficher les filtres",value=True,key="show_filters")
        if show_filters:
            unf=st.toggle("📁 Charger nouveaux fichiers",value=False,key="tf")
            ot_f=None
            av_f=None
            apm=[]
            if unf:
                ot_f=st.file_uploader("Fichier OT",type=["xlsx"],key="uot")
                av_f=st.file_uploader("Fichier AVIS",type=["xlsx"],key="uav")
            else:
                if os.path.exists("ot.xlsx"):
                    try:
                        _t=excr(pd.read_excel("ot.xlsx"))
                        apm=sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                    except Exception:
                        pass
                st.markdown("""<div style="background:rgba(255,255,255,.1);padding:6px 10px;border-radius:6px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Donnees</div><div style="font-size:14px;color:white;font-weight:600;margin-top:2px">📅 %s</div></div>"""%fichier_date,unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("**🎯 Postes**")
            sp=st.multiselect("Poste",["All"]+apm,["All"],key="sp")
            st.markdown("**🏭 Atelier**")
            sa=st.multiselect("Atelier",["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)"],["All"],key="sa")
            st.markdown("**🏢 Division**")
            sd=st.multiselect("Division",["All","SF1","SF2"],["All"],key="sd")
            st.markdown("---")
            st.markdown("**📅 Periode**")
            dr=st.date_input("Date debut planifiee",value=(datetime(2025,1,1).date(),datetime.today().date()),format="DD/MM/YYYY",key="dr")
        else:
            unf=False
            ot_f=None
            av_f=None
            apm=[]
            sp=["All"]
            sa=["All"]
            sd=["All"]
            dr=(datetime(2025,1,1).date(),datetime.today().date())
            if os.path.exists("ot.xlsx"):
                try:
                    _t=excr(pd.read_excel("ot.xlsx"))
                    apm=sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                except Exception:
                    pass

    # ===================== DATA LOADING =====================
    if not unf or (ot_f is not None and av_f is not None):
        try:
            if unf:
                raw_ot=pd.read_excel(ot_f)
                raw_av=pd.read_excel(av_f)
            else:
                raw_ot=pd.read_excel("ot.xlsx")
                raw_av=pd.read_excel("avis.xlsx")
            raw_ot=excr(raw_ot)
            raw_av=excr(raw_av)
            for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
                if c in raw_ot.columns:
                    raw_ot[c]=pd.to_datetime(raw_ot[c],errors="coerce")
            for c in ["Créé le","Début souhaité","Date de la clôture"]:
                if c in raw_av.columns:
                    raw_av[c]=pd.to_datetime(raw_av[c],errors="coerce")
            if not apm:
                apm=sorted(raw_ot[raw_ot["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
            if "All" in sp or not sp:
                sp_sel=apm
            else:
                sp_sel=[p for p in sp if p in apm]
            if "All" in sa or not sa:
                sa_sel=["All"]
            else:
                sa_sel=sa
            if "All" in sd or not sd:
                sd_sel=["All"]
            else:
                sd_sel=sd
            sdt=pd.to_datetime(dr[0]) if len(dr)==2 else pd.to_datetime(datetime(2025,1,1))
            edt=pd.to_datetime(dr[1]) if len(dr)==2 else pd.to_datetime(datetime.today())

            def mf(poste):
                p=str(poste).upper()
                if "All" not in sa_sel:
                    m=False
                    for a in sa_sel:
                        if a=="Sulfurique (PS)" and "PS" in p:
                            m=True
                        elif a=="Phosphorique (PP)" and "PP" in p:
                            m=True
                        elif a=="Engrais (TSP/REX)" and ("TSP" in p or "REX" in p):
                            m=True
                        elif a=="Feed (MCP/DCP)" and ("MCP" in p or "DCP" in p):
                            m=True
                    if not m:
                        return False
                if "All" not in sd_sel:
                    m=False
                for d in sd_sel:
                    if d in p:
                        m=True
                if "All" not in sd_sel and not m:
                    return False
                return True

            filtered_posts=[p for p in sp_sel if mf(p)]
            if not filtered_posts:
                filtered_posts=[p for p in apm if mf(p)]
            if not filtered_posts:
                filtered_posts=apm

            df_filt=raw_ot[(raw_ot["Poste travail princ."].isin(filtered_posts))&(raw_ot["Date de début planifiée"]>=sdt)&(raw_ot["Date de début planifiée"]<=edt)].copy()
            av_filt=raw_av[raw_av["Poste travail princ."].isin(filtered_posts)].copy()

            now=datetime.now()
            res=calc_kpis(df_filt,av_filt,now,filtered_posts)
            ckdf=res['ckdf']
            dfp=res['dfp']
            avf=res['avf']

            # ===================== SCORES =====================
            perf_scores={}
            qual_scores={}
            for p in filtered_posts:
                ps_sum=0
                qs_sum=0
                ps_n=0
                qs_n=0
                for k in QK:
                    if k in ckdf.columns and p in ckdf.index:
                        try:
                            v=float(ckdf.loc[p,k])
                            ps_sum+=gscore(k,v,CIBLE.get(k,100))
                            ps_n+=1
                        except Exception:
                            pass
                for k in PK:
                    if k in ckdf.columns and p in ckdf.index:
                        try:
                            v=float(ckdf.loc[p,k])
                            qs_sum+=gscore(k,v,CIBLE.get(k,100))
                            qs_n+=1
                        except Exception:
                            pass
                perf_scores[p]=(ps_sum/ps_n*100) if ps_n>0 else 0
                qual_scores[p]=(qs_sum/qs_n*100) if qs_n>0 else 0

            global_perf=sum(perf_scores.values())/len(perf_scores) if perf_scores else 0
            global_qual=sum(qual_scores.values())/len(qual_scores) if qual_scores else 0
            total_ot=len(df_filt)
            total_avis=len(avf)

            # ===================== HEADER AVEC V2.0 =====================
            st.markdown("""<div class="mh"><h1>📊 Dashboard KPI Maintenance [V2.0]</h1><span class="db">📅 %s</span></div>"""%fichier_date,unsafe_allow_html=True)

            # ===================== CARTES RESUME (SANS perf/qual periode global) =====================
            st.markdown('<div class="cr">',unsafe_allow_html=True)
            st.markdown('<div class="cc c1"><div class="cv">%d</div><div class="cl">Total OT</div></div>'%total_ot,unsafe_allow_html=True)
            st.markdown('<div class="cc c2"><div class="cv">%d</div><div class="cl">Postes</div></div>'%len(filtered_posts),unsafe_allow_html=True)
            st.markdown('<div class="cc c3"><div class="cv">%d</div><div class="cl">Avis sans OT</div></div>'%total_avis,unsafe_allow_html=True)
            nb_non_atteints=sum(1 for s in perf_scores.values() if s<80)+sum(1 for s in qual_scores.values() if s<80)
            st.markdown('<div class="cc c4"><div class="cv">%d</div><div class="cl">Postes Non Atteints</div></div>'%nb_non_atteints,unsafe_allow_html=True)
            st.markdown('</div>',unsafe_allow_html=True)

            # ===================== TABS — Synthese EN PREMIER =====================
            tabs=st.tabs(["📊 Synthèse & Actions","📈 Performance","📋 Qualité","⚠️ Anomalies","🏆 Classement","📉 Variations"])

            # ============ TAB 0 : SYNTHESE & ACTIONS ============
            with tabs[0]:
                st.markdown('<div class="stl a">🎯 Synthese & Actions Correctives</div>',unsafe_allow_html=True)

                # Top5/Bottom5 Performance
                st.markdown('<div class="stl p" style="margin-top:10px">🏆 Classement Performance — Top 5 & Bottom 5</div>',unsafe_allow_html=True)
                perf_sorted=sorted(perf_scores.items(),key=lambda x:x[1],reverse=True)
                perf_top5_df=pd.DataFrame(perf_sorted[:5],columns=["Poste","Score"]) if len(perf_sorted)>=1 else pd.DataFrame(columns=["Poste","Score"])
                perf_bot5_df=pd.DataFrame(perf_sorted[-5:][::-1],columns=["Poste","Score"]) if len(perf_sorted)>5 else pd.DataFrame(columns=["Poste","Score"])
                fig_pr=make_rank_bar_chart(perf_top5_df,perf_bot5_df,"Performance par Poste",80,"#38a169","#e53e3e")
                if fig_pr:
                    st.plotly_chart(fig_pr,use_container_width=True)
                else:
                    st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)

                # Top5/Bottom5 Qualite
                st.markdown('<div class="stl q" style="margin-top:10px">🏆 Classement Qualite — Top 5 & Bottom 5</div>',unsafe_allow_html=True)
                qual_sorted=sorted(qual_scores.items(),key=lambda x:x[1],reverse=True)
                qual_top5_df=pd.DataFrame(qual_sorted[:5],columns=["Poste","Score"]) if len(qual_sorted)>=1 else pd.DataFrame(columns=["Poste","Score"])
                qual_bot5_df=pd.DataFrame(qual_sorted[-5:][::-1],columns=["Poste","Score"]) if len(qual_sorted)>5 else pd.DataFrame(columns=["Poste","Score"])
                fig_qr=make_rank_bar_chart(qual_top5_df,qual_bot5_df,"Qualite par Poste",80,"#3182ce","#e53e3e")
                if fig_qr:
                    st.plotly_chart(fig_qr,use_container_width=True)
                else:
                    st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)

                # Pie OT par Statut
                st.markdown('<div class="stl s" style="margin-top:10px">🥧 Repartition des OT par Statut</div>',unsafe_allow_html=True)
                if "Statut OT" in dfp.columns and not dfp.empty:
                    ot_stat_df=dfp.groupby("Statut OT").size().reset_index(name="Nombre")
                    ot_stat_df=ot_stat_df[ot_stat_df["Nombre"]>0]
                    fig_op=anl_pie_chart(ot_stat_df,"Statut OT","Nombre","Nombre d'OT par Statut",
                        colors=["#38a169","#3182ce","#805ad5","#e53e3e","#d69e2e","#ed8936","#4299e1","#9f7aea"])
                    if fig_op:
                        st.plotly_chart(fig_op,use_container_width=True)
                    else:
                        st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Colonne Statut OT non disponible</div>',unsafe_allow_html=True)

                # Pie Avis par Statut
                st.markdown('<div class="stl c" style="margin-top:10px">🥧 Repartition des Avis par Statut</div>',unsafe_allow_html=True)
                if "Statut utilisateur" in avf.columns and not avf.empty:
                    av_stat_df=avf.groupby("Statut utilisateur").size().reset_index(name="Nombre")
                    av_stat_df=av_stat_df[av_stat_df["Nombre"]>0]
                    fig_ap=anl_pie_chart(av_stat_df,"Statut utilisateur","Nombre","Nombre d'Avis par Statut",
                        colors=["#38a169","#3182ce","#805ad5","#e53e3e","#d69e2e","#ed8936","#4299e1","#9f7aea"])
                    if fig_ap:
                        st.plotly_chart(fig_ap,use_container_width=True)
                    else:
                        st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Aucun avis sans OT</div>',unsafe_allow_html=True)

                # Barres Performance avec cible
                st.markdown('<div class="stl p" style="margin-top:10px">📊 Indicateurs Performance avec Objectifs</div>',unsafe_allow_html=True)
                perf_avg={k:ckdf[k].mean() for k in QK if k in ckdf.columns}
                fig_pb=make_kpi_bar_chart_with_target(QK,perf_avg,CIBLE,"Indicateurs Performance — Moyenne (ligne rouge = objectif)","#38a169","#e53e3e")
                if fig_pb:
                    st.plotly_chart(fig_pb,use_container_width=True)

                # Barres Qualite avec cible
                st.markdown('<div class="stl q" style="margin-top:10px">📊 Indicateurs Qualite avec Objectifs</div>',unsafe_allow_html=True)
                qual_avg={k:ckdf[k].mean() for k in PK if k in ckdf.columns}
                fig_qb=make_kpi_bar_chart_with_target(PK,qual_avg,CIBLE,"Indicateurs Qualite — Moyenne (ligne rouge = objectif)","#3182ce","#e53e3e")
                if fig_qb:
                    st.plotly_chart(fig_qb,use_container_width=True)

                # Tableau Actions
                st.markdown('<div class="stl a" style="margin-top:10px">🔧 Plan d\'Actions Correctives</div>',unsafe_allow_html=True)
                all_avg={**perf_avg,**qual_avg}
                st.markdown(html_actions_table(ALL_KPI,all_avg,CIBLE,ACT_MAP),unsafe_allow_html=True)

                col_e1,col_e2=st.columns([1,1])
                with col_e1:
                    export_btn(ckdf.round(2).reset_index(),"synthese_kpis.xlsx")
                with col_e2:
                    st.info("💡 Les lignes rouges pointillees = objectif cible de chaque indicateur.")

            # ============ TAB 1 : PERFORMANCE ============
            with tabs[1]:
                st.markdown('<div class="stl p">📈 Indicateurs de Performance</div>',unsafe_allow_html=True)
                pcols_t=["Poste de travail"]+QK+["Score Performance"]
                prows=[]
                for p in filtered_posts:
                    rd={"Poste de travail":p,"_t":""}
                    for k in QK:
                        if k in ckdf.columns and p in ckdf.index:
                            try:
                                rd[k]=round(float(ckdf.loc[p,k]),1)
                            except Exception:
                                rd[k]="-"
                        else:
                            rd[k]="-"
                    rd["Score Performance"]=round(perf_scores.get(p,0),2)
                    prows.append(rd)
                cr={"Poste de travail":"CIBLE","_t":"cible"}
                for k in QK:
                    cr[k]=CIBLE.get(k,"-")
                cr["Score Performance"]="80%"
                prows.append(cr)
                tr={"Poste de travail":"Moyenne","_t":"total"}
                for k in QK:
                    if k in ckdf.columns:
                        tr[k]=round(ckdf[k].mean(),1)
                    else:
                        tr[k]="-"
                tr["Score Performance"]=round(global_perf,2)
                prows.append(tr)
                st.markdown(html_table(prows,pcols_t,"pt",sc_col=["Score Performance"]),unsafe_allow_html=True)
                export_btn(pd.DataFrame(prows).drop(columns=["_t"],errors="ignore"),"performance_detail.xlsx")

            # ============ TAB 2 : QUALITE ============
            with tabs[2]:
                st.markdown('<div class="stl q">📋 Indicateurs de Qualite</div>',unsafe_allow_html=True)
                qcols_t=["Poste de travail"]+PK+["Score Qualite"]
                qrows=[]
                for p in filtered_posts:
                    rd={"Poste de travail":p,"_t":""}
                    for k in PK:
                        if k in ckdf.columns and p in ckdf.index:
                            try:
                                rd[k]=round(float(ckdf.loc[p,k]),1)
                            except Exception:
                                rd[k]="-"
                        else:
                            rd[k]="-"
                    rd["Score Qualite"]=round(qual_scores.get(p,0),2)
                    qrows.append(rd)
                cr2={"Poste de travail":"CIBLE","_t":"cible"}
                for k in PK:
                    cr2[k]=CIBLE.get(k,"-")
                cr2["Score Qualite"]="80%"
                qrows.append(cr2)
                tr2={"Poste de travail":"Moyenne","_t":"total"}
                for k in PK:
                    if k in ckdf.columns:
                        tr2[k]=round(ckdf[k].mean(),1)
                    else:
                        tr2[k]="-"
                tr2["Score Qualite"]=round(global_qual,2)
                qrows.append(tr2)
                st.markdown(html_table(qrows,qcols_t,"qt",sc_col=["Score Qualite"]),unsafe_allow_html=True)
                export_btn(pd.DataFrame(qrows).drop(columns=["_t"],errors="ignore"),"qualite_detail.xlsx")

            # ============ TAB 3 : ANOMALIES ============
            with tabs[3]:
                st.markdown('<div class="stl a">⚠️ Anomalies Detectees</div>',unsafe_allow_html=True)
                ano_p=[]
                ano_q=[]
                for p in filtered_posts:
                    np_c=0
                    nq_c=0
                    for k in QK:
                        if k in ckdf.columns and p in ckdf.index:
                            try:
                                v=float(ckdf.loc[p,k])
                                if is_lb(k):
                                    if v>CIBLE.get(k,100):
                                        np_c+=1
                                else:
                                    if v<CIBLE.get(k,0):
                                        np_c+=1
                            except Exception:
                                pass
                    for k in PK:
                        if k in ckdf.columns and p in ckdf.index:
                            try:
                                v=float(ckdf.loc[p,k])
                                if v<CIBLE.get(k,100):
                                    nq_c+=1
                            except Exception:
                                pass
                    if np_c>0:
                        ano_p.append({"Poste de travail":p,"Nombre anomalies":np_c,"_t":""})
                    if nq_c>0:
                        ano_q.append({"Poste de travail":p,"Nombre anomalies":nq_c,"_t":""})
                ano_p.sort(key=lambda x:x["Nombre anomalies"],reverse=True)
                ano_q.sort(key=lambda x:x["Nombre anomalies"],reverse=True)
                tp=sum(r["Nombre anomalies"] for r in ano_p)
                tq=sum(r["Nombre anomalies"] for r in ano_q)
                if ano_p:
                    ano_p.append({"Poste de travail":"TOTAL","Nombre anomalies":tp,"_t":"total"})
                if ano_q:
                    ano_q.append({"Poste de travail":"TOTAL","Nombre anomalies":tq,"_t":"total"})
                st.markdown('<div class="stl p" style="font-size:13px">Anomalies Performance</div>',unsafe_allow_html=True)
                if ano_p:
                    st.markdown(html_ano(ano_p,["Poste de travail","Nombre anomalies"]),unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Aucune anomalie performance</div>',unsafe_allow_html=True)
                st.markdown('<div class="stl q" style="font-size:13px;margin-top:8px">Anomalies Qualite</div>',unsafe_allow_html=True)
                if ano_q:
                    st.markdown(html_ano(ano_q,["Poste de travail","Nombre anomalies"]),unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Aucune anomalie qualite</div>',unsafe_allow_html=True)

            # ============ TAB 4 : CLASSEMENT ============
            with tabs[4]:
                st.markdown('<div class="stl c">🏆 Classement General des Postes</div>',unsafe_allow_html=True)
                st.markdown(html_classement(perf_scores,"#38a169"),unsafe_allow_html=True)
                st.markdown('<div style="height:8px"></div>',unsafe_allow_html=True)
                st.markdown(html_classement(qual_scores,"#3182ce"),unsafe_allow_html=True)
                st.markdown('<div style="height:8px"></div>',unsafe_allow_html=True)
                st.markdown(html_grouped_bars(filtered_posts,perf_scores,qual_scores,"Comparaison Performance / Qualite par Poste"),unsafe_allow_html=True)

            # ============ TAB 5 : VARIATIONS ============
            with tabs[5]:
                st.markdown('<div class="stl s">📉 Variations & Journal des Modifications</div>',unsafe_allow_html=True)
                kpis_path=os.path.join("kpis","indicateurs_kpis.xlsx")
                if os.path.exists(kpis_path):
                    hist_df=load_historical_kpis(kpis_path)
                    var_df=calculate_variations(hist_df)
                    if not var_df.empty:
                        st.markdown("**Journal des variations significatives (|ecart| >= 5%)**")
                        journal=generate_journal(var_df)
                        if not journal.empty:
                            st.dataframe(journal.drop(columns=["Significatif"],errors="ignore"),use_container_width=True,height=400)
                            export_btn(journal,"journal_variations.xlsx")
                        else:
                            st.markdown('<div class="es">Aucune variation significative</div>',unsafe_allow_html=True)
                        top5_v,bot5_v=calculate_rankings(var_df)
                        if not top5_v.empty:
                            st.markdown('<div class="stl p" style="margin-top:10px">Top 5 — Meilleure progression</div>',unsafe_allow_html=True)
                            st.dataframe(top5_v,use_container_width=True)
                        if not bot5_v.empty:
                            st.markdown('<div class="stl a" style="margin-top:10px">Bottom 5 — Plus forte degradation</div>',unsafe_allow_html=True)
                            st.dataframe(bot5_v,use_container_width=True)
                    else:
                        st.markdown('<div class="es">Pas assez de donnees historiques (minimum 2 periodes)</div>',unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Aucun fichier historique. Sauvegardez d\'abord les KPIs.</div>',unsafe_allow_html=True)
                st.markdown("---")
                if st.button("💾 Sauvegarder les KPIs pour historique",key="save_hist"):
                    pcols_s=["Poste de travail"]+QK+["Score Performance"]
                    prows_s=[]
                    for p in filtered_posts:
                        rd={"Poste de travail":p}
                        for k in QK:
                            if k in ckdf.columns and p in ckdf.index:
                                try:
                                    rd[k]=round(float(ckdf.loc[p,k]),1)
                                except Exception:
                                    rd[k]=0
                            else:
                                rd[k]=0
                        rd["Score Performance"]=round(perf_scores.get(p,0),2)
                        prows_s.append(rd)
                    qcols_s=["Poste de travail"]+PK+["Score Qualite"]
                    qrows_s=[]
                    for p in filtered_posts:
                        rd={"Poste de travail":p}
                        for k in PK:
                            if k in ckdf.columns and p in ckdf.index:
                                try:
                                    rd[k]=round(float(ckdf.loc[p,k]),1)
                                except Exception:
                                    rd[k]=0
                            else:
                                rd[k]=0
                        rd["Score Qualite"]=round(qual_scores.get(p,0),2)
                        qrows_s.append(rd)
                    save_kpis_to_excel(prows_s,pcols_s,qrows_s,qcols_s,
                        ano_p,["Poste de travail","Nombre anomalies"] if ano_p else [],
                        ano_q,["Poste de travail","Nombre anomalies"] if ano_q else [],
                        fichier_date)
                    st.success("✅ KPIs sauvegardes dans kpis/indicateurs_kpis.xlsx")

        except Exception as e:
            st.error("❌ Erreur de chargement : %s"%str(e))
            if unf:
                st.info("Chargez les fichiers OT et AVIS depuis la barre laterale.")
            else:
                st.info("Verifiez que ot.xlsx et avis.xlsx sont dans le meme repertoire.")
    else:
        if unf:
            st.markdown("""<div style="min-height:60vh;display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div style="font-size:64px;margin-bottom:20px">📁</div>
            <h2 style="font-size:28px;color:#1e3a5f;font-weight:800">Chargement des donnees</h2>
            <p style="font-size:16px;color:#718096;margin-top:10px">Veuillez charger les fichiers OT et AVIS depuis la barre laterale.</p></div>""",unsafe_allow_html=True)

if __name__ == "__main__":
    main()
