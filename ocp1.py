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
      "OT exécution 1mois< <3mois",
      "Performance Graissage","Performance Inspection","Performance Appels Systématiques"]
PK = ["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé",
      "Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL",
      "OT Fiabilité","Total Avis de Panne"]
ALL_KPI = QK + PK

CIBLE = {"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,
         "OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,
         "OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,
         "OT exécution 1mois< <3mois":15,"appel avis approuvé":95,"OT LANC ESTIME":100,
         "Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,
         "OT CONFIME":100,"OT_COR_EGAL":100,
         "Performance Graissage":95,"Performance Inspection":95,"Performance Appels Systématiques":95,
         "OT Fiabilité":100,"Total Avis de Panne":100}

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
           "OT exécution 1mois< <3mois":"Reduire les OT entre 1 et 3 mois.",
           "Performance Graissage":"Ameliorer le taux de realisation des OT de graissage (Type 350).",
           "Performance Inspection":"Ameliorer le taux de realisation des OT d'inspection (Types 290,300,310).",
           "Performance Appels Systématiques":"Ameliorer le taux de realisation des appels systematiques (Type 360).",
           "OT Fiabilité":"Maintenir la fiabilite des OT a 100%.",
           "Total Avis de Panne":"Maintenir le suivi des avis de panne a 100%."}

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

def compute_cache_key(file_date, filters_dict, now_str):
    key_data = {"file_date": file_date, "filters": filters_dict, "now": now_str}
    return hashlib.md5(json.dumps(key_data, sort_keys=True, default=str).encode()).hexdigest()

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

def load_historical_kpis(filepath):
    if not os.path.exists(filepath): return pd.DataFrame()
    try: wb=load_workbook(filepath,read_only=True,data_only=True)
    except Exception: return pd.DataFrame()
    records=[]; section=None; headers=None
    for sheet_name in wb.sheetnames:
        try:
            ws=wb[sheet_name]; rows_data=list(ws.iter_rows(values_only=True))
            for row in rows_data:
                cell0=str(row[0]).strip() if row[0] else ""
                if "INDICATEURS DE PERFORMANCE" in cell0.upper(): section="perf"; headers=None; continue
                elif "INDICATEURS DE QUALITE" in cell0.upper(): section="qual"; headers=None; continue
                elif "ANOMALIES" in cell0.upper(): section=None; continue
                if section and headers is None and cell0:
                    headers=[str(c).strip() if c else "" for c in row]; continue
                if section and headers and cell0 and cell0 not in ("CIBLE","Total general",""):
                    entry={"Date":sheet_name}
                    for j,h in enumerate(headers):
                        if j<len(row): entry[h]=row[j]
                    entry["_section"]=section; records.append(entry)
        except Exception: continue
    wb.close()
    if not records: return pd.DataFrame()
    df=pd.DataFrame(records)
    df["Date_parsed"]=pd.to_datetime(df["Date"].str.replace("-","/"),format="%d/%m/%Y",errors="coerce")
    return df.sort_values("Date_parsed").reset_index(drop=True)

def calculate_variations(hist_df):
    if hist_df.empty or "Date" not in hist_df.columns: return pd.DataFrame()
    dates=sorted(hist_df["Date"].unique())
    if len(dates)<2: return pd.DataFrame()
    perf_df=hist_df[hist_df["_section"]=="perf"].copy()
    qual_df=hist_df[hist_df["_section"]=="qual"].copy()
    variations=[]
    for i in range(1,len(dates)):
        prev_date,curr_date=dates[i-1],dates[i]
        prev_perf=perf_df[perf_df["Date"]==prev_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        curr_perf=perf_df[perf_df["Date"]==curr_date].set_index("Poste de travail") if "Poste de travail" in perf_df.columns else pd.DataFrame()
        prev_qual=qual_df[qual_df["Date"]==prev_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        curr_qual=qual_df[qual_df["Date"]==curr_date].set_index("Poste de travail") if "Poste de travail" in qual_df.columns else pd.DataFrame()
        for sec_name,prev_d,curr_d,kpi_list in [("Performance",prev_perf,curr_perf,QK+["Score Performance"]),("Qualite",prev_qual,curr_qual,PK+["Score Qualite"])]:
            for poste in set(prev_d.index)&set(curr_d.index):
                for kpi in kpi_list:
                    if kpi not in prev_d.columns or kpi not in curr_d.columns: continue
                    try: pv=float(prev_d.loc[poste,kpi])
                    except Exception: continue
                    try: cv=float(curr_d.loc[poste,kpi])
                    except Exception: continue
                    diff=cv-pv; pct=(diff/pv*100) if pv!=0 else (100 if cv!=0 else 0)
                    if abs(diff)<=0.5: trend="stabilite"
                    elif diff>0.5: trend="hausse"
                    else: trend="baisse"
                    variations.append({"Date precedente":prev_date,"Date actuelle":curr_date,"Poste":poste,
                        "Type":sec_name,"KPI":kpi,"Valeur precedente":round(pv,2),"Valeur actuelle":round(cv,2),
                        "Ecart":round(diff,2),"Ecart %":round(pct,2),"Tendance":trend})
    return pd.DataFrame(variations)

def generate_journal(var_df):
    if var_df.empty: return pd.DataFrame()
    j=var_df.copy(); j["Significatif"]=j["Ecart %"].abs()>=5
    j=j[j["Significatif"]].copy()
    j["Sens"]=j.apply(lambda r:"Amelioration" if ((r["Tendance"]=="hausse" and r["KPI"] not in LOWER_BETTER) or (r["Tendance"]=="baisse" and r["KPI"] in LOWER_BETTER)) else "Degradation",axis=1)
    return j.sort_values(["Date actuelle","Sens","Ecart %"],ascending=[True,False,False])

def calculate_rankings(var_df):
    if var_df.empty: return pd.DataFrame(),pd.DataFrame()
    scores={}
    for poste in var_df["Poste"].unique():
        pv=var_df[var_df["Poste"]==poste].copy()
        scores[poste]=sum((-r["Ecart %"] if r["KPI"] in LOWER_BETTER else r["Ecart %"]) for _,r in pv.iterrows())
    ranked=sorted(scores.items(),key=lambda x:x[1],reverse=True)
    return pd.DataFrame(ranked[:5],columns=["Poste","Score variation"]),pd.DataFrame(ranked[-5:][::-1],columns=["Poste","Score variation"])

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
    .cr{display:grid;gap:6px;margin-bottom:6px}
    .cr2{grid-template-columns:repeat(2,1fr)}
    .cr3{grid-template-columns:repeat(3,1fr)}
    .cr4{grid-template-columns:repeat(4,1fr)}
    .cc{background:#fff;border-radius:var(--r);padding:10px 12px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid var(--b);text-align:center}
    .cc .cv{font-size:26px;font-weight:900;line-height:1}
    .cc .cl{font-size:11px;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
    .cc.c1{border-top:3px solid #3182ce}.cc.c1 .cv{color:#2b6cb0}
    .cc.c2{border-top:3px solid #38a169}.cc.c2 .cv{color:#276749}
    .cc.c3{border-top:3px solid #805ad5}.cc.c3 .cv{color:#6b46c1}
    .cc.c4{border-top:3px solid #e53e3e}.cc.c4 .cv{color:#c53030}
    .sf-section{margin-bottom:8px}
    .sf-title{font-size:14px;font-weight:800;color:#1e3a5f;padding:4px 10px;background:linear-gradient(90deg,#e2e8f0,transparent);border-radius:6px 6px 0 0;border-left:4px solid #2c5282;margin-bottom:4px}
    .sf-title.sf1{border-left-color:#2b6cb0}
    .sf-title.sf2{border-left-color:#d69e2e}
    .stl{font-size:15px;font-weight:700;color:var(--p);margin:6px 0 2px 0;padding-left:10px;border-left:3px solid var(--pl)}
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}.stl.s{border-left-color:#d69e2e}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.3px;padding:5px 6px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw.st thead th{background:linear-gradient(135deg,#975a16,#d69e2e)}
    .tw.omt thead th{background:linear-gradient(135deg,#6b46c1,#805ad5)}
    .tw.tht thead th{background:linear-gradient(135deg,#9b2c2c,#e53e3e)}
    .tw.blg thead th{background:linear-gradient(135deg,#553c9a,#805ad5)}
    .tw tbody td{padding:4px 6px;border-bottom:1px solid #edf2f7;white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .tw tbody td.fc{background:#dbeafe!important;font-weight:700!important;color:#1e3a5f!important;min-width:180px;text-align:left!important;position:sticky;left:0;z-index:5}
    .tw tbody tr:nth-child(even) td.fc{background:#bfdbfe!important;color:#1e3a5f!important}
    .tw tbody tr:hover td.fc{background:#93c5fd!important;color:#1a202c!important}
    .tw.qt tbody td.fc{background:#dbeafe!important;color:#1e40af!important}
    .tw.qt tbody tr:nth-child(even) td.fc{background:#bfdbfe!important}
    .tw.pt tbody td.fc{background:#d1fae5!important;color:#065f46!important}
    .tw.pt tbody tr:nth-child(even) td.fc{background:#a7f3d0!important}
    .tw.at tbody td.fc{background:#fee2e2!important;color:#991b1b!important}
    .tw.at tbody tr:nth-child(even) td.fc{background:#fecaca!important}
    .tw.omt tbody td.fc{background:#ede9fe!important;color:#5b21b6!important}
    .tw.omt tbody tr:nth-child(even) td.fc{background:#ddd6fe!important}
    .tw.tht tbody td.fc{background:#fee2e2!important;color:#991b1b!important}
    .tw.tht tbody tr:nth-child(even) td.fc{background:#fecaca!important}
    .tw.blg tbody td.fc{background:#ede9fe!important;color:#5b21b6!important}
    .tw.blg tbody tr:nth-child(even) td.fc{background:#ddd6fe!important}
    .tw .trow td{background:#c9d4e2!important;font-weight:800!important;font-size:12px!important;border-top:2px solid #1e3a5f!important}
    .tw .trow td.fc{background:#a0aec0!important;color:#1a202c!important}
    .tw .tcol{background:#e8edf3!important;font-weight:700!important}
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
    .dgrid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:8px 14px;font-weight:700;font-size:15px;width:100%}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label,div[data-testid="stSidebar"] .stCheckbox label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:3px 8px;margin-bottom:3px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .es{text-align:center;padding:14px;color:#718096;font-size:14px}
    .synth-tbl{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px}
    .synth-tbl thead th{background:var(--p);color:#fff;font-weight:700;font-size:11px;padding:5px 8px;border:none;white-space:nowrap;position:sticky;top:0}
    .synth-tbl tbody td{padding:4px 8px;border-bottom:1px solid #edf2f7;text-align:center}
    .synth-tbl tbody tr:nth-child(even) td{background:#f7fafc}
    .synth-tbl tbody tr:hover td{background:#ebf8ff!important}
    .synth-tbl .poste-cell{text-align:left;font-weight:700;white-space:nowrap;min-width:140px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:17px}.cg,.dgrid{grid-template-columns:1fr}.car .cal{width:120px}.gbr-l{width:100px}}
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

    # ---- Helpers ----
    def contient_mot(t,lm):
        t=str(t); return any(m in t for l in lm for m in l.split())
    def cat_age(a):
        if pd.isna(a): return "Inconnu"
        if a<=1: return "<1 mois"
        elif a>=3: return ">3 mois"
        return "1 mois < <3 mois"
    def ckpi(n,d,sz=100): return np.where(d==0,sz,(n/d)*100)

    def cpiv(df,f,c,p):
        """Pivot table qui GARDE TOUS les postes de travail même si vide"""
        df_filtre = df[f]
        if df_filtre.empty:
            return pd.DataFrame(0, index=p, columns=["Inconnu"]).fillna(0)
        piv = pd.pivot_table(df_filtre, index="Poste travail princ.", columns=c, values="Ordre", aggfunc="count", fill_value=0)
        return piv.reindex(p, fill_value=0).fillna(0)

    def excr(df):
        if "Poste travail princ." in df.columns:
            return df[~df["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)].copy()
        return df

    def build_statut_pivot(df_sub, posts):
        if df_sub.empty:
            return pd.DataFrame(0, index=posts, columns=["CRÉÉ","LANC","CLOT","TCLO","Total"]).fillna(0).astype(int)
        piv=pd.pivot_table(df_sub, index="Poste travail princ.", columns="Statut OT", values="Ordre", aggfunc="count", fill_value=0)
        for s in ["CRÉÉ","LANC","CLOT","TCLO"]:
            if s not in piv.columns: piv[s]=0
        piv["Total"]=piv[["CRÉÉ","LANC","CLOT","TCLO"]].sum(axis=1)
        return piv.reindex(posts, fill_value=0).fillna(0).astype(int)

    def html_statut_pivot(piv_df, table_class):
        cols=["Poste de travail","CRÉÉ","LANC","CLOT","TCLO","Total"]
        h='<table class="tw %s"><thead><tr>'%table_class+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for poste,row in piv_df.iterrows():
            h+='<tr><td class="fc">%s</td>'%poste
            for c in ["CRÉÉ","LANC","CLOT","TCLO"]:
                h+='<td style="text-align:center">%d</td>'%int(row.get(c,0))
            h+='<td class="tcol" style="text-align:center;font-weight:800">%d</td>'%int(row.get("Total",0))
            h+='</tr>'
        h+='<tr class="trow"><td class="fc">Total</td>'
        for c in ["CRÉÉ","LANC","CLOT","TCLO"]:
            h+='<td class="tcol" style="text-align:center">%d</td>'%int(piv_df[c].sum())
        h+='<td class="tcol" style="text-align:center">%d</td>'%int(piv_df["Total"].sum())
        h+='</tr></tbody></table>'
        return h

    def show_pie_pair(piv_df, title_prefix):
        global_counts=piv_df[["CRÉÉ","LANC","CLOT","TCLO"]].sum()
        global_counts=global_counts[global_counts>0]
        realised=global_counts.get("CLOT",0)+global_counts.get("TCLO",0)
        not_realised=global_counts.sum()-realised
        c1,c2=st.columns(2)
        with c1:
            if not global_counts.empty:
                fig1=px.pie(global_counts, names=global_counts.index, values=global_counts.values,
                    title="%s — Par Statut OT"%title_prefix,
                    color_discrete_sequence=["#e53e3e","#d69e2e","#38a169","#3182ce"])
                fig1.update_traces(textposition='inside',textinfo='percent+value',textfont_size=11)
                fig1.update_layout(margin=dict(t=50,b=10,l=10,r=10),height=340,legend=dict(font_size=10,orientation="h",yanchor="bottom",y=-0.1))
                st.plotly_chart(fig1,use_container_width=True)
            else: st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)
        with c2:
            if global_counts.sum()>0:
                pie2_data=pd.DataFrame({"Statut":["Réalisés (CLOT+TCLO)","Non Réalisés"],"Nombre":[realised,not_realised]})
                fig2=px.pie(pie2_data, names="Statut", values="Nombre",
                    title="%s — Réalisés vs Non Réalisés"%title_prefix,
                    color="Statut", color_discrete_map={"Réalisés (CLOT+TCLO)":"#38a169","Non Réalisés":"#e53e3e"})
                fig2.update_traces(textposition='inside',textinfo='percent+value',textfont_size=11)
                fig2.update_layout(margin=dict(t=50,b=10,l=10,r=10),height=340,legend=dict(font_size=10,orientation="h",yanchor="bottom",y=-0.1))
                st.plotly_chart(fig2,use_container_width=True)
            else: st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)

    def calc_kpis(df_i, av_i, now, posts):
        res={}; df=df_i.copy(); av=av_i.copy()
        df["Backlog preparation"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MP_KW)),"CARACTERISE","NON CARACTERISE")
        df["Backlog planification"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,MPLAN_KW)),"CARACTERISE","NON CARACTERISE")
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
        df["_tw_num"]=pd.to_numeric(df.get("Type de travail",pd.Series(dtype=float)),errors="coerce")
        res['dfp']=df

        # Taux realisation correctif
        filt_corr=(df["Nº appel pl.entret."].fillna(0)==0)&(df["Contient SOPL"]==1)
        an=cpiv(df,filt_corr,"Statut OT",posts)
        for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
        an["OT_CLOTURES"]=an["CLOT"]+an["TCLO"]
        an["TOTAL_OT"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
        an["TAUX_REALISATION_CORRECTIF/PT"]=np.where(an["TOTAL_OT"]==0,100.0,ckpi(an["OT_CLOTURES"],an["TOTAL_OT"]))

        # Age preparation
        pr=cpiv(df,(df["Statut OT"]=="CRÉÉ")&(df["Statut utilisateur"].str.contains("CRPR",na=False)),"ap",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pr[c]=pr.get(c,0)
        pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"])
        pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0)
        pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)

        # Age planification
        pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Statut utilisateur"].str.contains("ATPL",case=False,na=False)),"alp",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: pl[c]=pl.get(c,0)
        pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"])
        pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0)
        pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)

        # Age execution
        ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
        for c in ["<1 mois",">3 mois","1 mois < <3 mois","Inconnu"]: ex[c]=ex.get(c,0)
        ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois","Inconnu"]].sum(axis=1)
        ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"])
        ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0)
        ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)

        # OT LANCE ESTIME
        la=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["OUI","NON"]: la[c]=la.get(c,0)
        la["Total"]=la["OUI"]+la["NON"]; la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"])

        # Backlog preparation caracterise
        pc=pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
        pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"])

        # Backlog planification caracterise
        plc=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
        plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])

        # OT CONFIME et OT_COR_EGAL
        for kn,cn in [("OT CONFIME","OT CONFIME"),("OT_COR_EGAL","OT_COR_EGAL")]:
            pv=pd.pivot_table(df,index="Poste travail princ.",columns=cn,values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
            for c in ["OUI","NON"]: pv[c]=pv.get(c,0)
            pv["Total"]=pv["OUI"]+pv["NON"]; pv[cn]=ckpi(pv["OUI"],pv["Total"]); res[kn.lower().replace(" ","_")]=pv

        # Appel avis approuve
        avf=av[(av["Ordre"].isna())|(av["Ordre"].astype(str).str.strip()=="")].copy(); res['avf']=avf
        tca=pd.pivot_table(avf,index="Poste travail princ.",columns="Statut utilisateur",values="Avis",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
        for c in ["APRQ","APRV","APRV AVAU","REJT"]: tca[c]=tca.get(c,0)
        tca["Total"]=tca[["APRQ","APRV","APRV AVAU","REJT"]].sum(axis=1); tca["appel avis approuvé"]=ckpi(tca["APRV"],tca["Total"])

        # Performance Graissage
        g_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&(df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
        g_den=df[(df["Contient SOPL"]==1)&(df["_tw_num"]==350)].groupby("Poste travail princ.")["Ordre"].count()
        g_df=pd.DataFrame({"_n":g_num,"_d":g_den}).reindex(posts,fill_value=0)
        g_df["Performance Graissage"]=np.where(g_df["_d"]==0,100.0,(g_df["_n"]/g_df["_d"])*100)

        # Performance Inspection
        ins_types=[290,300,310]
        ins_base=(df["_tw_num"].isin(ins_types))&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now)
        ins_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_den=df[(df["Contient SOPL"]==1)&ins_base].groupby("Poste travail princ.")["Ordre"].count()
        ins_df=pd.DataFrame({"_n":ins_num,"_d":ins_den}).reindex(posts,fill_value=0)
        ins_df["Performance Inspection"]=np.where(ins_df["_d"]==0,100.0,(ins_df["_n"]/ins_df["_d"])*100)

        # Performance Appels Systematiques
        sys_base=(df["_tw_num"]==360)&(df["Date de début planifiée"].notna())&(df["Date de début planifiée"]<=now)
        sys_num=df[(df["Statut OT"].isin(["CLOT","TCLO"]))&sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_den=df[(df["Contient SOPL"]==1)&sys_base].groupby("Poste travail princ.")["Ordre"].count()
        sys_df=pd.DataFrame({"_n":sys_num,"_d":sys_den}).reindex(posts,fill_value=0)
        sys_df["Performance Appels Systématiques"]=np.where(sys_df["_d"]==0,100.0,(sys_df["_n"]/sys_df["_d"])*100)

        # Fiabilite et Avis Panne toujours 100
        fiab_s=pd.Series(100.0,index=posts); avpan_s=pd.Series(100.0,index=posts)

        # Stocker les pivots de backlog pour la page Backlog
        res['backlog_prep_pivot'] = pc.copy()
        res['backlog_plan_pivot'] = plc.copy()
        res['backlog_exec_pivot'] = ex.copy()

        res['ckdf']=pd.DataFrame({
            "TAUX_REALISATION_CORRECTIF/PT":an["TAUX_REALISATION_CORRECTIF/PT"],
            "OT préparation <1 mois":pr["OT préparation <1 mois"],"OT préparation >3 mois":pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],
            "OT planification <1 mois":pl["OT planification <1 mois"],"OT planification >3 mois":pl["OT planification >3 mois"],"OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],
            "OT exécution <1 mois":ex["OT exécution <1 mois"],"OT exécution >3 mois":ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],
            "Performance Graissage":g_df["Performance Graissage"],"Performance Inspection":ins_df["Performance Inspection"],"Performance Appels Systématiques":sys_df["Performance Appels Systématiques"],
            "appel avis approuvé":tca["appel avis approuvé"],"OT LANC ESTIME":la["OT LANC ESTIME"],
            "Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],"Backlog planification caractérisé":plc["Backlog planification caractérisé"],
            "OT CONFIME":res['ot_confime']["OT CONFIME"],"OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"],
            "OT Fiabilité":fiab_s,"Total Avis de Panne":avpan_s
        })
        return res

    # ---- Cellule style ----
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
        if c in ["Performance Graissage","Performance Inspection","Performance Appels Systématiques"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val>=95 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>90 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT Fiabilité","Total Avis de Panne"]:
            return "background:#c6efce;color:#006100;font-weight:600" if val>=100 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=95 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        return ""

    def cs(v):
        try: val=float(str(v).replace(' %','').strip())
        except Exception: return ""
        return "background:#c6efce;color:#006100;font-weight:700" if val>=90 else ("background:#ffeb9c;color:#9c6500;font-weight:700" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:700")

    def kas(v):
        try: val=int(v)
        except Exception: return ""
        if val==0: return "background:#c6efce;color:#006100;font-weight:600"
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
        if k in ["Performance Graissage","Performance Inspection","Performance Appels Systématiques"]: return 1 if a>=95 else 0
        if k in ["OT Fiabilité","Total Avis de Panne"]: return 1 if a>=100 else 0
        return 0

    def is_lb(k): return k in LOWER_BETTER

    def is_anomaly(k, val):
        """Verifie si une valeur est une anomalie par rapport a la cible"""
        try: v=float(val)
        except Exception: return False
        t=CIBLE.get(k,100)
        if k in LOWER_BETTER: return v > t
        else: return v < t

    # ---- HTML builders ----
    def html_table(rows, cols, tc, sc_col=None):
        h='<table class="tw %s"><thead><tr>'%tc+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            rc="cb" if r.get("_t")=="cible" else ("trow" if r.get("_t")=="total" else "")
            h+='<tr class="%s">'%rc
            for j,c in enumerate(cols):
                v=r.get(c,"")
                if j==0 and r.get("_t") not in ("cible","total"):
                    s=ks(v,c) if sc_col and c in sc_col else ""
                    h+='<td class="fc" style="%s">%s</td>'%(s or "",v)
                elif r.get("_t")=="cible":
                    h+='<td>%s</td>'%v
                else:
                    is_last=(j==len(cols)-1)
                    s=ks(v,c) if sc_col and c in sc_col else ""
                    tcol_cls=" class=\"tcol\"" if is_last else ""
                    h+='<td%s style="%s">%s</td>'%(tcol_cls, s or "", v)
            h+='</tr>'
        return h+'</tbody></table>'

    def html_ano_table(ano_counts, kpi_list, posts, table_class, title, color_scheme="p"):
        """Tableau anomalies : compte le nombre d'anomalies par KPI par poste"""
        tc_map = {"p":"pt","q":"qt","a":"at"}
        tc = tc_map.get(color_scheme, "pt")
        h='<div class="ca"><div class="ct">%s</div>'%title
        h+='<table class="tw %s"><thead><tr><th style="min-width:280px">KPI</th>'%tc
        for p in posts: h+='<th>%s</th>'%p
        h+='<th>Total</th></tr></thead><tbody>'
        grand_total = 0
        for kpi in kpi_list:
            h+='<tr><td class="fc" style="white-space:normal;max-width:280px">%s</td>'%kpi
            total=0
            for p in posts:
                # Fiabilite et Avis Panne toujours 0
                if kpi in ["OT Fiabilité","Total Avis de Panne"]:
                    cnt=0
                else:
                    cnt=ano_counts.get(kpi,{}).get(p,0)
                total+=cnt
                s=kas(cnt)
                h+='<td style="%s;text-align:center">%d</td>'%(s or "",cnt)
            grand_total+=total
            s=kas(total)
            h+='<td class="tcol" style="%s;text-align:center;font-weight:800">%d</td>'%(s or "",total)
            h+='</tr>'
        # Ligne total
        h+='<tr class="trow"><td class="fc">Total Anomalies</td>'
        for p in posts:
            ptot=sum(ano_counts.get(k,{}).get(p,0) for k in kpi_list if k not in ["OT Fiabilité","Total Avis de Panne"])
            h+='<td class="tcol" style="text-align:center">%d</td>'%ptot
        h+='<td class="tcol" style="text-align:center">%d</td>'%grand_total
        h+='</tr></tbody></table></div>'
        return h

    def html_actions_table(kpi_list,actuals,targets,act_map):
        h='<table class="tw at"><thead><tr><th>KPI</th><th>Valeur Actuelle</th><th>Cible</th><th>Ecart</th><th>Statut</th><th>Action Recommandee</th></tr></thead><tbody>'
        for k in kpi_list:
            av=actuals.get(k,0); tv=targets.get(k,100); diff=av-tv
            met=av<=tv if is_lb(k) else av>=tv
            status="ATTEINT" if met else "NON ATTEINT"
            st_s="background:#c6efce;color:#006100;font-weight:700" if met else "background:#ffc7ce;color:#9c0006;font-weight:700"
            ec_clr="#276749" if met else "#c53030"
            action="Objectif atteint" if met else act_map.get(k,"")
            h+='<tr><td class="fc" style="%s">%s</td><td>%.1f%%</td><td>%.0f%%</td><td style="color:%s;font-weight:700">%+.1f%%</td><td style="%s">%s</td><td style="color:#4a5568">%s</td></tr>'%("",k,av,tv,ec_clr,diff,st_s,status,action)
        return h+'</tbody></table>'

    def html_classement(scores,accent):
        sp=sorted(scores.items(),key=lambda x:x[1],reverse=True)
        met_p=[(p,s) for p,s in sp if s>=80]; not_p=[(p,s) for p,s in sp if s<80]
        t5=met_p[:5]; b5=not_p[-5:] if len(not_p)>5 else not_p
        h='<div class="cg"><div><div class="ct" style="color:#38a169">Top 5 — Objectif Atteint</div>'
        if t5:
            for i,(p,s) in enumerate(t5): h+='<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(accent,i+1,p,cs("%.2f"%s),s)
        else: h+='<div style="padding:6px;font-size:12px;color:#718096">Aucun poste</div>'
        h+='</div><div><div class="ct" style="color:#e53e3e">Bottom 5 — Non Atteint</div>'
        if b5:
            for i,(p,s) in enumerate(reversed(b5)): h+='<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(len(b5)-i,p,cs("%.2f"%s),s)
        else: h+='<div style="padding:6px;font-size:12px;color:#38a169">Tous atteints</div>'
        h+='</div></div>'; return h

    def html_kpi_bars(kpi_list,actuals,targets,title,color_ok,color_fail):
        h='<div class="ca"><div class="ct" style="color:%s">%s</div>'%(color_ok,title)
        for k in kpi_list:
            av=actuals.get(k,0); tv=targets.get(k,100); met=av<=tv if is_lb(k) else av>=tv
            bw=min(max(av,0),100); bg=color_ok if met else color_fail
            h+='<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>'%(k,bw,bg,av)
        return h+'</div>'

    def html_bars(data,title,color):
        h='<div class="ca"><div class="ct" style="color:%s">%s</div>'%(color,title)
        for label,val in sorted(data,key=lambda x:x[1],reverse=True):
            bw=min(max(val,0),100)
            h+='<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>'%(label,bw,color,val)
        return h+'</div>'

    def html_grouped_bars(posts,pscores,qscores,title):
        h='<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>'%title
        h+='<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span></div>'
        for p in sorted(posts,key=lambda x:(pscores.get(x,0)+qscores.get(x,0))/2,reverse=True):
            pv,qv=pscores.get(p,0),qscores.get(p,0)
            h+='<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f gb-p" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w"><div class="gbr-f gb-q" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div></div></div>'%(p,min(max(pv,0),100),pv,min(max(qv,0),100),qv)
        return h+'</div>'

    def html_backlog_char_pivot(piv_df, char_col_car, char_col_non, table_class, title):
        """Tableau caractérisation backlog avec Total coloré"""
        h='<div class="ca"><div class="ct">%s</div>'%title
        h+='<table class="tw %s"><thead><tr><th>Poste de travail</th><th>%s</th><th>%s</th><th>Total</th></tr></thead><tbody>'%(table_class, char_col_car, char_col_non)
        for poste,row in piv_df.iterrows():
            car_val=int(row.get(char_col_car,0))
            non_val=int(row.get(char_col_non,0))
            tot=car_val+non_val
            h+='<tr><td class="fc">%s</td><td style="text-align:center">%d</td><td style="text-align:center">%d</td><td class="tcol" style="text-align:center;font-weight:800">%d</td></tr>'%(poste,car_val,non_val,tot)
        # Total row
        tot_car=int(piv_df[char_col_car].sum())
        tot_non=int(piv_df[char_col_non].sum())
        h+='<tr class="trow"><td class="fc">Total</td><td class="tcol" style="text-align:center">%d</td><td class="tcol" style="text-align:center">%d</td><td class="tcol" style="text-align:center">%d</td></tr>'%(tot_car,tot_non,tot_car+tot_non)
        h+='</tbody></table></div>'
        return h

    def show_backlog_char_charts(piv_df, char_col_car, char_col_non, title):
        """Pie charts pour caractérisation backlog"""
        tot_car=int(piv_df[char_col_car].sum())
        tot_non=int(piv_df[char_col_non].sum())
        c1,c2=st.columns(2)
        with c1:
            if tot_car+tot_non>0:
                pie_data=pd.DataFrame({"Statut":[char_col_car,char_col_non],"Nombre":[tot_car,tot_non]})
                fig=px.pie(pie_data,names="Statut",values="Nombre",title="%s — Global"%title,
                    color="Statut",color_discrete_map={char_col_car:"#38a169",char_col_non:"#e53e3e"})
                fig.update_traces(textposition='inside',textinfo='percent+value',textfont_size=11)
                fig.update_layout(margin=dict(t=50,b=10,l=10,r=10),height=340,legend=dict(font_size=10,orientation="h",yanchor="bottom",y=-0.1))
                st.plotly_chart(fig,use_container_width=True)
            else: st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)
        with c2:
            if tot_car+tot_non>0:
                # Bar chart par poste
                bar_df=piv_df[[char_col_car,char_col_non]].copy()
                bar_df.columns=["CARACTERISE","NON CARACTERISE"]
                bar_df=bar_df.sort_values("NON CARACTERISE",ascending=True)
                fig2=go.Figure()
                fig2.add_trace(go.Bar(y=bar_df.index,x=bar_df["CARACTERISE"],orientation='h',name='CARACTERISE',marker_color='#38a169'))
                fig2.add_trace(go.Bar(y=bar_df.index,x=bar_df["NON CARACTERISE"],orientation='h',name='NON CARACTERISE',marker_color='#e53e3e'))
                fig2.update_layout(barmode='stack',title="%s — Par Poste"%title,height=340,margin=dict(t=50,b=10,l=10,r=10),legend=dict(font_size=10,orientation="h",yanchor="bottom",y=-0.15),yaxis=dict(tickfont=dict(size=9)))
                st.plotly_chart(fig2,use_container_width=True)
            else: st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)

    # ============================================================
    # SIDEBAR
    # ============================================================
    with st.sidebar:
        st.markdown("### 📂 Fichier de donnees")
        uploaded=st.file_uploader("Importer le fichier Excel",type=["xlsx","xls"],key="file_up")
        if uploaded is not None:
            try:
                xl=pd.ExcelFile(uploaded)
                sheet_names=xl.sheet_names
                ot_sheet=st.selectbox("Feuille OT",sheet_names,index=0 if len(sheet_names)>0 else 0,key="ot_sh")
                av_sheet=st.selectbox("Feuille Avis",sheet_names,index=1 if len(sheet_names)>1 else 0,key="av_sh")
                df_raw=pd.read_excel(uploaded,sheet_name=ot_sheet)
                av_raw=pd.read_excel(uploaded,sheet_name=av_sheet)
                st.success("Fichier charge avec succes")
            except Exception as e:
                st.error("Erreur de lecture: %s"%e); st.stop()
        else:
            st.warning("Veuillez importer un fichier Excel"); st.stop()

        st.markdown("### 🔍 Filtres")
        if "Poste travail princ." in df_raw.columns:
            all_posts=sorted(df_raw["Poste travail princ."].dropna().unique().tolist())
            sel_posts=st.multiselect("Postes de travail",all_posts,default=all_posts,key="posts_sel")
        else:
            st.error("Colonne 'Poste travail princ.' introuvable"); st.stop()

        ref_date=st.date_input("Date de reference",datetime.now(),key="ref_d")
        now=pd.Timestamp(ref_date)

        excl_cress=st.checkbox("Exclure cresseurs",value=True,key="excl_cr")
        save_xl=st.button("💾 Sauvegarder KPIs en Excel",key="save_btn")

    # ============================================================
    # PREPARATION DES DONNEES
    # ============================================================
    df=df_raw.copy(); av=av_raw.copy()
    if excl_cress: df=excr(df)
    posts=sel_posts if sel_posts else []
    if not posts: st.warning("Aucun poste selectionne"); st.stop()

    res=calc_kpis(df,av,now,posts)
    ckdf=res['ckdf']

    # Separer SF1 et SF2
    sf1_posts=sorted([p for p in posts if str(p).upper().startswith("SF1")])
    sf2_posts=sorted([p for p in posts if str(p).upper().startswith("SF2")])
    other_posts=sorted([p for p in posts if not str(p).upper().startswith("SF1") and not str(p).upper().startswith("SF2")])

    # Calculs scores
    pscores={}; qscores={}
    for p in posts:
        psc=0; qsc=0
        for k in QK:
            a=ckdf.loc[p,k] if p in ckdf.index else 0
            t=CIBLE.get(k,100)
            psc+=gscore(k,a,t)
        for k in PK:
            a=ckdf.loc[p,k] if p in ckdf.index else 0
            t=CIBLE.get(k,100)
            qsc+=gscore(k,a,t)
        pscores[p]=(psc/len(QK)*100) if QK else 0
        qscores[p]=(qsc/len(PK)*100) if PK else 0

    # Anomalies comptage
    ano_perf_counts={}
    ano_qual_counts={}
    for k in QK:
        ano_perf_counts[k]={}
        for p in posts:
            val=ckdf.loc[p,k] if p in ckdf.index else 0
            ano_perf_counts[k][p]=1 if is_anomaly(k,val) else 0
    for k in PK:
        ano_qual_counts[k]={}
        for p in posts:
            if k in ["OT Fiabilité","Total Avis de Panne"]:
                ano_qual_counts[k][p]=0
            else:
                val=ckdf.loc[p,k] if p in ckdf.index else 0
                ano_qual_counts[k][p]=1 if is_anomaly(k,val) else 0

    # Sauvegarde Excel
    if save_xl:
        pcols=["Poste de travail"]+QK+["Score Performance"]
        prows=[]; qcols=["Poste de travail"]+PK+["Score Qualite"]; qrows=[]
        for p in posts:
            pr={"Poste de travail":p,"Score Performance":round(pscores.get(p,0),2)}
            qr={"Poste de travail":p,"Score Qualite":round(qscores.get(p,0),2)}
            for k in QK: pr[k]=round(ckdf.loc[p,k],2) if p in ckdf.index else 0
            for k in PK: qr[k]=round(ckdf.loc[p,k],2) if p in ckdf.index else 0
            prows.append(pr); qrows.append(qr)
        cr={"Poste de travail":"CIBLE"}
        for k in QK: cr[k]=CIBLE.get(k,"")
        cr["Score Performance"]=""
        prows.append({"_t":"cible",**cr})
        qr2={"Poste de travail":"CIBLE"}
        for k in PK: qr2[k]=CIBLE.get(k,"")
        qr2["Score Qualite"]=""
        qrows.append({"_t":"cible",**qr2})
        # Total row
        ptr={"Poste de travail":"Total general","_t":"total","Score Performance":round(np.mean([pscores.get(p,0) for p in posts]),2)}
        qtr={"Poste de travail":"Total general","_t":"total","Score Qualite":round(np.mean([qscores.get(p,0) for p in posts]),2)}
        for k in QK: ptr[k]=round(ckdf[k].mean(),2)
        for k in PK: qtr[k]=round(ckdf[k].mean(),2)
        prows.append(ptr); qrows.append(qtr)
        # Anomalies rows
        ano_p_r=[{"Poste de travail":p,**{k:ano_perf_counts[k].get(p,0) for k in QK}} for p in posts]
        ano_q_r=[{"Poste de travail":p,**{k:ano_qual_counts[k].get(p,0) for k in PK}} for p in posts]
        ano_p_c=["Poste de travail"]+QK; ano_q_c=["Poste de travail"]+PK
        save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_r,ano_p_c,ano_q_r,ano_q_c,fichier_date)
        st.sidebar.success("KPIs sauvegardes!")

    # ============================================================
    # HEADER
    # ============================================================
    st.markdown('<div class="mh"><h1>📊 Dashboard KPI Maintenance</h1><span class="db">📅 %s</span></div>'%fichier_date,unsafe_allow_html=True)

    # ============================================================
    # ICONES HAUT : SF1 et SF2 separes avec scores Performance et Qualite
    # ============================================================
    def render_sf_cards(sf_posts_list, sf_label, sf_class):
        if not sf_posts_list: return ""
        h='<div class="sf-section"><div class="sf-title %s">🏭 %s (%d postes)</div><div class="cr cr4">'%(sf_class, sf_label, len(sf_posts_list))
        for p in sf_posts_list:
            pv=pscores.get(p,0); qv=qscores.get(p,0)
            # Choisir couleur selon le plus bas des deux scores
            mn=min(pv,qv)
            if mn>=90: cc="c2"  # vert
            elif mn>=80: cc="c1"  # bleu
            elif mn>=70: cc="c3"  # violet
            else: cc="c4"  # rouge
            h+='<div class="cc %s"><div class="cv">%.0f%% / %.0f%%</div><div class="cl">%s<br><span style="font-size:9px;color:#a0aec0">P: %.0f%% | Q: %.0f%%</span></div></div>'%(cc,pv,qv,p,pv,qv)
        h+='</div></div>'
        return h

    cards_html = ""
    if sf1_posts: cards_html += render_sf_cards(sf1_posts, "SF1 — Performance / Qualité", "sf1")
    if sf2_posts: cards_html += render_sf_cards(sf2_posts, "SF2 — Performance / Qualité", "sf2")
    if other_posts: cards_html += render_sf_cards(other_posts, "Autres Postes — Performance / Qualité", "sf1")

    # Cartes resume globaux
    global_p=np.mean([pscores.get(p,0) for p in posts])
    global_q=np.mean([qscores.get(p,0) for p in posts])
    total_ot=len(df[df["Poste travail princ."].isin(posts)])
    total_anomalies=sum(ano_perf_counts[k].get(p,0) for k in QK for p in posts)+sum(ano_qual_counts[k].get(p,0) for k in PK if k not in ["OT Fiabilité","Total Avis de Panne"] for p in posts)

    cards_html+='<div class="cr cr4">'
    cards_html+='<div class="cc c1"><div class="cv">%.1f%%</div><div class="cl">Score Performance Global</div></div>'%global_p
    cards_html+='<div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Qualite Global</div></div>'%global_q
    cards_html+='<div class="cc c3"><div class="cv">%d</div><div class="cl">Total OT</div></div>'%total_ot
    cards_html+='<div class="cc c4"><div class="cv">%d</div><div class="cl">Total Anomalies</div></div>'%total_anomalies
    cards_html+='</div>'

    st.markdown(cards_html, unsafe_allow_html=True)

    # ============================================================
    # ONGLETS
    # ============================================================
    tab_syn, tab_perf, tab_qual, tab_ano, tab_backlog = st.tabs([
        "📋 Synthese","⚡ Performance","✅ Qualite","⚠️ Anomalies","📦 Backlog"
    ])

    # ---- SYNTHESE ----
    with tab_syn:
        st.markdown('<div class="stl p">Classement des Postes</div>',unsafe_allow_html=True)
        st.markdown(html_classement(pscores,"#2b6cb0"),unsafe_allow_html=True)
        st.markdown(html_grouped_bars(posts,pscores,qscores,"Scores Performance & Qualite par Poste"),unsafe_allow_html=True)
        st.markdown('<div class="stl p">Indicateurs de Performance — Barres</div>',unsafe_allow_html=True)
        perf_avgs={k:round(ckdf[k].mean(),2) for k in QK}
        st.markdown(html_kpi_bars(QK,perf_avgs,CIBLE,"Performance Globale","#38a169","#e53e3e"),unsafe_allow_html=True)
        st.markdown('<div class="stl q">Indicateurs de Qualite — Barres</div>',unsafe_allow_html=True)
        qual_avgs={k:round(ckdf[k].mean(),2) for k in PK}
        st.markdown(html_kpi_bars(PK,qual_avgs,CIBLE,"Qualite Globale","#3182ce","#e53e3e"),unsafe_allow_html=True)

    # ---- PERFORMANCE ----
    with tab_perf:
        st.markdown('<div class="stl p">Tableau de Performance par Poste</div>',unsafe_allow_html=True)
        pcols_t=["Poste de travail"]+QK+["Score Performance"]
        prows_t=[]
        for p in posts:
            r={"Poste de travail":p}
            for k in QK: r[k]="%.1f"%ckdf.loc[p,k] if p in ckdf.index else "0.0"
            r["Score Performance"]="%.2f"%pscores.get(p,0)
            prows_t.append(r)
        # Cible row
        cr={"_t":"cible","Poste de travail":"CIBLE"}
        for k in QK: cr[k]="%.0f"%CIBLE.get(k,0)
        cr["Score Performance"]="100"
        prows_t.append(cr)
        # Total row
        tr={"_t":"total","Poste de travail":"Total general"}
        for k in QK: tr[k]="%.1f"%ckdf[k].mean()
        tr["Score Performance"]="%.2f"%global_p
        prows_t.append(tr)
        st.markdown(html_table(prows_t,pcols_t,"pt",sc_col=set(QK+["Score Performance"])),unsafe_allow_html=True)
        st.markdown(html_kpi_bars(QK,perf_avgs,CIBLE,"Détail Performance","#38a169","#e53e3e"),unsafe_allow_html=True)
        st.markdown(html_actions_table(QK,perf_avgs,CIBLE,ACT_MAP),unsafe_allow_html=True)

    # ---- QUALITE ----
    with tab_qual:
        st.markdown('<div class="stl q">Tableau de Qualite par Poste</div>',unsafe_allow_html=True)
        qcols_t=["Poste de travail"]+PK+["Score Qualite"]
        qrows_t=[]
        for p in posts:
            r={"Poste de travail":p}
            for k in PK: r[k]="%.1f"%ckdf.loc[p,k] if p in ckdf.index else "0.0"
            r["Score Qualite"]="%.2f"%qscores.get(p,0)
            qrows_t.append(r)
        cr2={"_t":"cible","Poste de travail":"CIBLE"}
        for k in PK: cr2[k]="%.0f"%CIBLE.get(k,0)
        cr2["Score Qualite"]="100"
        qrows_t.append(cr2)
        tr2={"_t":"total","Poste de travail":"Total general"}
        for k in PK: tr2[k]="%.1f"%ckdf[k].mean()
        tr2["Score Qualite"]="%.2f"%global_q
        qrows_t.append(tr2)
        st.markdown(html_table(qrows_t,qcols_t,"qt",sc_col=set(PK+["Score Qualite"])),unsafe_allow_html=True)
        st.markdown(html_kpi_bars(PK,qual_avgs,CIBLE,"Détail Qualite","#3182ce","#e53e3e"),unsafe_allow_html=True)
        st.markdown(html_actions_table(PK,qual_avgs,CIBLE,ACT_MAP),unsafe_allow_html=True)

    # ---- ANOMALIES ----
    with tab_ano:
        st.markdown('<div class="stl a">Anomalies Performance</div>',unsafe_allow_html=True)
        st.markdown(html_ano_table(ano_perf_counts, QK, posts, "pt", "Anomalies Indicateurs de Performance", "p"),unsafe_allow_html=True)
        st.markdown('<div class="stl a">Anomalies Qualite</div>',unsafe_allow_html=True)
        st.markdown(html_ano_table(ano_qual_counts, PK, posts, "qt", "Anomalies Indicateurs de Qualite (Fiabilité & Avis Panne = 0)", "q"),unsafe_allow_html=True)

    # ---- BACKLOG ----
    with tab_backlog:
        df_backlog=res['dfp']
        df_filtered=df_backlog[df_backlog["Poste travail princ."].isin(posts)]

        # -- OT OMS --
        st.markdown('<div class="stl c">OT OMS par Poste et Statut OT</div>',unsafe_allow_html=True)
        oms_df=df_filtered[df_filtered["Type de travail"].astype(str).str.contains("OMS",case=False,na=False)]
        oms_piv=build_statut_pivot(oms_df,posts)
        st.markdown(html_statut_pivot(oms_piv,"omt"),unsafe_allow_html=True)
        show_pie_pair(oms_piv,"OT OMS")

        # -- OT Thermographie --
        st.markdown('<div class="stl c" style="margin-top:12px">OT Thermographie par Poste et Statut OT</div>',unsafe_allow_html=True)
        thermo_df=df_filtered[df_filtered["Type de travail"].astype(str).str.contains("hermo",case=False,na=False)]
        thermo_piv=build_statut_pivot(thermo_df,posts)
        st.markdown(html_statut_pivot(thermo_piv,"tht"),unsafe_allow_html=True)
        show_pie_pair(thermo_piv,"OT Thermographie")

        # -- Tous les OT --
        st.markdown('<div class="stl c" style="margin-top:12px">Tous les OT par Poste et Statut OT</div>',unsafe_allow_html=True)
        all_piv=build_statut_pivot(df_filtered,posts)
        st.markdown(html_statut_pivot(all_piv,"blg"),unsafe_allow_html=True)
        show_pie_pair(all_piv,"Tous les OT")

        # -- Caracterisation Backlog Preparation --
        st.markdown('<div class="stl c" style="margin-top:12px">Caracterisation Backlog Preparation (Statut=CRÉÉ)</div>',unsafe_allow_html=True)
        bp=res.get('backlog_prep_pivot')
        if bp is not None:
            st.markdown(html_backlog_char_pivot(bp,"CARACTERISE","NON CARACTERISE","blg","Backlog Preparation — Caracterisation"),unsafe_allow_html=True)
            show_backlog_char_charts(bp,"CARACTERISE","NON CARACTERISE","Backlog Preparation")
        else:
            st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)

        # -- Caracterisation Backlog Planification --
        st.markdown('<div class="stl c" style="margin-top:12px">Caracterisation Backlog Planification (Statut=LANC)</div>',unsafe_allow_html=True)
        bpl=res.get('backlog_plan_pivot')
        if bpl is not None:
            st.markdown(html_backlog_char_pivot(bpl,"CARACTERISE","NON CARACTERISE","blg","Backlog Planification — Caracterisation"),unsafe_allow_html=True)
            show_backlog_char_charts(bpl,"CARACTERISE","NON CARACTERISE","Backlog Planification")
        else:
            st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)

        # -- Caracterisation Backlog Execution (par age) --
        st.markdown('<div class="stl c" style="margin-top:12px">Caracterisation Backlog Execution (Age OT SOPL)</div>',unsafe_allow_html=True)
        bex=res.get('backlog_exec_pivot')
        if bex is not None:
            # Regrouper : <=1 mois = OK, >1 mois = ANOMALIE
            bex2=bex.copy()
            bex2["OK"]=bex2.get("<1 mois",0)
            bex2["ANOMALIE"]=bex2.get("1 mois < <3 mois",0)+bex2.get(">3 mois",0)+bex2.get("Inconnu",0)
            st.markdown(html_backlog_char_pivot(bex2,"OK","ANOMALIE","blg","Backlog Execution — Age (OK si <1 mois)"),unsafe_allow_html=True)
            show_backlog_char_charts(bex2,"OK","ANOMALIE","Backlog Execution")
        else:
            st.markdown('<div class="es">Aucune donnee</div>',unsafe_allow_html=True)

if __name__ == "__main__":
    main()
