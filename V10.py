# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os, re
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import plotly.graph_objects as go
from itertools import cycle

# ============================================================
st.set_page_config(layout="wide", page_title="Dashboard KPI")
# ============================================================

# --- Constantes KPI ---
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
ACT_MAP = {"TAUX_REALISATION_CORRECTIF/PT":"Améliorer le taux de réalisation des OT.",
           "OT préparation <1 mois":"Réduire l'âge de préparation des OT (< 1 mois).",
           "OT préparation >3 mois":"Traiter les OT avec préparation > 3 mois.",
           "OT préparation 1mois< <3mois":"Réduire les OT entre 1 et 3 mois de préparation.",
           "OT planification <1 mois":"Réduire l'âge de planification des OT (< 1 mois).",
           "OT planification >3 mois":"Traiter les OT avec planification > 3 mois.",
           "OT planification 1mois< <3mois":"Réduire les OT entre 1 et 3 mois de planification.",
           "OT exécution <1 mois":"Réduire l'âge d'exécution des OT (< 1 mois).",
           "OT exécution >3 mois":"Traiter les OT avec exécution > 3 mois.",
           "OT exécution 1mois< <3mois":"Réduire les OT entre 1 et 3 mois d'exécution.",
           "OT LANC ESTIME":"Estimer les coûts des OT lancés.",
           "Backlog préparation caractérisé":"Caractériser le backlog de préparation.",
           "Backlog planification caractérisé":"Caractériser le backlog de planification.",
           "OT CONFIME":"Confirmer les OT terminés.",
           "OT_COR_EGAL":"Rapprocher les coûts réels et budgétés.",
           "appel avis approuvé":"Créer un OT pour les avis sans ordre."}
LOWER_BETTER = ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois",
                "OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]
MP_KW = ["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
MPLAN_KW = ["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
CONSIGNES_HSE = [
    "Port obligatoire des EPI avant toute intervention.","Port obligatoire du casque de sécurité.",
    "Port obligatoire des lunettes de protection.","Port obligatoire des gants adaptés au travail.",
    "Utiliser les protections auditives dans les zones bruyantes.","Vérifier l'absence de tension avant toute intervention électrique.",
    "Respecter la procédure de consignation et déconsignation.","Ne jamais intervenir sur un équipement en marche.",
    "Baliser et sécuriser la zone de travail.","Maintenir le poste de travail propre et ordonné.",
    "Vérifier l'état des outils avant utilisation.","Utiliser uniquement du matériel homologué.",
    "Respecter les permis de travail en vigueur.","Identifier les risques avant de commencer une tâche.",
    "Signaler immédiatement toute situation dangereuse.","Ne jamais neutraliser un dispositif de sécurité.",
    "Vérifier les détecteurs de gaz avant utilisation.","Respecter les règles des espaces confinés.",
    "Utiliser les points d'ancrage pour les travaux en hauteur.","Ne pas travailler seul lors d'opérations à risque.",
    "Respecter les limites de charge des équipements.","Maintenir les voies de circulation dégagées.",
    "Respecter la signalisation de sécurité.","Connaître les issues de secours les plus proches.",
    "Aucun travail n'est plus urgent que la sécurité.","Zéro accident commence par un comportement sûr.",
    "La sécurité est l'affaire de tous.","Chaque incident peut être évité par la prévention."]

# Palette uniforme pour tous les graphiques
PIE_COLORS = ["#3182CE","#38A169","#D69E2E","#E53E3E","#805AD5",
              "#DD6B20","#319795","#D53F8C","#718096","#2B6CB0",
              "#276749","#B7791F","#9B2C2C","#6B46C1","#C05621",
              "#2C7A7B","#B83280","#4A5568","#ECC94B","#48BB78"]

# ============================================================
# Fonctions utilitaires
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
                if section and headers and cell0 and cell0 not in ("CIBLE","Total général",""):
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
        for sec_name,df_sec,kpi_list in [("Performance",perf_df,QK+["Score Performance"]),("Qualité",qual_df,PK+["Score Qualité"])]:
            prev_d=df_sec[df_sec["Date"]==prev_date].set_index("Poste de travail") if "Poste de travail" in df_sec.columns else pd.DataFrame()
            curr_d=df_sec[df_sec["Date"]==curr_date].set_index("Poste de travail") if "Poste de travail" in df_sec.columns else pd.DataFrame()
            for poste in set(prev_d.index)&set(curr_d.index):
                for kpi in kpi_list:
                    if kpi not in prev_d.columns or kpi not in curr_d.columns: continue
                    try: pv=float(prev_d.loc[poste,kpi])
                    except Exception: continue
                    try: cv=float(curr_d.loc[poste,kpi])
                    except Exception: continue
                    diff=cv-pv; pct=(diff/pv*100) if pv!=0 else (100 if cv!=0 else 0)
                    if abs(diff)<=0.5: trend="stabilité"
                    elif diff>0.5: trend="hausse"
                    else: trend="baisse"
                    variations.append({"Date précédente":prev_date,"Date actuelle":curr_date,"Poste":poste,
                        "Type":sec_name,"KPI":kpi,"Valeur précédente":round(pv,2),"Valeur actuelle":round(cv,2),
                        "Écart":round(diff,2),"Écart %":round(pct,2),"Tendance":trend})
    return pd.DataFrame(variations)

def generate_journal(var_df):
    if var_df.empty: return pd.DataFrame()
    j=var_df.copy(); j["Significatif"]=j["Écart %"].abs()>=5
    j=j[j["Significatif"]].copy()
    j["Sens"]=j.apply(lambda r:"Amélioration" if ((r["Tendance"]=="hausse" and r["KPI"] not in LOWER_BETTER) or (r["Tendance"]=="baisse" and r["KPI"] in LOWER_BETTER)) else "Dégradation",axis=1)
    return j.sort_values(["Date actuelle","Sens","Écart %"],ascending=[True,False,False])

def calculate_rankings(var_df):
    if var_df.empty: return pd.DataFrame(),pd.DataFrame()
    scores={}
    for poste in var_df["Poste"].unique():
        pv=var_df[var_df["Poste"]==poste].copy()
        scores[poste]=sum((-r["Écart %"] if r["KPI"] in LOWER_BETTER else r["Écart %"]) for _,r in pv.iterrows())
    ranked=sorted(scores.items(),key=lambda x:x[1],reverse=True)
    top5=pd.DataFrame(ranked[:5],columns=["Poste","Score variation"]) if len(ranked)>=5 else pd.DataFrame(ranked,columns=["Poste","Score variation"])
    bot5=pd.DataFrame(ranked[-5:][::-1],columns=["Poste","Score variation"]) if len(ranked)>=5 else pd.DataFrame()
    return top5,bot5

def contient_mot(t,lm):
    t=str(t); return any(m in t for l in lm for m in l.split())

def cat_age(a):
    if pd.isna(a): return "Inconnu"
    if a<=1: return "<1 mois"
    elif a>=3: return ">3 mois"
    return "1 mois < <3 mois"

def ckpi(n,d,sz=100): return np.where(d==0,sz,(n/d)*100)

def cpiv(df,f,c,p):
    return pd.pivot_table(df[f],index="Poste travail princ.",columns=c,values="Ordre",aggfunc="count",fill_value=0).reindex(p,fill_value=0)

def get_centrale(p):
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

def is_lb(k): return k in LOWER_BETTER

# ============================================================
# CSS
# ============================================================
def inject_custom_css():
    st.markdown("""<style>
    section[data-testid="stSidebar"]{width:250px!important}
    section[data-testid="stSidebar"][aria-expanded="false"]{width:0px!important}
    .main .block-container{max-width:100%!important;width:100%!important;padding-left:.5rem!important;padding-right:.5rem!important}
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
    .sr{display:flex;align-items:center;padding:6px 10px;background:#fff;border-radius:5px;margin-bottom:2px;border:1px solid var(--b);font-size:13px}
    .sr .sn{font-weight:700;color:var(--p);min-width:220px;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .sc{padding:3px 9px;border-radius:12px;font-weight:800;font-size:14px;min-width:50px;text-align:center;margin:0 8px;color:#fff}
    .sr .sa{color:#718096;font-size:12px;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .g-green{background:#c6efce;color:#006100;font-weight:600}
    .g-yellow{background:#ffeb9c;color:#9c6500;font-weight:600}
    .g-red{background:#ffc7ce;color:#9c0006;font-weight:600}
    .hbar-chart{background:#fff;border-radius:var(--r);padding:14px 18px;border:1px solid var(--b);box-shadow:0 2px 8px rgba(0,0,0,.04);margin-top:8px}
    .hbar-chart .hbar-title{font-size:15px;font-weight:800;color:var(--p);margin-bottom:4px;padding-bottom:6px;border-bottom:2px solid var(--b)}
    .hbar-chart .hbar-sub{font-size:11px;color:#718096;margin-bottom:10px}
    .hbar-row{display:flex;align-items:center;padding:4px 0;gap:8px}
    .hbar-label{width:260px;font-size:11px;font-weight:700;color:#2d3748;text-align:right;padding-right:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .hbar-track{flex:1;position:relative;height:26px;background:#edf2f7;border-radius:4px;overflow:visible}
    .hbar-fill{height:100%;border-radius:4px;position:relative;min-width:2px;transition:width .5s ease}
    .hbar-fill.perf{background:linear-gradient(90deg,#276749,#48bb78)}
    .hbar-fill.qual{background:linear-gradient(90deg,#2b6cb0,#4299e1)}
    .hbar-fill.ok{background:linear-gradient(90deg,#276749,#48bb78)}
    .hbar-fill.ko{background:linear-gradient(90deg,#c53030,#fc8181)}
    .hbar-fill.warn{background:linear-gradient(90deg,#d69e2e,#f6e05e)}
    .hbar-val{position:absolute;right:-58px;top:50%;transform:translateY(-50%);font-size:12px;font-weight:900;color:#1a202c;white-space:nowrap}
    .hbar-target{position:absolute;top:50%;transform:translateY(-50%);width:2px;height:20px;background:#e53e3e;z-index:2}
    .hbar-target-label{position:absolute;top:-14px;transform:translateX(-50%);font-size:9px;color:#e53e3e;font-weight:700;white-space:nowrap}
    .hbar-legend{display:flex;gap:16px;margin-top:10px;padding-top:8px;border-top:1px solid #edf2f7;font-size:11px;font-weight:600;color:#4a5568}
    .hbar-legend span{display:flex;align-items:center;gap:5px}
    .hbar-legend .lg-swatch{width:14px;height:14px;border-radius:3px;display:inline-block}
    .synth-box{background:linear-gradient(135deg,#f7fafc,#edf2f7);border-radius:var(--r);padding:14px 18px;margin-top:8px;border:1px solid #cbd5e0}
    .synth-box .synth-title{font-size:16px;font-weight:800;color:var(--p);margin-bottom:8px;display:flex;align-items:center;gap:8px}
    .synth-kpi-row{display:flex;align-items:center;padding:5px 0;border-bottom:1px solid #e2e8f0;font-size:12px;gap:8px}
    .synth-kpi-row:last-child{border:none}
    .synth-kpi-name{width:280px;font-weight:700;color:#2d3748;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .synth-kpi-val{min-width:65px;text-align:center;font-weight:800;font-size:13px;padding:2px 8px;border-radius:4px}
    .synth-kpi-target{min-width:55px;text-align:center;color:#718096;font-weight:600;font-size:11px}
    .synth-kpi-ecart{min-width:70px;text-align:center;font-weight:800;font-size:12px}
    .synth-kpi-badge{min-width:90px;text-align:center;padding:3px 8px;border-radius:12px;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.3px}
    .synth-kpi-action{flex:1;color:#4a5568;font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .badge-ok{background:#c6efce;color:#006100}
    .badge-ko{background:#ffc7ce;color:#9c0006}
    .badge-warn{background:#ffeb9c;color:#9c6500}
    .synth-resume{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:10px}
    .synth-resume-card{background:#fff;border-radius:8px;padding:10px 14px;text-align:center;border:1px solid var(--b)}
    .synth-resume-card .src-val{font-size:28px;font-weight:900}
    .synth-resume-card .src-lbl{font-size:11px;color:#718096;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
    .action-list{margin-top:6px}
    .action-item{display:flex;align-items:flex-start;padding:6px 10px;margin-bottom:4px;background:#fff;border-radius:6px;border-left:4px solid;font-size:12px;gap:8px}
    .action-item.act-ko{border-left-color:#e53e3e}
    .action-item.act-warn{border-left-color:#d69e2e}
    .action-item .act-kpi{font-weight:800;color:#2d3748;min-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .action-item .act-ecart{font-weight:700;min-width:60px;text-align:center}
    .action-item .act-text{flex:1;color:#4a5568;font-size:11px;line-height:1.4}
    .rank-card{background:#fff;border-radius:var(--r);padding:12px 16px;border:1px solid var(--b);box-shadow:0 2px 8px rgba(0,0,0,.04)}
    .rank-card .rank-title{font-size:15px;font-weight:800;margin-bottom:8px;padding-bottom:5px;border-bottom:2px solid var(--b)}
    .rank-row{display:flex;align-items:center;padding:5px 0;font-size:13px;border-bottom:1px solid #f7fafc}
    .rank-row:last-child{border:none}
    .rank-row .rank-num{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:13px;color:#fff;margin-right:10px;flex-shrink:0}
    .rank-row .rank-name{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .rank-row .rank-score{font-weight:900;min-width:70px;text-align:right}
    .explain-box{background:#fffbeb;border:1px solid #d69e2e;border-radius:var(--r);padding:16px 20px;margin-top:8px}
    .explain-box h3{color:#975a16;font-size:15px;margin-bottom:8px}
    .explain-box p,.explain-box li{color:#4a5568;font-size:13px;line-height:1.6}
    .explain-box ul{padding-left:20px}
    .toggle-btn{display:inline-flex;align-items:center;gap:8px;padding:8px 20px;border-radius:6px;font-weight:700;font-size:14px;cursor:pointer;border:2px solid transparent;transition:all .2s}
    .toggle-btn.active-p{background:#276749;color:#fff;border-color:#276749}
    .toggle-btn.active-q{background:#2b6cb0;color:#fff;border-color:#2b6cb0}
    .toggle-btn.inactive{background:#fff;color:#718096;border-color:#e2e8f0}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stFileUploader label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:3px 8px;margin-bottom:3px;border:1px solid rgba(255,255,255,.1)}
    .trend-up{color:#276749;font-weight:800}.trend-down{color:#c53030;font-weight:800}.trend-stable{color:#718096;font-weight:800}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:17px}.synth-resume{grid-template-columns:1fr}.hbar-label{width:140px}.synth-kpi-name{width:140px}}
    </style>""",unsafe_allow_html=True)

# ============================================================
# Fonctions de rendu HTML
# ============================================================
def ks(v,c):
    """Style cellule KPI avec couleur selon objectif"""
    try: val=float(v)
    except Exception: return ""
    if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]:
        if val>=CIBLE.get(c,80): return "background:#c6efce;color:#006100;font-weight:600"
        if val>=CIBLE.get(c,80)-5: return "background:#ffeb9c;color:#9c6500;font-weight:600"
        return "background:#ffc7ce;color:#9c0006;font-weight:600"
    if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]:
        if val<=CIBLE.get(c,15): return "background:#c6efce;color:#006100;font-weight:600"
        return "background:#ffc7ce;color:#9c0006;font-weight:600"
    if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]:
        if val<=CIBLE.get(c,5): return "background:#c6efce;color:#006100;font-weight:600"
        return "background:#ffc7ce;color:#9c0006;font-weight:600"
    if c=="TAUX_REALISATION_CORRECTIF/PT":
        if val>=85: return "background:#c6efce;color:#006100;font-weight:600"
        if val>=80: return "background:#ffeb9c;color:#9c6500;font-weight:600"
        return "background:#ffc7ce;color:#9c0006;font-weight:600"
    if c=="appel avis approuvé":
        if val>=95: return "background:#c6efce;color:#006100;font-weight:600"
        if val>=90: return "background:#ffeb9c;color:#9c6500;font-weight:600"
        return "background:#ffc7ce;color:#9c0006;font-weight:600"
    if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]:
        if val>=100: return "background:#c6efce;color:#006100;font-weight:600"
        if val>=95: return "background:#ffeb9c;color:#9c6500;font-weight:600"
        return "background:#ffc7ce;color:#9c0006;font-weight:600"
    return ""

def cs(v):
    try: val=float(str(v).replace(' %','').strip())
    except Exception: return ""
    if val>=90: return "background:#c6efce;color:#006100;font-weight:700"
    if val>=80: return "background:#ffeb9c;color:#9c6500;font-weight:700"
    return "background:#ffc7ce;color:#9c0006;font-weight:700"

def kas(v):
    try: val=int(v)
    except Exception: return ""
    if val==0: return "color:#cbd5e0"
    if val<=3: return "background:#ffeb9c;color:#9c6500;font-weight:600"
    if val<=10: return "background:#fed7d7;color:#c53030;font-weight:600"
    return "background:#fc8181;color:#742a2a;font-weight:800"

def gscore(k,a):
    if pd.isna(a): return 0
    if k in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return 1 if a>=75 else 0
    if k in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return 1 if a<=15 else 0
    if k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return 1 if a<=5 else 0
    if k=="TAUX_REALISATION_CORRECTIF/PT": return 1 if a>=80 else 0
    if k=="appel avis approuvé": return 1 if a>=90 else 0
    if k in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]: return 1 if a>=95 else 0
    return 0

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
    h+='</tbody></table>'
    return h

def html_ano(rows,cols):
    h='<table class="tw at"><thead><tr>'+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
    for r in rows:
        h+='<tr class="%s">'%("tr" if r.get("_t")=="total" else "")
        for c in cols: v=r.get(c,""); h+='<td style="%s">%s</td>'%(kas(v) or "",v)
        h+='</tr>'
    h+='</tbody></table>'
    return h

def html_hbar(kpi_list, actuals, targets, total_general, section_type, title):
    color_main = "#276749" if section_type=="perf" else "#2b6cb0"
    label = "Performance" if section_type=="perf" else "Qualité"
    h='<div class="hbar-chart"><div class="hbar-title">%s</div>'%title
    h+='<div class="hbar-sub">Comparaison au Total Général (%.1f%%) — %s</div>'%(total_general,label)
    all_v=[actuals.get(k,0) for k in kpi_list]+[total_general]+[targets.get(k,100) for k in kpi_list]
    mx=max(all_v)*1.15 if max(all_v)>0 else 100
    for k in kpi_list:
        av=actuals.get(k,0); tv=targets.get(k,100)
        met=av<=tv if is_lb(k) else av>=tv
        if met: fc="ok"
        elif is_lb(k): fc="warn" if av<=tv*1.5 else "ko"
        else: fc="warn" if av>=tv-5 else "ko"
        bw=max((av/mx)*100,0.5); tp=min((tv/mx)*100,100); tgp=min((total_general/mx)*100,100)
        h+='<div class="hbar-row"><div class="hbar-label">%s</div>'%k
        h+='<div class="hbar-track"><div class="hbar-fill %s" style="width:%.1f%%"><div class="hbar-val">%.1f%%</div></div>'%(fc,bw,av)
        h+='<div class="hbar-target" style="left:%.1f%%"><div class="hbar-target-label">C:%.0f</div></div>'%(tp,tv)
        h+='<div style="position:absolute;left:%.1f%%;top:50%%;transform:translateY(-50%%);width:2px;height:26px;border-left:2px dashed %s;z-index:1;opacity:.7"></div>'%(tgp,color_main)
        h+='</div></div>'
    h+='<div class="hbar-legend"><span><i class="lg-swatch" style="background:#276749"></i>Atteint</span>'
    h+='<span><i class="lg-swatch" style="background:#d69e2e"></i>Attention</span>'
    h+='<span><i class="lg-swatch" style="background:#c53030"></i>Non atteint</span>'
    h+='<span><i class="lg-swatch" style="border-left:2px dashed %s"></i>Total Général</span></div></div>'%color_main
    return h

def html_synthese(kpi_list, actuals, targets, section_type):
    icon="📊" if section_type=="perf" else "✅"
    label="PERFORMANCE" if section_type=="perf" else "QUALITÉ"
    nb_ok=nb_ko=nb_warn=0; det=""
    for k in kpi_list:
        av=actuals.get(k,0); tv=targets.get(k,100); diff=av-tv
        met=av<=tv if is_lb(k) else av>=tv
        if is_lb(k): iw=(not met) and av<=tv*1.5
        else: iw=(not met) and av>=tv-5
        if met: nb_ok+=1; bc="badge-ok"; bt="ATTEINT"; vs="background:#c6efce;color:#006100"
        elif iw: nb_warn+=1; bc="badge-warn"; bt="ATTENTION"; vs="background:#ffeb9c;color:#9c6500"
        else: nb_ko+=1; bc="badge-ko"; bt="NON ATTEINT"; vs="background:#ffc7ce;color:#9c0006"
        ec="#276749" if met else ("#9c6500" if iw else "#c53030")
        act="✅ Objectif atteint" if met else ACT_MAP.get(k,"À définir")
        det+='<div class="synth-kpi-row"><div class="synth-kpi-name">%s</div>'%k
        det+='<div class="synth-kpi-val" style="%s">%.1f%%</div>'%(vs,av)
        det+='<div class="synth-kpi-target">Cible: %.0f%%</div>'%tv
        det+='<div class="synth-kpi-ecart" style="color:%s">%+.1f%%</div>'%(ec,diff)
        det+='<div class="synth-kpi-badge %s">%s</div>'%(bc,bt)
        det+='<div class="synth-kpi-action">%s</div></div>'%act
    acts=""
    ko_kpis=[]; warn_kpis=[]
    for k in kpi_list:
        av=actuals.get(k,0); tv=targets.get(k,100); diff=av-tv; met=av<=tv if is_lb(k) else av>=tv
        if not met:
            iw=(av<=tv*1.5) if is_lb(k) else (av>=tv-5)
            (warn_kpis if iw else ko_kpis).append((k,diff))
    if ko_kpis or warn_kpis:
        acts='<div class="action-list"><div style="font-size:13px;font-weight:800;color:#c53030;margin-bottom:4px">🔴 Actions Correctives Prioritaires</div>'
        for k,d in ko_kpis:
            acts+='<div class="action-item act-ko"><div class="act-kpi">%s</div><div class="act-ecart" style="color:#c53030">%+.1f%%</div><div class="act-text">%s</div></div>'%(k,d,ACT_MAP.get(k,""))
        if warn_kpis:
            acts+='<div style="font-size:13px;font-weight:800;color:#d69e2e;margin:6px 0 4px 0">🟡 Actions d\'Amélioration</div>'
            for k,d in warn_kpis:
                acts+='<div class="action-item act-warn"><div class="act-kpi">%s</div><div class="act-ecart" style="color:#9c6500">%+.1f%%</div><div class="act-text">%s</div></div>'%(k,d,ACT_MAP.get(k,""))
        acts+='</div>'
    h='<div class="synth-box"><div class="synth-title"><span style="font-size:20px">%s</span> Synthèse & Actions — %s</div>'%(icon,label)
    h+='<div class="synth-resume">'
    h+='<div class="synth-resume-card"><div class="src-val" style="color:#276749">%d</div><div class="src-lbl">Atteints</div></div>'%nb_ok
    h+='<div class="synth-resume-card"><div class="src-val" style="color:#d69e2e">%d</div><div class="src-lbl">Attention</div></div>'%nb_warn
    h+='<div class="synth-resume-card"><div class="src-val" style="color:#c53030">%d</div><div class="src-lbl">Non Atteints</div></div>'%nb_ko
    h+='</div>%s%s</div>'%(det,acts)
    return h

# ============================================================
# Pie Chart avec Pie of Pie
# ============================================================
def make_pie_of_pie(labels, values, title, total_label=None):
    """Crée un pie chart; si des parts < 5%, utilise mode Pie of Pie (2 graphiques côte à côte)"""
    total = sum(values)
    if total == 0:
        fig = go.Figure()
        fig.add_annotation(text="Aucune donnée", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=16, color="#718096"))
        fig.update_layout(title=dict(text=f"{title}<br><sup>Total: 0</sup>", font=dict(size=14)), height=350, margin=dict(t=60,b=20,l=20,r=20))
        return fig

    pct = [v/total*100 for v in values]
    threshold = 5
    small_idx = [i for i,p in enumerate(pct) if p < threshold and p > 0]
    n_total = f"Total: {total:,}" if total_label is None else f"{total_label}: {total:,}"

    if len(small_idx) >= 2:
        # Mode Pie of Pie
        main_labels = []; main_vals = []; other_labels = []; other_vals = []
        other_sum = 0
        for i,(l,v) in enumerate(zip(labels,values)):
            if i in small_idx:
                other_labels.append(l); other_vals.append(v); other_sum += v
            else:
                main_labels.append(l); main_vals.append(v)
        main_labels.append("Autres"); main_vals.append(other_sum)

        fig = make_subplots(rows=1, cols=2, specs=[[{"type":"domain"},{"type":"domain"}]],
                           subplot_titles=("Répartition principale","Détail « Autres »"))
        clrs_main = list(PIE_COLORS[:len(main_labels)-1]) + ["#A0AEC0"]
        fig.add_trace(go.Pie(labels=main_labels, values=main_vals, marker_colors=clrs_main,
                             textinfo='label+percent+value', textposition='outside',
                             hole=0, pull=[0.05 if l=="Autres" else 0 for l in main_labels],
                             name="Principal"), row=1, col=1)
        clrs_other = [PIE_COLORS[labels.index(l) % len(PIE_COLORS)] for l in other_labels]
        fig.add_trace(go.Pie(labels=other_labels, values=other_vals, marker_colors=clrs_other,
                             textinfo='label+percent+value', textposition='outside',
                             hole=0, name="Détail"), row=1, col=2)
        fig.update_layout(title=dict(text=f"{title}<br><sup>{n_total}</sup>", font=dict(size=14)),
                         height=380, margin=dict(t=70,b=20,l=20,r=20),
                         legend=dict(font=dict(size=10), orientation="h", yanchor="bottom", y=-0.15),
                         showlegend=True)
    else:
        clrs = [PIE_COLORS[i % len(PIE_COLORS)] for i in range(len(labels))]
        fig = go.Figure(go.Pie(labels=labels, values=values, marker_colors=clrs,
                                textinfo='label+percent+value', textposition='outside', hole=0))
        fig.update_layout(title=dict(text=f"{title}<br><sup>{n_total}</sup>", font=dict(size=14)),
                         height=380, margin=dict(t=70,b=20,l=20,r=20),
                         legend=dict(font=dict(size=10), orientation="h", yanchor="bottom", y=-0.15),
                         showlegend=True)
    return fig

# ============================================================
# Calcul des KPIs
# ============================================================
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
    res['dfp']=df

    an=cpiv(df,df["Nº appel pl.entret."].fillna(0)==0,"Statut OT",posts)
    for c in ["CLOT","CRÉÉ","LANC","TCLO"]: an[c]=an.get(c,0)
    an["Total"]=an[["CLOT","CRÉÉ","LANC","TCLO"]].sum(axis=1)
    an["TAUX_REALISATION_CORRECTIF/PT"]=ckpi(an["TCLO"],an["Total"])

    pr=cpiv(df,df["Statut OT"]=="CRÉÉ","ap",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pr[c]=pr.get(c,0)
    pr["Total"]=pr[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    pr["OT préparation <1 mois"]=ckpi(pr["<1 mois"],pr["Total"])
    pr["OT préparation >3 mois"]=ckpi(pr[">3 mois"],pr["Total"],0)
    pr["OT préparation 1mois< <3mois"]=ckpi(pr["1 mois < <3 mois"],pr["Total"],0)

    pl=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==0),"alp",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: pl[c]=pl.get(c,0)
    pl["Total"]=pl[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    pl["OT planification <1 mois"]=ckpi(pl["<1 mois"],pl["Total"])
    pl["OT planification >3 mois"]=ckpi(pl[">3 mois"],pl["Total"],0)
    pl["OT planification 1mois< <3mois"]=ckpi(pl["1 mois < <3 mois"],pl["Total"],0)

    ex=cpiv(df,(df["Statut OT"]=="LANC")&(df["Contient SOPL"]==1),"aex",posts)
    for c in ["<1 mois",">3 mois","1 mois < <3 mois"]: ex[c]=ex.get(c,0)
    ex["Total"]=ex[["<1 mois","1 mois < <3 mois",">3 mois"]].sum(axis=1)
    ex["OT exécution <1 mois"]=ckpi(ex["<1 mois"],ex["Total"])
    ex["OT exécution >3 mois"]=ckpi(ex[">3 mois"],ex["Total"],0)
    ex["OT exécution 1mois< <3mois"]=ckpi(ex["1 mois < <3 mois"],ex["Total"],0)

    la=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="OT LANC ESTIME",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["OUI","NON"]: la[c]=la.get(c,0)
    la["Total"]=la["OUI"]+la["NON"]; la["OT LANC ESTIME"]=ckpi(la["OUI"],la["Total"])

    pc=pd.pivot_table(df[df["Statut OT"]=="CRÉÉ"],index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: pc[c]=pc.get(c,0)
    pc["Total"]=pc["CARACTERISE"]+pc["NON CARACTERISE"]; pc["Backlog préparation caractérisé"]=ckpi(pc["CARACTERISE"],pc["Total"])

    plc=pd.pivot_table(df[df["Statut OT"]=="LANC"],index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0).reindex(posts,fill_value=0)
    for c in ["CARACTERISE","NON CARACTERISE"]: plc[c]=plc.get(c,0)
    plc["Total"]=plc["CARACTERISE"]+plc["NON CARACTERISE"]; plc["Backlog planification caractérisé"]=ckpi(plc["CARACTERISE"],plc["Total"])

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
    res['an']=an; res['pr']=pr; res['pl']=pl; res['ex']=ex
    res['la']=la; res['pc']=pc; res['plc']=plc; res['tca']=tca
    return res

# ============================================================
# Main
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
        <h1 style="text-align:center;font-size:46px;color:#fff;font-weight:900;margin:0">HSE - CONSIGNE DE SÉCURITÉ</h1>
        <p style="text-align:center;color:rgba(255,255,255,.6);font-size:22px;margin-top:8px;letter-spacing:3px;text-transform:uppercase">Sécurité - Santé - Environnement</p>
        <div style="background:linear-gradient(135deg,#f6e05e,#ed8936);padding:36px 48px;border-radius:20px;font-size:32px;font-weight:700;text-align:center;margin:40px 0;color:#1a202c;max-width:800px;box-shadow:0 20px 60px rgba(0,0,0,.3)">⚠️ %s</div>
        <h2 style="text-align:center;color:#48bb78;font-size:36px;font-weight:900">Aucun travail n'est plus urgent que la sécurité</h2>
        <div style="margin-top:40px;width:200px;height:4px;background:rgba(255,255,255,.1);border-radius:2px;overflow:hidden"><div style="width:100%%;height:100%%;background:linear-gradient(90deg,#48bb78,#38a169);border-radius:2px;animation:ld 5.5s ease-in-out forwards"></div></div>
        <style>@keyframes ld{from{width:0}to{width:100%%}}</style></div>"""%c,unsafe_allow_html=True)
        time.sleep(6); st.session_state.hse_affiche=True; st.rerun(); st.stop()

    # --- Sidebar ---
    with st.sidebar:
        st.markdown("### 📁 Fichiers source")
        f_ot = st.file_uploader("Fichier OT (Excel)", type=["xlsx","xls"], key="ot")
        f_av = st.file_uploader("Fichier Avis (Excel)", type=["xlsx","xls"], key="av")
        st.markdown("---")
        st.markdown("### 🎛️ Filtres")
        filtre_centrale = st.multiselect("Centrale", ["Sulfurique","Phosphorique","Engrais","Feed","Autre"], [])
        filtre_division = st.multiselect("Division", ["SF1","SF2","Autre"], [])

    # --- Chargement données ---
    df_ot = pd.DataFrame(); df_av = pd.DataFrame()
    if f_ot:
        try: df_ot=pd.read_excel(f_ot); df_ot=df_ot[~df_ot["Poste travail princ."].astype(str).str.contains("cresseur",case=False,na=False)]
        except Exception: st.error("Erreur lecture fichier OT")
    if f_av:
        try: df_av=pd.read_excel(f_av)
        except Exception: st.error("Erreur lecture fichier Avis")

    if df_ot.empty:
        st.markdown('<div style="text-align:center;padding:60px;color:#718096;font-size:18px">📂 Veuillez charger le fichier OT dans la barre latérale pour afficher le dashboard.</div>',unsafe_allow_html=True)
        st.stop()

    if "Poste travail princ." not in df_ot.columns:
        st.error("Colonne 'Poste travail princ.' introuvable dans le fichier OT."); st.stop()

    posts = sorted(df_ot["Poste travail princ."].dropna().unique().tolist())
    # Appliquer filtres
    if filtre_centrale:
        posts = [p for p in posts if get_centrale(p) in filtre_centrale]
    if filtre_division:
        posts = [p for p in posts if get_division(p) in filtre_division]
    if not posts:
        st.warning("Aucun poste de travail sélectionné avec les filtres actuels."); st.stop()

    now = datetime.now()
    if fichier_date:
        try: now = pd.to_datetime(fichier_date, format="%d/%m/%Y")
        except Exception: pass

    kpis = calc_kpis(df_ot, df_av, now, posts)
    ckdf = kpis['ckdf']

    # Scores
    perf_scores = {}; qual_scores = {}
    for p in posts:
        sp = sum(gscore(k, ckdf.loc[p,k]) for k in QK if p in ckdf.index and k in ckdf.columns)
        sq = sum(gscore(k, ckdf.loc[p,k]) for k in PK if p in ckdf.index and k in ckdf.columns)
        perf_scores[p] = round(sp/len(QK)*100,1) if QK else 0
        qual_scores[p] = round(sq/len(PK)*100,1) if PK else 0

    tg_perf = np.mean(list(perf_scores.values())) if perf_scores else 0
    tg_qual = np.mean(list(qual_scores.values())) if qual_scores else 0

    # Anomalies
    ano_p = []; ano_q = []
    for p in posts:
        if p not in ckdf.index: continue
        for k in QK:
            v=ckdf.loc[p,k]
            if not gscore(k,v): ano_p.append({"Poste de travail":p,"KPI":k,"Valeur":round(v,1),"Cible":CIBLE.get(k,"-")})
        for k in PK:
            v=ckdf.loc[p,k]
            if not gscore(k,v): ano_q.append({"Poste de travail":p,"KPI":k,"Valeur":round(v,1),"Cible":CIBLE.get(k,"-")})

    # Sauvegarde Excel
    pcols=["Poste de travail"]+QK+["Score Performance"]
    prows=[{"Poste de travail":p,"_t":""} for p in posts]
    for i,p in enumerate(posts):
        for k in QK: prows[i][k]=round(ckdf.loc[p,k],1) if p in ckdf.index else 0
        prows[i]["Score Performance"]=perf_scores.get(p,0)
    prows.append({"Poste de travail":"Total général","_t":"total"})
    tot_row=prows[-1]
    for k in QK: tot_row[k]=round(ckdf[k].mean(),1) if k in ckdf.columns else 0
    tot_row["Score Performance"]=round(tg_perf,1)
    prows.append({"Poste de travail":"CIBLE","_t":"cible"})
    for k in QK: prows[-1][k]=CIBLE.get(k,"-")
    prows[-1]["Score Performance"]=100

    qcols=["Poste de travail"]+PK+["Score Qualité"]
    qrows=[{"Poste de travail":p,"_t":""} for p in posts]
    for i,p in enumerate(posts):
        for k in PK: qrows[i][k]=round(ckdf.loc[p,k],1) if p in ckdf.index else 0
        qrows[i]["Score Qualité"]=qual_scores.get(p,0)
    qrows.append({"Poste de travail":"Total général","_t":"total"})
    tot_row=qrows[-1]
    for k in PK: tot_row[k]=round(ckdf[k].mean(),1) if k in ckdf.columns else 0
    tot_row["Score Qualité"]=round(tg_qual,1)
    qrows.append({"Poste de travail":"CIBLE","_t":"cible"})
    for k in PK: qrows[-1][k]=CIBLE.get(k,"-")
    qrows[-1]["Score Qualité"]=100

    ano_p_cols=["Poste de travail","KPI","Valeur","Cible"]; ano_q_cols=ano_p_cols[:]
    ano_p_rows=[{"_t":""}]+ano_p
    if ano_p: ano_p_rows.append({"Poste de travail":"TOTAL","KPI":"","Valeur":len(ano_p),"Cible":"","_t":"total"})
    ano_q_rows=[{"_t":""}]+ano_q
    if ano_q: ano_q_rows.append({"Poste de travail":"TOTAL","KPI":"","Valeur":len(ano_q),"Cible":"","_t":"total"})

    save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_rows,ano_p_cols,ano_q_rows,ano_q_cols,fichier_date)

    # --- Header ---
    st.markdown('<div class="mh"><h1>📈 Dashboard KPI Maintenance</h1><div class="db">📅 %s</div></div>'%fichier_date,unsafe_allow_html=True)

    # Cartes résumé
    st.markdown('<div class="cr">'
        '<div class="cc c1"><div class="cv">%d</div><div class="cl">Postes de travail</div></div>'
        '<div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Performance</div></div>'
        '<div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Qualité</div></div>'
        '<div class="cc c4"><div class="cv">%d</div><div class="cl">Anomalies totales</div></div>'
        '</div>'%(len(posts),tg_perf,tg_qual,len(ano_p)+len(ano_q)),unsafe_allow_html=True)

    # ============================================================
    # ONGLETS
    # ============================================================
    tab1, tab2, tab3, tab4 = st.tabs(["📌 Présentation", "⚡ Performance & Qualité", "🔍 Analyse", "📉 Suivi des Améliorations"])

    # ============================================================
    # TAB 1 : PRÉSENTATION (anciennement Analyse + Classement intégré)
    # ============================================================
    with tab1:
        # --- Analyse Comparative avec couleur auto ---
        st.markdown('<div class="stl">Analyse Comparative — Mise en couleur selon atteinte des objectifs</div>',unsafe_allow_html=True)
        comp_cols=["Poste de travail"]+QK+PK+["Score Performance","Score Qualité"]
        comp_rows=[]
        for p in posts:
            r={"Poste de travail":p,"_t":""}
            for k in QK+PK: r[k]=round(ckdf.loc[p,k],1) if p in ckdf.index else 0
            r["Score Performance"]=perf_scores.get(p,0); r["Score Qualité"]=qual_scores.get(p,0)
            comp_rows.append(r)
        comp_rows.append({"Poste de travail":"CIBLE","_t":"cible"})
        for k in QK+PK: comp_rows[-1][k]=CIBLE.get(k,"-")
        comp_rows[-1]["Score Performance"]=100; comp_rows[-1]["Score Qualité"]=100
        st.markdown(html_table(comp_rows,comp_cols,"qt"),unsafe_allow_html=True)

        st.markdown('<div style="font-size:11px;color:#718096;margin-top:2px">🟢 Vert = objectif atteint ou dépassé &nbsp;|&nbsp; 🟡 Jaune = proche de l\'objectif &nbsp;|&nbsp; 🔴 Rouge = objectif non atteint</div>',unsafe_allow_html=True)

        # --- Répartition par centrale (remplace atelier) ---
        st.markdown('<div class="stl c">Répartition par Centrale</div>',unsafe_allow_html=True)
        df_centrale = kpis['dfp'].copy()
        df_centrale["Centrale"] = df_centrale["Poste travail princ."].apply(get_centrale)
        rep_c = df_centrale.groupby("Centrale")["Ordre"].count().sort_values(ascending=False)
        fig_c = make_pie_of_pie(rep_c.index.tolist(), rep_c.values.tolist(), "Répartition des OT par Centrale", "OT")
        st.plotly_chart(fig_c, use_container_width=True)

        # --- Répartition par division ---
        st.markdown('<div class="stl s">Répartition par Division</div>',unsafe_allow_html=True)
        df_div = kpis['dfp'].copy()
        df_div["Division"] = df_div["Poste travail princ."].apply(get_division)
        rep_d = df_div.groupby("Division")["Ordre"].count().sort_values(ascending=False)
        fig_d = make_pie_of_pie(rep_d.index.tolist(), rep_d.values.tolist(), "Répartition des OT par Division", "OT")
        st.plotly_chart(fig_d, use_container_width=True)

        # --- Classement intégré ---
        st.markdown('<div class="stl p">Classement des Postes de Travail</div>',unsafe_allow_html=True)

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.markdown('<div class="rank-card"><div class="rank-title" style="color:#276749">🏆 Classement Performance</div>' ,unsafe_allow_html=True)
            ranked_p = sorted(perf_scores.items(), key=lambda x:x[1], reverse=True)
            for i,(p,s) in enumerate(ranked_p):
                clr="#276749" if i<3 else ("#d69e2e" if i<len(ranked_p)-3 else "#c53030")
                st.markdown('<div class="rank-row"><div class="rank-num" style="background:%s">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:%s">%.1f%%</div></div>'%(clr,i+1,p,clr,s),unsafe_allow_html=True)
            st.markdown('</div>',unsafe_allow_html=True)

        with col_r2:
            st.markdown('<div class="rank-card"><div class="rank-title" style="color:#2b6cb0">🏆 Classement Qualité</div>',unsafe_allow_html=True)
            ranked_q = sorted(qual_scores.items(), key=lambda x:x[1], reverse=True)
            for i,(p,s) in enumerate(ranked_q):
                clr="#2b6cb0" if i<3 else ("#d69e2e" if i<len(ranked_q)-3 else "#c53030")
                st.markdown('<div class="rank-row"><div class="rank-num" style="background:%s">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:%s">%.1f%%</div>'%(clr,i+1,p,clr,s),unsafe_allow_html=True)
            st.markdown('</div>',unsafe_allow_html=True)

        # --- Explication du mode de calcul ---
        st.markdown('<div class="explain-box"><h3>📋 Mode de calcul du classement</h3>',unsafe_allow_html=True)
        st.markdown("""<ul>
        <li><strong>Classement Performance :</strong> Basé sur les {n_perf} indicateurs de performance (TAUX_RÉALISATION, âge préparation/planification/exécution).</li>
        <li><strong>Classement Qualité :</strong> Basé sur les {n_qual} indicateurs de qualité (appels avis, OT estimés, backlog caractérisé, OT confirmés, coûts).</li>
        <li><strong>Pondérations :</strong> Chaque indicateur a un poids égal (1 point si atteint, 0 sinon). Le score est la moyenne : (Points obtenus / Nombre d'indicateurs) × 100.</li>
        <li><strong>Seuils d'atteinte :</strong>
            <ul>
            <li>TAUX_RÉALISATION ≥ 80%, Prépa/Planif/Exéc &lt;1 mois ≥ 75%, &gt;3 mois ≤ 5%, 1-3 mois ≤ 15%</li>
            <li>Appel avis ≥ 90%, OT LANC ESTIME / Backlog caractérisé / OT CONFIME / OT_COR_EGAL ≥ 95%</li>
            </ul>
        </li>
        <li><strong>Critères de tri :</strong> Ordre décroissant du score. Les 3 premiers sont en vert, les 3 derniers en rouge.</li>
        <li><strong>Gestion des ex æquo :</strong> En cas de score identique, les postes sont classés par ordre alphabétique du nom.</li>
        </ul>""".format(n_perf=len(QK), n_qual=len(PK)),unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)

    # ============================================================
    # TAB 2 : PERFORMANCE & QUALITÉ (avec bouton bascule)
    # ============================================================
    with tab2:
        if "pq_toggle" not in st.session_state: st.session_state.pq_toggle = "perf"

        col_btn1, col_btn2, col_spacer = st.columns([1,1,4])
        with col_btn1:
            if st.button("📊 Performance", use_container_width=True, type="primary" if st.session_state.pq_toggle=="perf" else "secondary"):
                st.session_state.pq_toggle = "perf"
        with col_btn2:
            if st.button("✅ Qualité", use_container_width=True, type="primary" if st.session_state.pq_toggle=="qual" else "secondary"):
                st.session_state.pq_toggle = "qual"

        mode = st.session_state.pq_toggle

        if mode == "perf":
            kpi_list = QK; score_dict = perf_scores; tg = tg_perf
            label = "Performance"; stype = "perf"; tc = "pt"
            # Barres horizontales AVANT le tableau
            actuals = {k: round(ckdf[k].mean(),1) for k in QK} if not ckdf.empty else {}
            st.markdown(html_hbar(QK, actuals, CIBLE, tg, "perf", "Chart Bar — Performance vs Total Général"),unsafe_allow_html=True)
            # Tableau
            st.markdown('<div class="stl p">Indicateurs de Performance (%)</div>',unsafe_allow_html=True)
            st.markdown(html_table(prows,pcols,"pt",sc_col=set(QK)|{"Score Performance"}),unsafe_allow_html=True)
            # Anomalies
            st.markdown('<div class="stl a">Nombre d\'anomalies Performance</div>',unsafe_allow_html=True)
            if ano_p_rows: st.markdown(html_ano(ano_p_rows,ano_p_cols),unsafe_allow_html=True)
            else: st.markdown('<div style="padding:10px;color:#276749;font-weight:700">✅ Aucune anomalie performance</div>',unsafe_allow_html=True)
            # Synthèse
            st.markdown(html_synthese(QK, actuals, CIBLE, "perf"),unsafe_allow_html=True)
        else:
            kpi_list = PK; score_dict = qual_scores; tg = tg_qual
            label = "Qualité"; stype = "qual"; tc = "qt"
            actuals = {k: round(ckdf[k].mean(),1) for k in PK} if not ckdf.empty else {}
            st.markdown(html_hbar(PK, actuals, CIBLE, tg, "qual", "Chart Bar — Qualité vs Total Général"),unsafe_allow_html=True)
            st.markdown('<div class="stl q">Indicateurs de Qualité (%)</div>',unsafe_allow_html=True)
            st.markdown(html_table(qrows,qcols,"qt",sc_col=set(PK)|{"Score Qualité"}),unsafe_allow_html=True)
            st.markdown('<div class="stl a">Nombre d\'anomalies Qualité</div>',unsafe_allow_html=True)
            if ano_q_rows: st.markdown(html_ano(ano_q_rows,ano_q_cols),unsafe_allow_html=True)
            else: st.markdown('<div style="padding:10px;color:#276749;font-weight:700">✅ Aucune anomalie qualité</div>',unsafe_allow_html=True)
            st.markdown(html_synthese(PK, actuals, CIBLE, "qual"),unsafe_allow_html=True)

    # ============================================================
    # TAB 3 : ANALYSE (Backlog Préparation, Planification, OMS, Thermographie)
    # ============================================================
    with tab3:
        sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs(["📦 Backlog Préparation", "📅 Backlog Planification", "🔧 OMS", "🌡️ Thermographie"])

        # --- Backlog Préparation ---
        with sub_tab1:
            st.markdown('<div class="stl q">Backlog Préparation par Poste de Travail</div>',unsafe_allow_html=True)
            df_bp = kpis['dfp'][kpis['dfp']["Statut OT"]=="CRÉÉ"].copy()
            bp_data = []
            for p in posts:
                sub = df_bp[df_bp["Poste travail princ."]==p]
                total = len(sub)
                carac = len(sub[sub["Backlog preparation"]=="CARACTERISE"])
                pct = round(carac/total*100,1) if total>0 else 0
                bp_data.append({"Poste de travail":p,"Total Backlog Préparation":total,
                               "Backlog Préparation Caractérisé":carac,"% Caractérisé":pct})
            if bp_data:
                bp_df = pd.DataFrame(bp_data)
                # Style pour % caractérisé
                bp_html = '<table class="tw qt"><thead><tr><th>Poste de travail</th><th>Total Backlog Préparation</th><th>Backlog Préparation Caractérisé</th><th>% Caractérisé</th></tr></thead><tbody>'
                for _,r in bp_df.iterrows():
                    v=r["% Caractérisé"]
                    s = "background:#c6efce;color:#006100;font-weight:600" if v>=80 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if v>=50 else "background:#ffc7ce;color:#9c0006;font-weight:600")
                    bp_html += '<tr><td>%s</td><td>%d</td><td>%d</td><td style="%s">%.1f%%</td></tr>'%(r["Poste de travail"],r["Total Backlog Préparation"],r["Backlog Préparation Caractérisé"],s,v)
                tot_t=bp_df["Total Backlog Préparation"].sum(); tot_c=bp_df["Backlog Préparation Caractérisé"].sum(); tot_p=round(tot_c/tot_t*100,1) if tot_t>0 else 0
                bp_html += '<tr class="tr"><td>Total général</td><td>%d</td><td>%d</td><td>%.1f%%</td></tr></tbody></table>'%(tot_t,tot_c,tot_p)
                st.markdown(bp_html,unsafe_allow_html=True)

                # Pie 1: Caractérisé vs Non
                fig1 = make_pie_of_pie(["Caractérisé","Non caractérisé"],[tot_c,tot_t-tot_c],
                                       "Backlog Préparation : Caractérisé vs Non", "Total OT")
                st.plotly_chart(fig1, use_container_width=True)

                # Pie 2: Types de caractérisation
                df_carac = df_bp[df_bp["Backlog preparation"]=="CARACTERISE"]
                type_counts = {}
                for _,row in df_carac.iterrows():
                    s=str(row.get("Statut utilisateur","")).upper()
                    matched = [kw for kw in MP_KW if kw in s]
                    t = max(matched,key=len) if matched else "AUTRE"
                    type_counts[t] = type_counts.get(t,0)+1
                if type_counts:
                    fig2 = make_pie_of_pie(list(type_counts.keys()),list(type_counts.values()),
                                           "Répartition des Types de Caractérisation (Préparation)", "Total")
                    st.plotly_chart(fig2, use_container_width=True)

        # --- Backlog Planification ---
        with sub_tab2:
            st.markdown('<div class="stl p">Backlog Planification par Poste de Travail</div>',unsafe_allow_html=True)
            df_bl = kpis['dfp'][kpis['dfp']["Statut OT"]=="LANC"].copy()
            bl_data = []
            for p in posts:
                sub = df_bl[df_bl["Poste travail princ."]==p]
                total = len(sub)
                carac = len(sub[sub["Backlog planification"]=="CARACTERISE"])
                pct = round(carac/total*100,1) if total>0 else 0
                bl_data.append({"Poste de travail":p,"Total Backlog Planification":total,
                               "Backlog Planification Caractérisé":carac,"% Caractérisé":pct})
            if bl_data:
                bl_df = pd.DataFrame(bl_data)
                bl_html = '<table class="tw pt"><thead><tr><th>Poste de travail</th><th>Total Backlog Planification</th><th>Backlog Planification Caractérisé</th><th>% Caractérisé</th></tr></thead><tbody>'
                for _,r in bl_df.iterrows():
                    v=r["% Caractérisé"]
                    s = "background:#c6efce;color:#006100;font-weight:600" if v>=80 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if v>=50 else "background:#ffc7ce;color:#9c0006;font-weight:600")
                    bl_html += '<tr><td>%s</td><td>%d</td><td>%d</td><td style="%s">%.1f%%</td></tr>'%(r["Poste de travail"],r["Total Backlog Planification"],r["Backlog Planification Caractérisé"],s,v)
                tot_t=bl_df["Total Backlog Planification"].sum(); tot_c=bl_df["Backlog Planification Caractérisé"].sum(); tot_p=round(tot_c/tot_t*100,1) if tot_t>0 else 0
                bl_html += '<tr class="tr"><td>Total général</td><td>%d</td><td>%d</td><td>%.1f%%</td></tr></tbody></table>'%(tot_t,tot_c,tot_p)
                st.markdown(bl_html,unsafe_allow_html=True)

                fig1 = make_pie_of_pie(["Caractérisé","Non caractérisé"],[tot_c,tot_t-tot_c],
                                       "Backlog Planification : Caractérisé vs Non", "Total OT")
                st.plotly_chart(fig1, use_container_width=True)

                df_carac = df_bl[df_bl["Backlog planification"]=="CARACTERISE"]
                type_counts = {}
                for _,row in df_carac.iterrows():
                    s=str(row.get("Statut utilisateur","")).upper()
                    matched = [kw for kw in MPLAN_KW if kw in s]
                    t = max(matched,key=len) if matched else "AUTRE"
                    type_counts[t] = type_counts.get(t,0)+1
                if type_counts:
                    fig2 = make_pie_of_pie(list(type_counts.keys()),list(type_counts.values()),
                                           "Répartition des Types de Caractérisation (Planification)", "Total")
                    st.plotly_chart(fig2, use_container_width=True)

        # --- OMS ---
        with sub_tab3:
            st.markdown('<div class="stl a">Analyse OMS — OT contenant « OMS » dans la désignation</div>',unsafe_allow_html=True)
            df_oms = kpis['dfp'][kpis['dfp']["Désignation"].astype(str).str.contains("OMS",case=False,na=False)].copy() if "Désignation" in kpis['dfp'].columns else pd.DataFrame()
            oms_data = []
            for p in posts:
                sub = df_oms[df_oms["Poste travail princ."]==p]
                for statut in sub["Statut OT"].unique():
                    cnt = len(sub[sub["Statut OT"]==statut])
                    oms_data.append({"Poste de travail":p,"Statut OT":statut,"Nombre d'OT OMS":cnt})
            if oms_data:
                oms_df = pd.DataFrame(oms_data)
                oms_html = '<table class="tw at"><thead><tr><th>Poste de travail</th><th>Statut OT</th><th>Nombre d\'OT OMS</th></tr></thead><tbody>'
                for _,r in oms_df.iterrows():
                    oms_html += '<tr><td>%s</td><td>%s</td><td>%d</td></tr>'%(r["Poste de travail"],r["Statut OT"],r["Nombre d'OT OMS"])
                oms_html += '<tr class="tr"><td colspan="2">Total</td><td>%d</td></tr></tbody></table>'%oms_df["Nombre d'OT OMS"].sum()
                st.markdown(oms_html,unsafe_allow_html=True)

                statut_counts = df_oms.groupby("Statut OT")["Ordre"].count().to_dict()
                if statut_counts:
                    fig_oms = make_pie_of_pie(list(statut_counts.keys()),list(statut_counts.values()),
                                              "Répartition des OT OMS par Statut", "Total OMS")
                    st.plotly_chart(fig_oms, use_container_width=True)
            else:
                st.markdown('<div style="padding:20px;color:#718096;text-align:center">Aucun OT contenant « OMS » trouvé.</div>',unsafe_allow_html=True)

        # --- Thermographie ---
        with sub_tab4:
            st.markdown('<div class="stl s">Analyse Thermographie — OT liés à la Thermographie</div>',unsafe_allow_html=True)
            thermo_kw = ["THERMOGRAPHIE","THERMO","THERM"]
            mask = kpis['dfp']["Désignation"].astype(str).apply(lambda x: any(kw in x.upper() for kw in thermo_kw)) if "Désignation" in kpis['dfp'].columns else pd.Series([False]*len(kpis['dfp']))
            df_thermo = kpis['dfp'][mask].copy()
            thermo_data = []
            for p in posts:
                sub = df_thermo[df_thermo["Poste travail princ."]==p]
                for statut in sub["Statut OT"].unique():
                    cnt = len(sub[sub["Statut OT"]==statut])
                    thermo_data.append({"Poste de travail":p,"Statut OT":statut,"Nombre d'OT Thermographie":cnt})
            if thermo_data:
                th_df = pd.DataFrame(thermo_data)
                th_html = '<table class="tw st"><thead><tr><th>Poste de travail</th><th>Statut OT</th><th>Nombre d\'OT Thermographie</th></tr></thead><tbody>'
                for _,r in th_df.iterrows():
                    th_html += '<tr><td>%s</td><td>%s</td><td>%d</td></tr>'%(r["Poste de travail"],r["Statut OT"],r["Nombre d'OT Thermographie"])
                th_html += '<tr class="tr"><td colspan="2">Total</td><td>%d</td></tr></tbody></table>'%th_df["Nombre d'OT Thermographie"].sum()
                st.markdown(th_html,unsafe_allow_html=True)

                statut_counts = df_thermo.groupby("Statut OT")["Ordre"].count().to_dict()
                if statut_counts:
                    fig_th = make_pie_of_pie(list(statut_counts.keys()),list(statut_counts.values()),
                                             "Répartition des OT Thermographie par Statut", "Total Thermo")
                    st.plotly_chart(fig_th, use_container_width=True)
            else:
                st.markdown('<div style="padding:20px;color:#718096;text-align:center">Aucun OT lié à la Thermographie trouvé.</div>',unsafe_allow_html=True)

    # ============================================================
    # TAB 4 : SUIVI DES AMÉLIORATIONS
    # ============================================================
    with tab4:
        hist_path = os.path.join("kpis","indicateurs_kpis.xlsx")
        hist_df = load_historical_kpis(hist_path)

        if hist_df.empty:
            st.markdown('<div class="explain-box"><h3>📭 Aucun historique disponible</h3>'
                       '<p>Le système enregistre automatiquement un instantané des KPI à chaque nouvelle date détectée dans <code>date.txt</code>. '
                       'Modifiez la date du fichier pour créer un nouvel enregistrement, puis revenez consulter cette page.</p></div>',unsafe_allow_html=True)
        else:
            st.markdown('<div class="stl c">Historique des Enregistrements</div>',unsafe_allow_html=True)
            dates_dispo = sorted(hist_df["Date"].unique())
            st.markdown('<div style="font-size:13px;color:#4a5568;margin-bottom:8px">📅 %d dates d\'enregistrement disponibles : %s</div>'%(len(dates_dispo),", ".join(dates_dispo)),unsafe_allow_html=True)

            var_df = calculate_variations(hist_df)

            if var_df.empty:
                st.markdown('<div style="padding:20px;color:#718096;text-align:center">Pas assez de dates pour calculer des variations (minimum 2 requises).</div>',unsafe_allow_html=True)
            else:
                # Filtres période
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    sel_periode = st.selectbox("Période d'analyse", [f"{dates_dispo[i]} → {dates_dispo[i+1]}" for i in range(len(dates_dispo)-1)], index=len(dates_dispo)-2 if len(dates_dispo)>=2 else 0)
                pv_date = sel_periode.split(" → ")[0]; cv_date = sel_periode.split(" → ")[1]
                filt_var = var_df[(var_df["Date précédente"]==pv_date)&(var_df["Date actuelle"]==cv_date)].copy()

                # --- Variations ---
                st.markdown('<div class="stl p">Analyse des Variations (%s → %s)</div>'%(pv_date,cv_date),unsafe_allow_html=True)
                if not filt_var.empty:
                    var_html = '<table class="tw pt"><thead><tr><th>Poste</th><th>Type</th><th>KPI</th><th>Valeur préc.</th><th>Valeur act.</th><th>Écart</th><th>Écart %</th><th>Tendance</th></tr></thead><tbody>'
                    for _,r in filt_var.iterrows():
                        trend_icon = "▲" if r["Tendance"]=="hausse" else ("▼" if r["Tendance"]=="baisse" else "●")
                        trend_clr = "#276749" if r["Tendance"]=="hausse" else ("#c53030" if r["Tendance"]=="baisse" else "#718096")
                        var_html += '<tr><td>%s</td><td>%s</td><td>%s</td><td>%.1f</td><td>%.1f</td><td>%+.1f</td><td>%+.1f%%</td><td style="color:%s;font-weight:800">%s %s</td></tr>'%(
                            r["Poste"],r["Type"],r["KPI"],r["Valeur précédente"],r["Valeur actuelle"],r["Écart"],r["Écart %"],trend_clr,trend_icon,r["Tendance"])
                    var_html += '</tbody></table>'
                    st.markdown(var_html,unsafe_allow_html=True)

                    # --- Journal des évolutions ---
                    journal = generate_journal(var_df)
                    journal_filt = journal[(journal["Date précédente"]==pv_date)&(journal["Date actuelle"]==cv_date)] if not journal.empty else pd.DataFrame()
                    st.markdown('<div class="stl a">Journal des Évolutions Significatives (|Écart| ≥ 5%)</div>',unsafe_allow_html=True)
                    if not journal_filt.empty:
                        j_html = '<table class="tw at"><thead><tr><th>Date</th><th>Poste</th><th>Type</th><th>KPI</th><th>Écart %</th><th>Sens</th></tr></thead><tbody>'
                        for _,r in journal_filt.iterrows():
                            sens_clr = "#276749" if r["Sens"]=="Amélioration" else "#c53030"
                            sens_ico = "⬆" if r["Sens"]=="Amélioration" else "⬇"
                            j_html += '<tr><td>%s → %s</td><td>%s</td><td>%s</td><td>%s</td><td style="font-weight:800">%+.1f%%</td><td style="color:%s;font-weight:800">%s %s</td></td>'%(
                                r["Date précédente"],r["Date actuelle"],r["Poste"],r["Type"],r["KPI"],r["Écart %"],sens_clr,sens_ico,r["Sens"])
                        j_html += '</tbody></table>'
                        st.markdown(j_html,unsafe_allow_html=True)
                    else:
                        st.markdown('<div style="padding:10px;color:#276749;font-weight:700">✅ Aucune variation significative sur cette période.</div>',unsafe_allow_html=True)

                    # --- Top 5 Améliorations / Dégradations ---
                    top5, bot5 = calculate_rankings(filt_var)

                    col_t1, col_t2 = st.columns(2)
                    with col_t1:
                        st.markdown('<div class="rank-card"><div class="rank-title" style="color:#276749">🚀 Top 5 Améliorations</div>',unsafe_allow_html=True)
                        if not top5.empty:
                            for i,(_,r) in enumerate(top5.iterrows()):
                                clr = "#276749" if r["Score variation"]>0 else "#718096"
                                st.markdown('<div class="rank-row"><div class="rank-num" style="background:%s">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:%s">%+.1f pts</div></div>'%(clr,i+1,r["Poste"],clr,r["Score variation"]),unsafe_allow_html=True)
                        else:
                            st.markdown('<div style="padding:10px;color:#718096">Aucune amélioration significative</div>',unsafe_allow_html=True)
                        st.markdown('</div>',unsafe_allow_html=True)

                    with col_t2:
                        st.markdown('<div class="rank-card"><div class="rank-title" style="color:#c53030">⚠️ Top 5 Dégradations</div>',unsafe_allow_html=True)
                        if not bot5.empty:
                            for i,(_,r) in enumerate(bot5.iterrows()):
                                clr = "#c53030" if r["Score variation"]<0 else "#718096"
                                st.markdown('<div class="rank-row"><div class="rank-num" style="background:%s">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:%s">%+.1f pts</div></div>'%(clr,i+1,r["Poste"],clr,r["Score variation"]),unsafe_allow_html=True)
                        else:
                            st.markdown('<div style="padding:10px;color:#718096">Aucune dégradation significative</div>',unsafe_allow_html=True)
                        st.markdown('</div>',unsafe_allow_html=True)

                    # --- Sparklines / Courbes d'évolution ---
                    st.markdown('<div class="stl q">Courbes d\'Évolution par Poste et KPI</div>',unsafe_allow_html=True)
                    # Sélection poste
                    sel_poste = st.selectbox("Sélectionner un poste de travail", posts, key="sp_evolution")
                    sel_type = st.selectbox("Type d'indicateur", ["Performance","Qualité"], key="st_evolution")
                    kpi_ev_list = QK if sel_type=="Performance" else PK

                    for k in kpi_ev_list:
                        vals = []; dt_labels = []
                        for d in dates_dispo:
                            sub = hist_df[(hist_df["Date"]==d)&(hist_df["_section"]=="perf" if sel_type=="Performance" else "qual")]
                            if "Poste de travail" in sub.columns and sel_poste in sub["Poste de travail"].values and k in sub.columns:
                                try: v=float(sub[sub["Poste de travail"]==sel_poste][k].values[0])
                                except Exception: v=None
                            else: v=None
                            vals.append(v); dt_labels.append(d)
                        # Filtrer None
                        clean_v = [(d,v) for d,v in zip(dt_labels,vals) if v is not None]
                        if len(clean_v)>=2:
                            fig_ev = go.Figure()
                            fig_ev.add_trace(go.Scatter(x=[c[0] for c in clean_v], y=[c[1] for c in clean_v],
                                                        mode='lines+markers+text', text=[f"{c[1]:.1f}%" for c in clean_v],
                                                        textposition='top center', textfont=dict(size=10),
                                                        line=dict(color="#3182CE" if sel_type=="Qualité" else "#38a169", width=2.5),
                                                        marker=dict(size=8)))
                            tv = CIBLE.get(k,100)
                            fig_ev.add_hline(y=tv, line_dash="dash", line_color="#e53e3e", annotation_text=f"Cible: {tv}%", annotation_position="top right")
                            fig_ev.update_layout(title=dict(text=k, font=dict(size=13, color="#1e3a5f")), height=250,
                                                 margin=dict(t=40,b=30,l=50,r=30), xaxis=dict(tickfont=dict(size=9)), yaxis=dict(tickfont=dict(size=9)))
                            with st.expander(f"📈 {k}", expanded=False):
                                st.plotly_chart(fig_ev, use_container_width=True)

if __name__=="__main__":
    main()
