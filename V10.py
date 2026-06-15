# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os, json
from datetime import datetime
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
    .sr .stg{font-size:11px;color:#718096;min-width:60px;text-align:center;white-space:nowrap}
    .sr .sb{font-size:11px;font-weight:700;padding:2px 8px;border-radius:3px;white-space:nowrap}
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
    .anl-tbl{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:13px;margin:0}
    .anl-tbl thead th{background:var(--p);color:#fff;font-weight:700;font-size:12px;padding:6px 8px;border:none;white-space:nowrap;position:sticky;top:0}
    .anl-tbl tbody td{padding:5px 8px;border-bottom:1px solid #edf2f7}
    .anl-tbl tbody tr:nth-child(even) td{background:#f7fafc}
    .anl-tbl tbody tr:hover td{background:#ebf8ff!important}
    .anl-tbl .tot td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important}
    .g-green{background:#c6efce;color:#006100;font-weight:600}
    .g-yellow{background:#ffeb9c;color:#9c6500;font-weight:600}
    .g-red{background:#ffc7ce;color:#9c0006;font-weight:600}
    .trend-up{color:#276749;font-weight:800;font-size:16px}
    .trend-down{color:#c53030;font-weight:800;font-size:16px}
    .trend-stable{color:#718096;font-weight:800;font-size:16px}
    .spark-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:8px}
    .spark-card{background:#fff;border-radius:var(--r);padding:10px 12px;border:1px solid var(--b);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .spark-card .sp-title{font-size:13px;font-weight:800;color:var(--p);margin-bottom:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .spark-card .sp-sub{font-size:11px;color:#718096;margin-bottom:5px}
    .rank-card{background:#fff;border-radius:var(--r);padding:12px 16px;border:1px solid var(--b);box-shadow:0 2px 8px rgba(0,0,0,.04)}
    .rank-card .rank-title{font-size:15px;font-weight:800;margin-bottom:8px;padding-bottom:5px;border-bottom:2px solid var(--b)}
    .rank-row{display:flex;align-items:center;padding:5px 0;font-size:13px;border-bottom:1px solid #f7fafc}
    .rank-row:last-child{border:none}
    .rank-row .rank-num{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:13px;color:#fff;margin-right:10px;flex-shrink:0}
    .rank-row .rank-name{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .rank-row .rank-score{font-weight:900;min-width:70px;text-align:right}
    .synth-box{background:linear-gradient(135deg,#f7fafc,#edf2f7);border-radius:var(--r);padding:14px 18px;margin-top:8px;border:1px solid #cbd5e0}
    .synth-box .synth-title{font-size:16px;font-weight:800;color:var(--p);margin-bottom:8px;display:flex;align-items:center;gap:8px}
    .synth-box .synth-title .ico{font-size:20px}
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
    .hbar-chart{background:#fff;border-radius:var(--r);padding:14px 18px;border:1px solid var(--b);box-shadow:0 2px 8px rgba(0,0,0,.04);margin-top:8px}
    .hbar-chart .hbar-title{font-size:15px;font-weight:800;color:var(--p);margin-bottom:4px;padding-bottom:6px;border-bottom:2px solid var(--b)}
    .hbar-chart .hbar-sub{font-size:11px;color:#718096;margin-bottom:10px}
    .hbar-row{display:flex;align-items:center;padding:4px 0;gap:8px}
    .hbar-row:last-child{border:none}
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
    /* ===== TOGGLE BOUTON INDICATEURS ===== */
    .toggle-container{display:flex;align-items:center;gap:12px;margin:8px 0}
    .toggle-btn{padding:8px 22px;border-radius:20px;font-weight:800;font-size:14px;cursor:pointer;border:2px solid var(--b);background:#fff;color:#718096;transition:all .2s}
    .toggle-btn.active-perf{background:linear-gradient(135deg,#276749,#48bb78);color:#fff;border-color:#276749;box-shadow:0 3px 10px rgba(39,103,73,.3)}
    .toggle-btn.active-qual{background:linear-gradient(135deg,#2b6cb0,#4299e1);color:#fff;border-color:#2b6cb0;box-shadow:0 3px 10px rgba(43,108,176,.3)}
    .toggle-btn:hover{transform:translateY(-1px)}
    /* ===== METHODOLOGIE CLASSEMENT ===== */
    .methodo-box{background:#fff;border-radius:var(--r);padding:18px 22px;border:1px solid var(--b);box-shadow:0 2px 8px rgba(0,0,0,.04);margin-top:8px}
    .methodo-box .methodo-title{font-size:17px;font-weight:900;color:var(--p);margin-bottom:12px;padding-bottom:8px;border-bottom:3px solid var(--p);display:flex;align-items:center;gap:8px}
    .methodo-section{margin-bottom:14px}
    .methodo-section:last-child{margin-bottom:0}
    .methodo-section h4{font-size:14px;font-weight:800;color:var(--pl);margin-bottom:6px;display:flex;align-items:center;gap:6px}
    .methodo-section p,.methodo-section li{font-size:12px;color:#4a5568;line-height:1.7}
    .methodo-section ul{padding-left:18px;margin:4px 0}
    .methodo-section li{margin-bottom:3px}
    .methodo-section .formula{background:#f7fafc;border:1px solid #e2e8f0;border-radius:6px;padding:8px 14px;font-family:monospace;font-size:12px;color:#2d3748;margin:6px 0;font-weight:600}
    .methodo-section .weight-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:6px;margin:6px 0}
    .methodo-section .weight-item{display:flex;align-items:center;gap:8px;background:#f7fafc;padding:5px 10px;border-radius:5px;font-size:12px}
    .methodo-section .weight-item .wk{flex:1;font-weight:600;color:#2d3748}
    .methodo-section .weight-item .wv{font-weight:900;color:var(--p);min-width:40px;text-align:right}
    .methodo-section .tie-box{background:linear-gradient(135deg,#ffeb9c,#fefcbf);border:1px solid #d69e2e;border-radius:6px;padding:8px 14px;font-size:12px;color:#975a16;font-weight:600;margin-top:6px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:17px}.cg,.dgrid,.synth-resume{grid-template-columns:1fr}.car .cal{width:120px}.gbr-l{width:100px}.spark-grid{grid-template-columns:1fr}.hbar-label{width:140px}.synth-kpi-name{width:140px}.action-item .act-kpi{min-width:140px}.methodo-section .weight-grid{grid-template-columns:1fr}}
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
        return "Centrale"  # *** MODIFIE : "Autre" -> "Centrale" ***
    def get_division(p):
        p=str(p).upper()
        if "SF1" in p: return "SF1"
        if "SF2" in p: return "SF2"
        return "Autre"

    def calc_kpis(df_i,av_i,now,posts):
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
            return "background:#c6efce;color:#006100;font-weight:600" if val<=5 else "background:#ffc7ce;color:#9c0006;font-weight:600")
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

    def html_synthese_actions(kpi_list, actuals, targets, act_map, section_type):
        color_main = "#276749" if section_type == "perf" else "#2b6cb0"
        label = "PERFORMANCE" if section_type == "perf" else "QUALITE"
        icon = "📊" if section_type == "perf" else "✅"
        nb_total = len(kpi_list)
        nb_ok = 0; nb_ko = 0; nb_warn = 0
        details_html = ""
        for k in kpi_list:
            av = actuals.get(k, 0)
            tv = targets.get(k, 100)
            diff = av - tv
            met = av <= tv if is_lb(k) else av >= tv
            if is_lb(k):
                warn_threshold = tv * 1.5
                is_warn = (not met) and (av <= warn_threshold)
            else:
                warn_threshold = tv - 5
                is_warn = (not met) and (av >= warn_threshold)
            if met:
                nb_ok += 1; badge_class = "badge-ok"; badge_txt = "ATTEINT"
                val_style = "background:#c6efce;color:#006100"
            elif is_warn:
                nb_warn += 1; badge_class = "badge-warn"; badge_txt = "ATTENTION"
                val_style = "background:#ffeb9c;color:#9c6500"
            else:
                nb_ko += 1; badge_class = "badge-ko"; badge_txt = "NON ATTEINT"
                val_style = "background:#ffc7ce;color:#9c0006"
            ec_color = "#276749" if met else ("#9c6500" if is_warn else "#c53030")
            action = "✅ Objectif atteint — maintenir" if met else act_map.get(k, "A definir")
            details_html += '<div class="synth-kpi-row">'
            details_html += '<div class="synth-kpi-name">%s</div>' % k
            details_html += '<div class="synth-kpi-val" style="%s">%.1f%%</div>' % (val_style, av)
            details_html += '<div class="synth-kpi-target">Cible: %.0f%%</div>' % tv
            details_html += '<div class="synth-kpi-ecart" style="color:%s">%+.1f%%</div>' % (ec_color, diff)
            details_html += '<div class="synth-kpi-badge %s">%s</div>' % (badge_class, badge_txt)
            details_html += '<div class="synth-kpi-action">%s</div>' % action
            details_html += '</div>'
        actions_html = ""
        ko_kpis = []; warn_kpis = []
        for k in kpi_list:
            av = actuals.get(k, 0); tv = targets.get(k, 100)
            diff = av - tv; met = av <= tv if is_lb(k) else av >= tv
            if not met:
                if is_lb(k): is_warn = av <= tv * 1.5
                else: is_warn = av >= tv - 5
                if is_warn: warn_kpis.append((k, diff))
                else: ko_kpis.append((k, diff))
        if ko_kpis or warn_kpis:
            actions_html = '<div class="action-list">'
            actions_html += '<div style="font-size:13px;font-weight:800;color:#c53030;margin-bottom:4px">🔴 Actions Correctives Prioritaires</div>'
            for k, diff in ko_kpis:
                actions_html += '<div class="action-item act-ko"><div class="act-kpi">%s</div><div class="act-ecart" style="color:#c53030">%+.1f%%</div><div class="act-text">%s</div></div>' % (k, diff, act_map.get(k, ""))
            if warn_kpis:
                actions_html += '<div style="font-size:13px;font-weight:800;color:#d69e2e;margin:6px 0 4px 0">🟡 Actions d\'Amelioration</div>'
                for k, diff in warn_kpis:
                    actions_html += '<div class="action-item act-warn"><div class="act-kpi">%s</div><div class="act-ecart" style="color:#9c6500">%+.1f%%</div><div class="act-text">%s</div></div>' % (k, diff, act_map.get(k, ""))
            actions_html += '</div>'
        h = '<div class="synth-box">'
        h += '<div class="synth-title"><span class="ico">%s</span> Synthese & Actions — Indicateurs de %s</div>' % (icon, label)
        h += '<div class="synth-resume">'
        h += '<div class="synth-resume-card"><div class="src-val" style="color:#276749">%d</div><div class="src-lbl">Atteints</div></div>' % nb_ok
        h += '<div class="synth-resume-card"><div class="src-val" style="color:#d69e2e">%d</div><div class="src-lbl">Attention</div></div>' % nb_warn
        h += '<div class="synth-resume-card"><div class="src-val" style="color:#c53030">%d</div><div class="src-lbl">Non Atteints</div></div>' % nb_ko
        h += '</div>'
        h += details_html
        if actions_html: h += actions_html
        h += '</div>'
        return h

    def html_hbar_chart(kpi_list, actuals, targets, total_general, section_type, title):
        color_main = "#276749" if section_type == "perf" else "#2b6cb0"
        label = "Performance" if section_type == "perf" else "Qualite"
        h = '<div class="hbar-chart">'
        h += '<div class="hbar-title">%s</div>' % title
        h += '<div class="hbar-sub">Comparaison de chaque KPI par rapport au Total General (%.1f%%) — %s</div>' % (total_general, label)
        all_vals = [actuals.get(k, 0) for k in kpi_list]
        all_vals.append(total_general)
        all_vals.extend([targets.get(k, 100) for k in kpi_list])
        max_val = max(all_vals) * 1.15
        if max_val == 0: max_val = 100
        for k in kpi_list:
            av = actuals.get(k, 0)
            tv = targets.get(k, 100)
            met = av <= tv if is_lb(k) else av >= tv
            if met:
                fill_class = "ok"
            else:
                if is_lb(k):
                    fill_class = "warn" if av <= tv * 1.5 else "ko"
                else:
                    fill_class = "warn" if av >= tv - 5 else "ko"
            bar_width = max((av / max_val) * 100, 0.5)
            target_pos = min((tv / max_val) * 100, 100)
            total_pos = min((total_general / max_val) * 100, 100)
            h += '<div class="hbar-row">'
            h += '<div class="hbar-label">%s</div>' % k
            h += '<div class="hbar-track">'
            h += '<div class="hbar-fill %s" style="width:%.1f%%">' % (fill_class, bar_width)
            h += '<div class="hbar-val">%.1f%%</div>' % av
            h += '</div>'
            h += '<div class="hbar-target" style="left:%.1f%%"><div class="hbar-target-label">C:%.0f</div></div>' % (target_pos, tv)
            h += '<div style="position:absolute;left:%.1f%%;top:50%%;transform:translateY(-50%%);width:2px;height:26px;border-left:2px dashed %s;z-index:1;opacity:0.7"></div>' % (total_pos, color_main)
            h += '</div></div>'
        h += '<div class="hbar-legend">'
        h += '<span><span class="lg-swatch" style="background:linear-gradient(90deg,#276749,#48bb78)"></span> Atteint</span>'
        h += '<span><span class="lg-swatch" style="background:linear-gradient(90deg,#d69e2e,#f6e05e)"></span> Attention</span>'
        h += '<span><span class="lg-swatch" style="background:linear-gradient(90deg,#c53030,#fc8181)"></span> Non Atteint</span>'
        h += '<span><span class="lg-swatch" style="width:2px;height:14px;background:#e53e3e;border-radius:0"></span> Cible</span>'
        h += '<span><span class="lg-swatch" style="width:2px;height:14px;border-left:2px dashed %s;background:transparent;border-radius:0"></span> Total General</span>' % color_main
        h += '</div></div>'
        return h

    # ===== NOUVELLE FONCTION : Methodologie Classement HTML =====
    def html_methodo_classement():
        h = '<div class="methodo-box">'
        h += '<div class="methodo-title">📋 Methodologie de Calcul du Classement</div>'

        # 1. Classement Performance
        h += '<div class="methodo-section">'
        h += '<h4>🟢 1. Classement Performance</h4>'
        h += '<p>Le classement Performance est base sur les <strong>10 indicateurs de Performance</strong> suivants :</p>'
        h += '<div class="weight-grid">'
        perf_weights = [
            ("TAUX_REALISATION_CORRECTIF/PT", "Pondere x2"),
            ("OT préparation <1 mois", "Pondere x1"),
            ("OT préparation >3 mois", "Pondere x1"),
            ("OT préparation 1mois< <3mois", "Pondere x1"),
            ("OT planification <1 mois", "Pondere x1"),
            ("OT planification >3 mois", "Pondere x1"),
            ("OT planification 1mois< <3mois", "Pondere x1"),
            ("OT exécution <1 mois", "Pondere x1"),
            ("OT exécution >3 mois", "Pondere x1"),
            ("OT exécution 1mois< <3mois", "Pondere x1"),
        ]
        for kw, w in perf_weights:
            h += '<div class="weight-item"><span class="wk">%s</span><span class="wv">%s</span></div>' % (kw, w)
        h += '</div>'
        h += '<div class="formula">Score Perf = Σ (Score_KPI × Poids_KPI) / Σ(Poids_KPI) × 100</div>'
        h += '<p>ou Score_KPI = 1 si l\'indicateur atteint son seuil, 0 sinon. Les seuils sont : Taux realisation ≥ 80%%, OT &lt;1 mois ≥ 75%%, OT &gt;3 mois ≤ 5%%, OT 1-3 mois ≤ 15%%.</p>'
        h += '</div>'

        # 2. Classement Qualité
        h += '<div class="methodo-section">'
        h += '<h4>🔵 2. Classement Qualite</h4>'
        h += '<p>Le classement Qualite est base sur les <strong>6 indicateurs de Qualite</strong> suivants :</p>'
        h += '<div class="weight-grid">'
        qual_weights = [
            ("appel avis approuvé", "Pondere x1"),
            ("OT LANC ESTIME", "Pondere x1"),
            ("Backlog préparation caractérisé", "Pondere x1"),
            ("Backlog planification caractérisé", "Pondere x1"),
            ("OT CONFIME", "Pondere x1"),
            ("OT_COR_EGAL", "Pondere x1"),
        ]
        for kw, w in qual_weights:
            h += '<div class="weight-item"><span class="wk">%s</span><span class="wv">%s</span></div>' % (kw, w)
        h += '</div>'
        h += '<div class="formula">Score Qual = Σ (Score_KPI × Poids_KPI) / Σ(Poids_KPI) × 100</div>'
        h += '<p>ou Score_KPI = 1 si l\'indicateur atteint son seuil, 0 sinon. Les seuils sont : Appel avis ≥ 90%%, OT LANC ESTIME / Backlog / OT CONFIME / OT_COR_EGAL ≥ 95%%.</p>'
        h += '</div>'

        # 3. Pondérations
        h += '<div class="methodo-section">'
        h += '<h4>⚖️ 3. Pondérations Utilisees</h4>'
        h += '<p>Les pondérations reflètent l\'importance relative de chaque indicateur dans le classement global :</p>'
        h += '<ul>'
        h += '<li><strong>TAUX_REALISATION_CORRECTIF/PT</strong> : poids ×2 — indicateur cle de suivi de la realisation des OT correctifs.</li>'
        h += '<li><strong>Tous les autres indicateurs</strong> : poids ×1 — contribution égale au score global.</li>'
        h += '</ul>'
        h += '<p>Le score final est un pourcentage : <strong>100%%</strong> signifie que tous les KPI pondérés sont atteints, <strong>0%%</strong> signifie qu\'aucun n\'est atteint.</p>'
        h += '</div>'

        # 4. Critères de tri
        h += '<div class="methodo-section">'
        h += '<h4>🔀 4. Criteres de Tri</h4>'
        h += '<p>Le classement est établi selon les regles suivantes :</p>'
        h += '<ul>'
        h += '<li><strong>Tri principal</strong> : par score décroissant (le meilleur score en premier).</li>'
        h += '<li><strong>Classement Performance</strong> : trié sur le Score Performance calculé.</li>'
        h += '<li><strong>Classement Qualite</strong> : trié sur le Score Qualite calculé.</li>'
        h += '<li><strong>Top 5 / Flop 5</strong> : les 5 meilleurs et les 5 moins bons postes sont mis en évidence.</li>'
        h += '</ul>'
        h += '</div>'

        # 5. Gestion des ex æquo
        h += '<div class="methodo-section">'
        h += '<h4>🤝 5. Gestion des Ex Aequo</h4>'
        h += '<div class="tie-box">'
        h += '⚠️ En cas de scores identiques entre deux ou plusieurs postes, le classement est resolu par :<br>'
        h += '<strong>1.</strong> Le nombre total de KPI atteints (le plus grand nombre l\'emporte).<br>'
        h += '<strong>2.</strong> En cas d\'égalite persistante, le nombre d\'anomalies (le moins d\'anomalies l\'emporte).<br>'
        h += '<strong>3.</strong> En dernier recours, l\'ordre alphabétique du nom du poste est utilise comme critere de departage.'
        h += '</div>'
        h += '</div>'

        h += '</div>'
        return h

    # ===== NOUVELLE FONCTION : Classement HTML intégré =====
    def html_classement_section(posts, ckdf, perf_scores, qual_scores):
        h = ''
        # Calcul des scores par poste
        perf_rank = []
        qual_rank = []
        for p in posts:
            ps = 0; pp = 0; qs = 0; qp = 0
            for k in QK:
                w = 2 if k == "TAUX_REALISATION_CORRECTIF/PT" else 1
                ps += gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)) * w
                pp += w
            for k in PK:
                qs += gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)) * 1
                qp += 1
            perf_score = (ps / pp * 100) if pp > 0 else 0
            qual_score = (qs / qp * 100) if qp > 0 else 0
            perf_rank.append((p, perf_score, ps))
            qual_rank.append((p, qual_score, qs))

        # Tri avec gestion ex æquo
        def sort_key(item):
            return (-item[1], -item[2], item[0])
        perf_rank.sort(key=sort_key)
        qual_rank.sort(key=sort_key)

        # Top 5 et Flop 5 Performance
        top5_p = perf_rank[:5]
        flop5_p = perf_rank[-5:][::-1]
        top5_q = qual_rank[:5]
        flop5_q = qual_rank[-5:][::-1]

        rank_colors = ["#c53030","#d69e2e","#38a169","#3182ce","#805ad5"]

        # Top 5 Performance
        h += '<div class="dgrid">'
        h += '<div class="rank-card"><div class="rank-title" style="color:#276749;border-bottom-color:#276749">🏆 Top 5 Performance</div>'
        for i, (p, s, sc) in enumerate(top5_p):
            bg = rank_colors[i] if i < len(rank_colors) else "#718096"
            h += '<div class="rank-row"><div class="rank-num" style="background:%s">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:#276749">%.1f%%</div></div>' % (bg, i+1, p, s)
        h += '</div>'

        # Flop 5 Performance
        h += '<div class="rank-card"><div class="rank-title" style="color:#c53030;border-bottom-color:#c53030">⚠️ Flop 5 Performance</div>'
        for i, (p, s, sc) in enumerate(flop5_p):
            h += '<div class="rank-row"><div class="rank-num" style="background:#e53e3e">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:#c53030">%.1f%%</div></div>' % (i+1, p, s)
        h += '</div></div>'

        h += '<div style="height:8px"></div>'

        # Top 5 Qualité
        h += '<div class="dgrid">'
        h += '<div class="rank-card"><div class="rank-title" style="color:#2b6cb0;border-bottom-color:#2b6cb0">🏆 Top 5 Qualite</div>'
        for i, (p, s, sc) in enumerate(top5_q):
            bg = rank_colors[i] if i < len(rank_colors) else "#718096"
            h += '<div class="rank-row"><div class="rank-num" style="background:%s">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:#2b6cb0">%.1f%%</div></div>' % (bg, i+1, p, s)
        h += '</div>'

        # Flop 5 Qualité
        h += '<div class="rank-card"><div class="rank-title" style="color:#c53030;border-bottom-color:#c53030">⚠️ Flop 5 Qualite</div>'
        for i, (p, s, sc) in enumerate(flop5_q):
            h += '<div class="rank-row"><div class="rank-num" style="background:#e53e3e">%d</div><div class="rank-name">%s</div><div class="rank-score" style="color:#c53030">%.1f%%</div></div>' % (i+1, p, s)
        h += '</div></div>'

        return h

    # ============================================================
    # SIDEBAR
    # ============================================================
    with st.sidebar:
        st.markdown("<div style='padding:10px 0;text-align:center'><span style='font-size:28px'>⚙️</span><br><span style='font-weight:800;font-size:13px;text-transform:uppercase;letter-spacing:1px'>Filtres</span></div>",unsafe_allow_html=True)
        f_atelier = st.selectbox("Atelier", ["Tous","Sulfurique","Phosphorique","Engrais","Feed","Centrale"])  # *** MODIFIE : "Autre" -> "Centrale" ***
        f_metier = st.selectbox("Metier", ["Tous","Electrique","Mecanique","Instrumentation","Genie Civil","Autre"])
        f_division = st.selectbox("Division", ["Tous","SF1","SF2","Autre"])
        f_postes = st.multiselect("Postes de travail", [])
        uploaded_ot = st.file_uploader("Fichier OT (Excel)", type=["xlsx","xls"], key="ot_up")
        uploaded_av = st.file_uploader("Fichier Avis (Excel)", type=["xlsx","xls"], key="av_up")
        st.markdown("---")
        if st.button("📊 Telecharger KPIs", use_container_width=True):
            st.toast("Fichier genere dans le dossier kpis/")

    # ============================================================
    # LECTURE FICHIERS
    # ============================================================
    df_raw = pd.DataFrame(); av_raw = pd.DataFrame()
    if uploaded_ot is not None:
        try:
            df_raw = pd.read_excel(uploaded_ot)
            if "Poste travail princ." not in df_raw.columns:
                st.error("Colonne 'Poste travail princ.' introuvable dans le fichier OT."); st.stop()
        except Exception as e:
            st.error(f"Erreur lecture OT: {e}"); st.stop()
    if uploaded_av is not None:
        try:
            av_raw = pd.read_excel(uploaded_av)
            if "Poste travail princ." not in av_raw.columns:
                st.error("Colonne 'Poste travail princ.' introuvable dans le fichier Avis."); st.stop()
        except Exception as e:
            st.error(f"Erreur lecture Avis: {e}"); st.stop()

    if df_raw.empty:
        st.markdown('<div class="es">📁 Veuillez charger les fichiers Excel (OT et Avis) depuis le panneau de filtres a gauche.</div>',unsafe_allow_html=True)
        st.stop()

    df_filt = excr(df_raw.copy())
    av_filt = av_raw.copy() if not av_raw.empty else pd.DataFrame(columns=["Poste travail princ.","Statut utilisateur","Avis","Ordre"])

    # Application des filtres
    if f_atelier != "Tous":
        mask = df_filt["Poste travail princ."].apply(lambda p: get_atelier(p) == f_atelier)
        df_filt = df_filt[mask]
        if not av_filt.empty:
            av_filt = av_filt[av_filt["Poste travail princ."].apply(lambda p: get_atelier(p) == f_atelier)]
    if f_metier != "Tous":
        mask = df_filt["Poste travail princ."].apply(lambda p: get_metier(p) == f_metier)
        df_filt = df_filt[mask]
        if not av_filt.empty:
            av_filt = av_filt[av_filt["Poste travail princ."].apply(lambda p: get_metier(p) == f_metier)]
    if f_division != "Tous":
        mask = df_filt["Poste travail princ."].apply(lambda p: get_division(p) == f_division)
        df_filt = df_filt[mask]
        if not av_filt.empty:
            av_filt = av_filt[av_filt["Poste travail princ."].apply(lambda p: get_division(p) == f_division)]
    if f_postes:
        df_filt = df_filt[df_filt["Poste travail princ."].isin(f_postes)]
        if not av_filt.empty:
            av_filt = av_filt[av_filt["Poste travail princ."].isin(f_postes)]

    posts = sorted(df_filt["Poste travail princ."].dropna().unique().tolist())
    if not posts:
        st.markdown('<div class="es">🔍 Aucun poste de travail ne correspond aux filtres selectionnes.</div>',unsafe_allow_html=True)
        st.stop()

    now = datetime.now()
    kr = calc_kpis(df_filt, av_filt, now, posts)
    ckdf = kr['ckdf']

    # ============================================================
    # CALCULS DERIVES
    # ============================================================
    # Scores par poste
    perf_scores = {}; qual_scores = {}
    for p in posts:
        ps = 0; pp = 0; qs = 0; qp = 0
        for k in QK:
            w = 2 if k == "TAUX_REALISATION_CORRECTIF/PT" else 1
            ps += gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)) * w
            pp += w
        for k in PK:
            qs += gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)) * 1
            qp += 1
        perf_scores[p] = (ps / pp * 100) if pp > 0 else 0
        qual_scores[p] = (qs / qp * 100) if qp > 0 else 0

    avg_perf = np.mean(list(perf_scores.values())) if perf_scores else 0
    avg_qual = np.mean(list(qual_scores.values())) if qual_scores else 0
    tot_ot = len(df_filt)
    tot_anomalies_p = sum(1 for p in posts for k in QK if not gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)))
    tot_anomalies_q = sum(1 for p in posts for k in PK if not gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)))

    # Totaux généraux pour les KPI
    tg_perf = {}
    for k in QK:
        vals = [ckdf.loc[p, k] for p in posts if p in ckdf.index]
        tg_perf[k] = np.mean(vals) if vals else 0
    tg_qual = {}
    for k in PK:
        vals = [ckdf.loc[p, k] for p in posts if p in ckdf.index]
        tg_qual[k] = np.mean(vals) if vals else 0

    # Anomalies par poste
    ano_perf = {}
    ano_qual = {}
    for p in posts:
        ano_perf[p] = sum(1 for k in QK if not gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)))
        ano_qual[p] = sum(1 for k in PK if not gscore(k, ckdf.loc[p, k] if p in ckdf.index else 0, CIBLE.get(k, 100)))

    # ============================================================
    # HEADER
    # ============================================================
    st.markdown('<div class="mh"><h1>📊 Dashboard KPI — Suivi Maintenance</h1><span class="db">📅 %s</span></div>' % fichier_date, unsafe_allow_html=True)

    # CARTES RESUME
    st.markdown('<div class="cr">', unsafe_allow_html=True)
    st.markdown('<div class="cc c1"><div class="cv">%d</div><div class="cl">Total OT</div></div>' % tot_ot, unsafe_allow_html=True)
    st.markdown('<div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Perf. Moyen</div></div>' % avg_perf, unsafe_allow_html=True)
    st.markdown('<div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Qual. Moyen</div></div>' % avg_qual, unsafe_allow_html=True)
    st.markdown('<div class="cc c4"><div class="cv">%d</div><div class="cl">Total Anomalies</div></div>' % (tot_anomalies_p + tot_anomalies_q), unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ============================================================
    # ONGLETS (sans "Classement" — integre dans "Analyse")
    # ============================================================
    tab1, tab2, tab3 = st.tabs(["🔍 Analyse / Présentation", "📰 Journal des Variations", "📈 Tendances"])

    # ============================================================
    # ONGLET 1 : ANALYSE / PRESENTATION (avec Classement intégré)
    # ============================================================
    with tab1:

        # ----- SECTION 1 : Graphiques Barres Horizontales (AVANT les tableaux) -----
        st.markdown('<div class="stl p">Graphiques de Comparaison</div>', unsafe_allow_html=True)
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown(html_hbar_chart(QK, tg_perf, CIBLE, avg_perf, "perf", "Performance vs Total General"), unsafe_allow_html=True)
        with col_g2:
            st.markdown(html_hbar_chart(PK, tg_qual, CIBLE, avg_qual, "qual", "Qualite vs Total General"), unsafe_allow_html=True)

        # ----- SECTION 2 : Bouton Toggle Indicateurs -----
        st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)
        if "toggle_indicateurs" not in st.session_state:
            st.session_state.toggle_indicateurs = "perf"

        col_t1, col_t2, col_t3 = st.columns([1, 2, 1])
        with col_t2:
            toggle_html = '<div class="toggle-container" style="justify-content:center">'
            perf_class = "toggle-btn active-perf" if st.session_state.toggle_indicateurs == "perf" else "toggle-btn"
            qual_class = "toggle-btn active-qual" if st.session_state.toggle_indicateurs == "qual" else "toggle-btn"
            toggle_html += '<button class="%s" onclick="document.getElementById(\'toggle_perf\').click()">📊 Performance (%%)</button>' % perf_class
            toggle_html += '<button class="%s" onclick="document.getElementById(\'toggle_qual\').click()">✅ Qualite (%%)</button>' % qual_class
            toggle_html += '</div>'
            st.markdown(toggle_html, unsafe_allow_html=True)
            c_tog1, c_tog2 = st.columns(2)
            with c_tog1:
                if st.button("📊 Performance (%)", use_container_width=True, key="toggle_perf"):
                    st.session_state.toggle_indicateurs = "perf"
                    st.rerun()
            with c_tog2:
                if st.button("✅ Qualite (%)", use_container_width=True, key="toggle_qual"):
                    st.session_state.toggle_indicateurs = "qual"
                    st.rerun()

        # ----- SECTION 3 : Tableaux d'indicateurs (selon toggle) -----
        st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)

        if st.session_state.toggle_indicateurs == "perf":
            # Tableau Performance
            st.markdown('<div class="stl p">Indicateurs de Performance (%%)</div>', unsafe_allow_html=True)
            p_cols = ["Poste de travail"] + QK + ["Score Performance"]
            p_rows = []
            cible_row = {"_t": "cible", "Poste de travail": "CIBLE"}
            for k in QK: cible_row[k] = CIBLE.get(k, "-")
            cible_row["Score Performance"] = "-"
            p_rows.append(cible_row)
            for p in posts:
                row = {"Poste de travail": p}
                for k in QK:
                    row[k] = round(ckdf.loc[p, k], 1) if p in ckdf.index else 0
                row["Score Performance"] = round(perf_scores.get(p, 0), 1)
                p_rows.append(row)
            # Ligne total
            total_row = {"_t": "total", "Poste de travail": "Total general"}
            for k in QK: total_row[k] = round(tg_perf.get(k, 0), 1)
            total_row["Score Performance"] = round(avg_perf, 1)
            p_rows.append(total_row)
            st.markdown(html_table(p_rows, p_cols, "pt", sc_col=set(QK + ["Score Performance"])), unsafe_allow_html=True)

            # Anomalies Performance
            st.markdown('<div class="stl a">Nombre d\'Anomalies Performance</div>', unsafe_allow_html=True)
            ano_cols = ["Poste de travail", "Nombre anomalies"]
            ano_rows = []
            for p in posts:
                ano_rows.append({"Poste de travail": p, "Nombre anomalies": ano_perf.get(p, 0)})
            ano_rows.append({"_t": "total", "Poste de travail": "Total", "Nombre anomalies": tot_anomalies_p})
            st.markdown(html_ano(ano_rows, ano_cols), unsafe_allow_html=True)

            # Synthèse Performance
            st.markdown(html_synthese_actions(QK, tg_perf, CIBLE, ACT_MAP, "perf"), unsafe_allow_html=True)

        else:
            # Tableau Qualité
            st.markdown('<div class="stl q">Indicateurs de Qualite (%%)</div>', unsafe_allow_html=True)
            q_cols = ["Poste de travail"] + PK + ["Score Qualite"]
            q_rows = []
            cible_row = {"_t": "cible", "Poste de travail": "CIBLE"}
            for k in PK: cible_row[k] = CIBLE.get(k, "-")
            cible_row["Score Qualite"] = "-"
            q_rows.append(cible_row)
            for p in posts:
                row = {"Poste de travail": p}
                for k in PK:
                    row[k] = round(ckdf.loc[p, k], 1) if p in ckdf.index else 0
                row["Score Qualite"] = round(qual_scores.get(p, 0), 1)
                q_rows.append(row)
            total_row = {"_t": "total", "Poste de travail": "Total general"}
            for k in PK: total_row[k] = round(tg_qual.get(k, 0), 1)
            total_row["Score Qualite"] = round(avg_qual, 1)
            q_rows.append(total_row)
            st.markdown(html_table(q_rows, q_cols, "qt", sc_col=set(PK + ["Score Qualite"])), unsafe_allow_html=True)

            # Anomalies Qualité
            st.markdown('<div class="stl a">Nombre d\'Anomalies Qualite</div>', unsafe_allow_html=True)
            ano_cols = ["Poste de travail", "Nombre anomalies"]
            ano_rows = []
            for p in posts:
                ano_rows.append({"Poste de travail": p, "Nombre anomalies": ano_qual.get(p, 0)})
            ano_rows.append({"_t": "total", "Poste de travail": "Total", "Nombre anomalies": tot_anomalies_q})
            st.markdown(html_ano(ano_rows, ano_cols), unsafe_allow_html=True)

            # Synthèse Qualité
            st.markdown(html_synthese_actions(PK, tg_qual, CIBLE, ACT_MAP, "qual"), unsafe_allow_html=True)

        # ----- SECTION 4 : CLASSEMENT INTÉGRÉ -----
        st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
        st.markdown('<div class="stl c">🏅 Classement des Postes de Travail</div>', unsafe_allow_html=True)
        st.markdown(html_classement_section(posts, ckdf, perf_scores, qual_scores), unsafe_allow_html=True)

        # ----- SECTION 5 : MÉTHODOLOGIE CLASSEMENT -----
        st.markdown(html_methodo_classement(), unsafe_allow_html=True)

    # ============================================================
    # ONGLET 2 : JOURNAL DES VARIATIONS
    # ============================================================
    with tab2:
        st.markdown('<div class="stl s">Journal des Variations Significatives</div>', unsafe_allow_html=True)
        kpis_path = os.path.join("kpis", "indicateurs_kpis.xlsx")
        if os.path.exists(kpis_path):
            hist_df = load_historical_kpis(kpis_path)
            if not hist_df.empty:
                var_df = calculate_variations(hist_df)
                if not var_df.empty:
                    journal_df = generate_journal(var_df)
                    if not journal_df.empty:
                        j_cols = ["Date precedente","Date actuelle","Poste","Type","KPI","Valeur precedente","Valeur actuelle","Ecart","Ecart %","Sens"]
                        j_rows = []
                        for _, r in journal_df.iterrows():
                            j_rows.append({c: r.get(c, "") for c in j_cols})
                        st.markdown('<div style="overflow-x:auto">' + html_table(j_rows, j_cols, "st") + '</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="es">Aucune variation significative detectee entre les periodes.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Pas assez de donnees historiques pour calculer les variations (minimum 2 periodes requises).</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="es">Aucune donnee historique exploitable trouvee dans le fichier KPIs.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Le fichier historique des KPIs n\'existe pas encore. Il sera cree apres la premiere generation de KPIs.</div>', unsafe_allow_html=True)

    # ============================================================
    # ONGLET 3 : TENDANCES
    # ============================================================
    with tab3:
        st.markdown('<div class="stl">Tendances par KPI et Poste</div>', unsafe_allow_html=True)
        kpis_path = os.path.join("kpis", "indicateurs_kpis.xlsx")
        if os.path.exists(kpis_path):
            hist_df = load_historical_kpis(kpis_path)
            if not hist_df.empty:
                dates_list = sorted(hist_df["Date"].unique())
                if len(dates_list) >= 2:
                    spark_html = '<div class="spark-grid">'
                    for p in posts:
                        p_data = hist_df[hist_df["Poste de travail"] == p]
                        if p_data.empty: continue
                        # Score Perf trend
                        perf_data = p_data[p_data["_section"] == "perf"]
                        qual_data = p_data[p_data["_section"] == "qual"]
                        spark_html += '<div class="spark-card">'
                        spark_html += '<div class="sp-title">📌 %s</div>' % p
                        if not perf_data.empty and "Score Performance" in perf_data.columns:
                            latest = perf_data.iloc[-1].get("Score Performance", "-")
                            spark_html += '<div class="sp-sub">Score Perf: <strong>%.1f%%</strong></div>' % (float(latest) if isinstance(latest, (int, float, np.floating)) else 0)
                        else:
                            spark_html += '<div class="sp-sub">Score Perf: N/A</div>'
                        if not qual_data.empty and "Score Qualite" in qual_data.columns:
                            latest = qual_data.iloc[-1].get("Score Qualite", "-")
                            spark_html += '<div class="sp-sub">Score Qual: <strong>%.1f%%</strong></div>' % (float(latest) if isinstance(latest, (int, float, np.floating)) else 0)
                        else:
                            spark_html += '<div class="sp-sub">Score Qual: N/A</div>'
                        # Mini trend
                        if not perf_data.empty and "Score Performance" in perf_data.columns:
                            scores = perf_data["Score Performance"].dropna().astype(float).tolist()
                            if len(scores) >= 2:
                                diff = scores[-1] - scores[-2]
                                if diff > 0.5: spark_html += '<span class="trend-up">▲ +%.1f</span>' % diff
                                elif diff < -0.5: spark_html += '<span class="trend-down">▼ %.1f</span>' % diff
                                else: spark_html += '<span class="trend-stable">— 0.0</span>'
                        spark_html += '</div>'
                    spark_html += '</div>'
                    st.markdown(spark_html, unsafe_allow_html=True)
                else:
                    st.markdown('<div class="es">Minimum 2 periodes requises pour afficher les tendances.</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="es">Aucune donnee historique disponible.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="es">Le fichier historique des KPIs n\'existe pas encore.</div>', unsafe_allow_html=True)

    # ============================================================
    # EXPORT EXCEL
    # ============================================================
    if st.session_state.get("telecharger_kpi", False) or False:
        p_cols_xl = ["Poste de travail"] + QK + ["Score Performance"]
        p_rows_xl = [{"Poste de travail": p} for p in posts]
        for i, p in enumerate(posts):
            for k in QK:
                p_rows_xl[i][k] = round(ckdf.loc[p, k], 1) if p in ckdf.index else 0
            p_rows_xl[i]["Score Performance"] = round(perf_scores.get(p, 0), 1)

        q_cols_xl = ["Poste de travail"] + PK + ["Score Qualite"]
        q_rows_xl = [{"Poste de travail": p} for p in posts]
        for i, p in enumerate(posts):
            for k in PK:
                q_rows_xl[i][k] = round(ckdf.loc[p, k], 1) if p in ckdf.index else 0
            q_rows_xl[i]["Score Qualite"] = round(qual_scores.get(p, 0), 1)

        ano_p_rows = [{"Poste de travail": p, "Nombre anomalies": ano_perf.get(p, 0)} for p in posts]
        ano_p_cols = ["Poste de travail", "Nombre anomalies"]
        ano_q_rows = [{"Poste de travail": p, "Nombre anomalies": ano_qual.get(p, 0)} for p in posts]
        ano_q_cols = ["Poste de travail", "Nombre anomalies"]

        save_kpis_to_excel(p_rows_xl, p_cols_xl, q_rows_xl, q_cols_xl,
                          ano_p_rows, ano_p_cols, ano_q_rows, ano_q_cols, fichier_date)
        st.toast("✅ Fichier KPIs genere avec succes !")

if __name__ == "__main__":
    main()
