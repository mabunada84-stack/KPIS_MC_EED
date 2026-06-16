# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os, json, hashlib
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

PIE_COLORS = [
    "#1e3a5f","#2b6cb0","#3182ce","#4299e1","#63b3ed",
    "#276749","#38a169","#48bb78","#68d391","#9ae6b4",
    "#805ad5","#9f7aea","#b794f4","#d6bcfa","#e9d8fd",
    "#c53030","#e53e3e","#fc8181","#feb2b2","#fed7d7",
    "#d69e2e","#ecc94b","#f6e05e","#faf089","#fefcbf",
    "#744210","#975a16","#b7791f","#d69e2e","#ed8936",
    "#2d3748","#4a5568","#718096","#a0aec0","#cbd5e0"
]

# ============================================================
# === LECTURE EXCEL ULTRA-ROBUSTE ===
# Détecte le vrai format binaire (magic bytes) au lieu de
# se fier à l'extension du fichier.
# ============================================================
def safe_read_excel(filepath_or_buffer, **kwargs):
    """
    Lit un fichier tabulaire quel que soit son vrai format :
    - XLSX (ZIP: PK\\x03\\x04)
    - XLS  (OLE: D0 CF 11 E0)
    - HTML table (export SAP, commence par <)
    - CSV  (texte avec séparateurs ; , \\t)
    - XML  spreadsheet
    """
    is_buffer = isinstance(filepath_or_buffer, (io.BytesIO, io.BufferedReader, io.BufferedIOBase))
    
    # Lire les premiers octets pour identifier le vrai format
    header_bytes = b""
    if is_buffer:
        pos = filepath_or_buffer.tell()
        try:
            header_bytes = filepath_or_buffer.read(512)
            filepath_or_buffer.seek(pos)
        except Exception:
            try:
                filepath_or_buffer.seek(0)
                header_bytes = filepath_or_buffer.read(512)
                filepath_or_buffer.seek(0)
            except Exception:
                header_bytes = b""
    elif isinstance(filepath_or_buffer, str) and os.path.exists(filepath_or_buffer):
        try:
            with open(filepath_or_buffer, "rb") as f:
                header_bytes = f.read(512)
        except Exception:
            header_bytes = b""

    # --- Détection par magic bytes ---
    # XLSX = fichier ZIP (commence par PK\x03\x04)
    if header_bytes[:4] == b'PK\x03\x04':
        try:
            return pd.read_excel(filepath_or_buffer, engine="openpyxl", **kwargs)
        except Exception as e1:
            try:
                return pd.read_excel(filepath_or_buffer, engine="calamine", **kwargs)
            except Exception as e2:
                raise ValueError(f"Fichier XLSX detecte mais erreur de lecture: {e1} / {e2}")

    # XLS = OLE2 Compound (commence par D0 CF 11 E0)
    if header_bytes[:4] == b'\xd0\xcf\x11\xe0':
        try:
            return pd.read_excel(filepath_or_buffer, engine="xlrd", **kwargs)
        except Exception as e:
            raise ValueError(f"Fichier XLS detecte mais erreur de lecture: {e}")

    # --- Détection texte : HTML, CSV, TSV, XML ---
    try:
        header_text = header_bytes.decode("utf-8", errors="ignore").strip()
    except Exception:
        header_text = ""

    # HTML table (très courant avec les exports SAP)
    if header_text.lower().startswith("<!doctype") or header_text.lower().startswith("<html") or "<table" in header_text.lower():
        try:
            # pd.read_html retourne une liste de DataFrames
            dfs = pd.read_html(filepath_or_buffer, **kwargs)
            if dfs:
                return dfs[0]
        except Exception as e:
            raise ValueError(f"Fichier HTML detecte mais erreur de lecture: {e}")

    # XML spreadsheet
    if header_text.lower().startswith("<?xml") and "spreadsheet" in header_text.lower():
        try:
            return pd.read_excel(filepath_or_buffer, engine="openpyxl", **kwargs)
        except Exception:
            pass

    # --- Tentative CSV/TSV avec détection du séparateur ---
    # Essayer de lire comme texte et détecter le séparateur
    try:
        if is_buffer:
            pos2 = filepath_or_buffer.tell()
            text_sample = filepath_or_buffer.read(4096).decode("utf-8", errors="ignore")
            filepath_or_buffer.seek(pos2)
        elif isinstance(filepath_or_buffer, str):
            with open(filepath_or_buffer, "r", encoding="utf-8", errors="ignore") as f:
                text_sample = f.read(4096)
        else:
            text_sample = ""

        if text_sample:
            lines = [l for l in text_sample.split("\n") if l.strip()]
            if lines:
                # Compter les séparateurs potentiels dans les premières lignes
                for sep, name in [("\t", "TAB"), (";", "POINT-VIRGULE"), (",", "VIRGULE"), ("|", "PIPE")]:
                    counts = [lines[i].count(sep) for i in range(min(3, len(lines)))]
                    if all(c > 0 for c in counts) and len(set(counts)) <= 1:
                        # Séparateur trouvé de manière fiable
                        try:
                            if is_buffer:
                                filepath_or_buffer.seek(0)
                            df = pd.read_csv(filepath_or_buffer, sep=sep, **kwargs)
                            if df is not None and not df.empty and len(df.columns) > 1:
                                return df
                        except Exception:
                            continue

                # Dernier recours CSV avec sep=auto (sniffing)
                try:
                    if is_buffer:
                        filepath_or_buffer.seek(0)
                    df = pd.read_csv(filepath_or_buffer, sep=None, engine="python", **kwargs)
                    if df is not None and not df.empty and len(df.columns) > 1:
                        return df
                except Exception:
                    pass
    except Exception:
        pass

    # --- Dernier recours : essayer tous les moteurs Excel ---
    errors = []
    for eng in ["openpyxl", "xlrd", "calamine"]:
        try:
            if is_buffer:
                filepath_or_buffer.seek(0)
            df = pd.read_excel(filepath_or_buffer, engine=eng, **kwargs)
            if df is not None and not df.empty:
                return df
        except Exception as e:
            errors.append(f"{eng}: {str(e)[:80]}")

    # Essai sans moteur spécifié
    try:
        if is_buffer:
            filepath_or_buffer.seek(0)
        return pd.read_excel(filepath_or_buffer, **kwargs)
    except Exception as e:
        errors.append(f"auto: {str(e)[:80]}")

    # Tout a échoué — message d'erreur détaillé
    ext_info = ""
    if isinstance(filepath_or_buffer, str):
        ext = os.path.splitext(filepath_or_buffer)[1].lower()
        ext_info = f" (extension: {ext})"
        size = os.path.getsize(filepath_or_buffer) if os.path.exists(filepath_or_buffer) else 0
        ext_info += f" (taille: {size} octets)"
    
    hex_header = header_bytes[:16].hex(" ") if header_bytes else "vide"
    raise ValueError(
        f"Impossible de lire le fichier{ext_info}.\n"
        f"En-tete hex: [{hex_header}]\n"
        f"En-tete texte: [{header_text[:100]}]\n"
        f"Erreurs: {'; '.join(errors)}\n"
        f"Conseil: Ouvrez le fichier dans Excel puis 'Enregistrer sous' > Classeur Excel (.xlsx)"
    )


# ============================================================
CACHE_FILE = ".dashboard_cache.json"

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt","r",encoding="utf-8") as f: return f.read().strip()
        except Exception: pass
    return datetime.now().strftime("%d/%m/%Y")

def build_cache_key(fichier_date, sp, sa, sd, dr):
    raw = json.dumps({"date": fichier_date, "sp": sorted(sp), "sa": sorted(sa), "sd": sorted(sd), "dr": [str(dr[0]), str(dr[1])] if len(dr)==2 else []}, sort_keys=True)
    return hashlib.md5(raw.encode()).hexdigest()

def save_cache(key, data):
    try:
        cache = {}
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE,"r",encoding="utf-8") as f: cache = json.load(f)
        serializable = {}
        for k, v in data.items():
            if isinstance(v, pd.DataFrame):
                s = {"_type": "df", "columns": list(v.columns), "index": [str(i) for i in v.index], "data": v.reset_index(drop=True).to_dict(orient="split")["data"]}
                clean_data = []
                for row in s["data"]:
                    clean_row = []
                    for cell in row:
                        if isinstance(cell, (np.integer,)): clean_row.append(int(cell))
                        elif isinstance(cell, (np.floating,)): clean_row.append(float(cell))
                        elif isinstance(cell, np.bool_): clean_row.append(bool(cell))
                        elif pd.isna(cell): clean_row.append(None)
                        else: clean_row.append(cell)
                    clean_data.append(clean_row)
                s["data"] = clean_data
                serializable[k] = s
            elif isinstance(v, dict):
                clean_d = {}
                for dk, dv in v.items():
                    dk_s = str(dk)
                    if isinstance(dv, (np.integer,)): clean_d[dk_s] = int(dv)
                    elif isinstance(dv, (np.floating,)): clean_d[dk_s] = float(dv)
                    elif pd.isna(dv): clean_d[dk_s] = None
                    else: clean_d[dk_s] = dv
                serializable[k] = {"_type": "dict", "data": clean_d}
            elif isinstance(v, list): serializable[k] = {"_type": "list", "data": json.loads(json.dumps(v, default=str))}
            elif isinstance(v, (int, float, str, bool, type(None))): serializable[k] = {"_type": "val", "data": v}
            else: serializable[k] = {"_type": "val", "data": str(v)}
        cache[key] = serializable
        with open(CACHE_FILE,"w",encoding="utf-8") as f: json.dump(cache, f, ensure_ascii=False)
    except Exception: pass

def load_cache(key):
    try:
        if not os.path.exists(CACHE_FILE): return None
        with open(CACHE_FILE,"r",encoding="utf-8") as f: cache = json.load(f)
        if key not in cache: return None
        raw = cache[key]; result = {}
        for k, v in raw.items():
            t = v.get("_type","val"); d = v.get("data")
            if t == "df" and isinstance(d, dict):
                cols = d.get("columns",[]); idx = d.get("index",[]); rows = d.get("data",[])
                df = pd.DataFrame(rows, columns=cols)
                if idx and len(idx) == len(df): df.index = idx
                result[k] = df
            elif t == "dict": result[k] = d
            elif t == "list": result[k] = d
            else: result[k] = d
        return result
    except Exception: return None

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
                if section and headers is None and cell0: headers=[str(c).strip() if c else "" for c in row]; continue
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
    perf_df=hist_df[hist_df["_section"]=="perf"].copy(); qual_df=hist_df[hist_df["_section"]=="qual"].copy(); variations=[]
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
                    variations.append({"Date precedente":prev_date,"Date actuelle":curr_date,"Poste":poste,"Type":sec_name,"KPI":kpi,"Valeur precedente":round(pv,2),"Valeur actuelle":round(cv,2),"Ecart":round(diff,2),"Ecart %":round(pct,2),"Tendance":trend})
    return pd.DataFrame(variations)

def generate_journal(var_df):
    if var_df.empty: return pd.DataFrame()
    j=var_df.copy(); j["Significatif"]=j["Ecart %"].abs()>=5; j=j[j["Significatif"]].copy()
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

# ============================================================
# PIE CHART PROFESSIONNEL
# ============================================================
def create_professional_pie(labels, values, title="", colors=None, hole=0.45, pull_small=0.12, small_threshold=5, show_center_text=True, center_text="", height=480, font_size_label=12):
    total = sum(values)
    if total == 0:
        fig = go.Figure(); fig.add_annotation(text="Aucune donnée", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="#718096"))
        fig.update_layout(height=height, margin=dict(t=40,b=10,l=10,r=10)); return fig
    n = len(labels)
    if colors is None: colors = [PIE_COLORS[i % len(PIE_COLORS)] for i in range(n)]
    pulls = [pull_small if 0 < (v/total*100) < small_threshold else 0 for v in values]
    text_labels = [f"{lab}<br>{v/total*100:.1f}%<br>({int(v)})" if v > 0 else "" for lab, v in zip(labels, values)]
    text_positions = ["outside" if (v/total*100) < small_threshold else "inside" for v in values]
    fig = go.Figure(go.Pie(labels=labels, values=values, hole=hole, pull=pulls, marker=dict(colors=colors, line=dict(color='white', width=2.5)), text=text_labels, textposition=text_positions, textfont=dict(size=font_size_label, color="#1a202c"), hovertemplate='<b>%{label}</b><br>Nombre: <b>%{value}</b><br>Pourcentage: <b>%{percent}</b><extra></extra>', sort=False, direction='clockwise', rotation=0))
    if show_center_text and hole > 0:
        if not center_text: center_text = f"Total<br><b>{int(total)}</b>"
        fig.add_annotation(text=center_text, x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False, font=dict(size=20, color="#1e3a5f", weight="bold"), align="center")
    fig.update_layout(title=dict(text=title, font=dict(size=16, color="#1e3a5f", weight="bold"), x=0.5, xanchor="center", y=0.97, yanchor="top"), height=height, margin=dict(t=50, b=20, l=30, r=30), showlegend=True, legend=dict(font=dict(size=11, color="#4a5568"), orientation="h", yanchor="bottom", y=-0.08, xanchor="center", x=0.5), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    return fig

def create_status_pie_chart(df, status_col, title="", colors_map=None, height=480):
    if df.empty or status_col not in df.columns:
        fig = go.Figure(); fig.add_annotation(text="Aucune donnée", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="#718096"))
        fig.update_layout(height=height); return fig
    counts = df[status_col].value_counts(); labels = counts.index.tolist(); values = counts.values.tolist()
    dc = {"CLOT":"#276749","TCLO":"#38a169","CRÉÉ":"#2b6cb0","CRÉE":"#2b6cb0","LANC":"#d69e2e","ENCO":"#805ad5","LIBE":"#4299e1","CARACTERISE":"#276749","NON CARACTERISE":"#c53030","OUI":"#276749","NON":"#c53030","APRV":"#276749","APRQ":"#2b6cb0","REJT":"#c53030","<1 mois":"#276749",">3 mois":"#c53030","1 mois < <3 mois":"#d69e2e","APRV AVAU":"#38a169","Inconnu":"#a0aec0"}
    colors = [(colors_map or dc).get(str(lab).strip(), PIE_COLORS[len(colors) % len(PIE_COLORS)]) for lab in labels]
    return create_professional_pie(labels=labels, values=values, title=title, colors=colors, hole=0.42, pull_small=0.15, small_threshold=6, show_center_text=True, center_text=f"Total<br><b>{int(sum(values))}</b>", height=height, font_size_label=11)

def create_age_pie_chart(df, age_col, title="", height=480):
    return create_status_pie_chart(df, age_col, title=title, colors_map={"<1 mois":"#276749","1 mois < <3 mois":"#d69e2e",">3 mois":"#c53030","Inconnu":"#a0aec0"}, height=height)

def create_kpi_pie_by_poste(ckdf, kpi_name, title="", height=500):
    if kpi_name not in ckdf.columns:
        fig = go.Figure(); fig.add_annotation(text="KPI non disponible", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="#718096")); fig.update_layout(height=height); return fig
    vals = ckdf[kpi_name].dropna(); vals = vals[vals != 0]
    if vals.empty or vals.sum() == 0:
        fig = go.Figure(); fig.add_annotation(text="Aucune valeur", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="#718096")); fig.update_layout(height=height); return fig
    labels = [str(idx) for idx in vals.index]; values = vals.values.tolist()
    return create_professional_pie(labels=labels, values=values, title=title, colors=[PIE_COLORS[i%len(PIE_COLORS)] for i in range(len(labels))], hole=0.40, pull_small=0.18, small_threshold=4, center_text=f"{kpi_name[:25]}<br>Moy: <b>{vals.sum()/max(len(vals),1):.1f}%</b>", height=height, font_size_label=10)

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
    .tw tbody td{padding:4px 6px;border-bottom:1px solid #edf2f7;white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important}
    .stTabs [data-baseweb="tab-list"]{gap:3px;background:#e2e8f0;padding:3px;border-radius:6px;margin-bottom:4px}
    .stTabs [data-baseweb="tab"]{border-radius:5px;padding:6px 14px;font-weight:600;font-size:14px}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--p)!important;box-shadow:0 2px 5px rgba(0,0,0,.07)}
    .sr{display:flex;align-items:center;padding:6px 10px;background:#fff;border-radius:5px;margin-bottom:2px;border:1px solid var(--b);font-size:13px}
    .sr .sn{font-weight:700;color:var(--p);min-width:220px;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .sc{padding:3px 9px;border-radius:12px;font-weight:800;font-size:14px;min-width:50px;text-align:center;margin:0 8px;color:#fff}
    .sr .sa{color:#718096;font-size:12px;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .stg{font-size:11px;color:#718096;min-width:60px;text-align:center;white-space:nowrap}
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
    .rank-card{background:#fff;border-radius:var(--r);padding:12px 16px;border:1px solid var(--b);box-shadow:0 2px 8px rgba(0,0,0,.04)}
    .rank-card .rank-title{font-size:15px;font-weight:800;margin-bottom:8px;padding-bottom:5px;border-bottom:2px solid var(--b)}
    .rank-row{display:flex;align-items:center;padding:5px 0;font-size:13px;border-bottom:1px solid #f7fafc}
    .rank-row:last-child{border:none}
    .rank-row .rank-num{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:13px;color:#fff;margin-right:10px;flex-shrink:0}
    .rank-row .rank-name{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .rank-row .rank-score{font-weight:900;min-width:70px;text-align:right}
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
        res['ckdf']=pd.DataFrame({"TAUX_REALISATION_CORRECTIF/PT":an["TAUX_REALISATION_CORRECTIF/PT"],"OT préparation <1 mois":pr["OT préparation <1 mois"],"OT préparation >3 mois":pr["OT préparation >3 mois"],"OT préparation 1mois< <3mois":pr["OT préparation 1mois< <3mois"],"OT planification <1 mois":pl["OT planification <1 mois"],"OT planification >3 mois":pl["OT planification >3 mois"],"OT planification 1mois< <3mois":pl["OT planification 1mois< <3mois"],"OT exécution <1 mois":ex["OT exécution <1 mois"],"OT exécution >3 mois":ex["OT exécution >3 mois"],"OT exécution 1mois< <3mois":ex["OT exécution 1mois< <3mois"],"appel avis approuvé":tca["appel avis approuvé"],"OT LANC ESTIME":la["OT LANC ESTIME"],"Backlog préparation caractérisé":pc["Backlog préparation caractérisé"],"Backlog planification caractérisé":plc["Backlog planification caractérisé"],"OT CONFIME":res['ot_confime']["OT CONFIME"],"OT_COR_EGAL":res['ot_cor_egal']["OT_COR_EGAL"]})
        return res

    def ks(v,c):
        try: val=float(v)
        except Exception: return ""
        if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val>=80 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=75 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val<=15 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val<=5 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c=="TAUX_REALISATION_CORRECTIF/PT": return "background:#c6efce;color:#006100;font-weight:600" if val>=85 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c=="appel avis approuvé": return "background:#c6efce;color:#006100;font-weight:600" if val>=95 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=90 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]: return "background:#c6efce;color:#006100;font-weight:600" if val>=100 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=95 else "background:#ffc7ce;color:#9c0006;font-weight:600")
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
        h='<table class="tw at"><thead><tr><th>KPI</th><th>Valeur</th><th>Cible</th><th>Ecart</th><th>Statut</th><th>Action</th></tr></thead><tbody>'
        for k in kpi_list:
            av=actuals.get(k,0); tv=targets.get(k,100); diff=av-tv; met=av<=tv if is_lb(k) else av>=tv
            st_s="background:#c6efce;color:#006100;font-weight:700" if met else "background:#ffc7ce;color:#9c0006;font-weight:700"
            ec="#276749" if met else "#c53030"; action="Atteint" if met else act_map.get(k,"")
            h+='<tr><td style="font-weight:600">%s</td><td>%.1f%%</td><td>%.0f%%</td><td style="color:%s;font-weight:700">%+.1f%%</td><td style="%s">%s</td><td style="color:#4a5568">%s</td></tr>'%(k,av,tv,ec,diff,st_s,"ATTEINT" if met else "NON ATTEINT",action)
        return h+'</tbody></table>'
    def html_classement(scores,accent):
        sp=sorted(scores.items(),key=lambda x:x[1],reverse=True)
        not_p=[(p,s) for p,s in sp if s<80]; t5=[(p,s) for p,s in sp if s>=80][:5]; b5=not_p[-5:] if len(not_p)>5 else not_p
        h='<div class="cg"><div><div class="ct" style="color:#38a169">Top 5 - Atteint</div>'
        if t5:
            for i,(p,s) in enumerate(t5): h+='<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(accent,i+1,p,cs("%.2f"%s),s)
        else: h+='<div style="padding:6px;font-size:12px;color:#718096">Aucun</div>'
        h+='</div><div><div class="ct" style="color:#e53e3e">Bottom 5 - Non Atteint</div>'
        if b5:
            for i,(p,s) in enumerate(reversed(b5)): h+='<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(len(b5)-i,p,cs("%.2f"%s),s)
        else: h+='<div style="padding:6px;font-size:12px;color:#38a169">Tous atteints</div>'
        h+='</div></div>'; return h
    def html_kpi_bars(kpi_list,actuals,targets,title,color_ok,color_fail):
        h='<div class="ca"><div class="ct" style="color:%s">%s</div>'%(color_ok,title)
        for k in kpi_list:
            av=actuals.get(k,0); met=av<=targets.get(k,100) if is_lb(k) else av>=targets.get(k,100)
            bw=min(max(av,0),100); bg=color_ok if met else color_fail
            h+='<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>'%(k,bw,bg,av)
        return h+'</div>'
    def html_grouped_bars(posts,pscores,qscores,title):
        h='<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>'%title
        h+='<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span></div>'
        for p in sorted(posts,key=lambda x:(pscores.get(x,0)+qscores.get(x,0))/2,reverse=True):
            pv,qv=pscores.get(p,0),qscores.get(p,0)
            h+='<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f gb-p" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w"><div class="gbr-f gb-q" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div></div></div>'%(p,min(max(pv,0),100),pv,min(max(qv,0),100),qv)
        return h+'</div>'
    def export_btn(df,filename):
        buf=io.BytesIO(); df.to_excel(buf,index=False,engine='openpyxl'); buf.seek(0)
        st.download_button("📥 Exporter Excel",data=buf,file_name=filename,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===================== SIDEBAR =====================
    with st.sidebar:
        st.markdown("""<div style="padding:10px 0 4px 0"><div style="font-size:22px;margin-bottom:2px">⚙️</div><div style="font-size:14px;font-weight:800;color:white">Filtres</div></div>""",unsafe_allow_html=True)
        st.markdown("---")
        show_filters=st.checkbox("Afficher les filtres",value=True,key="show_filters")
        if show_filters:
            unf=st.toggle("📁 Charger nouveaux fichiers",value=False,key="tf")
            ot_f=av_f=None; apm=[]
            if unf:
                ot_f=st.file_uploader("Fichier OT",type=["xlsx","xls","csv"],key="uot")
                av_f=st.file_uploader("Fichier AVIS",type=["xlsx","xls","csv"],key="uav")
            else:
                if os.path.exists("ot.xlsx"):
                    try:
                        _t=excr(safe_read_excel("ot.xlsx"))
                        apm=sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                    except Exception: pass
                st.markdown("""<div style="background:rgba(255,255,255,.1);padding:6px 10px;border-radius:6px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Donnees</div><div style="font-size:14px;color:white;font-weight:600;margin-top:2px">📅 %s</div></div>"""%fichier_date,unsafe_allow_html=True)
            st.markdown("---"); st.markdown("**🎯 Postes**")
            sp=st.multiselect("Poste",["All"]+apm,["All"],key="sp")
            st.markdown("**🏭 Atelier**")
            sa=st.multiselect("Atelier",["All","Sulfurique (PS)","Phosphorique (PP)","Engrais (TSP/REX)","Feed (MCP/DCP)"],["All"],key="sa")
            st.markdown("**🏢 Division**")
            sd=st.multiselect("Division",["All","SF1","SF2"],["All"],key="sd")
            st.markdown("---"); st.markdown("**📅 Periode**")
            dr=st.date_input("Date debut planifiee",value=(datetime(2025,1,1).date(),datetime.today().date()),format="DD/MM/YYYY",key="dr")
        else:
            unf=False; ot_f=av_f=None; apm=[]; sp=["All"]; sa=["All"]; sd=["All"]
            dr=(datetime(2025,1,1).date(),datetime.today().date())
            if os.path.exists("ot.xlsx"):
                try:
                    _t=excr(safe_read_excel("ot.xlsx"))
                    apm=sorted(_t[_t["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                except Exception: pass

    # ===================== DATA LOADING =====================
    if not unf or (ot_f is not None and av_f is not None):
        cache_key = build_cache_key(fichier_date, sp, sa, sd, dr) if not unf else None
        cached_data = load_cache(cache_key) if cache_key else None
        if cached_data is not None:
            ckdf=dfp=avf=df_dash=None; pa=qa=pa_d=qa_d=pscores=qscores=pscores_d=qscores_d={}; vp=[]; all_ano=[]; ano_data={}
            for k in ['ckdf','dfp','avf','df_dash']: 
                v=cached_data.get(k); 
                if k=='ckdf': ckdf=v
                elif k=='dfp': dfp=v
                elif k=='avf': avf=v
                elif k=='df_dash': df_dash=v
            for k in ['pa','qa','pa_d','qa_d','pscores','qscores','pscores_d','qscores_d']:
                v=cached_data.get(k,{}); 
                if k=='pa': pa=v
                elif k=='qa': qa=v
                elif k=='pa_d': pa_d=v
                elif k=='qa_d': qa_d=v
                elif k=='pscores': pscores=v
                elif k=='qscores': qscores=v
                elif k=='pscores_d': pscores_d=v
                elif k=='qscores_d': qscores_d=v
            vp=cached_data.get('vp',[]); all_ano=cached_data.get('all_ano',[]); ano_data=cached_data.get('ano_data',{})
            _cache_hit = True
        else:
            _cache_hit = False
    if not _cache_hit:
            try:
                if unf:
                    if ot_f is None:
                        st.error("📁 Veuillez selectionner le fichier OT")
                        st.stop()
                    if av_f is None:
                        st.error("📁 Veuillez selectionner le fichier AVIS")
                        st.stop()
                    raw_ot=safe_read_excel(ot_f); raw_av=safe_read_excel(av_f)
                else:
                    # Vérifier que les fichiers existent et ne sont pas vides
                    for fname, label in [("ot.xlsx","OT"),("avis.xlsx","AVIS")]:
                        if not os.path.exists(fname):
                            st.error(f"📁 Fichier {label} introuvable : {fname}\n\nPlacez-le dans le même dossier que le script.")
                            st.stop()
                        fsize = os.path.getsize(fname)
                        if fsize < 100:
                            st.error(f"❌ Le fichier {label} ({fname}) est vide ou corrompu ({fsize} octets).\n\n"
                                     f"**Action requise :**\n"
                                     f"1. Ouvrez votre vrai fichier de données dans Excel\n"
                                     f"2. Faites **Fichier > Enregistrer sous**\n"
                                     f"3. Choisissez **Classeur Excel (.xlsx)**\n"
                                     f"4. Nommez-le **{fname}** et placez-le à côté du script")
                            st.stop()
                    raw_ot=safe_read_excel("ot.xlsx"); raw_av=safe_read_excel("avis.xlsx")

      
                raw_ot=excr(raw_ot); raw_av=excr(raw_av)
                for c in ["Créé le","Date de début planifiée","Date de clôture","Début réel","Fin réelle"]:
                    if c in raw_ot.columns: raw_ot[c]=pd.to_datetime(raw_ot[c],errors="coerce")
                for c in ["Créé le","Début souhaité","Date de la clôture"]:
                    if c in raw_av.columns: raw_av[c]=pd.to_datetime(raw_av[c],errors="coerce")
                if not apm: apm=sorted(raw_ot[raw_ot["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)]["Poste travail princ."].dropna().unique().tolist())
                if "All" in sp or not sp: sp=apm
                if "All" in sa or not sa: sa=["All"]
                if "All" in sd or not sd: sd=["All"]
                sdt=pd.to_datetime(dr[0]) if len(dr)==2 else pd.to_datetime(datetime(2025,1,1))
                edt=pd.to_datetime(dr[1]) if len(dr)==2 else pd.to_datetime(datetime.today())
                def mf(poste):
                    p=str(poste).upper()
                    if "All" not in sa:
                        m=False
                        if "Sulfurique (PS)" in sa and "PS" in p: m=True
                        if "Phosphorique (PP)" in sa and "PP" in p: m=True
                        if "Engrais (TSP/REX)" in sa and ("TSP" in p or "REX" in p): m=True
                        if "Feed (MCP/DCP)" in sa and ("MCP" in p or "DCP" in p): m=True
                        if not m: return False
                    if "All" not in sd:
                        m=False
                        if "SF1" in sd and "SF1" in p: m=True
                        if "SF2" in sd and "SF2" in p: m=True
                        if not m: return False
                    return True
                vp=[p for p in apm if mf(p) and p in sp]
                df=raw_ot[(raw_ot["Poste travail princ."].isin(vp))&(raw_ot["Date de début planifiée"].between(sdt,edt))].copy()
                avdf=raw_av[raw_av["Poste travail princ."].isin(vp)].copy()
                df=excr(df[df["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)].drop_duplicates())
                avdf=excr(avdf[(avdf["Ordre"].isna())|(avdf["Ordre"].astype(str).str.strip().eq(""))].drop_duplicates())
                if "Statut système" in df.columns: df["Statut OT"]=df["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]
                df_dash=raw_ot[raw_ot["Poste travail princ."].isin(vp)].copy()
                df_dash=excr(df_dash[df_dash["Poste travail princ."].astype(str).str.startswith(("SF1","SF2"),na=False)].drop_duplicates())
                if "Statut système" in df_dash.columns: df_dash["Statut OT"]=df_dash["Statut système"].fillna("").astype(str).str.strip().str.split().str[0]
                now=pd.Timestamp.now()
                res=calc_kpis(df,avdf,now,vp); ckdf=res['ckdf']; dfp=res['dfp']; avf=res['avf']
                res_d=calc_kpis(df_dash,avdf,now,vp); ckdf_d=res_d['ckdf']
                pa={k:round(ckdf[k].mean(),2) for k in QK}; qa={k:round(ckdf[k].mean(),2) for k in PK}
                pa_d={k:round(ckdf_d[k].mean(),2) for k in QK}; qa_d={k:round(ckdf_d[k].mean(),2) for k in PK}
                pscores={}; qscores={}
                for poste in ckdf.index:
                    r=ckdf.loc[poste]
                    pscores[poste]=(sum(gscore(k,r[k],CIBLE[k]) for k in QK if k in r.index)/len(QK)*100) if QK else 0
                    qscores[poste]=(sum(gscore(k,r[k],CIBLE[k]) for k in PK if k in r.index)/len(PK)*100) if PK else 0
                pscores_d={}; qscores_d={}
                for poste in ckdf_d.index:
                    r=ckdf_d.loc[poste]
                    pscores_d[poste]=(sum(gscore(k,r[k],CIBLE[k]) for k in QK if k in r.index)/len(QK)*100) if QK else 0
                    qscores_d[poste]=(sum(gscore(k,r[k],CIBLE[k]) for k in PK if k in r.index)/len(PK)*100) if PK else 0
                all_ano=[]; sub_p={"TAUX_REALISATION_CORRECTIF/PT":lambda d:d[(d["Nº appel pl.entret."].fillna(0)==0)&(~d["Statut OT"].isin(["CLOT","TCLO"]))],"OT préparation <1 mois":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]!="<1 mois")],"OT préparation >3 mois":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]==">3 mois")],"OT planification <1 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]!="<1 mois")],"OT planification >3 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]==">3 mois")],"OT exécution <1 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]!="<1 mois")],"OT exécution >3 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]==">3 mois")],"OT préparation 1mois< <3mois":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]=="1 mois < <3 mois")],"OT planification 1mois< <3mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]=="1 mois < <3 mois")],"OT exécution 1mois< <3mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]=="1 mois < <3 mois")]}
                sub_q={"OT LANC ESTIME":lambda d:d[(d["Statut OT"]=="LANC")&(d["OT LANC ESTIME"]=="NON")],"Backlog préparation caractérisé":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["Backlog preparation"]=="NON CARACTERISE")],"Backlog planification caractérisé":lambda d:d[(d["Statut OT"]=="LANC")&(d["Backlog planification"]=="NON CARACTERISE")],"OT COR Egal":lambda d:d[(d["OT COR EGAL"]=="NON")],"OT CONFIME":lambda d:d[(d["OT CONFIME"]=="NON")&(d["Statut OT"].isin(["CLOT","TCLO"]))],"appel avis approuvé":lambda d:d[(d["Statut utilisateur"].isin(["APRQ","REJT"]))]}
                ano_data={}
                for kn,fn in sub_p.items():
                    try:
                        sd2=fn(dfp); cnt=len(sd2)
                        if cnt>0:
                            grp=sd2.groupby("Poste travail princ.")["Ordre"].count().to_dict()
                            all_ano.extend([{"KPI":kn,"Poste":p,"Nb anomalies":n} for p,n in grp.items() if n>0]); ano_data[kn]=cnt
                    except Exception: pass
                for kn,fn in sub_q.items():
                    try:
                        sd2=fn(dfp); cnt=len(sd2)
                        if cnt>0:
                            grp=sd2.groupby("Poste travail princ.")["Ordre"].count().to_dict()
                            all_ano.extend([{"KPI":kn,"Poste":p,"Nb anomalies":n} for p,n in grp.items() if n>0]); ano_data[kn]=cnt
                    except Exception: pass
                if cache_key:
                    save_cache(cache_key, {'ckdf':ckdf,'dfp':dfp,'avf':avf,'pa':pa,'qa':qa,'pa_d':pa_d,'qa_d':qa_d,'pscores':pscores,'qscores':qscores,'pscores_d':pscores_d,'qscores_d':qscores_d,'vp':vp,'df_dash':df_dash,'all_ano':all_ano,'ano_data':ano_data})
            except Exception as e:
                st.error(f"Erreur de chargement: {str(e)}"); st.stop()
        else:
            pass

        # ===================== DASHBOARD =====================
        p_score=round(np.mean(list(pscores.values())),2) if pscores else 0
        q_score=round(np.mean(list(qscores.values())),2) if qscores else 0
        total_ot=len(dfp); total_anom=len(all_ano)
        p_score_d=round(np.mean(list(pscores_d.values())),2) if pscores_d else 0
        q_score_d=round(np.mean(list(qscores_d.values())),2) if qscores_d else 0
        total_ot_d=len(df_dash) if df_dash is not None else 0

        st.markdown('<div class="mh"><h1>📊 DASHBOARD KPI - SUIVI MAINTENANCE</h1><span class="db">📅 %s</span></div>'%fichier_date,unsafe_allow_html=True)
        st.markdown('<div class="cr"><div class="cc c1"><div class="cv">%d</div><div class="cl">OT (Periode)</div></div><div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Perf.</div></div><div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Qual.</div></div><div class="cc c4"><div class="cv">%d</div><div class="cl">Anomalies</div></div></div>'%(total_ot,p_score,q_score,total_anom),unsafe_allow_html=True)
        st.markdown('<div class="cr"><div class="cc c1"><div class="cv">%d</div><div class="cl">OT (Total)</div></div><div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Perf. Total</div></div><div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Qual. Total</div></div><div class="cc c4"><div class="cv">%d</div><div class="cl">Postes</div></div></div>'%(total_ot_d,p_score_d,q_score_d,len(vp)),unsafe_allow_html=True)

        tab1,tab2,tab3,tab4,tab5,tab6=st.tabs(["📋 Performance","🎯 Qualite","⚠️ Anomalies","📊 Graphiques","📈 Tendances","💾 Export"])

        with tab1:
            st.markdown('<div class="stl p">INDICATEURS DE PERFORMANCE</div>',unsafe_allow_html=True)
            pcols=["Poste de travail"]+QK+["Score Performance"]; prows=[]
            for poste in ckdf.index:
                r=ckdf.loc[poste]; row={"Poste de travail":poste,"_t":""}
                for k in QK: row[k]="%.1f"%r[k] if k in r.index and not pd.isna(r[k]) else "N/A"
                row["Score Performance"]="%.2f"%pscores.get(poste,0); prows.append(row)
            cible_row={"Poste de travail":"CIBLE","_t":"cible"}
            for k in QK: cible_row[k]=CIBLE.get(k,"")
            cible_row["Score Performance"]="100.00"; prows.insert(0,cible_row)
            tot_row={"Poste de travail":"Moyenne","_t":"total"}
            for k in QK: tot_row[k]="%.1f"%pa.get(k,0)
            tot_row["Score Performance"]="%.2f"%p_score; prows.append(tot_row)
            st.markdown(html_table(prows,pcols,"pt",sc_col={"Score Performance"}),unsafe_allow_html=True)
            st.markdown('<div class="stl p" style="margin-top:8px">BAREMES</div>',unsafe_allow_html=True)
            st.markdown(html_kpi_bars(QK,pa,CIBLE,"Performance Globale","#38a169","#e53e3e"),unsafe_allow_html=True)
            st.markdown('<div class="stl c" style="margin-top:8px">CLASSEMENT</div>',unsafe_allow_html=True)
            st.markdown(html_classement(pscores,"#276749"),unsafe_allow_html=True)

        with tab2:
            st.markdown('<div class="stl q">INDICATEURS DE QUALITE</div>',unsafe_allow_html=True)
            qcols=["Poste de travail"]+PK+["Score Qualite"]; qrows=[]
            for poste in ckdf.index:
                r=ckdf.loc[poste]; row={"Poste de travail":poste,"_t":""}
                for k in PK: row[k]="%.1f"%r[k] if k in r.index and not pd.isna(r[k]) else "N/A"
                row["Score Qualite"]="%.2f"%qscores.get(poste,0); qrows.append(row)
            cible_row2={"Poste de travail":"CIBLE","_t":"cible"}
            for k in PK: cible_row2[k]=CIBLE.get(k,"")
            cible_row2["Score Qualite"]="100.00"; qrows.insert(0,cible_row2)
            tot_row2={"Poste de travail":"Moyenne","_t":"total"}
            for k in PK: tot_row2[k]="%.1f"%qa.get(k,0)
            tot_row2["Score Qualite"]="%.2f"%q_score; qrows.append(tot_row2)
            st.markdown(html_table(qrows,qcols,"qt",sc_col={"Score Qualite"}),unsafe_allow_html=True)
            st.markdown('<div class="stl q" style="margin-top:8px">BAREMES</div>',unsafe_allow_html=True)
            st.markdown(html_kpi_bars(PK,qa,CIBLE,"Qualite Globale","#3182ce","#e53e3e"),unsafe_allow_html=True)
            st.markdown('<div class="stl c" style="margin-top:8px">CLASSEMENT</div>',unsafe_allow_html=True)
            st.markdown(html_classement(qscores,"#2b6cb0"),unsafe_allow_html=True)

        with tab3:
            st.markdown('<div class="stl a">ANOMALIES</div>',unsafe_allow_html=True)
            if all_ano:
                ano_df=pd.DataFrame(all_ano)
                ano_grp=ano_df.groupby("KPI")["Nb anomalies"].sum().sort_values(ascending=False).reset_index(); ano_grp.columns=["KPI","Total"]
                ano_pivot=ano_df.pivot_table(index="Poste",columns="KPI",values="Nb anomalies",aggfunc="sum",fill_value=0)
                ano_pivot["Total"]=ano_pivot.sum(axis=1); ano_pivot=ano_pivot.sort_values("Total",ascending=False)
                acols=["KPI"]+ano_pivot.columns.tolist()
                arows=[{"KPI":"Total","_t":"total","Total":int(ano_grp["Total"].sum())}]
                for k in ano_pivot.columns:
                    if k!="Total": arows.append({"KPI":k,"_t":"total","Total":int(ano_grp[ano_grp["KPI"]==k]["Total"].values[0]) if len(ano_grp[ano_grp["KPI"]==k])>0 else 0})
                for poste in ano_pivot.index:
                    row={"KPI":poste,"_t":""}; 
                    for c in ano_pivot.columns: row[c]=int(ano_pivot.loc[poste,c])
                    arows.append(row)
                st.markdown(html_ano(arows,acols),unsafe_allow_html=True)
                st.markdown('<div class="stl a" style="margin-top:8px">ACTIONS CORRECTIVES</div>',unsafe_allow_html=True)
                st.markdown(html_actions_table(list(ano_data.keys()),{**pa,**qa},CIBLE,ACT_MAP),unsafe_allow_html=True)
            else:
                st.markdown('<div class="es">✅ Aucune anomalie</div>',unsafe_allow_html=True)

        with tab4:
            st.markdown('<div class="stl c">PIE CHARTS - REPARTITION</div>',unsafe_allow_html=True)
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.plotly_chart(create_status_pie_chart(dfp,"Statut OT","Repartition par Statut OT",height=460),use_container_width=True,config={"displayModeBar":False})
                if "ap" in dfp.columns:
                    st.plotly_chart(create_age_pie_chart(dfp[dfp["Statut OT"]=="CRÉÉ"],"ap","Age Preparation",height=460),use_container_width=True,config={"displayModeBar":False})
                if "Backlog preparation" in dfp.columns:
                    st.plotly_chart(create_status_pie_chart(dfp[dfp["Statut OT"]=="CRÉÉ"],"Backlog preparation","Backlog Prep. Caracterise",height=460),use_container_width=True,config={"displayModeBar":False})
            with col_p2:
                if "alp" in dfp.columns:
                    st.plotly_chart(create_age_pie_chart(dfp[(dfp["Statut OT"]=="LANC")&(dfp["Contient SOPL"]==0)],"alp","Age Planification",height=460),use_container_width=True,config={"displayModeBar":False})
                if "aex" in dfp.columns:
                    st.plotly_chart(create_age_pie_chart(dfp[(dfp["Statut OT"]=="LANC")&(dfp["Contient SOPL"]==1)],"aex","Age Execution",height=460),use_container_width=True,config={"displayModeBar":False})
                if "Backlog planification" in dfp.columns:
                    st.plotly_chart(create_status_pie_chart(dfp[dfp["Statut OT"]=="LANC"],"Backlog planification","Backlog Plan. Caracterise",height=460),use_container_width=True,config={"displayModeBar":False})

            st.markdown('<div class="stl s" style="margin-top:10px">KPI PAR POSTE</div>',unsafe_allow_html=True)
            kpi_sel=st.selectbox("Choisir KPI",ALL_KPI,key="kpi_pie_sel")
            col_k1,col_k2=st.columns(2)
            with col_k1:
                st.plotly_chart(create_kpi_pie_by_poste(ckdf,kpi_sel,f"Repartition: {kpi_sel}",480),use_container_width=True,config={"displayModeBar":False})
            with col_k2:
                if "OT CONFIME" in dfp.columns:
                    st.plotly_chart(create_status_pie_chart(dfp,"OT CONFIME","OT Confirmes",480),use_container_width=True,config={"displayModeBar":False})

            st.markdown('<div class="stl p" style="margin-top:10px">ATELIER & DIVISION</div>',unsafe_allow_html=True)
            dfc=dfp.copy(); dfc["Atelier"]=dfc["Poste travail princ."].apply(get_atelier); dfc["Division"]=dfc["Poste travail princ."].apply(get_division); dfc["Metier"]=dfc["Poste travail princ."].apply(get_metier)
            col_a1,col_a2,col_a3=st.columns(3)
            with col_a1: st.plotly_chart(create_professional_pie(dfc["Atelier"].value_counts().index.tolist(),dfc["Atelier"].value_counts().values.tolist(),"Par Atelier",["#276749","#2b6cb0","#d69e2e","#805ad5","#a0aec0"],0.40,420),use_container_width=True,config={"displayModeBar":False})
            with col_a2: st.plotly_chart(create_professional_pie(dfc["Division"].value_counts().index.tolist(),dfc["Division"].value_counts().values.tolist(),"Par Division",["#1e3a5f","#4299e1","#a0aec0"],0.40,420),use_container_width=True,config={"displayModeBar":False})
            with col_a3: st.plotly_chart(create_professional_pie(dfc["Metier"].value_counts().index.tolist(),dfc["Metier"].value_counts().values.tolist(),"Par Metier",["#e53e3e","#2b6cb0","#805ad5","#d69e2e","#38a169"],0.40,420),use_container_width=True,config={"displayModeBar":False})

            st.markdown('<div class="stl p" style="margin-top:10px">SCORES PAR POSTE</div>',unsafe_allow_html=True)
            st.markdown(html_grouped_bars(vp,pscores,qscores,"Performance vs Qualite"),unsafe_allow_html=True)

            if "OT LANC ESTIME" in dfp.columns:
                col_e1,col_e2=st.columns(2)
                with col_e1: st.plotly_chart(create_status_pie_chart(dfp[dfp["Statut OT"]=="LANC"],"OT LANC ESTIME","OT Lances Estimes",420),use_container_width=True,config={"displayModeBar":False})
                with col_e2: st.plotly_chart(create_status_pie_chart(dfp,"OT_COR_EGAL","Couts Reels=Budgetes",420),use_container_width=True,config={"displayModeBar":False})
            if avf is not None and not avf.empty:
                st.markdown('<div class="stl q" style="margin-top:10px">APPELS AVIS</div>',unsafe_allow_html=True)
                col_v1,col_v2=st.columns(2)
                with col_v1: st.plotly_chart(create_status_pie_chart(avf,"Statut utilisateur","Statut Avis",420),use_container_width=True,config={"displayModeBar":False})
                with col_v2:
                    avp=avf["Poste travail princ."].value_counts()
                    st.plotly_chart(create_professional_pie(avp.index.tolist(),avp.values.tolist(),"Avis par Poste",hole=0.40,height=420),use_container_width=True,config={"displayModeBar":False})

        with tab5:
            st.markdown('<div class="stl s">TENDANCES</div>',unsafe_allow_html=True)
            hp=os.path.join("kpis","indicateurs_kpis.xlsx")
            if os.path.exists(hp):
                hdf=load_historical_kpis(hp); vdf=calculate_variations(hdf)
                if not vdf.empty:
                    jrn=generate_journal(vdf)
                    if not jrn.empty:
                        st.markdown('<div class="ca"><div class="ct">Variations Significatives (|e|>=5%)</div>',unsafe_allow_html=True)
                        for _,row in jrn.iterrows():
                            sc="#276749" if row["Sens"]=="Amelioration" else "#c53030"; si="▲" if row["Sens"]=="Amelioration" else "▼"
                            st.markdown('<div class="sr"><span class="sn">%s - %s</span><span class="sc" style="background:%s">%s %.1f%%</span><span class="sa">%s: %.1f → %.1f</span><span class="stg">%s → %s</span></div>'%(row["Poste"],row["Type"],sc,si,row["Ecart %"],row["KPI"],row["Valeur precedente"],row["Valeur actuelle"],row["Date precedente"],row["Date actuelle"]),unsafe_allow_html=True)
                        st.markdown('</div>',unsafe_allow_html=True)
                    t5,b5=calculate_rankings(vdf)
                    if not t5.empty:
                        st.markdown('<div class="dgrid">',unsafe_allow_html=True)
                        st.markdown('<div class="rank-card"><div class="rank-title" style="color:#276749">🏆 Top 5 Amelioration</div>',unsafe_allow_html=True)
                        for i,(_,r) in enumerate(t5.iterrows()): st.markdown('<div class="rank-row"><span class="rank-num" style="background:#276749">%s</span><span class="rank-name">%s</span><span class="rank-score" style="color:#276749">%+.1f</span></div>'%(i+1,r["Poste"],r["Score variation"]),unsafe_allow_html=True)
                        st.markdown('</div><div class="rank-card"><div class="rank-title" style="color:#c53030">⚠️ Top 5 Degradation</div>',unsafe_allow_html=True)
                        for i,(_,r) in enumerate(b5.iterrows()): st.markdown('<div class="rank-row"><span class="rank-num" style="background:#c53030">%s</span><span class="rank-name">%s</span><span class="rank-score" style="color:#c53030">%+.1f</span></div>'%(i+1,r["Poste"],r["Score variation"]),unsafe_allow_html=True)
                        st.markdown('</div></div>',unsafe_allow_html=True)
                else: st.markdown('<div class="es">Pas assez de donnees historiques (min. 2 periodes)</div>',unsafe_allow_html=True)
            else: st.markdown('<div class="es">Aucun historique dans kpis/indicateurs_kpis.xlsx</div>',unsafe_allow_html=True)

        with tab6:
            st.markdown('<div class="stl s">EXPORT</div>',unsafe_allow_html=True)
            col_e1,col_e2=st.columns(2)
            with col_e1:
                st.markdown("**Performance**")
                pcols_exp=["Poste de travail"]+QK+["Score Performance"]; pdf_exp=pd.DataFrame(prows)
                if not pdf_exp.empty: pdf_exp=pdf_exp[pcols_exp]
                export_btn(pdf_exp,"performance_kpis.xlsx")
                st.markdown("**Qualite**")
                qcols_exp=["Poste de travail"]+PK+["Score Qualite"]; qdf_exp=pd.DataFrame(qrows)
                if not qdf_exp.empty: qdf_exp=qdf_exp[qcols_exp]
                export_btn(qdf_exp,"qualite_kpis.xlsx")
            with col_e2:
                st.markdown("**Anomalies**")
                if all_ano: export_btn(pd.DataFrame(all_ano),"anomalies.xlsx")
                else: st.info("Aucune anomalie")
                st.markdown("**Sauvegarde historique**")
                if st.button("💾 Sauvegarder",key="save_hist"):
                    pcols_h=["Poste de travail"]+QK+["Score Performance"]; qcols_h=["Poste de travail"]+PK+["Score Qualite"]
                    pr_h=[{k:r[k] for k in pcols_h if k in r} for r in prows if r.get("_t")!="cible"]
                    qr_h=[{k:r[k] for k in qcols_h if k in r} for r in qrows if r.get("_t")!="cible"]
                    ano_p_h=[]; ano_q_h=[]
                    if all_ano:
                        adf=pd.DataFrame(all_ano)
                        for kn in QK:
                            for _,r in adf[adf["KPI"]==kn].iterrows(): ano_p_h.append({"KPI":kn,"Poste":r["Poste"],"Nb":int(r["Nb anomalies"])})
                        for kn in PK:
                            for _,r in adf[adf["KPI"]==kn].iterrows(): ano_q_h.append({"KPI":kn,"Poste":r["Poste"],"Nb":int(r["Nb anomalies"])})
                    save_kpis_to_excel(pr_h,pcols_h,qr_h,qcols_h,ano_p_h,["KPI","Poste","Nb"] if ano_p_h else [],ano_q_h,["KPI","Poste","Nb"] if ano_q_h else [],fichier_date)
                    st.success("✅ Sauvegarde effectuee!")
    else:
        st.markdown('<div class="es" style="margin-top:100px">📁 Chargez les fichiers OT et AVIS</div>',unsafe_allow_html=True)

if __name__=="__main__":
    main()
