# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# === 1. Page config WIDE ===
st.set_page_config(
    layout="wide",
    page_title="Dashboard KPI",
    page_icon="📊"
)

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt", "r", encoding="utf-8") as f:
                return f.read().strip()
        except: pass
    return datetime.now().strftime("%d/%m/%Y")

def save_kpis_to_excel(prows, pcols, qrows, qcols, ano_p_r, ano_p_c, ano_q_r, ano_q_c, sheet_name):
    kpis_dir = "kpis"
    os.makedirs(kpis_dir, exist_ok=True)
    filepath = os.path.join(kpis_dir, "indicateurs_kpis.xlsx")
    sn = str(sheet_name).replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","")[:31]
    hf = Font(bold=True, color="FFFFFF", size=10)
    hfl = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    tf = Font(bold=True, size=12, color="1E3A5F")
    tb = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    try: wb = load_workbook(filepath)
    except: wb = Workbook(); 
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    if sn in wb.sheetnames: del wb[sn]
    ws = wb.create_sheet(sn)
    rn = 1
    def ws_sec(title, cols, rows, sr):
        ws.cell(row=sr, column=1, value=title).font = tf; sr += 1
        for j, c in enumerate(cols, 1):
            cl = ws.cell(row=sr, column=j, value=c); cl.font = hf; cl.fill = hfl; cl.alignment = Alignment(horizontal='center'); cl.border = tb
        sr += 1
        for r in rows:
            for j, c in enumerate(cols, 1):
                cl = ws.cell(row=sr, column=j, value=r.get(c, "")); cl.border = tb; cl.alignment = Alignment(horizontal='center')
            sr += 1
        return sr + 1
    rn = ws_sec("INDICATEURS DE PERFORMANCE", pcols, prows, rn)
    if ano_p_c and ano_p_r: rn = ws_sec("ANOMALIES PERFORMANCE", ano_p_c, ano_p_r, rn)
    rn = ws_sec("INDICATEURS DE QUALITE", qcols, qrows, rn)
    if ano_q_c and ano_q_r: rn = ws_sec("ANOMALIES QUALITE", ano_q_c, ano_q_r, rn)
    try: wb.save(filepath)
    except: pass

def parse_date_sheet(name):
    for fmt in ["%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"]:
        try: return datetime.strptime(name, fmt)
        except: pass
    return datetime.min

def parse_kpis_excel(filepath):
    try: wb = load_workbook(filepath, data_only=True)
    except: return {}
    results = {}
    sec_keys = ["INDICATEURS DE PERFORMANCE","ANOMALIES PERFORMANCE","INDICATEURS DE QUALITE","ANOMALIES QUALITE"]
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        poste_data = {}
        cur_sec = None; headers = None
        for row in ws.iter_rows(min_row=1, values_only=True):
            fc = str(row[0]).strip() if row[0] else ""
            fu = fc.upper()
            if fu in sec_keys:
                cur_sec = fu; headers = None; continue
            if fc == "CIBLE" or fc == "Total general": continue
            if cur_sec and headers is None:
                headers = [str(c).strip() if c else "" for c in row]; continue
            if cur_sec and headers and fc:
                rd = {headers[i]: row[i] for i in range(min(len(headers), len(row)))}
                poste = fc
                if poste not in poste_data: poste_data[poste] = {"Score Performance":0,"Score Qualite":0,"Ano Perf":0,"Ano Qual":0}
                if "Score Performance" in rd and rd["Score Performance"]:
                    try: poste_data[poste]["Score Performance"] = float(str(rd["Score Performance"]).replace('%','').strip())
                    except: pass
                if "Score Qualite" in rd and rd["Score Qualite"]:
                    try: poste_data[poste]["Score Qualite"] = float(str(rd["Score Qualite"]).replace('%','').strip())
                    except: pass
                if "ANOMALIE" in cur_sec and "Total" in rd and rd["Total"] is not None:
                    try:
                        v = int(float(str(rd["Total"])))
                        if "PERFORMANCE" in cur_sec: poste_data[poste]["Ano Perf"] = v
                        else: poste_data[poste]["Ano Qual"] = v
                    except: pass
        results[sheet_name] = poste_data
    return results

def get_tendance(pe, qe, ae):
    imp = 0; deg = 0
    if pe > 1: imp += 1
    elif pe < -1: deg += 1
    if qe > 1: imp += 1
    elif qe < -1: deg += 1
    if ae < -1: imp += 1
    elif ae > 1: deg += 1
    if imp >= 2: return "🟢 Amélioration","amelioration"
    elif deg >= 2: return "🔴 Dégradation","degradation"
    return "🟡 Stable","stable"

def generate_ai_report(comp_df, dn, dn1):
    if comp_df.empty: return "Aucune donnée disponible pour l'analyse."
    total = len(comp_df)
    nb_imp = len(comp_df[comp_df["Tendance"]=="amelioration"])
    nb_deg = len(comp_df[comp_df["Tendance"]=="degradation"])
    nb_stb = len(comp_df[comp_df["Tendance"]=="stable"])
    mp = comp_df["Ecart Perf"].mean()
    mq = comp_df["Ecart Qual"].mean()
    ma = comp_df["Ecart Ano"].mean()
    r = f"""## 📋 Rapport d'Évolution des KPIs
**Période analysée : {dn1} → {dn}**

---

### 📊 Résumé Général

| Indicateur | Valeur |
|---|---|
| Nombre total de postes analysés | **{total}** |
| Postes en amélioration 🟢 | **{nb_imp}** ({round(nb_imp/total*100,1)}%) |
| Postes en dégradation 🔴 | **{nb_deg}** ({round(nb_deg/total*100,1)}%) |
| Postes stables 🟡 | **{nb_stb}** ({round(nb_stb/total*100,1)}%) |
| Évolution moyenne Performance | **{mp:+.2f} %** |
| Évolution moyenne Qualité | **{mq:+.2f} %** |
| Évolution moyenne Anomalies | **{ma:+.1f}** |

---

### 🟢 Analyse des Améliorations
"""
    if nb_imp > 0:
        imp_df = comp_df[comp_df["Tendance"]=="amelioration"].sort_values("Ecart Perf", ascending=False)
        for _, rw in imp_df.head(5).iterrows():
            r += f"- **{rw['Poste']}** : Performance {rw['Ecart Perf']:+.2f}%, Qualité {rw['Ecart Qual']:+.2f}%, Anomalies {rw['Ecart Ano']:+.0f}\n"
        r += f"\n**Conclusion** : {nb_imp} poste(s) montrent une tendance positive. "
        if mp > 0 and mq > 0: r += "Performance et Qualité en hausse, dynamique positive globale."
        elif mp > 0: r += "Amélioration portée par la Performance. Qualité à surveiller."
        else: r += "Amélioration portée par la Qualité. Performance à consolider."
    else: r += "Aucun poste n'a montré d'amélioration significative.\n"
    r += "\n---\n### 🔴 Analyse des Dégradations\n"
    if nb_deg > 0:
        deg_df = comp_df[comp_df["Tendance"]=="degradation"].sort_values("Ecart Perf", ascending=True)
        for _, rw in deg_df.head(5).iterrows():
            r += f"- **{rw['Poste']}** : Performance {rw['Ecart Perf']:+.2f}%, Qualité {rw['Ecart Qual']:+.2f}%, Anomalies {rw['Ecart Ano']:+.0f}\n"
        r += f"\n**Conclusion** : {nb_deg} poste(s) nécessitent une attention particulière. "
        crit = deg_df[(deg_df["Ecart Perf"]<0) & (deg_df["Ecart Ano"]>0)]
        if not crit.empty: r += f"**Postes critiques** : {', '.join(crit['Poste'].tolist())}. Plan d'action prioritaire requis."
        else: r += "Dégradation modérée mais suivi attentif nécessaire."
    else: r += "Aucun poste en dégradation. Bonne stabilité.\n"
    r += "\n---\n### 💡 Recommandations\n"
    recs = []
    if nb_deg > 0:
        recs.append("1. **Réduire les anomalies récurrentes** sur les postes en dégradation via analyses de causes racines.")
        recs.append("2. **Renforcer les actions préventives** sur les postes avec baisse Performance + hausse Anomalies.")
    if nb_imp > 0:
        top_imp = comp_df[comp_df["Tendance"]=="amelioration"].sort_values("Ecart Perf", ascending=False).head(3)
        recs.append(f"3. **Généraliser les bonnes pratiques** : {', '.join(top_imp['Poste'].tolist())}.")
    if nb_deg > 0 and nb_imp > 0:
        recs.append("4. **Organiser un partage d'expérience** entre postes améliorés et dégradés.")
    if ma > 0: recs.append("5. **Plan de réduction des anomalies** ciblé avec objectifs hebdomadaires.")
    recs.append("6. **Suivi hebdomadaire des KPIs** pour détecter toute dérive rapidement.")
    recs.append("7. **Prioriser les postes critiques** (baisse Perf + hausse Anomalies simultanée).")
    recs.append("8. **Valider la caractérisation complète du backlog** pour fiabiliser les indicateurs.")
    r += "\n".join(recs)
    return r

# === 2. CSS avec sidebar réduite et plein écran ===
def inject_custom_css():
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--p:#1e3a5f;--pl:#2c5282;--b:#e2e8f0;--r:10px}
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#edf2f7;font-family:'Inter',sans-serif}

    /* === 5. Vrai mode plein écran === */
    .main .block-container{
        max-width: 100% !important;
        width: 100% !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        padding-top: .5rem !important;
        padding-bottom: .5rem !important;
    }

    /* === 2. Largeur sidebar réduite === */
    section[data-testid="stSidebar"] {
        width: 250px !important;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] {
        width: 0px !important;
    }
    section[data-testid="stSidebar"] .stMarkdown,
    section[data-testid="stSidebar"] .stSelectbox,
    section[data-testid="stSidebar"] .stMultiSelect,
    section[data-testid="stSidebar"] .stDateInput,
    section[data-testid="stSidebar"] .stToggle {
        margin-left: 4px;
        margin-right: 4px;
    }

    .stTabs,.stTabs>div,.stTabs [data-baseweb="tab-list"]{width:100%!important;max-width:100%!important}
    .mh{background:linear-gradient(135deg,var(--p),var(--pl));padding:12px 20px;border-radius:var(--r);margin-bottom:6px;box-shadow:0 6px 20px rgba(0,0,0,.1);overflow:hidden}
    .mh h1{color:#fff;font-size:18px;font-weight:800;margin:0;display:inline}
    .mh .db{float:right;background:rgba(255,255,255,.15);padding:3px 12px;border-radius:14px;color:#fff;font-size:11px;font-weight:500;border:1px solid rgba(255,255,255,.2);margin-top:2px}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:6px}
    .cc{background:#fff;border-radius:var(--r);padding:10px 14px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid var(--b);text-align:center}
    .cc .cv{font-size:26px;font-weight:900;line-height:1}
    .cc .cl{font-size:8px;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
    .cc.c1{border-top:3px solid #3182ce}.cc.c1 .cv{color:#2b6cb0}
    .cc.c2{border-top:3px solid #38a169}.cc.c2 .cv{color:#276749}
    .cc.c3{border-top:3px solid #805ad5}.cc.c3 .cv{color:#6b46c1}
    .cc.c4{border-top:3px solid #e53e3e}.cc.c4 .cv{color:#c53030}
    .stl{font-size:12px;font-weight:700;color:var(--p);margin:5px 0 2px 0;padding-left:10px;border-left:3px solid var(--pl)}
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}.stl.e{border-left-color:#d69e2e}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:8px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:7px;text-transform:uppercase;letter-spacing:.3px;padding:4px 5px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw.et thead th{background:linear-gradient(135deg,#b7791f,#d69e2e)}
    .tw tbody td{padding:3px 5px;border-bottom:1px solid #edf2f7;white-space:nowrap}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important;font-size:8px!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important;font-size:8px!important}
    .stTabs [data-baseweb="tab-list"]{gap:3px;background:#e2e8f0;padding:3px;border-radius:6px;margin-bottom:4px}
    .stTabs [data-baseweb="tab"]{border-radius:5px;padding:6px 14px;font-weight:600;font-size:11px}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--p)!important;box-shadow:0 2px 5px rgba(0,0,0,.07)}
    .sr{display:flex;align-items:center;padding:5px 10px;background:#fff;border-radius:5px;margin-bottom:1px;border:1px solid var(--b);font-size:9px}
    .sr .sn{font-weight:700;color:var(--p);min-width:220px;font-size:9px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .sc{padding:2px 8px;border-radius:12px;font-weight:800;font-size:10px;min-width:45px;text-align:center;margin:0 8px;color:#fff}
    .sr .sa{color:#718096;font-size:8px;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .stg{font-size:7px;color:#718096;min-width:55px;text-align:center;white-space:nowrap}
    .sr .sb{font-size:7px;font-weight:700;padding:1px 6px;border-radius:3px;white-space:nowrap}
    .ca{background:#fff;border-radius:var(--r);padding:10px;margin-top:3px;border:1px solid var(--b);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .ca .ct{font-size:11px;font-weight:700;margin-bottom:5px;padding-bottom:4px;border-bottom:1px solid var(--b)}
    .car{display:flex;align-items:center;margin-bottom:3px;font-size:8px}
    .car:last-child{margin-bottom:0}
    .car .cal{width:180px;font-weight:600;color:var(--p);text-align:right;padding-right:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .car .cab{flex:1;height:24px;background:#edf2f7;border-radius:4px;overflow:hidden}
    .car .caf{height:100%;border-radius:4px;transition:width .3s}
    .car .cav-out{font-size:8px;font-weight:800;color:#1a202c;min-width:55px;text-align:right;padding-left:5px}
    .gbr{display:flex;align-items:center;padding:3px 0;font-size:8px;border-bottom:1px solid #f7fafc}
    .gbr:last-child{border:none}
    .gbr-l{width:160px;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:7px}
    .gbr-g{display:flex;align-items:center;gap:4px;flex:1}
    .gbr-w{flex:1;height:20px;background:#edf2f7;border-radius:3px;overflow:hidden}
    .gbr-f{height:100%;border-radius:3px}
    .gb-p{background:linear-gradient(90deg,#2b6cb0,#4299e1)}
    .gb-q{background:linear-gradient(90deg,#276749,#48bb78)}
    .gbr-v{font-size:7px;font-weight:800;min-width:48px;text-align:right;color:#1a202c}
    .gbr-legend{display:flex;gap:14px;margin-bottom:5px;font-size:8px;font-weight:700}
    .gbr-legend span{display:flex;align-items:center;gap:4px}
    .gbr-legend i{display:inline-block;width:14px;height:14px;border-radius:2px}
    .cg{display:grid;grid-template-columns:1fr 1fr;gap:6px}
    .cg>div{background:#fff;border-radius:var(--r);padding:8px 10px;border:1px solid var(--b)}
    .cg .ct{font-size:10px;font-weight:700;margin-bottom:3px;padding-bottom:3px;border-bottom:1px solid var(--b)}
    .cgr{display:flex;align-items:center;padding:3px 0;font-size:9px;border-bottom:1px solid #f7fafc}
    .cgr:last-child{border:none}
    .cgr .rk{width:16px;font-weight:800;text-align:center}
    .cgr .pn{flex:1;font-weight:600;color:#1a202c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cgr .ps{font-weight:800;min-width:50px;text-align:right}
    .anl-tbl{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:9px;margin:0}
    .anl-tbl thead th{background:var(--p);color:#fff;font-weight:700;font-size:8px;padding:5px 7px;border:none;white-space:nowrap;position:sticky;top:0}
    .anl-tbl tbody td{padding:4px 7px;border-bottom:1px solid #edf2f7}
    .anl-tbl tbody tr:nth-child(even) td{background:#f7fafc}
    .anl-tbl tbody tr:hover td{background:#ebf8ff!important}
    .anl-tbl .tot td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important}
    .g-green{background:#c6efce;color:#006100;font-weight:600}
    .g-yellow{background:#ffeb9c;color:#9c6500;font-weight:600}
    .g-red{background:#ffc7ce;color:#9c0006;font-weight:600}
    .es{text-align:center;padding:12px;color:#718096;font-size:11px}
    .rh{display:flex;align-items:center;justify-content:space-between;margin-bottom:0}
    .rh .stl{margin:0}
    div[data-testid="stSidebar"]{background:linear-gradient(180deg,var(--p),#0f2744)}
    div[data-testid="stSidebar"]*{color:rgba(255,255,255,.9)!important}
    div[data-testid="stSidebar"] .stSelectbox label,div[data-testid="stSidebar"] .stMultiSelect label,div[data-testid="stSidebar"] .stDateInput label{color:rgba(255,255,255,.8)!important;font-weight:600;font-size:9px;text-transform:uppercase;letter-spacing:.5px}
    div[data-testid="stSidebar"] div[data-testid="stWidget"]{background:rgba(255,255,255,.08);border-radius:6px;padding:2px 6px;margin-bottom:2px;border:1px solid rgba(255,255,255,.1)}
    div[data-testid="stSidebar"] .stSelectbox>div>div,div[data-testid="stSidebar"] .stMultiSelect>div>div,div[data-testid="stSidebar"] .stDateInput>div>div{background:rgba(255,255,255,.95)!important;border-radius:5px}
    .stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--p),var(--pl));border:none;border-radius:6px;padding:8px 16px;font-weight:700;font-size:12px}
    .legend-box{background:#fff;border-radius:var(--r);padding:12px 16px;margin-top:8px;border:1px solid var(--b);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .legend-box .lt{font-size:11px;font-weight:800;color:var(--p);margin-bottom:8px;padding-bottom:5px;border-bottom:2px solid var(--pl)}
    .legend-box .ls{font-size:9px;font-weight:700;color:#2c5282;margin:5px 0 3px 0}
    .legend-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px 24px}
    .legend-item{display:flex;align-items:center;gap:8px;font-size:9px;padding:3px 0}
    .legend-item .lk{display:inline-block;min-width:60px;padding:2px 6px;border-radius:3px;font-weight:800;color:#fff;text-align:center;font-size:8px}
    .legend-item .ld{color:#4a5568;font-size:9px}
    ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#f1f1f1}::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:3px}
    @media(max-width:768px){.cr{grid-template-columns:repeat(2,1fr)}.mh h1{font-size:14px}.mh .db{float:none;display:block;margin-top:3px}.cg{grid-template-columns:1fr}.car .cal{width:110px}.gbr-l{width:100px}.legend-grid{grid-template-columns:1fr}}
    </style>""", unsafe_allow_html=True)

def main():
    try: locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    except:
        try: locale.setlocale(locale.LC_ALL, 'fr_FR')
        except: pass
    inject_custom_css()
    fichier_date = get_date_from_file()

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

    def calc_kpis(df_i, av_i, now, posts):
        res={}; df=df_i.copy(); av=av_i.copy()
        mp=["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
        df["Backlog preparation"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,mp)),"CARACTERISE","NON CARACTERISE")
        mplan=["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
        df["Backlog planification"]=np.where(df["Statut utilisateur"].apply(lambda x:contient_mot(x,mplan)),"CARACTERISE","NON CARACTERISE")
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
        except: return ""
        if c in ["OT préparation <1 mois","OT planification <1 mois","OT exécution <1 mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val>=80 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=75 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val<=15 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois"]: return "background:#c6efce;color:#006100;font-weight:600" if val<=5 else "background:#ffc7ce;color:#9c0006;font-weight:600"
        if c=="TAUX_REALISATION_CORRECTIF/PT": return "background:#c6efce;color:#006100;font-weight:600" if val>=85 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c=="appel avis approuvé": return "background:#c6efce;color:#006100;font-weight:600" if val>=95 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=90 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        if c in ["OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]: return "background:#c6efce;color:#006100;font-weight:600" if val>=100 else ("background:#ffeb9c;color:#9c6500;font-weight:600" if val>=95 else "background:#ffc7ce;color:#9c0006;font-weight:600")
        return ""
    def cs(v):
        try: val=float(str(v).replace(' %','').strip())
        except: return ""
        return "background:#c6efce;color:#006100;font-weight:700" if val>=90 else ("background:#ffeb9c;color:#9c6500;font-weight:700" if val>=80 else "background:#ffc7ce;color:#9c0006;font-weight:700")
    def kas(v):
        try: val=int(v)
        except: return ""
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
    def is_lb(k): return k in ["OT préparation >3 mois","OT planification >3 mois","OT exécution >3 mois","OT préparation 1mois< <3mois","OT planification 1mois< <3mois","OT exécution 1mois< <3mois"]

    def html_table(rows,cols,tc,sc_col=None):
        h='<table class="tw %s"><thead><tr>'%tc+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            rc="cb" if r.get("_t")=="cible" else ("tr" if r.get("_t")=="total" else "")
            h+='<tr class="%s">'%rc
            for c in cols:
                v=r.get(c,"")
                if r.get("_t")=="cible": h+='<td>%s</td>'%v
                else:
                    s=cs(v) if sc_col and c in sc_col else ks(v,c)
                    h+='<td style="%s">%s</td>'%(s or "",v)
            h+='</tr>'
        return h+'</tbody></table>'
    def html_ano(rows,cols):
        h='<table class="tw at"><thead><tr>'+''.join('<th>%s</th>'%c for c in cols)+'</tr></thead><tbody>'
        for r in rows:
            h+='<tr class="%s">'%("tr" if r.get("_t")=="total" else "")
            for c in cols: v=r.get(c,""); h+='<td style="%s">%s</td>'%(kas(v) or "",v)
            h+='</tr>'
        return h+'</tbody></table>'
    def html_synth(kpi_list,actuals,targets,act_map,accent):
        h=''
        for k in kpi_list:
            av,tv=actuals.get(k,0),targets.get(k,100)
            met=av<=tv if is_lb(k) else av>=tv
            sbg,sclr=("c6efce","#006100") if met else ("ffc7ce","#9c0006")
            scbg=accent if met else "#e53e3e"
            act="Objectif atteint" if met else act_map.get(k,"")
            h+='<div class="sr"><div class="sn">%s</div><div class="sc" style="background:%s">%.1f%%</div><div class="stg">Cible: %s%%</div><div class="sb" style="color:%s;background:%s">%s</div><div class="sa">%s</div></div>'%(k,scbg,av,tv,sclr,sbg,"ATTEINT" if met else "NON ATTEINT",act)
        return h
    def html_classement(scores,accent):
        sp=sorted(scores.items(),key=lambda x:x[1],reverse=True)
        met_p,not_p=[(p,s) for p,s in sp if s>=80],[(p,s) for p,s in sp if s<80]
        t5,b5=met_p[:5],not_p[-5:] if len(not_p)>5 else not_p
        h='<div class="cg"><div><div class="ct" style="color:#38a169">Top 5 — Objectif Atteint</div>'
        if t5:
            for i,(p,s) in enumerate(t5): h+='<div class="cgr"><span class="rk" style="color:%s">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(accent,i+1,p,cs("%.2f"%s),s)
        else: h+='<div style="padding:4px;font-size:8px;color:#718096">Aucun poste</div>'
        h+='</div><div><div class="ct" style="color:#e53e3e">Bottom 5 — Non Atteint</div>'
        if b5:
            for i,(p,s) in enumerate(reversed(b5)): h+='<div class="cgr"><span class="rk" style="color:#e53e3e">%s</span><span class="pn">%s</span><span class="ps" style="%s">%.2f%%</span></div>'%(len(b5)-i,p,cs("%.2f"%s),s)
        else: h+='<div style="padding:4px;font-size:8px;color:#38a169">Tous atteints</div>'
        h+='</div></div>'; return h
    def html_kpi_bars(kpi_list,actuals,targets,title,color_ok,color_fail):
        h='<div class="ca"><div class="ct" style="color:%s">%s</div>'%(color_ok,title)
        for k in kpi_list:
            av,tv=actuals.get(k,0),targets.get(k,100)
            met=av<=tv if is_lb(k) else av>=tv
            bw=min(max(av,0),100); bg=color_ok if met else color_fail
            h+='<div class="car"><div class="cal" style="width:260px">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>'%(k,bw,bg,av)
        h+='</div>'; return h
    def html_bars(data,title,color):
        h='<div class="ca"><div class="ct" style="color:%s">%s</div>'%(color,title)
        for label,val in sorted(data,key=lambda x:x[1],reverse=True):
            bw=min(max(val,0),100)
            h+='<div class="car"><div class="cal">%s</div><div class="cab"><div class="caf" style="width:%s%%;background:%s"></div></div><div class="cav-out">%.1f%%</div></div>'%(label,bw,color,val)
        h+='</div>'; return h
    def html_grouped_bars(posts,pscores,qscores,title):
        h='<div class="ca"><div class="ct" style="color:#1e3a5f">%s</div>'%title
        h+='<div class="gbr-legend"><span><i style="background:linear-gradient(90deg,#2b6cb0,#4299e1)"></i> Performance</span><span><i style="background:linear-gradient(90deg,#276749,#48bb78)"></i> Qualite</span></div>'
        sp2=sorted(posts,key=lambda x:(pscores.get(x,0)+qscores.get(x,0))/2,reverse=True)
        for p in sp2:
            pv,qv=pscores.get(p,0),qscores.get(p,0)
            pw,qw=min(max(pv,0),100),min(max(qv,0),100)
            h+='<div class="gbr"><div class="gbr-l">%s</div><div class="gbr-g"><div class="gbr-w"><div class="gbr-f gb-p" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div><div class="gbr-w"><div class="gbr-f gb-q" style="width:%s%%"></div></div><div class="gbr-v">%.1f%%</div></div></div>'%(p,pw,pv,qw,qv)
        h+='</div>'; return h

    # === 4. Chart donut eclaté avec height=450 et autosize ===
    def anl_pie_chart(data, names_col, values_col, title, colors=None):
        if data.empty: return None
        vals = data[values_col].values
        total = vals.sum()
        if total == 0: return None
        pct = vals / total * 100
        pull = []
        for p in pct:
            if p < 2: pull.append(0.18)
            elif p < 5: pull.append(0.12)
            elif p < 10: pull.append(0.06)
            else: pull.append(0.02)
        fig = px.pie(data, names=names_col, values=values_col, title=title,
                     color_discrete_sequence=colors or px.colors.qualitative.Set2,
                     hole=0.38, pull=pull)
        fig.update_traces(textposition='outside', textinfo='label+percent+value',
                          textfont_size=9, insidetextorientation='radial',
                          marker=dict(line=dict(color='#fff', width=1.5)))
        fig.update_layout(margin=dict(t=50,b=70,l=10,r=10), height=450, autosize=True,
                          title_font_size=12,
                          legend=dict(font_size=8, orientation="h", yanchor="bottom", y=-0.22,
                                      itemwidth=28, itemheight=14),
                          uniformtext_minsize=8, uniformtext_mode='hide')
        return fig

    def anl_html_table(df_out,pct_col=None,pct_thresh=(80,60)):
        h='<table class="anl-tbl"><thead><tr>'
        for c in df_out.columns: h+='<th>%s</th>'%c
        h+='</tr></thead><tbody>'
        for idx,row in df_out.iterrows():
            is_tot=str(idx)=="TOTAL" or str(row.iloc[0])=="TOTAL"
            rc="tot" if is_tot else ""
            h+='<tr class="%s">'%rc
            for c in df_out.columns:
                v=row[c]; s=""
                if pct_col and c==pct_col and not is_tot:
                    try:
                        pv=float(str(v).replace('%','').strip())
                        s="g-green" if pv>=pct_thresh[0] else ("g-yellow" if pv>=pct_thresh[1] else "g-red")
                    except: pass
                if isinstance(v,float): v=round(v,1)
                h+='<td class="%s">%s</td>'%(s,v)
            h+='</tr>'
        return h+'</tbody></table>'
    def export_btn(df,filename):
        buf=io.BytesIO()
        df.to_excel(buf,index=False,engine='openpyxl')
        buf.seek(0)
        st.download_button("📥 Exporter Excel",data=buf,file_name=filename,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===================== SIDEBAR =====================
    with st.sidebar:
        # === 3. Bouton pour masquer/afficher les filtres ===
        show_filters = st.checkbox("⚙️ Afficher les filtres", value=True, key="show_filters")
        if not show_filters:
            st.markdown('<div style="text-align:center;padding:20px 0;color:rgba(255,255,255,.5);font-size:10px">Cliquez sur la flèche ▶ en haut<br>pour rouvrir les filtres</div>', unsafe_allow_html=True)
            st.stop()

        st.markdown("""<div style="padding:8px 0 4px 0"><div style="font-size:18px;margin-bottom:2px">⚙️</div><div style="font-size:12px;font-weight:800;color:white">Filtres & Parametres</div><div style="font-size:8px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Configuration</div></div>""",unsafe_allow_html=True)
        st.markdown("---")
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
                except: pass
            st.markdown("""<div style="background:rgba(255,255,255,.1);padding:5px 8px;border-radius:6px;border:1px solid rgba(255,255,255,.15)"><div style="font-size:7px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:1px">Donnees</div><div style="font-size:10px;color:white;font-weight:600;margin-top:1px">📅 %s</div></div>"""%fichier_date,unsafe_allow_html=True)
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

    if not unf or (ot_f is not None and av_f is not None):
        try:
            if unf:
                raw_ot=pd.read_excel(ot_f); raw_av=pd.read_excel(av_f)
            else:
                raw_ot=pd.read_excel("ot.xlsx"); raw_av=pd.read_excel("avis.xlsx")
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
            res=calc_kpis(df,avdf,now,vp)
            ckdf=res['ckdf']; dfp=res['dfp']
            res_d=calc_kpis(df_dash,avdf,now,vp)
            ckdf_d=res_d['ckdf']
            qk=["TAUX_REALISATION_CORRECTIF/PT","OT préparation <1 mois","OT préparation >3 mois","OT préparation 1mois< <3mois","OT planification <1 mois","OT planification >3 mois","OT planification 1mois< <3mois","OT exécution <1 mois","OT exécution >3 mois","OT exécution 1mois< <3mois"]
            pk=["appel avis approuvé","OT LANC ESTIME","Backlog préparation caractérisé","Backlog planification caractérisé","OT CONFIME","OT_COR_EGAL"]
            cible={"TAUX_REALISATION_CORRECTIF/PT":85,"OT préparation <1 mois":80,"OT préparation >3 mois":5,"OT préparation 1mois< <3mois":15,"OT planification <1 mois":80,"OT planification >3 mois":5,"OT planification 1mois< <3mois":15,"OT exécution <1 mois":80,"OT exécution >3 mois":5,"OT exécution 1mois< <3mois":15,"appel avis approuvé":95,"OT LANC ESTIME":100,"Backlog préparation caractérisé":100,"Backlog planification caractérisé":100,"OT CONFIME":100,"OT_COR_EGAL":100}
            act_map={"TAUX_REALISATION_CORRECTIF/PT":"Ameliorer le taux de realisation des OT.","OT préparation <1 mois":"Reduire l'age de preparation des OT (< 1 mois).","OT préparation >3 mois":"Traiter les OT avec preparation > 3 mois.","OT planification <1 mois":"Reduire l'age de planification des OT (< 1 mois).","OT planification >3 mois":"Traiter les OT avec planification > 3 mois.","OT exécution <1 mois":"Reduire l'age d'execution des OT (< 1 mois).","OT exécution >3 mois":"Traiter les OT avec execution > 3 mois.","OT LANC ESTIME":"Estimer les couts des OT lances.","Backlog préparation caractérisé":"Caracteriser le backlog de preparation.","Backlog planification caractérisé":"Caracteriser le backlog de planification.","OT CONFIME":"Confirmer les OT termines.","OT_COR_EGAL":"Rapprocher les couts reels et budgetes.","appel avis approuvé":"Creer un OT pour les avis sans ordre."}

            pscores={}; qscores={}
            for poste in ckdf.index:
                r=ckdf.loc[poste]
                pscores[poste]=(sum(gscore(k,r[k],cible[k]) for k in qk if k in r.index)/len(qk)*100) if qk else 0
                qscores[poste]=(sum(gscore(k,r[k],cible[k]) for k in pk if k in r.index)/len(pk)*100) if pk else 0
            pa={k:round(ckdf[k].mean(),2) for k in qk}
            qa={k:round(ckdf[k].mean(),2) for k in pk}
            pscores_d={}; qscores_d={}
            for poste in ckdf_d.index:
                r=ckdf_d.loc[poste]
                pscores_d[poste]=(sum(gscore(k,r[k],cible[k]) for k in qk if k in r.index)/len(qk)*100) if qk else 0
                qscores_d[poste]=(sum(gscore(k,r[k],cible[k]) for k in pk if k in r.index)/len(pk)*100) if pk else 0
            pa_d={k:round(ckdf_d[k].mean(),2) for k in qk}
            qa_d={k:round(ckdf_d[k].mean(),2) for k in pk}

            all_ano=[]
            sub_p={"TAUX_REALISATION_CORRECTIF/PT":lambda d:d[(d["Nº appel pl.entret."].fillna(0)==0)&(~d["Statut OT"].isin(["CLOT","TCLO"]))],"OT préparation <1 mois":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]!="<1 mois")],"OT préparation >3 mois":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["ap"]==">3 mois")],"OT planification <1 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]!="<1 mois")],"OT planification >3 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)&(d["alp"]==">3 mois")],"OT exécution <1 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]!="<1 mois")],"OT exécution >3 mois":lambda d:d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)&(d["aex"]==">3 mois")]}
            sub_q={"OT LANC ESTIME":lambda d:d[(d["Statut OT"]=="LANC")&(d["OT LANC ESTIME"]=="NON")],"Backlog préparation caractérisé":lambda d:d[(d["Statut OT"]=="CRÉÉ")&(d["Backlog preparation"]=="NON CARACTERISE")],"Backlog planification caractérisé":lambda d:d[(d["Statut OT"]=="LANC")&(d["Backlog planification"]=="NON CARACTERISE")],"OT CONFIME":lambda d:d[d["OT CONFIME"]=="NON"],"OT_COR_EGAL":lambda d:d[d["OT_COR_EGAL"]=="NON"]}
            for poste in vp:
                if poste not in dfp["Poste travail princ."].values: continue
                dp=dfp[dfp["Poste travail princ."]==poste]
                for kn,sf in sub_p.items():
                    vk=ckdf.loc[poste,kn] if poste in ckdf.index else 100
                    if pd.notna(vk) and vk<cible[kn]:
                        cnt=len(sf(dp))
                        if cnt>0: all_ano.append({"Poste":poste,"KPI":kn,"Nb":cnt,"Type":"P"})
                for kn,sf in sub_q.items():
                    vk=ckdf.loc[poste,kn] if poste in ckdf.index else 100
                    if pd.notna(vk) and vk<cible[kn]:
                        cnt=len(sf(dp))
                        if cnt>0: all_ano.append({"Poste":poste,"KPI":kn,"Nb":cnt,"Type":"Q"})
                vk_av=ckdf.loc[poste,"appel avis approuvé"] if poste in ckdf.index else 100
                if pd.notna(vk_av) and vk_av<cible["appel avis approuvé"]:
                    cnt=len(res['avf'][res['avf']["Poste travail princ."]==poste])
                    if cnt>0: all_ano.append({"Poste":poste,"KPI":"appel avis approuvé","Nb":cnt,"Type":"Q"})

            def build_ano(ano_list,kpi_list):
                if not ano_list: return [],[]
                adf=pd.DataFrame(ano_list)
                pv=adf.pivot_table(index="Poste",columns="KPI",values="Nb",aggfunc="sum",fill_value=0).astype(int)
                pv["Total"]=pv.sum(axis=1); tot=pv.sum()
                cols=[c for c in kpi_list if c in pv.columns]+["Total"]; rows=[]
                for idx in pv.index:
                    r={"_t":"n","Poste de travail":idx}
                    for c in cols: r[c]=pv.loc[idx,c]
                    rows.append(r)
                tr={"_t":"total","Poste de travail":"Total general"}
                for c in cols: tr[c]=int(tot[c])
                rows.append(tr); return ["Poste de travail"]+cols,rows
            ano_p_c,ano_p_r=build_ano([a for a in all_ano if a["Type"]=="P"],qk)
            ano_q_c,ano_q_r=build_ano([a for a in all_ano if a["Type"]=="Q"],pk)

            def build_kpi(kpi_list,scores,sname):
                sp2=sorted(scores.keys(),key=lambda x:scores[x],reverse=True)
                cols=["Poste de travail"]+kpi_list+[sname]; rows=[]
                cr={"_t":"cible","Poste de travail":"CIBLE"}
                for k in kpi_list: cr[k]=cible[k]
                cr[sname]="100.00 %"; rows.append(cr)
                for p in sp2:
                    r={"_t":"n","Poste de travail":p}
                    for k in kpi_list: r[k]=round(ckdf.loc[p,k],2) if p in ckdf.index else ""
                    r[sname]="%.2f %%"%scores[p]; rows.append(r)
                tr={"_t":"total","Poste de travail":"Total general"}
                for k in kpi_list: tr[k]=round(ckdf[k].mean(),2)
                tr[sname]="%.2f %%"%(np.mean(list(scores.values())) if scores else 0)
                rows.append(tr); return cols,rows
            pcols,prows=build_kpi(qk,pscores,"Score Performance")
            qcols,qrows=build_kpi(pk,qscores,"Score Qualite")
            save_kpis_to_excel(prows,pcols,qrows,qcols,ano_p_r,ano_p_c,ano_q_r,ano_q_c,fichier_date)

            df_sc_d=pd.DataFrame([{"Poste":p,"Perf":pscores_d[p],"Qual":qscores_d[p],"Metier":get_metier(p),"Atelier":get_atelier(p),"Division":get_division(p)} for p in vp if p in pscores_d])
            by_at=df_sc_d.groupby("Atelier")[["Perf","Qual"]].mean().round(1) if not df_sc_d.empty else pd.DataFrame()
            by_mt=df_sc_d.groupby("Metier")[["Perf","Qual"]].mean().round(1) if not df_sc_d.empty else pd.DataFrame()
            by_div=df_sc_d.groupby("Division")[["Perf","Qual"]].mean().round(1) if not df_sc_d.empty else pd.DataFrame()
            total_ot=len(df); avg_p=np.mean(list(pscores.values())) if pscores else 0
            avg_q=np.mean(list(qscores.values())) if qscores else 0; total_ano=sum(a["Nb"] for a in all_ano)

            desig_col=None
            for cn in ["Désignation du travail","Designation du travail","Désignation","Designation","Description"]:
                if cn in dfp.columns: desig_col=cn; break

            st.markdown('<div class="mh"><h1>📊 KPI Dashboard MC & FEED</h1><div class="db">📅 %s</div></div>'%fichier_date,unsafe_allow_html=True)
            st.markdown("""<div class="cr">
            <div class="cc c1"><div class="cv">%s</div><div class="cl">Total OT Analyses</div></div>
            <div class="cc c2"><div class="cv">%.1f%%</div><div class="cl">Score Performance</div></div>
            <div class="cc c3"><div class="cv">%.1f%%</div><div class="cl">Score Qualite</div></div>
            <div class="cc c4"><div class="cv">%s</div><div class="cl">Total Anomalies</div></div>
            </div>"""%(total_ot,avg_p,avg_q,total_ano),unsafe_allow_html=True)

            tab0,tab1,tab2,tab3,tab4=st.tabs(["📊 TABLEAU DE BORD","📈 INDICATEURS PERFORMANCE","✅ INDICATEUR QUALITE","🔬 ANALYSE","📉 SUIVI ÉVOLUTION"])

            with tab0:
                st.markdown('<div class="stl p">Vue d\'ensemble par poste</div>',unsafe_allow_html=True)
                st.markdown(html_grouped_bars(vp,pscores_d,qscores_d,"Performance & Qualite par Poste de Travail"),unsafe_allow_html=True)
                st.markdown('<div class="stl p">🏭 Par Atelier</div>',unsafe_allow_html=True)
                if not by_at.empty:
                    c1,c2=st.columns(2)
                    with c1: st.markdown(html_bars([(i,r["Perf"]) for i,r in by_at.iterrows()],"Performance par Atelier","#2b6cb0"),unsafe_allow_html=True)
                    with c2: st.markdown(html_bars([(i,r["Qual"]) for i,r in by_at.iterrows()],"Qualite par Atelier","#276749"),unsafe_allow_html=True)
                st.markdown('<div class="stl p">🔧 Par Metier</div>',unsafe_allow_html=True)
                if not by_mt.empty:
                    c1,c2=st.columns(2)
                    with c1: st.markdown(html_bars([(i,r["Perf"]) for i,r in by_mt.iterrows()],"Performance par Metier","#2b6cb0"),unsafe_allow_html=True)
                    with c2: st.markdown(html_bars([(i,r["Qual"]) for i,r in by_mt.iterrows()],"Qualite par Metier","#276749"),unsafe_allow_html=True)
                st.markdown('<div class="stl p">🏢 Par Division</div>',unsafe_allow_html=True)
                if not by_div.empty:
                    c1,c2=st.columns(2)
                    with c1: st.markdown(html_bars([(i,r["Perf"]) for i,r in by_div.iterrows()],"Performance par Division","#2b6cb0"),unsafe_allow_html=True)
                    with c2: st.markdown(html_bars([(i,r["Qual"]) for i,r in by_div.iterrows()],"Qualite par Division","#276749"),unsafe_allow_html=True)
                st.markdown('<div class="stl p">Synthese globale</div>',unsafe_allow_html=True)
                st.markdown(html_synth(qk,pa_d,cible,act_map,"#2b6cb0"),unsafe_allow_html=True)
                st.markdown(html_synth(pk,qa_d,cible,act_map,"#276749"),unsafe_allow_html=True)
                st.markdown('<div class="stl p">Classement</div>',unsafe_allow_html=True)
                st.markdown(html_classement(pscores_d,"#2b6cb0"),unsafe_allow_html=True)
                st.markdown(html_classement(qscores_d,"#276749"),unsafe_allow_html=True)

            with tab1:
                choix_p=st.radio("Choix",["📈 Indicateurs","⚠️ Anomalies"],horizontal=True,key="choix_p")
                if choix_p=="📈 Indicateurs":
                    st.markdown('<div class="rh"><div class="stl p">Indicateurs de Performance par Poste</div></div>',unsafe_allow_html=True)
                    st.markdown(html_table(prows,pcols,"pt",["Score Performance"]),unsafe_allow_html=True)
                else:
                    if ano_p_c:
                        st.markdown('<div class="rh"><div class="stl a">Anomalies Performance</div></div>',unsafe_allow_html=True)
                        st.markdown(html_ano(ano_p_r,ano_p_c),unsafe_allow_html=True)
                    else: st.markdown('<div class="es">✅ Aucune anomalie.</div>',unsafe_allow_html=True)
                st.markdown('<div class="stl p">Synthese</div>',unsafe_allow_html=True)
                st.markdown(html_synth(qk,pa,cible,act_map,"#2b6cb0"),unsafe_allow_html=True)
                st.markdown(html_kpi_bars(qk,pa,cible,"Performance Globale","#2b6cb0","#e53e3e"),unsafe_allow_html=True)
                st.markdown(html_classement(pscores,"#2b6cb0"),unsafe_allow_html=True)

            with tab2:
                choix_q=st.radio("Choix",["✅ Indicateurs","⚠️ Anomalies"],horizontal=True,key="choix_q")
                if choix_q=="✅ Indicateurs":
                    st.markdown('<div class="rh"><div class="stl q">Indicateurs de Qualite par Poste</div></div>',unsafe_allow_html=True)
                    st.markdown(html_table(qrows,qcols,"qt",["Score Qualite"]),unsafe_allow_html=True)
                else:
                    if ano_q_c:
                        st.markdown('<div class="rh"><div class="stl a">Anomalies Qualite</div></div>',unsafe_allow_html=True)
                        st.markdown(html_ano(ano_q_r,ano_q_c),unsafe_allow_html=True)
                    else: st.markdown('<div class="es">✅ Aucune anomalie.</div>',unsafe_allow_html=True)
                st.markdown('<div class="stl q">Synthese</div>',unsafe_allow_html=True)
                st.markdown(html_synth(pk,qa,cible,act_map,"#276749"),unsafe_allow_html=True)
                st.markdown(html_kpi_bars(pk,qa,cible,"Qualite Globale","#276749","#e53e3e"),unsafe_allow_html=True)
                st.markdown(html_classement(qscores,"#276749"),unsafe_allow_html=True)

            with tab3:
                st.markdown('<div class="stl c">1. Analyse des OT OMS</div>',unsafe_allow_html=True)
                if desig_col:
                    oms_df=dfp[dfp[desig_col].astype(str).str.contains("OMS",case=False,na=False)]
                    if not oms_df.empty:
                        oms_pv=pd.pivot_table(oms_df,index="Poste travail princ.",columns="Statut OT",values="Ordre",aggfunc="count",fill_value=0)
                        oms_pv["Total"]=oms_pv.sum(axis=1); oms_pv=oms_pv.sort_values("Total",ascending=False)
                        oms_tot=oms_pv.sum(); oms_tot.name="TOTAL"; oms_pv=pd.concat([oms_pv,oms_tot.to_frame().T])
                        oms_exp=oms_pv.reset_index().rename(columns={"Poste travail princ.":"Poste de Travail"})
                        c1,c2=st.columns([1.2,1])
                        with c1:
                            st.markdown(anl_html_table(oms_exp),unsafe_allow_html=True)
                            export_btn(oms_exp,"analyse_oms.xlsx")
                        with c2:
                            op=oms_df["Statut OT"].value_counts().reset_index(); op.columns=["Statut","Nombre"]
                            fig=anl_pie_chart(op,"Statut","Nombre","Repartition OT OMS")
                            if fig: st.plotly_chart(fig,use_container_width=True)
                    else: st.markdown('<div class="es">Aucun OT OMS.</div>',unsafe_allow_html=True)
                else: st.markdown('<div class="es">Colonne Designation non trouvee.</div>',unsafe_allow_html=True)
                st.markdown("---")
                st.markdown('<div class="stl c">2. OT Thermographiques</div>',unsafe_allow_html=True)
                if desig_col:
                    th_df=dfp[dfp[desig_col].astype(str).str.contains("THERMO",case=False,na=False)]
                    if not th_df.empty:
                        th_pv=pd.pivot_table(th_df,index="Poste travail princ.",columns="Statut OT",values="Ordre",aggfunc="count",fill_value=0)
                        th_pv["Total"]=th_pv.sum(axis=1); th_pv=th_pv.sort_values("Total",ascending=False)
                        th_tot=th_pv.sum(); th_tot.name="TOTAL"; th_pv=pd.concat([th_pv,th_tot.to_frame().T])
                        th_exp=th_pv.reset_index().rename(columns={"Poste travail princ.":"Poste de Travail"})
                        c1,c2=st.columns([1.2,1])
                        with c1:
                            st.markdown(anl_html_table(th_exp),unsafe_allow_html=True)
                            export_btn(th_exp,"analyse_thermo.xlsx")
                        with c2:
                            tp=th_df["Statut OT"].value_counts().reset_index(); tp.columns=["Statut","Nombre"]
                            fig=anl_pie_chart(tp,"Statut","Nombre","Repartition OT Thermo")
                            if fig: st.plotly_chart(fig,use_container_width=True)
                    else: st.markdown('<div class="es">Aucun OT Thermo.</div>',unsafe_allow_html=True)
                st.markdown("---")
                st.markdown('<div class="stl c">3. Backlog Preparation</div>',unsafe_allow_html=True)
                bl_prep=dfp[dfp["Statut OT"]=="CRÉÉ"]
                if not bl_prep.empty:
                    bl_p_pv=pd.pivot_table(bl_prep,index="Poste travail princ.",columns="Backlog preparation",values="Ordre",aggfunc="count",fill_value=0)
                    for c in ["CARACTERISE","NON CARACTERISE"]:
                        if c not in bl_p_pv.columns: bl_p_pv[c]=0
                    bl_p_pv["Total"]=bl_p_pv.sum(axis=1)
                    bl_p_pv["% Caracterisation"]=np.where(bl_p_pv["Total"]==0,0,(bl_p_pv.get("CARACTERISE",0)/bl_p_pv["Total"]*100)).round(1)
                    bl_p_pv=bl_p_pv[["CARACTERISE","NON CARACTERISE","Total","% Caracterisation"]].sort_values("Total",ascending=False)
                    bl_p_tot=bl_p_pv.sum(); bl_p_tot["% Caracterisation"]=round(bl_p_tot["CARACTERISE"]/bl_p_tot["Total"]*100,1) if bl_p_tot["Total"]>0 else 0
                    bl_p_tot.name="TOTAL"; bl_p_pv=pd.concat([bl_p_pv,bl_p_tot.to_frame().T])
                    bl_p_exp=bl_p_pv.reset_index().rename(columns={"Poste travail princ.":"Poste de Travail"})
                    c1,c2=st.columns([1.2,1])
                    with c1:
                        st.markdown(anl_html_table(bl_p_exp,"% Caracterisation",(80,60)),unsafe_allow_html=True)
                        export_btn(bl_p_exp,"analyse_backlog_prep.xlsx")
                    with c2:
                        mp_t=["CRPR ATPD","CRPR ATMR","CRPR ATER","CRPR ATRS","CRPR ATMO","ATPD","ATMR","ATER","ATRS","ATMO"]
                        def gcp(su):
                            s=str(su).upper()
                            for t in mp_t:
                                if t in s: return t
                            return "NON CARACTERISE"
                        bpc=bl_prep.copy(); bpc["Type Caract."]=bpc["Statut utilisateur"].apply(gcp)
                        bp_pie=bpc["Type Caract."].value_counts().reset_index(); bp_pie.columns=["Caracterisation","Nombre"]
                        fig=anl_pie_chart(bp_pie,"Caracterisation","Nombre","Types Caract. Backlog Prep.")
                        if fig: st.plotly_chart(fig,use_container_width=True)
                else: st.markdown('<div class="es">Aucun backlog preparation.</div>',unsafe_allow_html=True)
                st.markdown("---")
                st.markdown('<div class="stl c">4. Backlog Planification</div>',unsafe_allow_html=True)
                bl_plan=dfp[dfp["Statut OT"]=="LANC"]
                if not bl_plan.empty:
                    bl_pl_pv=pd.pivot_table(bl_plan,index="Poste travail princ.",columns="Backlog planification",values="Ordre",aggfunc="count",fill_value=0)
                    for c in ["CARACTERISE","NON CARACTERISE"]:
                        if c not in bl_pl_pv.columns: bl_pl_pv[c]=0
                    bl_pl_pv["Total"]=bl_pl_pv.sum(axis=1)
                    bl_pl_pv["% Caracterisation"]=np.where(bl_pl_pv["Total"]==0,0,(bl_pl_pv.get("CARACTERISE",0)/bl_pl_pv["Total"]*100)).round(1)
                    bl_pl_pv=bl_pl_pv[["CARACTERISE","NON CARACTERISE","Total","% Caracterisation"]].sort_values("Total",ascending=False)
                    bl_pl_tot=bl_pl_pv.sum(); bl_pl_tot["% Caracterisation"]=round(bl_pl_tot["CARACTERISE"]/bl_pl_tot["Total"]*100,1) if bl_pl_tot["Total"]>0 else 0
                    bl_pl_tot.name="TOTAL"; bl_pl_pv=pd.concat([bl_pl_pv,bl_pl_tot.to_frame().T])
                    bl_pl_exp=bl_pl_pv.reset_index().rename(columns={"Poste travail princ.":"Poste de Travail"})
                    c1,c2=st.columns([1.2,1])
                    with c1:
                        st.markdown(anl_html_table(bl_pl_exp,"% Caracterisation",(80,60)),unsafe_allow_html=True)
                        export_btn(bl_pl_exp,"analyse_backlog_plan.xlsx")
                    with c2:
                        mpl_t=["ATPL ATEI","ATPL ATAL","ATPL ATER","ATPL AGAR","ATPL ATHS","ATEI","ATAL","ATAS","AGAR","ATHS"]
                        def gcl(su):
                            s=str(su).upper()
                            for t in mpl_t:
                                if t in s: return t
                            return "NON CARACTERISE"
                        blc=bl_plan.copy(); blc["Type Caract."]=blc["Statut utilisateur"].apply(gcl)
                        bl_pie=blc["Type Caract."].value_counts().reset_index(); bl_pie.columns=["Caracterisation","Nombre"]
                        fig=anl_pie_chart(bl_pie,"Caracterisation","Nombre","Types Caract. Backlog Plan.")
                        if fig: st.plotly_chart(fig,use_container_width=True)
                else: st.markdown('<div class="es">Aucun backlog planification.</div>',unsafe_allow_html=True)
                st.markdown("---")
                st.markdown('<div class="stl c">5. Backlog Execution</div>',unsafe_allow_html=True)
                if not dfp.empty:
                    def ces(d):
                        ln=d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==0)].shape[0]
                        sp_=d[(d["Statut OT"]=="LANC")&(d["Contient SOPL"]==1)].shape[0]
                        cl=d[d["Statut OT"]=="CLOT"].shape[0]; tc=d[d["Statut OT"]=="TCLO"].shape[0]
                        t=len(d); pc=round((cl+tc)/t*100,1) if t>0 else 0
                        return pd.Series({"LANC":ln,"SOPL":sp_,"CLOT":cl,"TCLO":tc,"Total":t,"% Cloture":pc})
                    bl_ex=dfp.groupby("Poste travail princ.").apply(ces).astype(int)
                    bl_ex["% Cloture"]=dfp.groupby("Poste travail princ.").apply(ces)["% Cloture"]
                    bl_ex=bl_ex.sort_values("Total",ascending=False)
                    bl_ex_tot=bl_ex.sum(); bl_ex_tot["% Cloture"]=round((bl_ex_tot["CLOT"]+bl_ex_tot["TCLO"])/bl_ex_tot["Total"]*100,1) if bl_ex_tot["Total"]>0 else 0
                    bl_ex_tot.name="TOTAL"; bl_ex=pd.concat([bl_ex,bl_ex_tot.to_frame().T])
                    bl_ex_exp=bl_ex.reset_index().rename(columns={"Poste travail princ.":"Poste de Travail"})
                    c1,c2=st.columns([1.2,1])
                    with c1:
                        st.markdown(anl_html_table(bl_ex_exp,"% Cloture",(80,60)),unsafe_allow_html=True)
                        export_btn(bl_ex_exp,"analyse_backlog_exec.xlsx")
                    with c2:
                        epd=pd.DataFrame({"Statut":["LANC","SOPL","CLOT","TCLO"],"Nombre":[int(bl_ex_tot["LANC"]),int(bl_ex_tot["SOPL"]),int(bl_ex_tot["CLOT"]),int(bl_ex_tot["TCLO"])]})
                        epd=epd[epd["Nombre"]>0]
                        if not epd.empty:
                            fig=anl_pie_chart(epd,"Statut","Nombre","Repartition Backlog Exec.",["#4299e1","#48bb78","#38a169","#2b6cb0"])
                            if fig: st.plotly_chart(fig,use_container_width=True)
                else: st.markdown('<div class="es">Aucune donnee.</div>',unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("""<div class="legend-box">
                <div class="lt">📖 Légende des Codes de Caractérisation</div>
                <div class="legend-grid">
                <div>
                <div class="ls" style="color:#276749">🔧 Caractérisation Préparation</div>
                <div class="legend-item"><span class="lk" style="background:#2b6cb0">ATPD</span><span class="ld">Attente PDR (Pièce Détachée Rechange)</span></div>
                <div class="legend-item"><span class="lk" style="background:#2c5282">ATMR</span><span class="ld">Attente Marche</span></div>
                <div class="legend-item"><span class="lk" style="background:#4299e1">ATRS</span><span class="ld">Attente Ressources</span></div>
                <div class="legend-item"><span class="lk" style="background:#63b3ed">ATMO</span><span class="ld">Attente Moyens ou Outillage</span></div>
                <div class="legend-item"><span class="lk" style="background:#90cdf4">ATER</span><span class="ld">Attente Équipement de Rechange</span></div>
                </div>
                <div>
                <div class="ls" style="color:#2b6cb0">📋 Caractérisation Planification</div>
                <div class="legend-item"><span class="lk" style="background:#d69e2e">ATEI</span><span class="ld">Attente Arrêt Équipement</span></div>
                <div class="legend-item"><span class="lk" style="background:#ecc94b">ATAL</span><span class="ld">Attente Arrêt Ligne</span></div>
                <div class="legend-item"><span class="lk" style="background:#f6e05e;color:#744210">ATAS</span><span class="ld">Attente Arrêt Site</span></div>
                <div class="legend-item"><span class="lk" style="background:#ed8936">AGAR</span><span class="ld">Attente Grand Arrêt de Révision</span></div>
                <div class="legend-item"><span class="lk" style="background:#e53e3e">ATHS</span><span class="ld">Attente HSE (Hygiène Sécurité Environnement)</span></div>
                </div>
                </div>
                </div>""", unsafe_allow_html=True)

            with tab4:
                kpis_path = os.path.join("kpis", "indicateurs_kpis.xlsx")
                if not os.path.exists(kpis_path):
                    st.markdown('<div class="es">📁 Fichier kpis/indicateurs_kpis.xlsx introuvable.</div>',unsafe_allow_html=True)
                else:
                    all_data = parse_kpis_excel(kpis_path)
                    if len(all_data) < 2:
                        st.markdown('<div class="es">📊 Au moins 2 dates necessaires. Dates dispo : %d</div>' % len(all_data),unsafe_allow_html=True)
                    else:
                        sorted_dates = sorted(all_data.keys(), key=parse_date_sheet, reverse=True)
                        date_n = sorted_dates[0]; date_n1 = sorted_dates[1]
                        data_n = all_data[date_n]; data_n1 = all_data[date_n1]

                        st.markdown('<div class="ca" style="margin-bottom:6px"><div class="ct" style="color:#d69e2e">📉 Comparaison : %s → %s</div><div style="font-size:9px;color:#718096">Analyse automatique des ecarts entre les deux derniers enregistrements</div></div>' % (date_n1, date_n), unsafe_allow_html=True)

                        all_postes = sorted(set(list(data_n.keys()) + list(data_n1.keys())))
                        comp_rows = []
                        for p in all_postes:
                            dn_=data_n.get(p,{"Score Performance":0,"Score Qualite":0,"Ano Perf":0,"Ano Qual":0})
                            dn1_=data_n1.get(p,{"Score Performance":0,"Score Qualite":0,"Ano Perf":0,"Ano Qual":0})
                            ep=round(dn_["Score Performance"]-dn1_["Score Performance"],2)
                            eq=round(dn_["Score Qualite"]-dn1_["Score Qualite"],2)
                            ea=int(dn_["Ano Perf"]+dn_["Ano Qual"])-int(dn1_["Ano Perf"]+dn1_["Ano Qual"])
                            ti,tk=get_tendance(ep,eq,ea)
                            comp_rows.append({"Poste":p,"Perf N-1":round(dn1_["Score Performance"],2),"Perf N":round(dn_["Score Performance"],2),"Ecart Perf":ep,"Qual N-1":round(dn1_["Score Qualite"],2),"Qual N":round(dn_["Score Qualite"],2),"Ecart Qual":eq,"Ano N-1":int(dn1_["Ano Perf"]+dn1_["Ano Qual"]),"Ano N":int(dn_["Ano Perf"]+dn_["Ano Qual"]),"Ecart Ano":ea,"Tendance":tk,"Tendance Icon":ti})
                        comp_df=pd.DataFrame(comp_rows)

                        nb_imp=len(comp_df[comp_df["Tendance"]=="amelioration"])
                        nb_deg=len(comp_df[comp_df["Tendance"]=="degradation"])
                        nb_stb=len(comp_df[comp_df["Tendance"]=="stable"])
                        st.markdown("""<div class="cr">
                        <div class="cc c2"><div class="cv">%s</div><div class="cl">Amelioration</div></div>
                        <div class="cc c4"><div class="cv">%s</div><div class="cl">Degradation</div></div>
                        <div class="cc" style="border-top:3px solid #d69e2e"><div class="cv" style="color:#d69e2e">%s</div><div class="cl">Stable</div></div>
                        <div class="cc c1"><div class="cv">%s</div><div class="cl">Total Postes</div></div>
                        </div>"""%(nb_imp,nb_deg,nb_stb,len(comp_df)),unsafe_allow_html=True)

                        st.markdown('<div class="stl e">Tableau Comparatif par Poste de Travail</div>',unsafe_allow_html=True)
                        h='<table class="tw et"><thead><tr><th>Poste</th><th>Perf N-1</th><th>Perf N</th><th>Ecart Perf</th><th>Qual N-1</th><th>Qual N</th><th>Ecart Qual</th><th>Ano N-1</th><th>Ano N</th><th>Ecart Ano</th><th>Tendance</th></tr></thead><tbody>'
                        for _,rw in comp_df.iterrows():
                            eps="g-green" if rw["Ecart Perf"]>0 else ("g-red" if rw["Ecart Perf"]<0 else "g-yellow")
                            eqs="g-green" if rw["Ecart Qual"]>0 else ("g-red" if rw["Ecart Qual"]<0 else "g-yellow")
                            eas="g-green" if rw["Ecart Ano"]<0 else ("g-red" if rw["Ecart Ano"]>0 else "g-yellow")
                            tc="#38a169" if rw["Tendance"]=="amelioration" else ("#e53e3e" if rw["Tendance"]=="degradation" else "#d69e2e")
                            h+='<tr><td style="font-weight:700">%s</td><td>%.2f%%</td><td>%.2f%%</td><td class="%s">%+.2f%%</td><td>%.2f%%</td><td>%.2f%%</td><td class="%s">%+.2f%%</td><td>%d</td><td>%d</td><td class="%s">%+d</td><td style="color:%s;font-weight:800">%s</td></tr>'%(rw["Poste"],rw["Perf N-1"],rw["Perf N"],eps,rw["Ecart Perf"],rw["Qual N-1"],rw["Qual N"],eqs,rw["Ecart Qual"],rw["Ano N-1"],rw["Ano N"],eas,rw["Ecart Ano"],tc,rw["Tendance Icon"])
                        h+='</tbody></table>'
                        st.markdown(h,unsafe_allow_html=True)

                        st.markdown('<div class="stl" style="border-left-color:#38a169">🟢 Top 5 Postes en Amélioration</div>',unsafe_allow_html=True)
                        top5_up=comp_df[comp_df["Tendance"]=="amelioration"].copy()
                        if not top5_up.empty:
                            top5_up["Score Gain"]=top5_up["Ecart Perf"]+top5_up["Ecart Qual"]-top5_up["Ecart Ano"]
                            top5_up=top5_up.nlargest(5,"Score Gain")
                            c1,c2=st.columns([1.2,1])
                            with c1:
                                uh='<table class="anl-tbl"><thead><tr><th>Rang</th><th>Poste</th><th>Gain Perf</th><th>Gain Qual</th><th>Reduction Ano</th></tr></thead><tbody>'
                                for i,(_,rw) in enumerate(top5_up.iterrows()):
                                    uh+='<tr><td style="font-weight:800;color:#38a169">%d</td><td style="font-weight:600">%s</td><td class="g-green">%+.2f%%</td><td class="g-green">%+.2f%%</td><td class="g-green">%+d</td></tr>'%(i+1,rw["Poste"],rw["Ecart Perf"],rw["Ecart Qual"],-rw["Ecart Ano"])
                                uh+='</tbody></table>'
                                st.markdown(uh,unsafe_allow_html=True)
                            with c2:
                                fig=go.Figure()
                                fig.add_trace(go.Bar(y=top5_up["Poste"][::-1],x=top5_up["Ecart Perf"][::-1],name="Perf",marker_color="#2b6cb0",orientation="h"))
                                fig.add_trace(go.Bar(y=top5_up["Poste"][::-1],x=top5_up["Ecart Qual"][::-1],name="Qual",marker_color="#276749",orientation="h"))
                                fig.update_layout(barmode="group",height=300,autosize=True,margin=dict(l=120,t=20,b=20,r=20),font_size=9,legend=dict(orientation="h",yanchor="bottom",y=1.02))
                                st.plotly_chart(fig,use_container_width=True)
                        else: st.markdown('<div class="es">Aucun poste en amélioration.</div>',unsafe_allow_html=True)

                        st.markdown('<div class="stl" style="border-left-color:#e53e3e">🔴 Top 5 Postes en Dégradation</div>',unsafe_allow_html=True)
                        top5_down=comp_df[comp_df["Tendance"]=="degradation"].copy()
                        if not top5_down.empty:
                            top5_down["Score Perte"]=-(top5_down["Ecart Perf"]+top5_down["Ecart Qual"])+top5_down["Ecart Ano"]
                            top5_down=top5_down.nlargest(5,"Score Perte")
                            c1,c2=st.columns([1.2,1])
                            with c1:
                                dh='<table class="anl-tbl"><thead><tr><th>Rang</th><th>Poste</th><th>Perte Perf</th><th>Perte Qual</th><th>Hausse Ano</th></tr></thead><tbody>'
                                for i,(_,rw) in enumerate(top5_down.iterrows()):
                                    dh+='<tr><td style="font-weight:800;color:#e53e3e">%d</td><td style="font-weight:600">%s</td><td class="g-red">%+.2f%%</td><td class="g-red">%+.2f%%</td><td class="g-red">%+d</td></tr>'%(i+1,rw["Poste"],rw["Ecart Perf"],rw["Ecart Qual"],rw["Ecart Ano"])
                                dh+='</tbody></table>'
                                st.markdown(dh,unsafe_allow_html=True)
                            with c2:
                                fig=go.Figure()
                                fig.add_trace(go.Bar(y=top5_down["Poste"][::-1],x=top5_down["Ecart Perf"][::-1],name="Perf",marker_color="#e53e3e",orientation="h"))
                                fig.add_trace(go.Bar(y=top5_down["Poste"][::-1],x=top5_down["Ecart Qual"][::-1],name="Qual",marker_color="#c53030",orientation="h"))
                                fig.update_layout(barmode="group",height=300,autosize=True,margin=dict(l=120,t=20,b=20,r=20),font_size=9,legend=dict(orientation="h",yanchor="bottom",y=1.02))
                                st.plotly_chart(fig,use_container_width=True)
                        else: st.markdown('<div class="es">Aucun poste en dégradation.</div>',unsafe_allow_html=True)

                        st.markdown('<div class="stl" style="border-left-color:#d69e2e">🟡 Top 5 Postes Stables</div>',unsafe_allow_html=True)
                        comp_df["Var Abs"]=abs(comp_df["Ecart Perf"])+abs(comp_df["Ecart Qual"])+abs(comp_df["Ecart Ano"])
                        top5_stable=comp_df.nsmallest(5,"Var Abs")
                        sh='<table class="anl-tbl"><thead><tr><th>Rang</th><th>Poste</th><th>Ecart Perf</th><th>Ecart Qual</th><th>Ecart Ano</th></tr></thead><tbody>'
                        for i,(_,rw) in enumerate(top5_stable.iterrows()):
                            sh+='<tr><td style="font-weight:800;color:#d69e2e">%d</td><td style="font-weight:600">%s</td><td>%+.2f%%</td><td>%+.2f%%</td><td>%+d</td></tr>'%(i+1,rw["Poste"],rw["Ecart Perf"],rw["Ecart Qual"],rw["Ecart Ano"])
                        sh+='</tbody></table>'
                        st.markdown(sh,unsafe_allow_html=True)

                        st.markdown('<div class="stl e">Graphiques d\'Évolution</div>',unsafe_allow_html=True)
                        c1,c2,c3=st.columns(3)
                        with c1:
                            fig=go.Figure()
                            fig.add_trace(go.Bar(x=comp_df["Poste"],y=comp_df["Perf N-1"],name=date_n1,marker_color="#a0aec0"))
                            fig.add_trace(go.Bar(x=comp_df["Poste"],y=comp_df["Perf N"],name=date_n,marker_color="#2b6cb0"))
                            fig.update_layout(title="Performance",height=450,autosize=True,margin=dict(t=40,b=80,l=10,r=10),font_size=8,xaxis_tickangle=-45,legend=dict(orientation="h",yanchor="bottom",y=-0.35),barmode="group")
                            st.plotly_chart(fig,use_container_width=True)
                        with c2:
                            fig=go.Figure()
                            fig.add_trace(go.Bar(x=comp_df["Poste"],y=comp_df["Qual N-1"],name=date_n1,marker_color="#a0aec0"))
                            fig.add_trace(go.Bar(x=comp_df["Poste"],y=comp_df["Qual N"],name=date_n,marker_color="#276749"))
                            fig.update_layout(title="Qualite",height=450,autosize=True,margin=dict(t=40,b=80,l=10,r=10),font_size=8,xaxis_tickangle=-45,legend=dict(orientation="h",yanchor="bottom",y=-0.35),barmode="group")
                            st.plotly_chart(fig,use_container_width=True)
                        with c3:
                            fig=go.Figure()
                            fig.add_trace(go.Bar(x=comp_df["Poste"],y=comp_df["Ano N-1"],name=date_n1,marker_color="#a0aec0"))
                            fig.add_trace(go.Bar(x=comp_df["Poste"],y=comp_df["Ano N"],name=date_n,marker_color="#e53e3e"))
                            fig.update_layout(title="Anomalies",height=450,autosize=True,margin=dict(t=40,b=80,l=10,r=10),font_size=8,xaxis_tickangle=-45,legend=dict(orientation="h",yanchor="bottom",y=-0.35),barmode="group")
                            st.plotly_chart(fig,use_container_width=True)

                        st.markdown("---")
                        st.markdown('<div class="stl e">🤖 Analyse Automatique</div>',unsafe_allow_html=True)
                        if st.button("🧠 Générer le Rapport d'Évolution",type="primary",key="gen_report"):
                            st.session_state["ai_report"]=generate_ai_report(comp_df,date_n,date_n1)
                        if "ai_report" in st.session_state:
                            st.markdown(st.session_state["ai_report"])

                        st.markdown("---")
                        st.markdown('<div class="stl e">📥 Exports</div>',unsafe_allow_html=True)
                        ec1,ec2=st.columns(2)
                        with ec1:
                            exp_df=comp_df[["Poste","Perf N-1","Perf N","Ecart Perf","Qual N-1","Qual N","Ecart Qual","Ano N-1","Ano N","Ecart Ano","Tendance Icon"]].copy()
                            exp_df.columns=["Poste","Perf N-1","Perf N","Ecart Perf","Qual N-1","Qual N","Ecart Qual","Ano N-1","Ano N","Ecart Ano","Tendance"]
                            buf=io.BytesIO()
                            exp_df.to_excel(buf,index=False,engine="openpyxl")
                            buf.seek(0)
                            st.download_button("📥 Exporter Comparaison Excel",data=buf,file_name="suivi_evolution_%s_vs_%s.xlsx"%(date_n1.replace("/","-"),date_n.replace("/","-")),mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        with ec2:
                            if "ai_report" in st.session_state:
                                html_r="""<!DOCTYPE html><html><head><meta charset="utf-8"><title>Rapport Evolution</title>
                                <style>body{font-family:Arial,sans-serif;max-width:900px;margin:30px auto;padding:20px;color:#1a202c}
                                h1{color:#1e3a5f;border-bottom:3px solid #2c5282;padding-bottom:10px}
                                h2{color:#2c5282;margin-top:30px}h3{color:#4a5568}
                                table{width:100%;border-collapse:collapse;margin:15px 0}
                                th{background:#1e3a5f;color:#fff;padding:8px;text-align:left}
                                td{padding:6px 8px;border-bottom:1px solid #e2e8f0}
                                tr:nth-child(even) td{background:#f7fafc}
                                hr{border:none;border-top:1px solid #e2e8f0;margin:20px 0}
                                ul{padding-left:20px}li{margin:5px 0}</style></head><body>
                                <h1>📉 Rapport d'Évolution des KPIs</h1>
                                <p><strong>Période :</strong> %s → %s</p><hr>
                                %s
                                <hr><p style="font-size:10px;color:#a0aec0;text-align:center">Genere automatiquement - KPI Dashboard MC & FEED</p>
                                </body></html>"""%(date_n1,date_n,st.session_state["ai_report"].replace("\n","<br>").replace("**","<strong>").replace("## ","<h2>").replace("### ","<h3>").replace("---","<hr>"))
                                st.download_button("📄 Exporter Rapport HTML (PDF)",data=html_r.encode("utf-8"),file_name="rapport_evolution_%s.html"%date_n.replace("/","-"),mime="text/html")

        except Exception as e:
            st.error("Erreur: %s"%str(e))
            import traceback; st.code(traceback.format_exc())
    else:
        st.markdown('<div class="es">📁 Veuillez charger les fichiers OT et AVIS.</div>',unsafe_allow_html=True)

if __name__ == "__main__":
    main()
