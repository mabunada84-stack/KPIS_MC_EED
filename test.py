# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, locale, random, time, os
from datetime import datetime
import plotly.express as px
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt", "r", encoding="utf-8") as f:
                return f.read().strip()
        except: pass
    return datetime.now().strftime("%d/%m/%Y")

def save_kpis_to_excel(prows, pcols, qrows, qcols, ano_p_r, ano_p_c, ano_q_r, ano_q_c, var_p_r, var_p_c, var_q_r, var_q_c, sheet_name):
    kpis_dir = "kpis"
    os.makedirs(kpis_dir, exist_ok=True)
    filepath = os.path.join(kpis_dir, "indicateurs_kpis.xlsx")
    sn = str(sheet_name).replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","")[:31]
    hf = Font(bold=True, color="FFFFFF", size=10)
    hfl = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    hfl2 = PatternFill(start_color="E65100", end_color="E65100", fill_type="solid")
    tf = Font(bold=True, size=12, color="1E3A5F")
    tb = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    try: wb = load_workbook(filepath)
    except:
        wb = Workbook()
        if "Sheet" in wb.sheetnames: del wb["Sheet"]
    if sn in wb.sheetnames: del wb[sn]
    ws = wb.create_sheet(sn)
    rn = 1
    def ws_section(title, cols, rows, sr, hdr_fill=None):
        ws.cell(row=sr, column=1, value=title).font = tf; sr += 1
        for j, c in enumerate(cols, 1):
            cl = ws.cell(row=sr, column=j, value=c); cl.font = hf; cl.fill = hdr_fill or hfl; cl.alignment = Alignment(horizontal='center'); cl.border = tb
        sr += 1
        for r in rows:
            for j, c in enumerate(cols, 1):
                cl = ws.cell(row=sr, column=j, value=r.get(c, "")); cl.border = tb; cl.alignment = Alignment(horizontal='center')
            sr += 1
        return sr + 1
    rn = ws_section("INDICATEURS DE PERFORMANCE", pcols, prows, rn)
    if ano_p_c and ano_p_r: rn = ws_section("ANOMALIES PERFORMANCE", ano_p_c, ano_p_r, rn)
    rn = ws_section("INDICATEURS DE QUALITE", qcols, qrows, rn)
    if ano_q_c and ano_q_r: rn = ws_section("ANOMALIES QUALITE", ano_q_c, ano_q_r, rn)
    if var_p_c and var_p_r: rn = ws_section("VARIANCE PERFORMANCE (Periode vs Reference)", var_p_c, var_p_r, rn, hfl2)
    if var_q_c and var_q_r: rn = ws_section("VARIANCE QUALITE (Periode vs Reference)", var_q_c, var_q_r, rn, hfl2)
    try: wb.save(filepath)
    except: pass

def inject_custom_css():
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root{--p:#1e3a5f;--pl:#2c5282;--b:#e2e8f0;--r:10px}
    *{box-sizing:border-box;margin:0;padding:0}
    .stApp{background:#edf2f7;font-family:'Inter',sans-serif}
    .main .block-container{
        max-width: 100% !important;
        width: 100% !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        padding-top:.6rem;padding-bottom:.6rem;
    }
    .stTabs,.stTabs>div,.stTabs [data-baseweb="tab-list"]{width:100%!important;max-width:100%!important}
    .mh{background:linear-gradient(135deg,var(--p),var(--pl));padding:10px 16px;border-radius:var(--r);margin-bottom:4px;box-shadow:0 6px 20px rgba(0,0,0,.1);overflow:hidden}
    .mh h1{font-size:24px;color:#fff;font-weight:800;margin:0;display:inline}
    .mh .db{float:right;background:rgba(255,255,255,.15);padding:2px 10px;border-radius:14px;color:#fff;font-size:10px;font-weight:500;border:1px solid rgba(255,255,255,.2);margin-top:2px}
    .cr{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;margin-bottom:4px}
    .cc{background:#fff;border-radius:var(--r);padding:8px 10px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid var(--b);text-align:center}
    .cc .cv{font-size:32px;font-weight:900;line-height:1}
    .cc .cl{font-size:12px;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:1px}
    .cc.c1{border-top:3px solid #3182ce}.cc.c1 .cv{color:#2b6cb0}
    .cc.c2{border-top:3px solid #38a169}.cc.c2 .cv{color:#276749}
    .cc.c3{border-top:3px solid #805ad5}.cc.c3 .cv{color:#6b46c1}
    .cc.c4{border-top:3px solid #e53e3e}.cc.c4 .cv{color:#c53030}
    .cc.c5{border-top:3px solid #e65100}.cc.c5 .cv{color:#e65100}
    .cc.c6{border-top:3px solid #00838f}.cc.c6 .cv{color:#00838f}
    .stl{font-size:16px;font-weight:700;color:var(--p);margin:4px 0 1px 0;padding-left:8px;border-left:3px solid var(--pl)}
    .stl.q{border-left-color:#3182ce}.stl.p{border-left-color:#38a169}.stl.a{border-left-color:#e53e3e}.stl.c{border-left-color:#805ad5}.stl.v{border-left-color:#e65100}
    .tw{width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:12px;display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0}
    .tw thead th{background:var(--p);color:#fff;font-weight:700;font-size:12px;text-transform:uppercase;letter-spacing:.3px;padding:6px;border:none;white-space:nowrap;position:sticky;top:0;z-index:10}
    .tw.qt thead th{background:linear-gradient(135deg,#2b6cb0,#3182ce)}
    .tw.pt thead th{background:linear-gradient(135deg,#276749,#38a169)}
    .tw.at thead th{background:linear-gradient(135deg,#c53030,#e53e3e)}
    .tw.vt thead th{background:linear-gradient(135deg,#e65100,#ff8f00)}
    .tw tbody td{padding:4px 6px;border-bottom:1px solid #edf2f7;white-space:nowrap;font-size:12px}
    .tw tbody tr:nth-child(even) td{background:#f7fafc}
    .tw tbody tr:hover td{background:#ebf8ff!important}
    .cb td{background:#2b6cb0!important;color:#fff!important;font-weight:700!important;font-size:12px!important}
    .tr td{background:#e2e8f0!important;font-weight:800!important;font-size:12px!important}
    .stTabs [data-baseweb="tab-list"]{gap:2px;background:#e2e8f0;padding:2px;border-radius:6px;margin-bottom:3px}
    .stTabs [data-baseweb="tab"]{border-radius:5px;padding:5px 10px;font-weight:600;font-size:14px}
    .stTabs [aria-selected="true"]{background:#fff!important;color:var(--p)!important;box-shadow:0 2px 5px rgba(0,0,0,.07)}
    .sr{display:flex;align-items:center;padding:4px 8px;background:#fff;border-radius:5px;margin-bottom:1px;border:1px solid var(--b);font-size:12px}
    .sr .sn{font-weight:700;color:var(--p);min-width:200px;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .sc{padding:2px 7px;border-radius:12px;font-weight:800;font-size:13px;min-width:40px;text-align:center;margin:0 6px;color:#fff}
    .sr .sa{color:#718096;font-size:12px;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sr .stg{font-size:12px;color:#718096;min-width:50px;text-align:center;white-space:nowrap}
    .sr .sb{font-size:12px;font-weight:700;padding:1px 5px;border-radius:3px;white-space:nowrap}
    .ca{background:#fff;border-radius:var(--r);padding:8px;margin-top:2px;border:1px solid var(--b);box-shadow:0 1px 4px rgba(0,0,0,.02)}
    .ca .ct{font-size:10px;font-weight:700;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid var(--b)}
    .car{display:flex;align-items:center;margin-bottom:3px;font-size:8px}
    .car:last-child{margin-bottom:0}
    .car .cal{width:160px;font-weight:600
