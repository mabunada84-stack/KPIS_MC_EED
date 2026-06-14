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

# ============= CONFIGURATION PAGE =============
st.set_page_config(
    layout="wide",
    page_title="Dashboard KPI MC & FEED",
    page_icon="📊"
)

# ============= OPTIMISATIONS - CACHE =============
@st.cache_data(ttl=3600)
def load_excel_cached(filepath):
    """Cache le chargement des fichiers Excel"""
    return pd.read_excel(filepath)

@st.cache_data
def excr_cached(df):
    """Cache l'exclusion des compresseurs"""
    if "Poste travail princ." not in df.columns:
        return df
    return df[~df["Poste travail princ."].astype(str).str.contains("cresseur", case=False, na=False)].copy()

@st.cache_data
def process_dates(df, date_columns):
    """Cache le traitement des dates"""
    df_copy = df.copy()
    for col in date_columns:
        if col in df_copy.columns:
            df_copy[col] = pd.to_datetime(df_copy[col], errors="coerce")
    return df_copy

def get_date_from_file():
    if os.path.exists("date.txt"):
        try:
            with open("date.txt", "r", encoding="utf-8") as f:
                return f.read().strip()
        except: pass
    return datetime.now().strftime("%d/%m/%Y")

# ============= GESTION HISTORIQUE =============
def get_historique_filepath():
    """Retourne le chemin du fichier d'historique"""
    kpis
