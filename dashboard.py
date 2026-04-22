import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM REFINADA ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Navegação Lateral */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 12px 20px !important; border-radius: 10px !important;
        margin-bottom: 8px !important; color: white !important;
        transition: all 0.3s ease; cursor: pointer;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
        border: none !important;
    }

    /* Cards de Métricas Reduzidos */
    .metric-card {
        background: #1e293b; padding: 15px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
    }
    .metric-title { color: #94a3b8; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { color: #10b981; font-size: 1.5rem; font-weight: 900; }

    /* Calendário */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; }
    .day-card { 
        background: #0f172a; border-radius: 8px; padding: 10px; 
        min-height: 90px; border: 1px solid rgba(255,255,255,0.05); 
    }
    .day-number { font-size: 0.9rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.7rem; font-weight: 600; color: #ffffff; text-align: right; line-height: 1.2; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Tratamento numérico
    for df in [df_order, df_stops]:
        for col in df.columns:
            if col not in ['Data', 'Máquina', 'Turno', 'Problema']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    return df_order, df_stops

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 INDUSTRIAL HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂
