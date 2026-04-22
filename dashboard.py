import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS (MENU LATERAL E LAYOUT) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Botões de Navegação Lateral */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 18px 25px !important; border-radius: 15px !important;
        margin-bottom: 12px !important; color: white !important;
        transition: all 0.3s ease; cursor: pointer; width: 100%;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
        border: none !important; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4);
    }
    
    /* Cards de Métricas */
    .metric-card {
        background: #1e293b; padding: 30px; border-radius: 20px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        margin-bottom: 20px;
    }
    .metric-value { color: #10b981; font-size: 2.8rem; font-weight: 900; }
    .metric-title { color: #94a3b8; font-size: 0.9rem; font-weight: 700; text-transform: uppercase; }

    /* Calendário */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 10px; width: 100%; }
    .day-card { 
        background: #0f172a; border-radius: 12px; padding: 15px; 
        min-height: 110px; border: 1px solid rgba(255,255,255,0.05); 
    }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNÇÃO DE CARREGAMENTO
@st.cache_data
def load_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    
    # Limpeza profunda de colunas
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Conversão Numérica Robusta
    for df in [df_order, df_stops]:
        for col in df.columns:
            if col not in ['Data', 'Máquina', 'Turno', 'Problema']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order['Máquina'] = df_order['Máquina'].astype(float).fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].astype(float).fillna(0).astype(int).astype(str)
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    
    return df_order, df_stops

# --- SIDEBAR FIXA ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 INDUSTRIAL HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Arquivo .xlsm", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu_choice = st.radio("NAVEGAÇÃO", 
                               ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "🎯 OEE & CALENDÁRIO"])

# --- LÓGICA PRINCIPAL ---
if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)
    
    # --- ABA 1: PERFORMANCE ---
    if menu_choice == "📈 PERFORMANCE":
        st.sidebar.header("Filtros Performance")
        f_data = st.sidebar.date_input("Data", [df_order['Data'].min(), df_order['Data'].max()], key='d1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        
        # Processamento
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq))]
        
        mc, est = df_f['Machine Counter'].sum(), df_f['Peças Estoque - Ajuste'].sum()
        loss = ((mc-est)/mc*100) if mc > 0 else 0
        
        st.markdown("## 📈 Performance Geral")
        c1, c2, c3 = st.columns(3)
        c1.markdown(f'<div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{mc:,.0f}</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-card"><div class="metric-title">Enviado Estoque</div><div class="metric-value">{est:,.0f}</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-card"><div class="metric-title">Loss %</div><div class="metric-value" style="color:#f43f5e">{loss:.2f}%</div></div>', unsafe_allow_html=True)
        
        st.markdown("### 🏆 Ranking de Máquinas")
        rk = df_order.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        rk['Mov %'] = (rk['Run Time']/rk['Horário Padrão']*100).round(2)
        st.dataframe(rk.sort_values('Mov %', ascending=False), use_container_width=True)

    # --- ABA 2: TOP 10 PARADAS ---
    elif menu_choice == "🛑 TOP 10 PARADAS":
        st.sidebar.header("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Data", [df_stops['Data'].min(), df_stops['Data'].max()], key='d2')
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1])]
        
        st.markdown("## 🛑 Análise de Top 10 Paradas")
        col1, col2 = st.columns(2)
        
        with col1:
            # Top 10 Minutos
            top_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
            fig_min = px.bar(top_min, orientation='h', title="Top 10 por Minutos", color_discrete_sequence=['#f43f5e'])
            fig_min.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color':'white'})
            st.plotly_chart(fig_min, use_container_width=True)
            
        with col2:
            # Top 10 Quantidade
            top_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
            fig_qtd = px.bar(top_qtd, orientation='h', title="Top 10 por Ocorrências", color_discrete_sequence=['#3b82f6'])
            fig_qtd.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color':'white'})
            st.plotly_chart(fig_qtd, use_container_width=True)

    # --- ABA 3: OEE & CALENDÁRIO ---
    elif menu_choice == "🎯 OEE & CALENDÁRIO":
        st.sidebar.header("Filtros OEE")
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t3')
        
        df_oee_f = df_order[df_order['Turno'].isin(f_turno)]
        
        st.markdown("## 🎯 Eficiência OEE e Calendário")
        
        # Calendário
        df_oee_f['Dia'] = df_oee_f['Data'].dt.day
        cal_data = df_oee_f.groupby('Dia').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time']/cal_data['Horário Padrão']*100).fillna(0)
        
        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, datetime.now().month))
        html = '<div class="calendar-grid">'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                val = cal_data[cal_data['Dia']==d]['Mov'].values[0] if d in cal_data['Dia'].values else 0
                cor = "#059669" if val > 85 else "#dc2626" if val > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><b>{d}</b><br><small>MOV: {val:.1f}%</small></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

else:
    st.info("💡 Por favor, carregue o arquivo .xlsm no menu lateral para iniciar.")
