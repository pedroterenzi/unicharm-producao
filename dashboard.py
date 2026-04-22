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
    uploaded_file = st.file_uploader("📂 Carregar Relatório (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO"])

# --- PROCESSAMENTO POR ABA ---
if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    # =========================================================
    # ABA 1: PERFORMANCE GERAL (Incluindo OEE)
    # =========================================================
    if menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()])
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq))]
        
        # Cálculos OEE
        df_f['Tempo_Disponivel'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        oee_data = df_f.groupby('Máquina').agg({'Run Time':'sum', 'Tempo_Disponivel':'sum', 'Peças Estoque - Ajuste':'sum', 'Machine Counter':'sum', 'Average Speed':'mean'}).reset_index()
        oee_data['Disp'] = (oee_data['Run Time'] / oee_data['Tempo_Disponivel']).clip(0,1)
        oee_data['Qual'] = (oee_data['Peças Estoque - Ajuste'] / oee_data['Machine Counter']).fillna(0).clip(0,1)
        oee_data['Perf'] = (oee_data['Machine Counter'] / (oee_data['Average Speed'] * oee_data['Run Time'])).fillna(0).clip(0,1)
        oee_data['OEE'] = (oee_data['Disp'] * oee_data['Perf'] * oee_data['Qual'] * 100).round(1)

        # UI
        st.markdown("## 📈 Performance e Eficiência (OEE)")
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f["Machine Counter"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f["Peças Estoque - Ajuste"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-card"><div class="metric-title">OEE Médio</div><div class="metric-value">{oee_data["OEE"].mean():.1f}%</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f["Run Time"].sum():,.0f}m</div></div>', unsafe_allow_html=True)

        st.markdown("### 🏆 Ranking de Eficiência por Máquina")
        st.dataframe(oee_data[['Máquina', 'Disp', 'Perf', 'Qual', 'OEE']].sort_values('OEE', ascending=False), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS (Um gráfico por linha)
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()])
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1])]
        
        st.markdown("## 🛑 Análise de Paradas")
        
        # Gráfico 1 - Minutos
        top_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
        fig_min = px.bar(top_min, orientation='h', title="Top 10 por Minutos Totais", color_discrete_sequence=['#f43f5e'])
        st.plotly_chart(fig_min, use_container_width=True)
        
        st.markdown("---")
        
        # Gráfico 2 - Quantidade
        top_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
        fig_qtd = px.bar(top_qtd, orientation='h', title="Top 10 por Frequência (Quantidade)", color_discrete_sequence=['#3b82f6'])
        st.plotly_chart(fig_qtd, use_container_width=True)

    # =========================================================
    # ABA 3: CALENDÁRIO (Filtros Mês, Máquina, Turno)
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        mes_idx = list(calendar.month_name).index(mes_sel)
        
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        f_turno_c = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        
        # Filtragem do calendário
        df_c = df_order[(df_order['Data'].dt.month == mes_idx) & (df_order['Máquina'].isin(f_maq_c)) & (df_order['Turno'].isin(f_turno_c))]
        
        st.markdown(f"## 📅 Calendário Operacional - {mes_sel}")
        
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'] * 100).fillna(0)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'] * 100).fillna(0)

        # Montagem do Grid
        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, mes_idx))
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#64748b; font-weight:900; font-size:0.7rem;">{n}</div>'
        
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_val = row['Mov'].values[0] if not row.empty else 0
                l_val = row['Loss'].values[0] if not row.empty else 0
                cor = "#059669" if m_val > 85 else "#dc2626" if m_val > 0 else "#1e293b"
                html += f'''<div class="day-card" style="background:{cor}">
                            <span class="day-number">{d}</span>
                            <div class="day-status">MOV: {m_val:.1f}%<br>LOSS: {l_val:.1f}%</div>
                          </div>'''
        st.markdown(html + '</div>', unsafe_allow_html=True)
else:
    st.info("💡 Por favor, carregue o arquivo Excel (.xlsm) para visualizar os dados.")
