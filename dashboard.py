import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM REFINADA ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Navegação Lateral Estilo Trading */
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

    /* Cartões de Métricas Padronizados */
    .metric-container {
        display: flex; justify-content: space-between; gap: 10px; margin-bottom: 20px;
    }
    .metric-card {
        background: #1e293b; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 100px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 4px; }
    .metric-value { color: #10b981; font-size: 1.2rem; font-weight: 900; line-height: 1; }

    /* Calendário */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; }
    .day-card { 
        background: #0f172a; border-radius: 8px; padding: 10px; 
        min-height: 90px; border: 1px solid rgba(255,255,255,0.05); 
    }
    .day-number { font-size: 0.9rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.65rem; font-weight: 600; color: #ffffff; text-align: right; line-height: 1.2; }

    /* Área 5 Porquês */
    .five-why-box {
        border: 2px solid #334155; border-radius: 10px; padding: 15px; background: #0f172a; margin-top: 10px;
    }
    .five-why-line { border-bottom: 1px solid #334155; padding: 8px 0; color: #94a3b8; font-size: 0.9rem; }

    /* Ajustes para Impressão */
    @media print {
        .stApp { background-color: white !important; color: black !important; }
        [data-testid="stSidebar"], header { display: none !important; }
        .metric-card { border: 1px solid black !important; background: white !important; }
        .metric-value, .metric-title, .day-number, .day-status { color: black !important; }
        .five-why-box { border: 2px solid black !important; background: white !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file_obj):
    # Lendo as abas
    df_order = pd.read_excel(file_obj, sheet_name="Result by order")
    df_stops = pd.read_excel(file_obj, sheet_name="Stop machine item")
    
    # Limpeza de nomes e conversão de colunas
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Tratamento numérico robusto (converte erros/texto em 0)
    cols_num = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed', 'Minutos', 'QTD']
    for df in [df_order, df_stops]:
        for col in cols_num:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    
    return df_order, df_stops

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 INDUSTRIAL HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Relatório (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 QUADRO DE MÁQUINA"])

# --- PROCESSAMENTO ---
if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    # =========================================================
    # ABA 1: PERFORMANCE GERAL
    # =========================================================
    if menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t1')
        
        # Filtro de dados
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & 
                        (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        # Cálculo OEE simplificado por turno disponível
        df_f['T_Dispo'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        
        st.markdown("## 📈 Performance e Eficiência")
        c1, c2, c3 = st.columns(3)
        c1.markdown(f'<div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f["Machine Counter"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        
        disponibilidade = (df_f['Run Time'].sum() / df_f['T_Dispo'].sum() * 100) if df_f['T_Dispo'].sum() > 0 else 0
        c2.markdown(f'<div class="metric-card"><div class="metric-title">Disponibilidade Média</div><div class="metric-value">{disponibilidade:.1f}%</div></div>', unsafe_allow_html=True)
        
        c3.markdown(f'<div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f["Run Time"].sum():,.0f}m</div></div>', unsafe_allow_html=True)

        # Gráficos de Velocímetro
        hp_sum = df_f['Horário Padrão'].sum()
        mov_p = (df_f["Run Time"].sum() / hp_sum * 100) if hp_sum > 0 else 0
        loss_p = ((df_f["Machine Counter"].sum() - df_f["Peças Estoque - Ajuste"].sum()) / df_f["Machine Counter"].sum() * 100) if df_f["Machine Counter"].sum() > 0 else 0

        col_v1, col_v2 = st.columns(2)
        def create_gauge(label, value, color, target):
            fig = go.Figure(go.Indicator(mode="gauge+number", value=value, title={'text': label},
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': color}, 'threshold': {'line': {'color': "white", 'width': 4}, 'value': target}}))
            fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=250)
            return fig

        col_v1.plotly_chart(create_gauge("Movimentação (%)", mov_p, "#10b981", 85), use_container_width=True)
        col_v2.plotly_chart(create_gauge("Loss (%)", loss_p, "#e74c3c", 5), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1])]
        
        st.markdown("## 🛑 Análise de Paradas")
        
        top_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
        fig_min = px.bar(top_min, orientation='h', title="Top 10 por Minutos Totais", color_discrete_sequence=['#f43f5e'])
        st.plotly_chart(fig_min, use_container_width=True)
        
        st.markdown("---")
        
        top_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
        fig_qtd = px.bar(top_qtd, orientation='h', title="Top 10 por Quantidade", color_discrete_sequence=['#3b82f6'])
        st.plotly_chart(fig_qtd, use_container_width=True)

    # =========================================================
    # ABA 3: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        mes_idx = list(calendar.month_name).index(mes_sel)
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m3')
        
        df_c = df_order[(df_order['Data'].dt.month == mes_idx) & (df_order['Máquina'].isin(f_maq_c))]
        
        st.markdown(f"## 📅 Calendário - {mes_sel}")
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'] * 100).fillna(0)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'] * 100).fillna(0)

        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, mes_idx))
        html = '<div class="calendar-grid">'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_v, l_v = (row['Mov'].values[0], row['Loss'].values[0]) if not row.empty else (0,0)
                cor = "#059669" if m_v > 85 else "#dc2626" if m_v > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">M: {m_v:.1f}%<br>L: {l_v:.1f}%</div></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA 4: QUADRO DE MÁQUINA (Board de Impressão)
    # =========================================================
    elif menu == "📋 QUADRO DE MÁQUINA":
        st.sidebar.subheader("Configuração")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        data_fim = df_order['Data'].max()
        data_ini = data_fim - timedelta(days=7)
        
        df_b = df_order[(df_order['Data'] >= data_ini) & (df_order['Máquina'] == maq_b)]
        df_sb = df_stops[(df_stops['Data'] >= data_ini) & (df_stops['Máquina'] == maq_b)]

        st.markdown(f"### 📋 QUADRO SEMANAL - MÁQUINA {maq_b}")
        st.write(f"Período: {data_ini.strftime('%d/%m')} a {data_fim.strftime('%d/%m/%Y')}")

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].sum()*100) if df_b["Horário Padrão"].sum()>0 else 0
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].sum()*100) if df_b["Machine Counter"].sum()>0 else 0
        
        st.markdown(f"""<div class="metric-container">
            <div class="metric-card"><div class="metric-title">Movimentação</div><div class="metric-value">{m_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Loss</div><div class="metric-value" style="color:#f43f5e">{l_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Peças</div><div class="metric-value">{df_b["Peças Estoque - Ajuste"].sum():,.0f}</div></div></div>""", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            st.write("🛑 Impacto das Paradas (%)")
            stop_imp = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True)
            if not stop_imp.empty:
                fig = px.bar(stop_imp/stop_imp.sum()*100, orientation='h', color_discrete_sequence=['#10b981'])
                fig.update_layout(height=350, margin=dict(l=0,r=0,t=0,b=0))
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.markdown("""<div class="five-why-box">
                <div style="color:#10b981; font-weight:bold;">5 PORQUÊS</div>
                <div class="five-why-line">1. _________________</div><div class="five-why-line">2. _________________</div>
                <div class="five-why-line">3. _________________</div><div class="five-why-line">4. _________________</div>
                <div class="five-why-line">5. _________________</div>
                <div style="margin-top:10px; font-weight:bold;">AÇÃO: _________________</div></div>""", unsafe_allow_html=True)

else:
    st.info("💡 Por favor, carregue o arquivo Excel (.xlsm) para visualizar os dados.")
