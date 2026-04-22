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

    /* Calendário Operacional */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; }
    .day-card { 
        background: #0f172a; border-radius: 8px; padding: 10px; 
        min-height: 90px; border: 1px solid rgba(255,255,255,0.05); 
    }
    .day-number { font-size: 0.9rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.65rem; font-weight: 600; color: #ffffff; text-align: right; line-height: 1.2; }

    /* Estilo 5 Porquês */
    .five-why-box {
        border: 2px solid #334155; border-radius: 10px; padding: 15px; background: #0f172a; margin-top: 10px;
    }
    .five-why-line { border-bottom: 1px solid #334155; padding: 8px 0; color: #94a3b8; font-size: 0.9rem; }

    /* Ajustes para Impressão */
    @media print {
        .stApp { background-color: white !important; color: black !important; }
        [data-testid="stSidebar"], header { display: none !important; }
        .metric-card { border: 1px solid #000 !important; background: white !important; }
        .metric-value, .metric-title, .day-number, .day-status { color: black !important; }
        .five-why-box { border: 2px solid black !important; background: white !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order") [cite: 1675, 1691]
    df_stops = pd.read_excel(file, sheet_name="Stop machine item") [cite: 1716, 1756]
    
    # Limpeza de espaços em branco nos nomes das colunas
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Tratamento numérico (converte texto sujo para 0)
    cols_num = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed', 'Minutos', 'QTD']
    for df in [df_order, df_stops]:
        for col in cols_num:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Tipagem de IDs e Datas
    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    
    return df_order, df_stops

# --- MENU LATERAL ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 INDUSTRIAL HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Relatório (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 QUADRO DE MÁQUINA"])

# --- LÓGICA POR ABA ---
if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    # =========================================================
    # ABA 1: PERFORMANCE GERAL
    # =========================================================
    if menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t1')
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & 
                        (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        # Cálculos OEE
        df_f['Tempo_Disponivel'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        oee_data = df_f.groupby('Máquina').agg({'Run Time':'sum', 'Tempo_Disponivel':'sum', 'Peças Estoque - Ajuste':'sum', 'Machine Counter':'sum', 'Average Speed':'mean', 'Horário Padrão':'sum'}).reset_index()
        
        oee_data['Disp'] = (oee_data['Run Time'] / oee_data['Tempo_Disponivel']).clip(0,1)
        oee_data['Qual'] = (oee_data['Peças Estoque - Ajuste'] / oee_data['Machine Counter']).fillna(0).clip(0,1)
        oee_data['Perf'] = (oee_data['Machine Counter'] / (oee_data['Average Speed'] * oee_data['Run Time'])).fillna(0).clip(0,1)
        oee_data['OEE'] = (oee_data['Disp'] * oee_data['Perf'] * oee_data['Qual'] * 100).round(1)

        st.markdown("## 📈 Performance e Eficiência")
        
        total_mc = df_f["Machine Counter"].sum()
        total_est = df_f["Peças Estoque - Ajuste"].sum()
        mov_p = (df_f["Run Time"].sum() / df_f["Horário Padrão"].sum() * 100) if df_f["Horário Padrão"].sum() > 0 else 0
        loss_p = ((total_mc - total_est) / total_mc * 100) if total_mc > 0 else 0

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{total_mc:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{total_est:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">OEE Médio</div><div class="metric-value">{oee_data["OEE"].mean():.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Movimentação</div><div class="metric-value">{mov_p:.1f}%</div></div>
            </div>
            """, unsafe_allow_html=True)

        col_v1, col_v2 = st.columns(2)
        def create_gauge(label, value, color, target):
            fig = go.Figure(go.Indicator(mode="gauge+number", value=value, title={'text': label},
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': color}, 'threshold': {'line': {'color': "white", 'width': 4}, 'value': target}}))
            fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=250)
            return fig

        col_v1.plotly_chart(create_gauge("Velocímetro Movimentação (%)", mov_p, "#10b981", 85), use_container_width=True)
        col_v2.plotly_chart(create_gauge("Velocímetro Loss (%)", loss_p, "#e74c3c", 5), use_container_width=True)

        st.markdown("### 🏆 Ranking por Máquina (Baseado em Movimentação)")
        st.dataframe(oee_data[['Máquina', 'Disp', 'Perf', 'Qual', 'OEE']].sort_values('OEE', ascending=False), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        f_maq_s = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()), key='m2')
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1]) & (df_stops['Máquina'].isin(f_maq_s))]
        
        st.markdown("## 🛑 Top 10 Paradas")
        
        # Gráfico 1: Minutos
        top_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
        fig_min = px.bar(top_min, orientation='h', title="Top 10 Paradas por Minutos Totais", color_discrete_sequence=['#f43f5e'])
        st.plotly_chart(fig_min, use_container_width=True)
        
        # Gráfico 2: Quantidade
        top_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
        fig_qtd = px.bar(top_qtd, orientation='h', title="Top 10 Paradas por Quantidade", color_discrete_sequence=['#3b82f6'])
        st.plotly_chart(fig_qtd, use_container_width=True)

    # =========================================================
    # ABA 3: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        mes_idx = list(calendar.month_name).index(mes_sel)
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m3')
        f_turno_c = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t3')
        
        df_c = df_order[(df_order['Data'].dt.month == mes_idx) & (df_order['Máquina'].isin(f_maq_c)) & (df_order['Turno'].isin(f_turno_c))]
        
        st.markdown(f"## 📅 Calendário Operacional - {mes_sel}")
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'] * 100).fillna(0)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'] * 100).fillna(0)

        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, mes_idx))
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#64748b; font-weight:900; font-size:0.7rem;">{n}</div>'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_v, l_v = (row['Mov'].values[0], row['Loss'].values[0]) if not row.empty else (0,0)
                cor = "#059669" if m_v > 85 else "#dc2626" if m_v > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">MOV: {m_v:.1f}%<br>LOSS: {l_v:.1f}%</div></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA 4: QUADRO DE MÁQUINA (Board de Impressão)
    # =========================================================
    elif menu == "📋 QUADRO DE MÁQUINA":
        st.sidebar.subheader("Configuração do Board")
        maq_b = st.sidebar.selectbox("Máquina para Board", sorted(df_order['Máquina'].unique()))
        data_fim = df_order['Data'].max()
        data_ini = data_fim - timedelta(days=7)
        
        df_b = df_order[(df_order['Data'] >= data_ini) & (df_order['Máquina'] == maq_b)]
        df_sb = df_stops[(df_stops['Data'] >= data_ini) & (df_stops['Máquina'] == maq_b)]

        st.markdown(f"""<div style="text-align:center; border-bottom:2px solid #10b981; padding-bottom:10px; margin-bottom:20px;">
            <h1 style="color:white; margin:0;">QUADRO SEMANAL - MÁQUINA {maq_b}</h1>
            <p style="color:#94a3b8; margin:5px 0;">{data_ini.strftime('%d/%m')} a {data_fim.strftime('%d/%m/%Y')}</p></div>""", unsafe_allow_html=True)

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].sum()*100) if df_b["Horário Padrão"].sum()>0 else 0
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].sum()*100) if df_b["Machine Counter"].sum()>0 else 0
        
        st.markdown(f"""<div class="metric-container">
            <div class="metric-card"><div class="metric-title">Movimentação</div><div class="metric-value">{m_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Loss (Perda)</div><div class="metric-value" style="color:#f43f5e">{l_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Peças Enviadas</div><div class="metric-value">{df_b["Peças Estoque - Ajuste"].sum():,.0f}</div></div>
            <div class="metric-card"><div class="metric-title">Minutos Produzidos</div><div class="metric-value">{df_b["Run Time"].sum():,.0f}m</div></div></div>""", unsafe_allow_html=True)

        c_b1, c_b2 = st.columns(2)
        with c_b1:
            st.markdown("### 🛑 Impacto das Paradas (%)")
            stop_imp = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True)
            if not stop_imp.empty:
                fig_b = px.bar(stop_imp/stop_imp.sum()*100, orientation='h', labels={'value':'% Impacto','Problema':''}, color_discrete_sequence=['#10b981'])
                fig_b.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color':'white'}, height=400)
                st.plotly_chart(fig_b, use_container_width=True)
        with c_b2:
            st.markdown("### 🧠 Análise de Causa Raiz (5 Porquês)")
            st.markdown("""<div class="five-why-box">
                <div style="color:#10b981; font-weight:bold; margin-bottom:10px;">PROBLEMA PRINCIPAL: ________________________</div>
                <div class="five-why-line"><b>1. Por que?</b> ____________________________________</div>
                <div class="five-why-line"><b>2. Por que?</b> ____________________________________</div>
                <div class="five-why-line"><b>3. Por que?</b> ____________________________________</div>
                <div class="five-why-line"><b>4. Por que?</b> ____________________________________</div>
                <div class="five-why-line"><b>5. Por que?</b> ____________________________________</div>
                <div style="margin-top:15px; color:#10b981; font-weight:bold;">AÇÃO CORRETIVA:</div>
                <div style="height:60px; border:1px dashed #334155; margin-top:5px;"></div></div>""", unsafe_allow_html=True)
        st.info("💡 Para imprimir: Ctrl+P e selecione 'Salvar como PDF'.")

else:
    st.info("Aguardando upload do arquivo Excel para processar os indicadores.")
