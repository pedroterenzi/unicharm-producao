import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 12px 20px !important; border-radius: 10px !important;
        margin-bottom: 8px !important; color: white !important; cursor: pointer;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
    }

    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 20px; }
    .metric-card {
        background: #1e293b; padding: 15px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 100px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { color: #10b981; font-size: 1.6rem; font-weight: 900; line-height: 1; }

    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 10px; width: 100%; }
    .day-card { 
        background: #0f172a; border-radius: 10px; padding: 12px; 
        min-height: 100px; border: 1px solid rgba(255,255,255,0.05); 
    }
    .day-number { font-size: 1rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.75rem; font-weight: 600; color: #ffffff; text-align: right; line-height: 1.3; }

    .five-why-box {
        border: 2px solid #cbd5e1; border-radius: 10px; padding: 25px; 
        background: #ffffff; margin-top: 20px; color: #1e293b;
    }
    .five-why-line { border-bottom: 1px dotted #94a3b8; padding: 12px 0; font-size: 1rem; color: #334155; }
    .five-why-title { color: #059669; font-weight: 900; font-size: 1.2rem; margin-bottom: 15px; border-bottom: 2px solid #059669; padding-bottom: 5px; }

    @media print {
        .stApp { background-color: white !important; color: black !important; }
        [data-testid="stSidebar"], header, .stButton { display: none !important; }
        .metric-card { border: 2px solid black !important; background: white !important; }
        .metric-value, .metric-title { color: black !important; }
        .five-why-box { border: 2px solid black !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E TRATAMENTO
@st.cache_data
def load_data(file_obj):
    df_order = pd.read_excel(file_obj, sheet_name="Result by order")
    df_stops = pd.read_excel(file_obj, sheet_name="Stop machine item")
    
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Tratamento de nulos e tipos
    nums = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed', 'Minutos', 'QTD']
    for df in [df_order, df_stops]:
        for col in nums:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    df_stops['Turno'] = df_stops['Turno'].fillna(0).astype(int).astype(str)
    
    return df_order, df_stops

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 INDUSTRIAL HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Relatório (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📝 ANÁLISE SEMANAL"])

# --- PROCESSAMENTO ---
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
        
        # Correção do OEE (Calculando T_Dispo primeiro)
        df_f = df_f.copy()
        df_f['T_Dispo'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        
        oee_df = df_f.groupby('Máquina').agg({
            'Run Time':'sum', 
            'Peças Estoque - Ajuste':'sum', 
            'Machine Counter':'sum', 
            'Average Speed':'mean', 
            'Horário Padrão':'sum',
            'T_Dispo':'sum'
        }).reset_index()
        
        oee_df['Disp'] = (oee_df['Run Time'] / oee_df['T_Dispo']).clip(0,1)
        oee_df['Qual'] = (oee_df['Peças Estoque - Ajuste'] / oee_df['Machine Counter']).fillna(0).clip(0,1)
        oee_df['Perf'] = (oee_df['Machine Counter'] / (oee_df['Average Speed'] * oee_df['Run Time'].replace(0,1))).fillna(0).clip(0,1)
        oee_df['OEE'] = (oee_df['Disp'] * oee_df['Perf'] * oee_df['Qual'] * 100).round(1)

        st.markdown("## 📈 Performance e Eficiência")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f["Machine Counter"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f["Peças Estoque - Ajuste"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-card"><div class="metric-title">OEE Médio</div><div class="metric-value">{oee_df["OEE"].mean():.1f}%</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f["Run Time"].sum():,.0f}m</div></div>', unsafe_allow_html=True)

        col_v1, col_v2 = st.columns(2)
        hp_sum = df_f['Horário Padrão'].sum()
        mov_p = (df_f["Run Time"].sum() / hp_sum * 100) if hp_sum > 0 else 0
        loss_p = ((df_f["Machine Counter"].sum() - df_f["Peças Estoque - Ajuste"].sum()) / df_f["Machine Counter"].sum() * 100) if df_f["Machine Counter"].sum() > 0 else 0

        with col_v1:
            fig1 = go.Figure(go.Indicator(mode="gauge+number", value=mov_p, title={'text': "Movimentação (%)"},
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#10b981"}, 'threshold': {'line': {'color': "white", 'width': 4}, 'value': 85}}))
            fig1.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=400)
            st.plotly_chart(fig1, use_container_width=True)
        with col_v2:
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=loss_p, title={'text': "Loss (%)"},
                gauge={'axis': {'range': [0, 15]}, 'bar': {'color': "#e74c3c"}, 'threshold': {'line': {'color': "white", 'width': 4}, 'value': 5}}))
            fig2.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=400)
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### 🏆 Ranking OEE por Máquina")
        st.dataframe(oee_df[['Máquina', 'Disp', 'Perf', 'Qual', 'OEE']].sort_values('OEE', ascending=False), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        f_maq_s = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()), key='m2')
        f_turno_s = st.sidebar.multiselect("Turnos", sorted(df_stops['Turno'].unique()), default=sorted(df_stops['Turno'].unique()), key='t2')
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1]) & 
                          (df_stops['Máquina'].isin(f_maq_s)) & (df_stops['Turno'].isin(f_turno_s))]
        
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
        # Criar seletor de Mês e Ano baseado nos dados
        df_order['Mes_Ano'] = df_order['Data'].dt.strftime('%m/%Y')
        mes_ano_lista = sorted(df_order['Mes_Ano'].unique(), key=lambda x: datetime.strptime(x, '%m/%Y'), reverse=True)
        mes_sel_ref = st.sidebar.selectbox("Mês/Ano", mes_ano_lista)
        
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m3')
        f_turno_c = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t3')
        
        # Filtragem
        df_c = df_order[(df_order['Mes_Ano'] == mes_sel_ref) & (df_order['Máquina'].isin(f_maq_c)) & (df_order['Turno'].isin(f_turno_c))]
        
        st.markdown(f"## 📅 Operação - {mes_sel_ref}")
        
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'].replace(0,1) * 100).fillna(0)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'].replace(0,1) * 100).fillna(0)

        # Gerar calendário
        sel_date_obj = datetime.strptime(mes_sel_ref, '%m/%Y')
        days = list(calendar.Calendar(0).itermonthdays(sel_date_obj.year, sel_date_obj.month))
        
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#64748b; font-weight:900; font-size:0.8rem;">{n}</div>'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_v, l_v = (row['Mov'].values[0], row['Loss'].values[0]) if not row.empty else (0,0)
                cor = "#059669" if m_v > 85 else "#dc2626" if m_v > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">MOV: {m_v:.1f}%<br>LOSS: {l_v:.1f}%</div></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA 4: ANÁLISE SEMANAL
    # =========================================================
    elif menu == "📝 ANÁLISE SEMANAL":
        st.sidebar.subheader("Filtros Análise")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_b = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='tb')
        periodo_b = st.sidebar.date_input("Período da Semana", [df_order['Data'].max() - timedelta(days=7), df_order['Data'].max()])
        
        # Filtragem
        df_b = df_order[(df_order['Data'].dt.date >= periodo_b[0]) & (df_order['Data'].dt.date <= periodo_b[1]) & 
                        (df_order['Máquina'] == maq_b) & (df_order['Turno'].isin(turno_b))]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo_b[0]) & (df_stops['Data'].dt.date <= periodo_b[1]) & 
                         (df_stops['Máquina'] == maq_b) & (df_stops['Turno'].isin(turno_b))]

        st.markdown(f"""<div style="text-align:center; border-bottom:3px solid #10b981; padding-bottom:10px; margin-bottom:20px;">
            <h1 style="color:white; margin:0;">ANÁLISE SEMANAL DE PERFORMANCE - MÁQUINA {maq_b}</h1>
            <p style="color:#94a3b8; font-size:1.2rem;">Período: {periodo_b[0].strftime('%d/%m/%Y')} a {periodo_b[1].strftime('%d/%m/%Y')}</p></div>""", unsafe_allow_html=True)

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].replace(0,1).sum()*100)
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].replace(0,1).sum()*100)
        
        st.markdown(f"""<div class="metric-container">
            <div class="metric-card"><div class="metric-title">Movimentação</div><div class="metric-value">{m_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Loss (Perda)</div><div class="metric-value" style="color:#f43f5e">{l_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Peças Enviadas</div><div class="metric-value">{df_b["Peças Estoque - Ajuste"].sum():,.0f}</div></div>
            <div class="metric-card"><div class="metric-title">Minutos Produzidos</div><div class="metric-value">{df_b["Run Time"].sum():,.0f}m</div></div></div>""", unsafe_allow_html=True)

        st.markdown("### 🛑 Top 5 Piores Paradas (Impacto %)")
        stop_imp = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True)
        if not stop_imp.empty:
            pior_parada = stop_imp.index[-1] # A última do sort (maior)
            fig_b = px.bar(stop_imp.tail(5)/stop_imp.sum()*100, orientation='h', labels={'value':'% Impacto','Problema':''}, color_discrete_sequence=['#10b981'])
            st.plotly_chart(fig_b, use_container_width=True)
        else:
            pior_parada = "NENHUMA PARADA REGISTRADA"

        st.markdown(f"""
            <div class="five-why-box">
                <div class="five-why-title">ANÁLISE DE CAUSA RAIZ - FERRAMENTA 5 PORQUÊS</div>
                <div style="font-weight:bold; margin-bottom:10px;">PROBLEMA FOCO: <span style="color:#e74c3c">{pior_parada}</span></div>
                <div class="five-why-line"><b>1º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>2º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>3º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>4º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>5º Por que?</b> ________________________________________________________________________________________</div>
                <br>
                <div style="display:flex; gap:20px;">
                    <div style="flex:1; border:1px solid #94a3b8; padding:10px; min-height:100px;"><b>CAUSA RAIZ:</b></div>
                    <div style="flex:2; border:1px solid #94a3b8; padding:10px; min-height:100px;"><b>AÇÃO CORRETIVA / PLANO DE AÇÃO:</b></div>
                </div>
            </div>
        """, unsafe_allow_html=True)
else:
    st.info("💡 Carregue o arquivo Excel para iniciar.")
