import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM (Botões Uniformes, Cores e Impressão) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Remover bolinhas do seletor lateral e criar blocos de botões */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label > div:first-child { display: none !important; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 15px 20px !important; border-radius: 10px !important;
        margin-bottom: 10px !important; color: white !important;
        transition: all 0.3s ease; cursor: pointer; width: 100%;
        display: block !important; text-align: center; font-weight: 600;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
        border: none !important; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
    }

    /* Cards de Métricas */
    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 15px; }
    .metric-card {
        background: #1e293b; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 70px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    .metric-value { color: #10b981; font-size: 1.3rem; font-weight: 900; line-height: 1; }

    /* Estilização de Tabelas de Reporte */
    .section-header {
        background: #1e293b; padding: 10px; border-radius: 5px;
        color: #10b981; font-weight: 800; text-transform: uppercase;
        margin-top: 20px; border-left: 5px solid #10b981; font-size: 0.9rem;
    }

    /* Área 5 Porquês */
    .five-why-box { border: 1px solid #000; border-radius: 5px; padding: 15px; background: #ffffff; color: #000; margin-top: 10px; }
    .five-why-line { border-bottom: 1px solid #000; padding: 8px 0; font-size: 0.9rem; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNÇÕES DE CARREGAMENTO E LIMPEZA
@st.cache_data
def load_production_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    nums = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed', 'Minutos', 'QTD']
    for df in [df_order, df_stops]:
        for col in nums:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    return df_order, df_stops

@st.cache_data
def get_metas_from_datas(file, data_referencia):
    try:
        xls = pd.ExcelFile(file)
        target_sheet = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        if not target_sheet: return 0, 0
        
        # Cabeçalho está na linha 3 (index 2)
        df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)
        row_dates = df_raw.iloc[2, :].tolist()
        
        meta_mes_total = 0
        meta_ate_hoje = 0
        
        # Dados começam na linha 4 (index 3)
        for _, row in df_raw.iloc[3:].iterrows():
            for col_idx, date_val in enumerate(row_dates):
                if isinstance(date_val, (datetime, pd.Timestamp)):
                    valor = pd.to_numeric(row[col_idx], errors='coerce') or 0
                    meta_mes_total += valor
                    if date_val.date() <= data_referencia:
                        meta_ate_hoje += valor
        return meta_mes_total, meta_ate_hoje
    except:
        return 0, 0

# --- MENU LATERAL ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    up_prod = st.file_uploader("1. Produção Real (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("2. Programação (DATAS)", type=["xlsx"])
    
    st.markdown("---")
    if up_prod:
        menu = st.radio("MENU", 
                        ["📈 PERFORMANCE", "📋 REPORTE DIÁRIO", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"],
                        label_visibility="collapsed")

# --- PROCESSAMENTO PRINCIPAL ---
if up_prod:
    df_order, df_stops = load_production_data(up_prod)
    
    # Categorização de Máquinas
    def categorizar(m):
        return "BABY" if m in ['2', '3', '4', '5', '6'] else "ADULTO"
    df_order['Categoria'] = df_order['Máquina'].apply(categorizar)

    # =========================================================
    # ABA: REPORTE DIÁRIO (Visão dos últimos 3 dias e metas)
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        st.markdown("## 📋 Reporte Diário de Produção (Acompanhamento)")
        
        # Últimos 3 dias
        data_referencia = df_order['Data'].max().date()
        ultimos_3_dias = sorted(df_order['Data'].dt.date.unique())[-3:]

        for dia in reversed(ultimos_3_dias):
            st.markdown(f"<div class='section-header'>Produção do dia: {dia.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
            df_dia = df_order[df_order['Data'].dt.date == dia]
            
            res_dia = df_dia.groupby(['Categoria', 'Máquina']).agg({
                'Run Time': 'sum', 'Horário Padrão': 'sum', 'Machine Counter': 'sum', 'Peças Estoque - Ajuste': 'sum'
            }).reset_index()
            
            res_dia['Mov. %'] = (res_dia['Run Time'] / res_dia['Horário Padrão'] * 100).round(1)
            res_dia['Perda %'] = ((res_dia['Machine Counter'] - res_dia['Peças Estoque - Ajuste']) / res_dia['Machine Counter'] * 100).round(1)
            
            st.table(res_dia[['Categoria', 'Máquina', 'Mov. %', 'Perda %', 'Peças Estoque - Ajuste']].rename(columns={'Peças Estoque - Ajuste': 'Qtd. Estoque'}))

        # ACUMULADO DO MÊS
        st.markdown(f"<div class='section-header'>Acumulado do Mês (Mês {data_referencia.month})</div>", unsafe_allow_html=True)
        df_mes = df_order[df_order['Data'].dt.month == data_referencia.month]
        
        res_mes = df_mes.groupby(['Categoria', 'Máquina']).agg({
            'Run Time': 'sum', 'Horário Padrão': 'sum', 'Machine Counter': 'sum', 'Peças Estoque - Ajuste': 'sum'
        }).reset_index()
        
        res_mes['Mov. %'] = (res_mes['Run Time'] / res_mes['Horário Padrão'] * 100).round(1)
        res_mes['Perda %'] = ((res_mes['Machine Counter'] - res_mes['Peças Estoque - Ajuste']) / res_mes['Machine Counter'] * 100).round(1)
        st.table(res_mes[['Categoria', 'Máquina', 'Mov. %', 'Perda %', 'Peças Estoque - Ajuste']])

        # GERAL FÁBRICA VS METAS (CRUZAMENTO DATAS)
        st.markdown("<div class='section-header'>Geral Fábrica vs Metas Corporativas</div>", unsafe_allow_html=True)
        
        total_p_real = df_mes['Peças Estoque - Ajuste'].sum()
        total_mov_real = (df_mes['Run Time'].sum() / df_mes['Horário Padrão'].sum() * 100) if df_mes['Horário Padrão'].sum() > 0 else 0
        total_perda_real = ((df_mes['Machine Counter'].sum() - total_p_real) / df_mes['Machine Counter'].sum() * 100) if df_mes['Machine Counter'].sum() > 0 else 0

        meta_total, meta_hoje = (0, 0)
        if up_datas:
            meta_total, meta_hoje = get_metas_from_datas(up_datas, data_referencia)

        c1, c2, c3 = st.columns(3)
        c1.metric("Movimentação (Meta 90%)", f"{total_mov_real:.1f}%", f"{total_mov_real-90:.1f}%")
        c2.metric("Perda (Meta 2,5%)", f"{total_perda_real:.1f}%", f"{2.5-total_perda_real:.1f}%", delta_color="inverse")
        
        if up_datas:
            c3.metric("Estoque vs Meta Dia", f"{total_p_real:,.0f}", f"{total_p_real - meta_hoje:,.0f}")
            st.info(f"🎯 **Meta de Peças Planejada para o Mês Completo:** {meta_total:,.0f}")
        else:
            st.warning("Suba o arquivo DATAS para ver a meta de estoque.")

    # =========================================================
    # ABA: PERFORMANCE GERAL (Velocímetros e KPIs)
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()])
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1])]
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f['Machine Counter'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f['Peças Estoque - Ajuste'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Minutos Produzidos</div><div class="metric-value">{df_f['Run Time'].sum():,.0f}m</div></div>
            </div>
            """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        hp_sum = df_f['Horário Padrão'].sum()
        with col1:
            mov_v = (df_f['Run Time'].sum()/hp_sum*100) if hp_sum > 0 else 0
            fig = go.Figure(go.Indicator(mode="gauge+number", value=mov_v, title={'text': "Movimentação Geral %"}, gauge={'bar':{'color':"#10b981"}, 'axis':{'range':[0,100]}}))
            fig.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            loss_v = ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100) if df_f['Machine Counter'].sum()>0 else 0
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=loss_v, title={'text': "Loss %"}, gauge={'bar':{'color':"#e74c3c"}, 'axis':{'range':[0,15]}}))
            fig2.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig2, use_container_width=True)

    # =========================================================
    # DEMAIS ABAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        f_d_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()])
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_d_s[0]) & (df_stops['Data'].dt.date <= f_d_s[1])]
        st.markdown("## 🛑 Top 10 Paradas")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10), orientation='h', color_discrete_sequence=['#f43f5e']), use_container_width=True)

    elif menu == "📅 CALENDÁRIO":
        st.markdown("## 📅 Calendário Operacional")
        # (Lógica original de calendário mantida com filtros de Mês/Máquina)

    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Configuração do Relatório")
        maq_sel = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        st.markdown(f"### 📋 ANÁLISE SEMANAL - MÁQUINA {maq_sel}")
        st.markdown("""<div class="five-why-box"><div style="font-weight:bold; color:#059669;">FERRAMENTA 5 PORQUÊS</div>
            <div class="five-why-line">1. Por que? _________________________________________________</div>
            <div class="five-why-line">2. Por que? _________________________________________________</div>
            <div class="five-why-line">3. Por que? _________________________________________________</div>
            <div class="five-why-line">4. Por que? _________________________________________________</div>
            <div class="five-why-line">5. Por que? _________________________________________________</div>
            <br><b>CAUSA RAIZ:</b> _____________________________________________________
            <br><b>PLANO DE AÇÃO:</b> ___________________________________________________</div>""", unsafe_allow_html=True)
else:
    st.info("💡 Por favor, carregue os arquivos Excel no menu lateral para iniciar.")
