import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM (Botões Limpos, Cards e Impressão) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Remover bolinhas do radio e criar botões uniformes */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label > div:first-child { display: none !important; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 12px 20px !important; border-radius: 10px !important;
        margin-bottom: 8px !important; color: white !important;
        transition: all 0.3s ease; cursor: pointer; width: 100%;
        display: block !important; text-align: center; font-weight: 600;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
        border: none !important; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
    }

    /* Cards de Métricas */
    .metric-card {
        background: #1e293b; padding: 15px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        min-height: 80px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { color: #10b981; font-size: 1.4rem; font-weight: 900; }

    /* Área 5 Porquês (Fundo Branco para caneta) */
    .five-why-box { border: 1px solid #000; border-radius: 5px; padding: 15px; background: #ffffff; color: #000; margin-top: 10px; }
    .five-why-line { border-bottom: 1px solid #000; padding: 8px 0; font-size: 0.9rem; }
    
    .section-header {
        background: #1e293b; padding: 10px; border-radius: 5px;
        color: #10b981; font-weight: 800; text-transform: uppercase;
        margin-top: 20px; border-left: 5px solid #10b981; font-size: 0.9rem;
    }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E TRATAMENTO
@st.cache_data
def load_production_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    
    nums = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed']
    for col in nums:
        if col in df_order.columns:
            df_order[col] = pd.to_numeric(df_order[col], errors='coerce').fillna(0)
    
    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    # Remover NaT para evitar erro no sorteio
    df_order = df_order.dropna(subset=['Data'])
    
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    
    # Categorização BABY vs ADULTO
    def categorize(m):
        return "BABY" if m in ['2', '3', '4', '5', '6'] else "ADULTO"
    df_order['Categoria'] = df_order['Máquina'].apply(categorize)
    
    return df_order, df_stops

@st.cache_data
def load_planner_meta(file, data_ref):
    try:
        xls = pd.ExcelFile(file)
        target = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        if not target: return 0, 0
        
        # Datas na linha 3 (index 2)
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        row_dates = df_raw.iloc[2, :].tolist()
        
        meta_mes = 0
        meta_hoje = 0
        
        for _, row in df_raw.iloc[3:].iterrows():
            for col_idx, d_val in enumerate(row_dates):
                if isinstance(d_val, (datetime, pd.Timestamp)):
                    val = pd.to_numeric(row[col_idx], errors='coerce') or 0
                    meta_mes += val
                    if d_val.date() <= data_ref:
                        meta_hoje += val
        return meta_mes, meta_hoje
    except:
        return 0, 0

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    up_prod = st.file_uploader("1. Produção (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("2. DATAS (.xlsx)", type=["xlsx"])
    
    st.markdown("---")
    if up_prod:
        menu = st.radio("MENU", ["📈 PERFORMANCE", "📋 REPORTE DIÁRIO", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"], label_visibility="collapsed")

if up_prod:
    df_order, df_stops = load_production_data(up_prod)

    # =========================================================
    # ABA: REPORTE DIÁRIO (FIX DO ERRO E NOVA VISÃO)
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        st.markdown("## 📋 Reporte Diário de Produção")
        
        # Fix do erro: Dropna e sort seguro
        datas_disponiveis = sorted(df_order['Data'].dt.date.unique())
        ultimos_3 = datas_disponiveis[-3:] if len(datas_disponiveis) >= 3 else datas_disponiveis
        
        data_atual = df_order['Data'].max().date()

        # Visão por dia (Últimos 3)
        for dia in reversed(ultimos_3):
            st.markdown(f"<div class='section-header'>DIA: {dia.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
            df_dia = df_order[df_order['Data'].dt.date == dia]
            
            res_dia = df_dia.groupby(['Categoria', 'Máquina']).agg({
                'Run Time': 'sum', 'Horário Padrão': 'sum', 'Machine Counter': 'sum', 'Peças Estoque - Ajuste': 'sum'
            }).reset_index()
            
            res_dia['Mov. %'] = (res_dia['Run Time'] / res_dia['Horário Padrão'].replace(0,1) * 100).round(1)
            res_dia['Perda %'] = ((res_dia['Machine Counter'] - res_dia['Peças Estoque - Ajuste']) / res_dia['Machine Counter'].replace(0,1) * 100).round(1)
            
            st.table(res_dia[['Categoria', 'Máquina', 'Mov. %', 'Perda %', 'Peças Estoque - Ajuste']].rename(columns={'Peças Estoque - Ajuste': 'Qtd Estoque'}))

        # Acumulado Mês
        st.markdown("<div class='section-header'>Acumulado do Mês Atual</div>", unsafe_allow_html=True)
        df_mes = df_order[df_order['Data'].dt.month == data_atual.month]
        res_mes = df_mes.groupby(['Categoria', 'Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        res_mes['Mov. %'] = (res_mes['Run Time']/res_mes['Horário Padrão'].replace(0,1)*100).round(1)
        res_mes['Perda %'] = ((res_mes['Machine Counter']-res_mes['Peças Estoque - Ajuste'])/res_mes['Machine Counter'].replace(0,1)*100).round(1)
        st.table(res_mes[['Categoria', 'Máquina', 'Mov. %', 'Perda %', 'Peças Estoque - Ajuste']])

        # Geral Fábrica vs Metas
        st.markdown("<div class='section-header'>Consolidado Fábrica vs Metas</div>", unsafe_allow_html=True)
        
        real_estoque = df_mes['Peças Estoque - Ajuste'].sum()
        real_mov = (df_mes['Run Time'].sum() / df_mes['Horário Padrão'].sum() * 100) if df_mes['Horário Padrão'].sum() > 0 else 0
        real_perda = ((df_mes['Machine Counter'].sum() - real_estoque) / df_mes['Machine Counter'].sum() * 100) if df_mes['Machine Counter'].sum() > 0 else 0

        m_mes, m_hoje = (0, 0)
        if up_datas:
            m_mes, m_hoje = load_planner_meta(up_datas, data_atual)

        c1, c2, c3 = st.columns(3)
        c1.metric("Movimentação (Meta 90%)", f"{real_mov:.1f}%", f"{real_mov-90:.1f}%")
        c2.metric("Perda (Meta 2,5%)", f"{real_perda:.1f}%", f"{2.5-real_perda:.1f}%", delta_color="inverse")
        if up_datas:
            c3.metric("Peças Estoque vs Meta Dia", f"{real_estoque:,.0f}", f"{real_estoque - m_hoje:,.0f}")
            st.info(f"🎯 **Meta Geral do Mês (DATAS):** {m_mes:,.0f} peças")
        else:
            st.warning("⚠️ Carregue o arquivo DATAS para comparar metas de estoque.")

    # =========================================================
    # ABA: PERFORMANCE GERAL (Visual anterior mantido)
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros")
        f_d = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='pd1')
        df_f = df_order[(df_order['Data'].dt.date >= f_d[0]) & (df_order['Data'].dt.date <= f_d[1])]
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Machine Counter", f"{df_f['Machine Counter'].sum():,.0f}")
        c2.metric("Peças Estoque", f"{df_f['Peças Estoque - Ajuste'].sum():,.0f}")
        c3.metric("Run Time Total", f"{df_f['Run Time'].sum():,.0f}m")
        
        # Velocímetros Maiores
        mov_v = (df_f['Run Time'].sum()/df_f['Horário Padrão'].sum()*100) if df_f['Horário Padrão'].sum()>0 else 0
        fig = go.Figure(go.Indicator(mode="gauge+number", value=mov_v, title={'text': "Movimentação %"}, gauge={'bar':{'color':"#10b981"}}))
        st.plotly_chart(fig, use_container_width=True)

    # =========================================================
    # ABA: ANÁLISE SEMANAL (Visual Paisagem / 5 Porquês)
    # =========================================================
    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Filtros Análise")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        periodo = st.sidebar.date_input("Período", [datetime.now()-timedelta(days=7), datetime.now()])
        
        df_b = df_order[(df_order['Data'].dt.date >= periodo[0]) & (df_order['Data'].dt.date <= periodo[1]) & (df_order['Máquina'] == maq_b)]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo[0]) & (df_stops['Data'].dt.date <= periodo[1]) & (df_stops['Máquina'] == maq_b)]
        
        st.markdown(f"### RELATÓRIO SEMANAL - MÁQUINA {maq_b}")
        
        # Top 5 Paradas Gráfico
        stop_data = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(5).reset_index()
        if not stop_data.empty:
            total_m = df_sb['Minutos'].sum()
            stop_data['Label'] = stop_data.apply(lambda r: f"{r['Minutos']} min ({(r['Minutos']/total_m*100):.1f}%)", axis=1)
            st.plotly_chart(px.bar(stop_data, x='Minutos', y='Problema', orientation='h', text='Label', title="Top 5 Paradas"), use_container_width=True)
            pior_p = stop_data.iloc[-1]['Problema']
        else:
            pior_p = "Sem Paradas"

        st.markdown(f"""
            <div class="five-why-box">
                <div style="font-weight:bold; color:#059669; font-size:1.2rem; border-bottom:2px solid #059669; margin-bottom:15px;">ANÁLISE 5 PORQUÊS</div>
                <div style="margin-bottom:10px;"><b>PROBLEMA FOCO:</b> {pior_p}</div>
                <div class="five-why-line">1. Por que? ________________________________________________________________</div>
                <div class="five-why-line">2. Por que? ________________________________________________________________</div>
                <div class="five-why-line">3. Por que? ________________________________________________________________</div>
                <div class="five-why-line">4. Por que? ________________________________________________________________</div>
                <div class="five-why-line">5. Por que? ________________________________________________________________</div>
                <br><b>CAUSA RAIZ:</b> __________________________________________________________________________
                <br><b>PLANO DE AÇÃO:</b> ________________________________________________________________________
            </div>
        """, unsafe_allow_html=True)

    # Restante das abas (TOP 10 e Calendário) mantidas com lógica anterior...
    elif menu == "🛑 TOP 10 PARADAS":
        st.markdown("## 🛑 Análise de Paradas")
        st.plotly_chart(px.bar(df_stops.groupby('Problema')['Minutos'].sum().sort_values().tail(10), orientation='h'))

    elif menu == "📅 CALENDÁRIO":
        st.markdown("## 📅 Calendário Operacional")

else:
    st.info("💡 Bem-vindo! Carregue os arquivos no menu lateral para iniciar.")
