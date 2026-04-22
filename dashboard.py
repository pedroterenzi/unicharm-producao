import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Performance Hub", page_icon="🏭")

# --- ESTILIZAÇÃO CSS PREMIUM (Ajustada para não espremer) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Estilo dos Cards de Métricas */
    .metric-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 30px 20px;
        border-radius: 15px;
        text-align: center;
        border: 1px solid rgba(255,255,255,0.1);
        margin-bottom: 10px;
    }
    .metric-title { color: #94a3b8; font-size: 0.9rem; font-weight: 700; text-transform: uppercase; margin-bottom: 10px; }
    .metric-value { color: #10b981; font-size: 2.2rem; font-weight: 900; }

    /* Estilo do Calendário */
    .calendar-grid { 
        display: grid; 
        grid-template-columns: repeat(7, 1fr); 
        gap: 12px; 
        margin-top: 20px; 
        width: 100%;
    }
    .day-card { 
        background: #1e293b; 
        border-radius: 12px; 
        padding: 15px; 
        min-height: 110px; 
        border: 1px solid rgba(255,255,255,0.05); 
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .day-number { font-size: 1.2rem; font-weight: 900; color: #f8fafc; }
    .day-status { font-size: 0.85rem; font-weight: 600; color: #ffffff; text-align: right; line-height: 1.4; }
    
    /* Tabelas */
    .stDataFrame { background-color: #0f172a; border-radius: 10px; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA ROBUSTA
@st.cache_data
def load_data(file):
    # Lendo as abas
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    
    # Limpeza absoluta de nomes de colunas
    df_order.columns = [str(c).strip() for c in df_order.columns]
    df_stops.columns = [str(c).strip() for c in df_stops.columns]
    
    # Lista de colunas críticas para converter em número
    cols_order = [
        'Machine Counter', 'Peças Estoque - Ajuste', 'Run Time', 'Horário Padrão', 'Average Speed',
        'Manutenção', 'Limpeza', 'Ajuste de Partida de Máquina', 'Troca de Tamanho de Máquina',
        'Ajuste Após Troca de Tamanho Máquina', 'Troca de Optima', 'Troca de Dosetec',
        'Checagem de Liberação do Operador', 'Parada Programada', 'Parada por Falta / Problema de MP',
        'Liberação de Linha Qualidade', 'Amostragem (Sampling)', 'Segurança do Trabalho', 'Outros'
    ]
    
    for col in cols_order:
        if col in df_order.columns:
            df_order[col] = pd.to_numeric(df_order[col], errors='coerce').fillna(0)

    # Limpeza aba de Paradas
    df_stops['Minutos'] = pd.to_numeric(df_stops['Minutos'], errors='coerce').fillna(0)
    df_stops['QTD'] = pd.to_numeric(df_stops['QTD'], errors='coerce').fillna(0)
    
    # Ajuste de Datas e IDs
    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order = df_order.dropna(subset=['Data'])
    
    # Forçar Máquina e Turno para inteiro sem .0
    df_order['Máquina'] = df_order['Máquina'].astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].astype(int).astype(str)
    
    return df_order, df_stops

# --- SIDEBAR ---
st.sidebar.title("💎 Industrial Analytics")
uploaded_file = st.sidebar.file_uploader("Suba o arquivo .xlsm", type=["xlsm"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)
    
    # Filtros
    st.sidebar.header("Filtros")
    min_d = df_order['Data'].min().date()
    max_d = df_order['Data'].max().date()
    data_sel = st.sidebar.date_input("Período", [min_d, max_d])
    
    # Turno e Máquina
    turno_list = sorted(df_order['Turno'].unique())
    turno_sel = st.sidebar.multiselect("Turnos", options=turno_list, default=turno_list)
    
    maq_list = sorted(df_order['Máquina'].unique())
    maq_sel = st.sidebar.multiselect("Máquinas (Abas 1 e 2)", options=maq_list, default=maq_list)
    
    # Aplicação dos Filtros
    if len(data_sel) == 2:
        df_f = df_order[(df_order['Data'].dt.date >= data_sel[0]) & (df_order['Data'].dt.date <= data_sel[1])]
        df_s_f = df_stops[(df_stops['Data'].dt.date >= data_sel[0]) & (df_stops['Data'].dt.date <= data_sel[1])]
    else:
        df_f, df_s_f = df_order.copy(), df_stops.copy()

    # Filtro de Turno (Aplica a tudo)
    df_f = df_f[df_f['Turno'].isin(turno_sel)]
    df_s_f = df_s_f[df_s_f['Turno'].isin(turno_sel)]
    
    # DataFrame filtrado por máquina para abas 1 e 2
    df_maq_f = df_f[df_f['Máquina'].isin(maq_sel)]
    df_s_maq_f = df_s_f[df_s_f['Máquina'].isin(maq_sel)]

    # ABAS
    tab1, tab2, tab3 = st.tabs(["📈 Performance Geral", "⏰ Análise de Tempos", "🎯 Eficiência (OEE)"])

    # --- ABA 1: PERFORMANCE ---
    with tab1:
        # Cálculos
        mc = df_maq_f['Machine Counter'].sum()
        est = df_maq_f['Peças Estoque - Ajuste'].sum()
        rt = df_maq_f['Run Time'].sum()
        hp = df_maq_f['Horário Padrão'].sum()
        
        mov = (rt/hp*100) if hp > 0 else 0
        loss = ((mc-est)/mc*100) if mc > 0 else 0
        media_turno = df_maq_f.groupby(['Data', 'Turno', 'Máquina'])['Peças Estoque - Ajuste'].sum().mean()

        # Layout Metrics
        m1, m2, m3, m4 = st.columns(4)
        with m1: st.markdown(f'<div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{mc:,.0f}</div></div>', unsafe_allow_html=True)
        with m2: st.markdown(f'<div class="metric-card"><div class="metric-title">Horário Padrão</div><div class="metric-value">{hp:,.0f}</div></div>', unsafe_allow_html=True)
        with m3: st.markdown(f'<div class="metric-card"><div class="metric-title">Enviado Estoque</div><div class="metric-value">{est:,.0f}</div></div>', unsafe_allow_html=True)
        with m4: st.markdown(f'<div class="metric-card"><div class="metric-title">Peças/Turno</div><div class="metric-value">{media_turno:,.0f}</div></div>', unsafe_allow_html=True)

        c_v1, c_v2 = st.columns(2)
        with c_v1:
            fig1 = go.Figure(go.Indicator(mode="gauge+number", value=mov, title={'text': "Movimentação %"}, gauge={'bar':{'color':"#10b981"}, 'axis':{'range':[0,100]}}))
            fig1.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350)
            st.plotly_chart(fig1, use_container_width=True)
        with c_v2:
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=loss, title={'text': "Loss %"}, gauge={'bar':{'color':"#f43f5e"}, 'axis':{'range':[0,20]}}))
            fig2.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350)
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### 🏆 Ranking de Máquinas (Todas as Máquinas)")
        rk = df_f.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        rk['Movimentação %'] = (rk['Run Time']/rk['Horário Padrão']*100)
        rk['Loss %'] = ((rk['Machine Counter']-rk['Peças Estoque - Ajuste'])/rk['Machine Counter']*100)
        st.dataframe(rk[['Máquina', 'Movimentação %', 'Loss %']].sort_values('Movimentação %', ascending=False), use_container_width=True)

    # --- ABA 2: TEMPOS ---
    with tab2:
        st.markdown("<h3 style='color:white'>⏰ Detalhamento de Paradas</h3>", unsafe_allow_html=True)
        t1, t2 = st.columns(2)
        with t1:
            p_min = df_s_maq_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
            st.plotly_chart(px.bar(p_min, orientation='h', title="Top 10 Minutos", color_discrete_sequence=['#f43f5e']), use_container_width=True)
        with t2:
            p_qtd = df_s_maq_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
            st.plotly_chart(px.bar(p_qtd, orientation='h', title="Top 10 Quantidade", color_discrete_sequence=['#3b82f6']), use_container_width=True)

        # Tempos Não Operacionais
        cols_perdas = ['Manutenção', 'Limpeza', 'Ajuste de Partida de Máquina', 'Troca de Tamanho de Máquina', 'Parada Programada', 'Parada por Falta / Problema de MP', 'Outros']
        perdas = df_maq_f[cols_perdas].sum().sort_values(ascending=True)
        st.plotly_chart(px.bar(perdas, orientation='h', title="Impacto por Categoria (Min/Kg)", color_discrete_sequence=['#fbbf24']), use_container_width=True)

    # --- ABA 3: OEE E CALENDÁRIO ---
    with tab3:
        # Cálculo OEE
        df_f['Tempo_Disponivel'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        oee = df_f.groupby('Máquina').agg({'Run Time':'sum', 'Tempo_Disponivel':'sum', 'Peças Estoque - Ajuste':'sum', 'Machine Counter':'sum', 'Average Speed':'mean'}).reset_index()
        oee['Disponibilidade'] = (oee['Run Time'] / oee['Tempo_Disponivel']).clip(0,1)
        oee['Qualidade'] = (oee['Peças Estoque - Ajuste'] / oee['Machine Counter']).fillna(0).clip(0,1)
        oee['Performance'] = (oee['Machine Counter'] / (oee['Average Speed'] * oee['Run Time'])).fillna(0).clip(0,1)
        oee['OEE %'] = (oee['Disponibilidade'] * oee['Performance'] * oee['Qualidade'] * 100).round(2)
        
        st.markdown("### 🎯 Eficiência OEE por Máquina")
        st.dataframe(oee[['Máquina', 'Disponibilidade', 'Performance', 'Qualidade', 'OEE %']].sort_values('OEE %', ascending=False), use_container_width=True)

        st.markdown("---")
        st.markdown("### 📅 Calendário de Operação")
        
        # Dados Calendário
        df_f['Dia'] = df_f['Data'].dt.day
        cal_df = df_f.groupby('Dia').agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'})
        cal_df['Mov'] = (cal_df['Run Time']/cal_df['Horário Padrão']*100).round(1)
        cal_df['Loss'] = ((cal_df['Machine Counter']-cal_df['Peças Estoque - Ajuste'])/cal_df['Machine Counter']*100).round(1)

        cal = calendar.Calendar(firstweekday=0)
        now = datetime.now()
        days = list(cal.itermonthdays(now.year, now.month))
        
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: 
            html += f'<div style="text-align:center; color:#64748b; font-weight:900; font-size:0.8rem;">{n}</div>'
        
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                m_val = cal_df['Mov'].get(d, 0)
                l_val = cal_df['Loss'].get(d, 0)
                bg = "#059669" if m_val > 85 else "#dc2626" if m_val > 0 else "#0f172a"
                html += f'''<div class="day-card" style="background:{bg}">
                            <span class="day-number">{d}</span>
                            <div class="day-status">MOV: {m_val}%<br>LOSS: {l_val}%</div>
                          </div>'''
        st.markdown(html + '</div>', unsafe_allow_html=True)

else:
    st.info("Aguardando upload do arquivo Excel.")
