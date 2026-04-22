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
    .metric-card { background: #1e293b; padding: 20px; border-radius: 15px; text-align: center; color: white; border: 1px solid rgba(255,255,255,0.1); }
    .metric-value { font-size: 2rem; font-weight: 900; color: #10b981; }
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; margin-top: 15px; }
    .day-card { background: #0f172a; border-radius: 10px; padding: 8px; min-height: 80px; border: 1px solid rgba(255,255,255,0.05); }
    .day-number { font-size: 0.9rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.7rem; font-weight: 700; text-align: right; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file):
    # Carregar abas
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    
    # Limpeza de nomes (espaços extras)
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Conversão de Tipos
    cols_num = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed']
    for c in cols_num:
        df_order[c] = pd.to_numeric(df_order[c], errors='coerce').fillna(0)
    
    df_order['Data'] = pd.to_datetime(df_order['Data'])
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int)
    
    return df_order, df_stops

uploaded_file = st.sidebar.file_uploader("📂 Carregar Relatório Oficial (.xlsm)", type=["xlsm"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)
    
    # --- FILTROS LATERAIS ---
    st.sidebar.header("Filtros")
    
    # Filtro de Data Simples
    min_d, max_d = df_order['Data'].min().date(), df_order['Data'].max().date()
    data_sel = st.sidebar.date_input("Período", [min_d, max_d])
    
    if len(data_sel) == 2:
        start_d, end_d = data_sel
        mask = (df_order['Data'].dt.date >= start_d) & (df_order['Data'].dt.date <= end_d)
        df_f = df_order[mask]
        df_s_f = df_stops[(df_stops['Data'].dt.date >= start_d) & (df_stops['Data'].dt.date <= end_d)]
    else:
        df_f = df_order
        df_s_f = df_stops

    maq_list = sorted(df_f['Máquina'].unique())
    maq_sel = st.sidebar.multiselect("Máquinas", options=maq_list, default=maq_list)
    
    turno_list = sorted(df_f['Turno'].unique())
    turno_sel = st.sidebar.multiselect("Turnos", options=turno_list, default=turno_list)
    
    df_f = df_f[df_f['Máquina'].isin(maq_sel) & df_f['Turno'].isin(turno_sel)]
    df_s_f = df_s_f[df_s_f['Máquina'].isin(maq_sel) & df_s_f['Turno'].isin(turno_sel)]

    # --- ABAS ---
    tab_perf, tab_tempos, tab_eficiencia = st.tabs(["📈 Performance Geral", "⏰ Análise de Tempos", "🎯 Eficiência & OEE"])

    # =========================================================
    # ABA 1: PERFORMANCE GERAL
    # =========================================================
    with tab_perf:
        st.markdown("<h2 style='color: white;'>📈 Indicadores de Performance</h2>", unsafe_allow_html=True)
        
        # Cálculos Consolidados
        total_mc = df_f['Machine Counter'].sum()
        total_est = df_f['Peças Estoque - Ajuste'].sum()
        total_rt = df_f['Run Time'].sum()
        total_hp = df_f['Horário Padrão'].sum()
        
        mov_perc = (total_rt / total_hp * 100) if total_hp > 0 else 0
        loss_perc = ((total_mc - total_est) / total_mc * 100) if total_mc > 0 else 0
        
        # Cards
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="metric-card"><small>MACHINE COUNTER</small><div class="metric-value">{total_mc:,.0f}</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-card"><small>HORÁRIO PADRÃO (MIN)</small><div class="metric-value">{total_hp:,.0f}</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-card"><small>ENVIADO ESTOQUE</small><div class="metric-value">{total_est:,.0f}</div></div>', unsafe_allow_html=True)
        
        # Média peças/turno (Agrupado conforme solicitado)
        media_p_turno = df_f.groupby(['Data', 'Turno', 'Máquina'])['Peças Estoque - Ajuste'].sum().mean()
        c4.markdown(f'<div class="metric-card"><small>MÉDIA PEÇAS/TURNO</small><div class="metric-value">{media_p_turno:,.0f}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Velocímetros
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            fig_mov = go.Figure(go.Indicator(mode="gauge+number", value=mov_perc, title={'text': "Movimentação %"},
                                            gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#10b981"}}))
            fig_mov.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"})
            st.plotly_chart(fig_mov, use_container_width=True)
        
        with col_v2:
            fig_loss = go.Figure(go.Indicator(mode="gauge+number", value=loss_perc, title={'text': "Loss %"},
                                             gauge={'axis': {'range': [0, 20]}, 'bar': {'color': "#f43f5e"}}))
            fig_loss.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"})
            st.plotly_chart(fig_loss, use_container_width=True)

        # Ranking Máquinas (Melhor para Pior por Movimentação)
        st.markdown("<h3 style='color: white;'>🏆 Ranking de Máquinas</h3>", unsafe_allow_html=True)
        ranking = df_f.groupby('Máquina').agg({
            'Run Time': 'sum',
            'Horário Padrão': 'sum',
            'Machine Counter': 'sum',
            'Peças Estoque - Ajuste': 'sum'
        }).reset_index()
        
        ranking['Movimentação %'] = (ranking['Run Time'] / ranking['Horário Padrão'] * 100).round(2)
        ranking['Loss %'] = ((ranking['Machine Counter'] - ranking['Peças Estoque - Ajuste']) / ranking['Machine Counter'] * 100).round(2)
        
        st.dataframe(ranking[['Máquina', 'Movimentação %', 'Loss %']].sort_values('Movimentação %', ascending=False), use_container_width=True)

    # =========================================================
    # ABA 2: ANÁLISE DE TEMPOS (PARADAS)
    # =========================================================
    with tab_tempos:
        st.markdown("<h2 style='color: white;'>⚠️ Análise de Paradas e Perdas</h2>", unsafe_allow_html=True)
        
        col_t1, col_t2 = st.columns(2)
        
        # Paradas por Minutos (Maior para Cima)
        with col_t1:
            st.write("Top 10 Paradas (Minutos)")
            stop_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
            fig_stop_min = px.bar(stop_min, orientation='h', color_discrete_sequence=['#f43f5e'])
            st.plotly_chart(fig_stop_min, use_container_width=True)
            
        # Paradas por Quantidade
        with col_t2:
            st.write("Top 10 Paradas (Quantidade)")
            stop_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
            fig_stop_qtd = px.bar(stop_qtd, orientation='h', color_discrete_sequence=['#3b82f6'])
            st.plotly_chart(fig_stop_qtd, use_container_width=True)

        # Tabela de Tempos Não Operacionais
        st.write("Distribuição de Perdas (Kg) e Tempos (Min)")
        cols_perdas = ['Manutenção', 'Limpeza', 'Ajuste de Partida de Máquina', 'Troca de Tamanho de Máquina', 
                       'Parada Programada', 'Parada por Falta / Problema de MP', 'Outros']
        perdas_total = df_f[cols_perdas].sum().sort_values(ascending=False)
        st.bar_chart(perdas_total)

    # =========================================================
    # ABA 3: EFICIÊNCIA & OEE & CALENDÁRIO
    # =========================================================
    with tab_eficiencia:
        st.markdown("<h2 style='color: white;'>🎯 OEE e Calendário Industrial</h2>", unsafe_allow_html=True)
        
        # Lógica OEE por Máquina
        # Tempo Total do Turno (Conforme regra fornecida)
        df_f['Tempo_Turno'] = df_f['Turno'].map({1: 455, 2: 440, 3: 415})
        
        oee_df = df_f.groupby('Máquina').agg({
            'Run Time': 'sum',
            'Tempo_Turno': 'sum',
            'Peças Estoque - Ajuste': 'sum',
            'Machine Counter': 'sum',
            'Average Speed': 'mean'
        }).reset_index()
        
        oee_df['Disponibilidade'] = (oee_df['Run Time'] / oee_df['Tempo_Turno']).clip(0, 1)
        oee_df['Qualidade'] = (oee_df['Peças Estoque - Ajuste'] / oee_df['Machine Counter']).clip(0, 1)
        # Performance: Peças / (Velocidade * Tempo de Corrida)
        oee_df['Performance'] = (oee_df['Machine Counter'] / (oee_df['Average Speed'] * oee_df['Run Time'])).clip(0, 1)
        
        oee_df['OEE %'] = (oee_df['Disponibilidade'] * oee_df['Performance'] * oee_df['Qualidade'] * 100).round(2)
        
        st.write("Visão OEE por Máquina")
        st.dataframe(oee_df[['Máquina', 'Disponibilidade', 'Performance', 'Qualidade', 'OEE %']].sort_values('OEE %', ascending=False))

        st.markdown("---")
        
        # VISÃO CALENDÁRIO
        st.write("📅 Calendário de Movimentação (Média Diária)")
        
        cal_data = df_f.groupby(df_f['Data'].dt.day).agg({'Run Time': 'sum', 'Horário Padrão': 'sum', 'Machine Counter': 'sum', 'Peças Estoque - Ajuste': 'sum'})
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'] * 100).round(1)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'] * 100).round(1)

        # Montagem do grid do calendário (Simplificado do exemplo de trading)
        today = datetime.now()
        cal_obj = calendar.Calendar(firstweekday=0)
        dias = list(cal_obj.itermonthdays(today.year, today.month))
        
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#94a3b8; font-weight:900; font-size:0.7rem;">{n}</div>'
        
        for d in dias:
            if d == 0: html += '<div style="opacity:0"></div>'
            else:
                val_mov = cal_data['Mov'].get(d, 0)
                val_loss = cal_data['Loss'].get(d, 0)
                cor_bg = "#059669" if val_mov > 85 else "#dc2626" if val_mov > 0 else "#0f172a"
                html += f'''<div class="day-card" style="background:{cor_bg}">
                            <span class="day-number">{d}</span>
                            <div class="day-status">MOV: {val_mov}%<br>LOSS: {val_loss}%</div>
                          </div>'''
        st.markdown(html + '</div>', unsafe_allow_html=True)

else:
    st.info("Aguardando upload do arquivo para processar.")
