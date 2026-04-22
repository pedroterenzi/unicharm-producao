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
    .metric-card { background: #1e293b; padding: 25px; border-radius: 15px; text-align: center; color: white; border: 1px solid rgba(255,255,255,0.1); margin-bottom: 20px;}
    .metric-value { font-size: 2.5rem; font-weight: 900; color: #10b981; }
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 10px; margin-top: 20px; }
    .day-card { background: #0f172a; border-radius: 10px; padding: 12px; min-height: 90px; border: 1px solid rgba(255,255,255,0.05); }
    .day-number { font-size: 1rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.8rem; font-weight: 700; text-align: right; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Lista de colunas que DEVEM ser numéricas
    cols_numericas = [
        'Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed',
        'Manutenção', 'Limpeza', 'Ajuste de Partida de Máquina', 'Troca de Tamanho de Máquina', 
        'Parada Programada', 'Parada por Falta / Problema de MP', 'Outros', 'Minutos', 'QTD'
    ]
    
    for df in [df_order, df_stops]:
        for c in cols_numericas:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    
    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order = df_order.dropna(subset=['Data'])
    
    # Converter IDs para Inteiro (remove o .0)
    df_order['Máquina'] = df_order['Máquina'].astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].astype(int).astype(str)
    df_stops['Máquina'] = df_stops['Máquina'].astype(float).fillna(0).astype(int).astype(str)
    
    return df_order, df_stops

uploaded_file = st.sidebar.file_uploader("📂 Carregar Relatório Oficial (.xlsm)", type=["xlsm"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)
    
    # --- FILTROS LATERAIS ---
    st.sidebar.header("⚙️ Filtros")
    
    min_d, max_d = df_order['Data'].min().date(), df_order['Data'].max().date()
    data_sel = st.sidebar.date_input("Período", [min_d, max_d])
    
    if len(data_sel) == 2:
        start_d, end_d = data_sel
        # Filtro Global de Data e Turno
        mask_order_base = (df_order['Data'].dt.date >= start_d) & (df_order['Data'].dt.date <= end_d)
        mask_stops_base = (df_stops['Data'].dt.date >= start_d) & (df_stops['Data'].dt.date <= end_d)
    else:
        st.warning("Selecione o início e o fim do período.")
        st.stop()

    turno_list = sorted(df_order['Turno'].unique())
    turno_sel = st.sidebar.multiselect("Turnos", options=turno_list, default=turno_list)
    
    # Filtro de Máquina (Apenas para as Abas de Performance e Tempos)
    maq_list = sorted(df_order['Máquina'].unique())
    maq_sel = st.sidebar.multiselect("Máquinas (Abas 1 e 2)", options=maq_list, default=maq_list)
    
    # Aplicando filtros para as abas 1 e 2
    df_f = df_order[mask_order_base & df_order['Turno'].isin(turno_sel) & df_order['Máquina'].isin(maq_sel)]
    df_s_f = df_stops[mask_stops_base & df_stops['Turno'].isin(turno_sel) & df_stops['Máquina'].isin(maq_sel)]
    
    # DataFrame para o Ranking (Ignora filtro de máquina, respeita Turno e Data)
    df_ranking_base = df_order[mask_order_base & df_order['Turno'].isin(turno_sel)]

    # --- ABAS ---
    tab_perf, tab_tempos, tab_eficiencia = st.tabs(["📈 Performance Geral", "⏰ Análise de Tempos", "🎯 Eficiência & OEE"])

    # =========================================================
    # ABA 1: PERFORMANCE GERAL
    # =========================================================
    with tab_perf:
        # KPIs Principais
        total_mc = df_f['Machine Counter'].sum()
        total_est = df_f['Peças Estoque - Ajuste'].sum()
        total_hp = df_f['Horário Padrão'].sum()
        total_rt = df_f['Run Time'].sum()
        
        mov_perc = (total_rt / total_hp * 100) if total_hp > 0 else 0
        loss_perc = ((total_mc - total_est) / total_mc * 100) if total_mc > 0 else 0
        media_p_turno = df_f.groupby(['Data', 'Turno', 'Máquina'])['Peças Estoque - Ajuste'].sum().mean()

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="metric-card"><small>MACHINE COUNTER</small><div class="metric-value">{total_mc:,.0f}</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-card"><small>HORÁRIO PADRÃO (MIN)</small><div class="metric-value">{total_hp:,.0f}</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-card"><small>ENVIADO ESTOQUE</small><div class="metric-value">{total_est:,.0f}</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="metric-card"><small>MÉDIA PEÇAS/TURNO</small><div class="metric-value">{media_p_turno:,.0f}</div></div>', unsafe_allow_html=True)

        col_v1, col_v2 = st.columns(2)
        with col_v1:
            fig_mov = go.Figure(go.Indicator(mode="gauge+number", value=mov_perc, title={'text': "Movimentação %", 'font': {'color': 'white'}},
                                            gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#10b981"}}))
            fig_mov.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350)
            st.plotly_chart(fig_mov, use_container_width=True)
        with col_v2:
            fig_loss = go.Figure(go.Indicator(mode="gauge+number", value=loss_perc, title={'text': "Loss %", 'font': {'color': 'white'}},
                                             gauge={'axis': {'range': [0, 20]}, 'bar': {'color': "#f43f5e"}}))
            fig_loss.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=350)
            st.plotly_chart(fig_loss, use_container_width=True)

        st.markdown("<h3 style='color: white;'>🏆 Ranking Geral de Máquinas (Todas as Máquinas)</h3>", unsafe_allow_html=True)
        ranking = df_ranking_base.groupby('Máquina').agg({
            'Run Time': 'sum', 'Horário Padrão': 'sum', 'Machine Counter': 'sum', 'Peças Estoque - Ajuste': 'sum'
        }).reset_index()
        ranking['Movimentação %'] = (ranking['Run Time'] / ranking['Horário Padrão'] * 100).round(2)
        ranking['Loss %'] = ((ranking['Machine Counter'] - ranking['Peças Estoque - Ajuste']) / ranking['Machine Counter'] * 100).round(2)
        
        st.dataframe(ranking[['Máquina', 'Movimentação %', 'Loss %']].sort_values('Movimentação %', ascending=False), use_container_width=True)

    # =========================================================
    # ABA 2: ANÁLISE DE TEMPOS
    # =========================================================
    with tab_tempos:
        st.markdown("<h3 style='color: white;'>⚠️ Top 10 Piores Paradas</h3>", unsafe_allow_html=True)
        col_t1, col_t2 = st.columns(2)
        
        with col_t1:
            stop_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
            fig_min = px.bar(stop_min, orientation='h', title="Minutos Totais", color_discrete_sequence=['#f43f5e'])
            fig_min.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color': 'white'})
            st.plotly_chart(fig_min, use_container_width=True)
            
        with col_t2:
            stop_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
            fig_qtd = px.bar(stop_qtd, orientation='h', title="Quantidade de Ocorrências", color_discrete_sequence=['#3b82f6'])
            fig_qtd.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color': 'white'})
            st.plotly_chart(fig_qtd, use_container_width=True)

        st.markdown("<h3 style='color: white;'>🛠️ Perdas e Tempos Não Operacionais</h3>", unsafe_allow_html=True)
        cols_perdas = [
            'Manutenção', 'Limpeza', 'Ajuste de Partida de Máquina', 'Troca de Tamanho de Máquina', 
            'Ajuste Após Troca de Tamanho Máquina', 'Troca de Optima', 'Troca de Dosetec', 
            'Checagem de Liberação do Operador', 'Parada Programada', 'Parada por Falta / Problema de MP', 
            'Liberação de Linha Qualidade', 'Amostragem (Sampling)', 'Segurança do Trabalho', 'Outros'
        ]
        # Garantir que apenas colunas existentes no DF sejam somadas
        cols_existentes = [c for c in cols_perdas if c in df_f.columns]
        perdas_total = df_f[cols_existentes].sum().sort_values(ascending=True)
        
        fig_perdas = px.bar(perdas_total, orientation='h', labels={'value': 'Total (Minutos/Kg)', 'index': 'Categoria'})
        fig_perdas.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color': 'white'})
        st.plotly_chart(fig_perdas, use_container_width=True)

    # =========================================================
    # ABA 3: EFICIÊNCIA & OEE & CALENDÁRIO
    # =========================================================
    with tab_eficiencia:
        st.markdown("<h3 style='color: white;'>🎯 Cálculo OEE por Máquina</h3>", unsafe_allow_html=True)
        
        df_f['Tempo_Turno'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        
        # OEE baseado na produtividade real
        oee_df = df_f.groupby('Máquina').agg({
            'Run Time': 'sum', 'Tempo_Turno': 'sum', 'Peças Estoque - Ajuste': 'sum', 
            'Machine Counter': 'sum', 'Average Speed': 'mean'
        }).reset_index()
        
        oee_df['Disponibilidade'] = (oee_df['Run Time'] / oee_df['Tempo_Turno']).clip(0, 1)
        oee_df['Qualidade'] = (oee_df['Peças Estoque - Ajuste'] / oee_df['Machine Counter']).fillna(0).clip(0, 1)
        # Performance: Peças produzidas vs Capacidade da velocidade média
        oee_df['Performance'] = (oee_df['Machine Counter'] / (oee_df['Average Speed'] * oee_df['Run Time'])).fillna(0).clip(0, 1)
        oee_df['OEE %'] = (oee_df['Disponibilidade'] * oee_df['Performance'] * oee_df['Qualidade'] * 100).round(2)
        
        st.dataframe(oee_df[['Máquina', 'Disponibilidade', 'Performance', 'Qualidade', 'OEE %']].sort_values('OEE %', ascending=False), use_container_width=True)

        st.markdown("---")
        st.markdown("<h3 style='color: white;'>📅 Calendário Operacional (Movimentação e Loss)</h3>", unsafe_allow_html=True)
        
        # Agrupar por dia para o calendário
        df_f['Dia'] = df_f['Data'].dt.day
        cal_data = df_f.groupby('Dia').agg({
            'Run Time': 'sum', 'Horário Padrão': 'sum', 'Machine Counter': 'sum', 'Peças Estoque - Ajuste': 'sum'
        })
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'] * 100).round(1)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'] * 100).round(1)

        cal_obj = calendar.Calendar(firstweekday=0)
        curr_date = datetime.now()
        dias = list(cal_obj.itermonthdays(curr_date.year, curr_date.month))
        
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#94a3b8; font-weight:900;">{n}</div>'
        
        for d in dias:
            if d == 0: html += '<div style="opacity:0"></div>'
            else:
                m = cal_data['Mov'].get(d, 0)
                l = cal_data['Loss'].get(d, 0)
                cor = "#059669" if m > 85 else "#dc2626" if m > 0 else "#0f172a"
                html += f'''<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">MOV: {m}%<br>LOSS: {l}%</div></div>'''
        st.markdown(html + '</div>', unsafe_allow_html=True)

else:
    st.info("Aguardando upload do arquivo Excel para processar os indicadores.")
