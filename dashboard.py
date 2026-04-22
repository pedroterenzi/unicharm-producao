import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Dashboard Produção Oficial", layout="wide")

# Estilização para melhorar o visual
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNÇÃO DE CARREGAMENTO INTELIGENTE
@st.cache_data
def load_uploaded_data(file):
    # Analisa o arquivo para identificar os nomes das abas
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
    
    # Busca por palavras-chave para evitar erro de nomes sensíveis a maiúsculas/minúsculas
    # Procura aba que contenha 'result' e 'order'
    sheet_order = next((s for s in sheet_names if "result" in s.lower() and "order" in s.lower()), None)
    # Procura aba que contenha 'stop' e 'item'
    sheet_stops = next((s for s in sheet_names if "stop" in s.lower() and "item" in s.lower()), None)

    if not sheet_order or not sheet_stops:
        st.error(f"❌ Erro: Abas necessárias não encontradas. As abas lidas foram: {sheet_names}")
        st.info("Dica: Verifique se as abas se chamam 'Result by order' e 'Stop Machine Item'.")
        st.stop()

    # Lendo os dados
    df_order = pd.read_excel(file, sheet_name=sheet_order)
    df_stops = pd.read_excel(file, sheet_name=sheet_stops)
    
    # Convertendo colunas de data
    df_order['Data'] = pd.to_datetime(df_order['Data'])
    df_stops['Data'] = pd.to_datetime(df_stops['Data'])
    
    return df_order, df_stops

# 3. INTERFACE DE ENTRADA
st.title("📊 Dashboard Operacional Unicharm")
st.write("Suba o arquivo Excel oficial para gerar os indicadores automáticos.")

uploaded_file = st.file_uploader("Selecione o arquivo .xlsm", type=["xlsm"])

if uploaded_file is not None:
    try:
        df_order, df_stops = load_uploaded_data(uploaded_file)
        st.success("✅ Dados processados com sucesso!")

        # --- SIDEBAR - FILTROS ---
        st.sidebar.header("⚙️ Filtros de Análise")
        
        # Filtro de Data
        min_date = df_order['Data'].min().date()
        max_date = df_order['Data'].max().date()
        
        dr = st.sidebar.date_input("Selecione o Período", [min_date, max_date])
        start_date = dr[0] if len(dr) > 0 else min_date
        end_date = dr[1] if len(dr) > 1 else max_date

        maquinas = st.sidebar.multiselect("Máquinas", options=sorted(df_order['Máquina'].unique()))
        turnos = st.sidebar.multiselect("Turnos", options=sorted(df_order['Turno'].unique()))

        # Aplicando os filtros nos DataFrames
        df_o_f = df_order[(df_order['Data'].dt.date >= start_date) & (df_order['Data'].dt.date <= end_date)]
        df_s_f = df_stops[(df_stops['Data'].dt.date >= start_date) & (df_stops['Data'].dt.date <= end_date)]

        if maquinas:
            df_o_f = df_o_f[df_o_f['Máquina'].isin(maquinas)]
            df_s_f = df_s_f[df_s_f['Máquina'].isin(maquinas)]
        if turnos:
            df_o_f = df_o_f[df_o_f['Turno'].isin(turnos)]
            df_s_f = df_s_f[df_s_f['Turno'].isin(turnos)]

        # --- CÁLCULOS DOS KPIs ---
        counter_total = df_o_f['Machine Counter'].sum()
        estoque_total = df_o_f['Peças Estoque'].sum()
        runtime_total = df_o_f['Run Time'].sum()
        std_time_total = df_o_f['Horário Padrão'].sum()

        movimentacao = (runtime_total / std_time_total * 100) if std_time_total > 0 else 0
        loss = ((counter_total - estoque_total) / counter_total * 100) if counter_total > 0 else 0

        # Média por Turno (Agrupado por dia/turno/máquina para precisão)
        df_t_sum = df_o_f.groupby(['Data', 'Turno', 'Máquina'])['Peças Estoque'].sum().reset_index()
        media_p_turno = df_t_sum['Peças Estoque'].mean() if not df_t_sum.empty else 0

        # --- EXIBIÇÃO VISUAL ---
        
        # KPIs - Big Numbers
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Machine Counter Total", f"{counter_total:,.0f}")
        k2.metric("Horário Padrão (Min)", f"{std_time_total:,.0f}")
        k3.metric("Peças p/ Estoque", f"{estoque_total:,.0f}")
        k4.metric("Média Peças/Turno", f"{media_p_turno:,.0f}")

        st.markdown("---")

        # Velocímetros (Gauges)
        col_m, col_l = st.columns(2)

        def draw_gauge(label, value, color, target):
            return go.Figure(go.Indicator(
                mode="gauge+number", value=value,
                title={'text': label, 'font': {'size': 20}},
                gauge={'axis': {'range': [None, 100]},
                       'bar': {'color': color},
                       'threshold': {'line': {'color': "black", 'width': 4}, 'value': target}}
            )).update_layout(height=280, margin=dict(l=30, r=30, t=50, b=20))

        col_m.plotly_chart(draw_gauge("Movimentação (%)", movimentacao, "#2ecc71", 85), use_container_width=True)
        col_l.plotly_chart(draw_gauge("Loss (%)", loss, "#e74c3c", 5), use_container_width=True)

        # Visão de Paradas
        st.subheader("⚠️ Análise de Paradas (Top 10)")
        if not df_s_f.empty:
            df_p = df_s_f.groupby('Problema').agg(Minutos=('Minutos', 'sum'), Qtd=('QTD Parada', 'sum')).reset_index()
            df_p = df_p.sort_values('Minutos', ascending=False).head(10)
            total_m = df_s_f['Minutos'].sum()
            df_p['% Influência'] = (df_p['Minutos'] / total_m * 100).round(1)

            fig_p = px.bar(df_p, x='Minutos', y='Problema', orientation='h', text='Minutos',
                           color='Minutos', color_continuous_scale='Reds', title="Piores Paradas (Minutos Totais)")
            fig_p.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_p, use_container_width=True)
        else:
            st.info("Sem dados de paradas registrados neste filtro.")

        # Ranking Final
        st.subheader("🏆 Ranking de Máquinas")
        df_rank = df_o_f.groupby('Máquina').agg(
            Movimentacao=('Run Time', lambda x: (x.sum() / df_o_f.loc[x.index, 'Horário Padrão'].sum() * 100)),
            Loss=('Machine Counter', lambda x: ((x.sum() - df_o_f.loc[x.index, 'Peças Estoque'].sum()) / x.sum() * 100))
        ).reset_index()
        
        # Adicionando média por turno por máquina
        med_m = df_t_sum.groupby('Máquina')['Peças Estoque'].mean().reset_index()
        df_rank = df_rank.merge(med_m, on='Máquina')
        df_rank.columns = ['Máquina', 'Movimentação (%)', 'Loss (%)', 'Média Peças/Turno']
        
        st.dataframe(df_rank.sort_values('Movimentação (%)', ascending=False).style.format({
            'Movimentação (%)': '{:.2f}%', 'Loss (%)': '{:.2f}%', 'Média Peças/Turno': '{:,.0f}'
        }).background_gradient(subset=['Movimentação (%)'], cmap='RdYlGn'), use_container_width=True)

    except Exception as e:
        st.error(f"Erro inesperado: {e}")

else:
    st.info("Aguardando upload do arquivo Excel...")
