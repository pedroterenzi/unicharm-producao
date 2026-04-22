import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Dashboard Produção Unicharm", layout="wide")

# 2. CARREGAMENTO DOS DADOS
@st.cache_data
def load_uploaded_data(file):
    xls = pd.ExcelFile(file)
    
    # Carregando as abas (nomes exatos informados)
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    
    # Limpeza de espaços em branco nos nomes das colunas (Resolve o problema do "Máquina ")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Conversão de Datas
    df_order['Data'] = pd.to_datetime(df_order['Data'])
    df_stops['Data'] = pd.to_datetime(df_stops['Data'])
    
    return df_order, df_stops

# 3. INTERFACE
st.title("📊 Dashboard Operacional Unicharm")
st.markdown("---")

uploaded_file = st.file_uploader("Suba o arquivo 'Relatório das paradas Operacionais(Oficial).xlsm'", type=["xlsm"])

if uploaded_file is not None:
    try:
        df_order, df_stops = load_uploaded_data(uploaded_file)
        
        # --- SIDEBAR - FILTROS ---
        st.sidebar.header("⚙️ Filtros")
        
        # Filtro de Data
        min_date = df_order['Data'].min().date()
        max_date = df_order['Data'].max().date()
        dr = st.sidebar.date_input("Período de Análise", [min_date, max_date])
        
        # Proteção para garantir intervalo de datas
        start_date = dr[0] if len(dr) > 0 else min_date
        end_date = dr[1] if len(dr) > 1 else max_date

        maquinas = st.sidebar.multiselect("Filtrar Máquina", options=sorted(df_order['Máquina'].unique()))
        turnos = st.sidebar.multiselect("Filtrar Turno", options=sorted(df_order['Turno'].unique()))

        # Aplicando os filtros
        df_o_f = df_order[(df_order['Data'].dt.date >= start_date) & (df_order['Data'].dt.date <= end_date)]
        df_s_f = df_stops[(df_stops['Data'].dt.date >= start_date) & (df_stops['Data'].dt.date <= end_date)]

        if maquinas:
            df_o_f = df_o_f[df_o_f['Máquina'].isin(maquinas)]
            df_s_f = df_s_f[df_s_f['Máquina'].isin(maquinas)]
        if turnos:
            df_o_f = df_o_f[df_o_f['Turno'].isin(turnos)]
            df_s_f = df_s_f[df_s_f['Turno'].isin(turnos)]

        # --- CÁLCULOS DOS INDICADORES ---
        
        # 1. Movimentação: (Run Time / Horário Padrão)
        rt_sum = df_o_f['Run Time'].sum()
        hp_sum = df_o_f['Horário Padrão'].sum()
        movimentacao = (rt_sum / hp_sum * 100) if hp_sum > 0 else 0

        # 2. Loss: (Machine Counter - Peças Estoque - Ajuste) / Machine Counter
        mc_sum = df_o_f['Machine Counter'].sum()
        estoque_sum = df_o_f['Peças Estoque - Ajuste'].sum()
        loss = ((mc_sum - estoque_sum) / mc_sum * 100) if mc_sum > 0 else 0

        # 3. Média de Peças por Turno (Soma por dia/turno/máquina e tira a média)
        df_turno_agrupado = df_o_f.groupby(['Data', 'Turno', 'Máquina'])['Peças Estoque - Ajuste'].sum().reset_index()
        media_turno = df_turno_agrupado['Peças Estoque - Ajuste'].mean() if not df_turno_agrupado.empty else 0

        # --- EXIBIÇÃO ---
        
        # Cards de Resumo
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Machine Counter", f"{mc_sum:,.0f}")
        kpi2.metric("Horário Padrão (Min)", f"{hp_sum:,.0f}")
        kpi3.metric("Peças (Ajuste)", f"{estoque_sum:,.0f}")
        kpi4.metric("Média Peças/Turno", f"{media_turno:,.0f}")

        st.markdown("---")

        # Velocímetros
        col_m, col_l = st.columns(2)

        def create_gauge(title, value, color, threshold):
            return go.Figure(go.Indicator(
                mode="gauge+number", value=value,
                title={'text': title, 'font': {'size': 20}},
                gauge={'axis': {'range': [None, 100]},
                       'bar': {'color': color},
                       'threshold': {'line': {'color': "black", 'width': 4}, 'value': threshold}}
            )).update_layout(height=300)

        col_m.plotly_chart(create_gauge("Movimentação (%)", movimentacao, "#2ecc71", 85), use_container_width=True)
        col_l.plotly_chart(create_gauge("Loss (%)", loss, "#e74c3c", 5), use_container_width=True)

        # Análise de Paradas
        st.subheader("⚠️ Piores Paradas do Período")
        if not df_s_f.empty:
            df_p = df_s_f.groupby('Problema').agg({'Minutos': 'sum', 'QTD': 'sum'}).reset_index()
            df_p = df_p.sort_values('Minutos', ascending=False).head(10)
            
            total_min_p = df_s_f['Minutos'].sum()
            df_p['Influência (%)'] = (df_p['Minutos'] / total_min_p * 100).round(2)

            fig_p = px.bar(df_p, x='Minutos', y='Problema', orientation='h', text='Minutos',
                           color='Minutos', color_continuous_scale='Reds')
            fig_p.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_p, use_container_width=True)
            st.dataframe(df_p[['Problema', 'QTD', 'Influência (%)']], hide_index=True, use_container_width=True)
        else:
            st.info("Sem dados de paradas para este filtro.")

        # Ranking de Máquinas
        st.subheader("🏆 Ranking de Eficiência por Máquina")
        
        df_rank = df_o_f.groupby('Máquina').agg({
            'Run Time': 'sum',
            'Horário Padrão': 'sum',
            'Machine Counter': 'sum',
            'Peças Estoque - Ajuste': 'sum'
        }).reset_index()

        df_rank['Movimentação (%)'] = (df_rank['Run Time'] / df_rank['Horário Padrão'] * 100).round(2)
        df_rank['Loss (%)'] = ((df_rank['Machine Counter'] - df_rank['Peças Estoque - Ajuste']) / df_rank['Machine Counter'] * 100).round(2)
        
        # Média por turno por máquina específica
        med_rank = df_turno_agrupado.groupby('Máquina')['Peças Estoque - Ajuste'].mean().reset_index()
        df_rank = df_rank.merge(med_rank, on='Máquina')
        df_rank.columns = ['Máquina', 'RT', 'HP', 'MC', 'P_EST', 'Movimentação (%)', 'Loss (%)', 'Média Peças/Turno']

        st.dataframe(df_rank[['Máquina', 'Movimentação (%)', 'Loss (%)', 'Média Peças/Turno']]
                     .sort_values('Movimentação (%)', ascending=False).style.format({
                         'Movimentação (%)': '{:.2f}%', 'Loss (%)': '{:.2f}%', 'Média Peças/Turno': '{:,.0f}'
                     }).background_gradient(subset=['Movimentação (%)'], cmap='RdYlGn'), use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
        st.write("Verifique se as abas 'Result by order' e 'Stop machine item' existem com as colunas corretas.")
else:
    st.info("Aguardando upload do arquivo Excel.")
