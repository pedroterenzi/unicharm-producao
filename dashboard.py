import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import unicodedata

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Dashboard Produção Oficial", layout="wide")

# Função para remover acentos e padronizar nomes de colunas
def normalize_column_name(name):
    n = str(name).strip()
    n = "".join(c for c in unicodedata.normalize('NFD', n) if unicodedata.category(c) != 'Mn')
    return n.lower()

# 2. FUNÇÃO DE CARREGAMENTO INTELIGENTE
@st.cache_data
def load_uploaded_data(file):
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
    
    sheet_order = next((s for s in sheet_names if "result" in s.lower() and "order" in s.lower()), None)
    sheet_stops = next((s for s in sheet_names if "stop" in s.lower() and "item" in s.lower()), None)

    if not sheet_order or not sheet_stops:
        st.error(f"❌ Abas não encontradas. Detectadas: {sheet_names}")
        st.stop()

    df_order = pd.read_excel(file, sheet_name=sheet_order)
    df_stops = pd.read_excel(file, sheet_name=sheet_stops)
    
    # Normalizar nomes das colunas (tira acento, espaço e põe minúsculo)
    df_order.columns = [normalize_column_name(col) for col in df_order.columns]
    df_stops.columns = [normalize_column_name(col) for col in df_stops.columns]
    
    # Mapeamento para o código funcionar independente do acento no Excel
    # Procura 'data', 'maquina', 'turno', 'run time', 'horario padrao', 'machine counter', 'pecas estoque', 'problema', 'minutos', 'qtd parada'
    
    df_order['data'] = pd.to_datetime(df_order['data'])
    df_stops['data'] = pd.to_datetime(df_stops['data'])
    
    return df_order, df_stops

# 3. INTERFACE
st.title("📊 Dashboard Operacional Unicharm")
uploaded_file = st.file_uploader("Selecione o arquivo .xlsm", type=["xlsm"])

if uploaded_file is not None:
    try:
        df_order, df_stops = load_uploaded_data(uploaded_file)
        
        # Filtros usando nomes normalizados
        st.sidebar.header("⚙️ Filtros")
        
        min_date = df_order['data'].min().date()
        max_date = df_order['data'].max().date()
        dr = st.sidebar.date_input("Período", [min_date, max_date])
        start_date = dr[0] if len(dr) > 0 else min_date
        end_date = dr[1] if len(dr) > 1 else max_date

        maquinas = st.sidebar.multiselect("Máquinas", options=sorted(df_order['maquina'].unique()))
        turnos = st.sidebar.multiselect("Turnos", options=sorted(df_order['turno'].unique()))

        # Aplicando Filtros
        df_o_f = df_order[(df_order['data'].dt.date >= start_date) & (df_order['data'].dt.date <= end_date)]
        df_s_f = df_stops[(df_stops['data'].dt.date >= start_date) & (df_stops['data'].dt.date <= end_date)]

        if maquinas:
            df_o_f = df_o_f[df_o_f['maquina'].isin(maquinas)]
            df_s_f = df_s_f[df_s_f['maquina'].isin(maquinas)]
        if turnos:
            df_o_f = df_o_f[df_o_f['turno'].isin(turnos)]
            df_s_f = df_s_f[df_s_f['turno'].isin(turnos)]

        # --- CÁLCULOS (Usando nomes normalizados) ---
        counter_total = df_o_f['machine counter'].sum()
        estoque_total = df_o_f['pecas estoque'].sum()
        runtime_total = df_o_f['run time'].sum()
        std_time_total = df_o_f['horario padrao'].sum()

        movimentacao = (runtime_total / std_time_total * 100) if std_time_total > 0 else 0
        loss = ((counter_total - estoque_total) / counter_total * 100) if counter_total > 0 else 0

        df_t_sum = df_o_f.groupby(['data', 'turno', 'maquina'])['pecas estoque'].sum().reset_index()
        media_p_turno = df_t_sum['pecas estoque'].mean() if not df_t_sum.empty else 0

        # --- EXIBIÇÃO ---
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Machine Counter", f"{counter_total:,.0f}")
        k2.metric("Horário Padrão (Min)", f"{std_time_total:,.0f}")
        k3.metric("Peças Estoque", f"{estoque_total:,.0f}")
        k4.metric("Média Peças/Turno", f"{media_p_turno:,.0f}")

        col_m, col_l = st.columns(2)

        def draw_gauge(label, value, color, target):
            return go.Figure(go.Indicator(
                mode="gauge+number", value=value,
                title={'text': label},
                gauge={'axis': {'range': [None, 100]},
                       'bar': {'color': color},
                       'threshold': {'line': {'color': "black", 'width': 4}, 'value': target}}
            )).update_layout(height=280)

        col_m.plotly_chart(draw_gauge("Movimentação (%)", movimentacao, "#2ecc71", 85), use_container_width=True)
        col_l.plotly_chart(draw_gauge("Loss (%)", loss, "#e74c3c", 5), use_container_width=True)

        st.subheader("⚠️ Top 10 Paradas")
        if not df_s_f.empty:
            df_p = df_s_f.groupby('problema').agg(minutos=('minutos', 'sum'), qtd=('qtd parada', 'sum')).reset_index()
            df_p = df_p.sort_values('minutos', ascending=False).head(10)
            fig_p = px.bar(df_p, x='minutos', y='problema', orientation='h', color='minutos', color_continuous_scale='Reds')
            st.plotly_chart(fig_p, use_container_width=True)

        st.subheader("🏆 Ranking Máquinas")
        df_rank = df_o_f.groupby('maquina').agg(
            mov=('run time', lambda x: (x.sum() / df_o_f.loc[x.index, 'horario padrao'].sum() * 100)),
            ls=('machine counter', lambda x: ((x.sum() - df_o_f.loc[x.index, 'pecas estoque'].sum()) / x.sum() * 100))
        ).reset_index()
        
        med_m = df_t_sum.groupby('maquina')['pecas estoque'].mean().reset_index()
        df_rank = df_rank.merge(med_m, on='maquina')
        df_rank.columns = ['Máquina', 'Movimentação (%)', 'Loss (%)', 'Média Peças/Turno']
        
        st.dataframe(df_rank.sort_values('Movimentação (%)', ascending=False).style.format({
            'Movimentação (%)': '{:.2f}%', 'Loss (%)': '{:.2f}%', 'Média Peças/Turno': '{:,.0f}'
        }).background_gradient(subset=['Movimentação (%)'], cmap='RdYlGn'), use_container_width=True)

    except Exception as e:
        st.error(f"Erro detectado: {e}. Verifique se as colunas 'Máquina', 'Data', 'Turno', etc., existem no arquivo.")

else:
    st.info("Aguardando upload do arquivo Excel...")
