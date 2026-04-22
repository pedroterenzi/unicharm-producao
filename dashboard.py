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
    sheet_names = xls.sheet_names
    
    sheet_order = next((s for s in sheet_names if "result" in s.lower() and "order" in s.lower()), None)
    sheet_stops = next((s for s in sheet_names if "stop" in s.lower() and "item" in s.lower()), None)

    if not sheet_order or not sheet_stops:
        st.error(f"❌ Abas não encontradas! Detectadas: {sheet_names}")
        st.stop()

    df_order = pd.read_excel(file, sheet_name=sheet_order)
    df_stops = pd.read_excel(file, sheet_name=sheet_stops)
    
    # Limpeza de nomes de colunas
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # --- TRATAMENTO DE ERROS DE TIPO (Resolve o erro do 'int' and 'str') ---
    cols_numericas_order = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste']
    for col in cols_numericas_order:
        if col in df_order.columns:
            # Converte para número, e o que for texto vira "NaN" (nulo), depois vira 0
            df_order[col] = pd.to_numeric(df_order[col], errors='coerce').fillna(0)

    cols_numericas_stops = ['Minutos', 'QTD']
    for col in cols_numericas_stops:
        if col in df_stops.columns:
            df_stops[col] = pd.to_numeric(df_stops[col], errors='coerce').fillna(0)
    
    # Conversão de Data com tratamento de erro
    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    
    # Remove linhas onde a data ficou nula (linhas vazias no final do Excel)
    df_order = df_order.dropna(subset=['Data'])
    
    return df_order, df_stops

# 3. INTERFACE
st.title("📊 Dashboard Operacional Unicharm")
uploaded_file = st.file_uploader("Selecione o arquivo .xlsm", type=["xlsm"])

if uploaded_file is not None:
    try:
        df_order, df_stops = load_uploaded_data(uploaded_file)
        
        # --- SIDEBAR FILTROS ---
        st.sidebar.header("⚙️ Filtros")
        min_date = df_order['Data'].min().date()
        max_date = df_order['Data'].max().date()
        
        # Correção para o date_input não quebrar sem seleção
        periodo = st.sidebar.date_input("Período", [min_date, max_date])
        
        if isinstance(periodo, list) or isinstance(periodo, tuple):
            if len(periodo) == 2:
                start_date, end_date = periodo
            else:
                start_date = end_date = periodo[0]
        else:
            start_date = end_date = periodo

        maquinas_lista = sorted(df_order['Máquina'].unique().tolist())
        maquinas = st.sidebar.multiselect("Máquinas", options=maquinas_lista)
        
        turnos_lista = sorted(df_order['Turno'].unique().tolist())
        turnos = st.sidebar.multiselect("Turnos", options=turnos_lista)

        # Aplicar Filtros
        mask_o = (df_order['Data'].dt.date >= start_date) & (df_order['Data'].dt.date <= end_date)
        mask_s = (df_stops['Data'].dt.date >= start_date) & (df_stops['Data'].dt.date <= end_date)
        
        df_o_f = df_order.loc[mask_o]
        df_s_f = df_stops.loc[mask_s]

        if maquinas:
            df_o_f = df_o_f[df_o_f['Máquina'].isin(maquinas)]
            df_s_f = df_s_f[df_s_f['Máquina'].isin(maquinas)]
        if turnos:
            df_o_f = df_o_f[df_o_f['Turno'].isin(turnos)]
            df_s_f = df_s_f[df_s_f['Turno'].isin(turnos)]

        # --- CÁLCULOS ---
        mc_sum = df_o_f['Machine Counter'].sum()
        estoque_sum = df_o_f['Peças Estoque - Ajuste'].sum()
        rt_sum = df_o_f['Run Time'].sum()
        hp_sum = df_o_f['Horário Padrão'].sum()

        mov = (rt_sum / hp_sum * 100) if hp_sum > 0 else 0
        loss = ((mc_sum - estoque_sum) / mc_sum * 100) if mc_sum > 0 else 0

        # Média por Turno
        df_t_agrupado = df_o_f.groupby(['Data', 'Turno', 'Máquina'])['Peças Estoque - Ajuste'].sum().reset_index()
        media_turno = df_t_agrupado['Peças Estoque - Ajuste'].mean() if not df_t_agrupado.empty else 0

        # --- EXIBIÇÃO ---
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Machine Counter", f"{mc_sum:,.0f}")
        k2.metric("Horário Padrão", f"{hp_sum:,.0f}")
        k3.metric("Peças (Ajuste)", f"{estoque_sum:,.0f}")
        k4.metric("Média Peças/Turno", f"{media_turno:,.0f}")

        col_m, col_l = st.columns(2)
        def draw_gauge(label, value, color, target):
            return go.Figure(go.Indicator(
                mode="gauge+number", value=value, title={'text': label},
                gauge={'axis': {'range': [None, 100]}, 'bar': {'color': color},
                       'threshold': {'line': {'color': "black", 'width': 4}, 'value': target}}
            )).update_layout(height=280)

        col_m.plotly_chart(draw_gauge("Movimentação (%)", mov, "#2ecc71", 85), use_container_width=True)
        col_l.plotly_chart(draw_gauge("Loss (%)", loss, "#e74c3c", 5), use_container_width=True)

        st.subheader("⚠️ Top 10 Paradas")
        if not df_s_f.empty:
            df_p = df_s_f.groupby('Problema').agg({'Minutos': 'sum', 'QTD': 'sum'}).reset_index()
            df_p = df_p.sort_values('Minutos', ascending=False).head(10)
            fig_p = px.bar(df_p, x='Minutos', y='Problema', orientation='h', color='Minutos', color_continuous_scale='Reds')
            st.plotly_chart(fig_p, use_container_width=True)
        
        st.subheader("🏆 Ranking de Máquinas")
        if not df_o_f.empty:
            df_rank = df_o_f.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
            df_rank['Movimentação (%)'] = (df_rank['Run Time'] / df_rank['Horário Padrão'] * 100)
            df_rank['Loss (%)'] = ((df_rank['Machine Counter'] - df_rank['Peças Estoque - Ajuste']) / df_rank['Machine Counter'] * 100)
            
            med_rank = df_t_agrupado.groupby('Máquina')['Peças Estoque - Ajuste'].mean().reset_index()
            df_rank = df_rank.merge(med_rank, on='Máquina')
            df_rank = df_rank[['Máquina', 'Movimentação (%)', 'Loss (%)', 'Peças Estoque - Ajuste_y']].rename(columns={'Peças Estoque - Ajuste_y': 'Média Peças/Turno'})

            st.dataframe(df_rank.sort_values('Movimentação (%)', ascending=False).style.format({
                'Movimentação (%)': '{:.2f}%', 'Loss (%)': '{:.2f}%', 'Média Peças/Turno': '{:,.0f}'
            }), use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
else:
    st.info("Aguardando upload do arquivo Excel.")
