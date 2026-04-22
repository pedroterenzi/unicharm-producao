import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Upload & Analyze - Performance Industrial", layout="wide")

# Estilização customizada para os velocímetros
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# 2. INTERFACE DE UPLOAD
st.title("📊 Dashboard Operacional")
st.subheader("Suba seu arquivo Excel para gerar as visões")

uploaded_file = st.file_uploader("Arraste ou selecione o arquivo 'Relatório das paradas Operacionais(Oficial).xlsm'", type=["xlsm"])

if uploaded_file is not None:
    # Função para carregar os dados do arquivo subido
    @st.cache_data
    def load_uploaded_data(file):
        # Lendo as abas específicas
        df_order = pd.read_excel(file, sheet_name="Result by order")
        df_stops = pd.read_excel(file, sheet_name="Stop Machine Item")
        
        # Conversão de datas
        df_order['Data'] = pd.to_datetime(df_order['Data'])
        df_stops['Data'] = pd.to_datetime(df_stops['Data'])
        
        return df_order, df_stops

    try:
        df_order, df_stops = load_uploaded_data(uploaded_file)
        st.success("Arquivo carregado com sucesso!")
        
        # --- SIDEBAR - FILTROS ---
        st.sidebar.header("Filtros de Análise")
        
        # Filtro de Data
        min_date = df_order['Data'].min().date()
        max_date = df_order['Data'].max().date()
        
        dr = st.sidebar.date_input("Período", [min_date, max_date])
        # Verificação para evitar erro caso o usuário limpe a data
        start_date = dr[0] if len(dr) > 0 else min_date
        end_date = dr[1] if len(dr) > 1 else max_date

        maquinas = st.sidebar.multiselect("Filtrar Máquinas", options=sorted(df_order['Máquina'].unique()))
        turnos = st.sidebar.multiselect("Filtrar Turnos", options=sorted(df_order['Turno'].unique()))

        # Aplicando Filtros
        df_o_f = df_order[(df_order['Data'].dt.date >= start_date) & (df_order['Data'].dt.date <= end_date)]
        df_s_f = df_stops[(df_stops['Data'].dt.date >= start_date) & (df_stops['Data'].dt.date <= end_date)]

        if maquinas:
            df_o_f = df_o_f[df_o_f['Máquina'].isin(maquinas)]
            df_s_f = df_s_f[df_s_f['Máquina'].isin(maquinas)]
        if turnos:
            df_o_f = df_o_f[df_o_f['Turno'].isin(turnos)]
            df_s_f = df_s_f[df_s_f['Turno'].isin(turnos)]

        # --- CÁLCULOS ---
        counter_total = df_o_f['Machine Counter'].sum()
        estoque_total = df_o_f['Peças Estoque'].sum()
        runtime_total = df_o_f['Run Time'].sum()
        std_time_total = df_o_f['Horário Padrão'].sum()

        perc_mov = (runtime_total / std_time_total * 100) if std_time_total > 0 else 0
        perc_loss = ((counter_total - estoque_total) / counter_total * 100) if counter_total > 0 else 0

        # Média por Turno Real (Agrupado)
        df_t_sum = df_o_f.groupby(['Data', 'Turno', 'Máquina'])['Peças Estoque'].sum().reset_index()
        media_p_turno = df_t_sum['Peças Estoque'].mean() if not df_t_sum.empty else 0

        # --- EXIBIÇÃO ---
        
        # KPIs em destaque
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Machine Counter", f"{counter_total:,.0f}")
        c2.metric("Horário Padrão", f"{std_time_total:,.0f}")
        c3.metric("Enviado Estoque", f"{estoque_total:,.0f}")
        c4.metric("Média Peças/Turno", f"{media_p_turno:,.0f}")

        st.markdown("---")

        # Velocímetros
        col_m, col_l = st.columns(2)

        def gauge_plot(label, value, color, threshold):
            return go.Figure(go.Indicator(
                mode="gauge+number", value=value,
                title={'text': label, 'font': {'size': 24}},
                gauge={'axis': {'range': [None, 100]},
                       'bar': {'color': color},
                       'threshold': {'line': {'color': "black", 'width': 4}, 'value': threshold}}
            )).update_layout(height=280, margin=dict(l=30, r=30, t=50, b=20))

        col_m.plotly_chart(gauge_plot("Movimentação (%)", perc_mov, "royalblue", 85), use_container_width=True)
        col_l.plotly_chart(gauge_plot("Loss (%)", perc_loss, "crimson", 5), use_container_width=True)

        # Análise de Paradas
        st.subheader("⚠️ Ranking de Paradas (Influência no Período)")
        if not df_s_f.empty:
            df_p = df_s_f.groupby('Problema').agg(Minutos=('Minutos', 'sum'), Qtd=('QTD Parada', 'sum')).reset_index()
            df_p = df_p.sort_values('Minutos', ascending=False).head(10)
            total_m = df_s_f['Minutos'].sum()
            df_p['% Influência'] = (df_p['Minutos'] / total_m * 100).round(1)

            fig_p = px.bar(df_p, x='Minutos', y='Problema', orientation='h', text='% Influência',
                           color='Minutos', color_continuous_scale='Reds')
            fig_p.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_p, use_container_width=True)
        else:
            st.info("Sem dados de paradas para os filtros aplicados.")

        # Ranking de Máquinas
        st.subheader("🏆 Performance por Máquina")
        df_rank = df_o_f.groupby('Máquina').agg(
            Movimentacao= ('Run Time', lambda x: (x.sum() / df_o_f.loc[x.index, 'Horário Padrão'].sum() * 100)),
            Loss= ('Machine Counter', lambda x: ((x.sum() - df_o_f.loc[x.index, 'Peças Estoque'].sum()) / x.sum() * 100))
        ).reset_index()
        
        # Mesclar com a média por turno calculada separadamente para precisão
        med_m = df_t_sum.groupby('Máquina')['Peças Estoque'].mean().reset_index()
        df_rank = df_rank.merge(med_m, on='Máquina')
        df_rank.columns = ['Máquina', 'Movimentação (%)', 'Loss (%)', 'Média Peças/Turno']
        
        st.dataframe(df_rank.sort_values('Movimentação (%)', ascending=False).style.format({
            'Movimentação (%)': '{:.2f}%', 'Loss (%)': '{:.2f}%', 'Média Peças/Turno': '{:,.0f}'
        }).background_gradient(subset=['Movimentação (%)'], cmap='RdYlGn'), use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar as abas do Excel. Verifique se os nomes das abas estão corretos. Erro: {e}")

else:
    st.info("Aguardando upload do arquivo para processar os indicadores.")
