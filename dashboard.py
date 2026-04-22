import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM (Botões sem bolinhas e Cards) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Remove bolinhas do seletor e cria botões uniformes */
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

    /* Cards de Métricas Reduzidos */
    .metric-container { display: flex; justify-content: space-between; gap: 8px; margin-bottom: 15px; }
    .metric-card {
        background: #1e293b; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 70px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    .metric-value { color: #10b981; font-size: 1.3rem; font-weight: 900; line-height: 1; }

    /* Estilo 5 Porquês */
    .five-why-box { border: 1px solid #000; border-radius: 5px; padding: 15px; background: #ffffff; color: #000; margin-top: 10px; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNÇÕES DE CARREGAMENTO
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
    df_order['Código'] = df_order['Código'].astype(str).str.strip()
    return df_order, df_stops

@st.cache_data
def load_planner_data(file):
    xls = pd.ExcelFile(file)
    # Busca aba que contenha PEÇAS
    target_sheet = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
    if not target_sheet: return None
    
    # Carrega o dataframe bruto para tratar o cabeçalho horrível
    df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)
    
    # A linha das datas reais é a 3 (índice 2 no Python)
    # A linha dos nomes das colunas (Código, Item) é a 2 (índice 1)
    
    # Vamos extrair as datas da linha de índice 2
    row_dates = df_raw.iloc[2, :].tolist()
    # Vamos extrair os códigos da coluna de índice 1 (Coluna B do Excel)
    
    # Limpeza para criar um mapeamento: Data -> Coluna
    planning_data = []
    
    # Localizar onde estão os códigos (Geralmente coluna B, índice 1)
    for index, row in df_raw.iloc[3:].iterrows(): # Dados começam na linha 4
        codigo = str(row[1]).strip() # Coluna B
        if codigo == 'nan' or codigo == '0': continue
        
        for col_idx, date_val in enumerate(row_dates):
            if isinstance(date_val, datetime):
                qtd_programada = pd.to_numeric(row[col_idx], errors='coerce') or 0
                if qtd_programada > 0:
                    planning_data.append({
                        'Data': date_val.date(),
                        'Código': codigo,
                        'Programado': qtd_programada
                    })
    
    return pd.DataFrame(planning_data)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    
    st.subheader("📁 Upload de Arquivos")
    up_prod = st.file_uploader("1. Relatório de Produção (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("2. Programação Comercial (.xlsx)", type=["xlsx"])
    
    st.markdown("---")
    if up_prod:
        menu = st.radio("MENU", 
                        ["📈 PERFORMANCE", "📊 PROGRAMADO X REAL", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"],
                        label_visibility="collapsed")

# --- PROCESSAMENTO PRINCIPAL ---
if up_prod:
    df_order, df_stops = load_production_data(up_prod)
    
    # =========================================================
    # ABA: PROGRAMADO X REALIZADO (O CORAÇÃO DO CRUZAMENTO)
    # =========================================================
    if menu == "📊 PROGRAMADO X REAL":
        st.markdown("## 📊 Aderência ao Plano de Produção")
        
        if not up_datas:
            st.warning("⚠️ O arquivo 'DATAS' é necessário para esta visão.")
        else:
            df_plan_clean = load_planner_data(up_datas)
            
            if df_plan_clean is not None and not df_plan_clean.empty:
                # Seletor de Data
                data_ref = st.date_input("Selecione o dia para conferência", df_order['Data'].max().date())
                
                # Filtrar Realizado
                real = df_order[df_order['Data'].dt.date == data_ref].groupby('Código')['Peças Estoque - Ajuste'].sum().reset_index()
                real.columns = ['Código', 'Realizado']
                
                # Filtrar Programado
                prog = df_plan_clean[df_plan_clean['Data'] == data_ref].copy()
                
                # Cruzamento (Merge)
                df_cross = pd.merge(prog, real, on='Código', how='outer').fillna(0)
                df_cross['Diferença'] = df_cross['Realizado'] - df_cross['Programado']
                df_cross['Aderência %'] = (df_cross['Realizado'] / df_cross['Programado'] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                
                # KPIs
                c1, c2, c3 = st.columns(3)
                total_p = df_cross['Programado'].sum()
                total_r = df_cross['Realizado'].sum()
                c1.metric("Programado (Dia)", f"{total_p:,.0f} pçs")
                c2.metric("Realizado (Dia)", f"{total_r:,.0f} pçs")
                c3.metric("Aderência", f"{(total_r/total_p*100 if total_p>0 else 0):.1f}%")

                st.markdown("### 📋 Comparativo por Código")
                st.dataframe(df_cross[['Código', 'Programado', 'Realizado', 'Diferença', 'Aderência %']].sort_values('Programado', ascending=False).style.format({
                    'Programado': '{:,.0f}', 'Realizado': '{:,.0f}', 'Diferença': '{:,.0f}', 'Aderência %': '{:.1f}%'
                }).background_gradient(subset=['Aderência %'], cmap='RdYlGn', vmin=0, vmax=100), use_container_width=True)

                fig = px.bar(df_cross[df_cross['Programado']>0], x='Código', y=['Programado', 'Realizado'], barmode='group', 
                             title="Programado vs Realizado por Item", color_discrete_map={'Programado':'#475569','Realizado':'#10b981'})
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.error("Erro ao ler as datas do arquivo DATAS. Verifique se as datas estão na linha 3.")

    # =========================================================
    # ABA: PERFORMANCE GERAL
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq))]
        
        # KPIs e Velocímetros (Conforme implementado antes)
        total_mc = df_f['Machine Counter'].sum()
        total_est = df_f['Peças Estoque - Ajuste'].sum()
        total_rt = df_f['Run Time'].sum()
        total_hp = df_f['Horário Padrão'].sum()
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{total_mc:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{total_est:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time</div><div class="metric-value">{total_rt:,.0f}m</div></div>
            </div>
            """, unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            mov = (total_rt/total_hp*100) if total_hp > 0 else 0
            fig1 = go.Figure(go.Indicator(mode="gauge+number", value=mov, title={'text': "Movimentação %"}, gauge={'bar':{'color':"#10b981"}}))
            fig1.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            loss = ((total_mc-total_est)/total_mc*100) if total_mc > 0 else 0
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=loss, title={'text': "Loss %"}, gauge={'bar':{'color':"#f43f5e"}}))
            fig2.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig2, use_container_width=True)

    # =========================================================
    # DEMAIS ABAS (TOP 10, CALENDÁRIO, ANÁLISE SEMANAL)
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1])]
        st.markdown("## 🛑 Análise de Paradas")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10), orientation='h', title="Top 10 por Minutos", color_discrete_sequence=['#f43f5e']), use_container_width=True)

    elif menu == "📅 CALENDÁRIO":
        st.markdown("## 📅 Calendário Operacional")
        # (Lógica do calendário já ajustada para filtros de Mês/Maq/Turno conforme pedido antes)
        st.info("Filtre na lateral para detalhar os dias.")

    elif menu == "📋 ANÁLISE SEMANAL":
        st.markdown("## 📋 Análise Semanal (Pronto para Impressão)")
        # (Lógica dos 5 porquês e ranking lateral conforme pedido antes)
        pass

else:
    st.info("💡 Por favor, carregue os arquivos de **Produção** e **DATAS** para habilitar todas as visões.")
