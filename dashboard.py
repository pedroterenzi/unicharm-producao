import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM (Botões Uniformes, Sem Bolinhas e Cards de Status) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Remove bolinhas do radio e cria botões uniformes */
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

    /* Cards de Métricas Estilo Status Dashboard */
    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 15px; }
    .status-card {
        background: #1e293b; padding: 15px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1); flex: 1;
    }
    .status-title { color: #94a3b8; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; }
    .status-value { color: #ffffff; font-size: 1.6rem; font-weight: 900; }

    /* Estilo 5 Porquês */
    .five-why-box { border: 1px solid #000; border-radius: 5px; padding: 15px; background: #ffffff; color: #000; margin-top: 10px; }
    .five-why-line { border-bottom: 1px solid #000; padding: 8px 0; font-size: 0.9rem; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNÇÕES DE TRATAMENTO DE DADOS
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
    target_sheet = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
    if not target_sheet: return None
    
    # Lê a planilha bruta. Linha 3 (índice 2) contém as datas
    df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)
    row_dates = df_raw.iloc[2, :].tolist() # Linha das datas
    
    planning_list = []
    # Dados começam na linha 4 (índice 3)
    # Coluna A=Máquina, B=Código
    for _, row in df_raw.iloc[3:].iterrows():
        maquina = str(row[0]).strip()
        codigo = str(row[1]).strip()
        if maquina == 'nan' or codigo == 'nan' or maquina == '0': continue
        
        for col_idx, date_val in enumerate(row_dates):
            if isinstance(date_val, (datetime, pd.Timestamp)):
                prog_val = pd.to_numeric(row[col_idx], errors='coerce') or 0
                if prog_val > 0:
                    planning_list.append({
                        'Data': date_val.date(),
                        'Máquina': maquina.replace('MQ ', '').replace('MQ', ''),
                        'Código': codigo,
                        'Programado': prog_val
                    })
    return pd.DataFrame(planning_list)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    st.subheader("📁 Upload de Arquivos")
    up_prod = st.file_uploader("1. Relatório de Produção (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("2. Programação (DATAS)", type=["xlsx"])
    
    st.markdown("---")
    if up_prod:
        menu = st.radio("NAVEGAÇÃO", 
                        ["📈 PERFORMANCE", "📊 PROGRAMADO X REAL", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"],
                        label_visibility="collapsed")

# --- PROCESSAMENTO PRINCIPAL ---
if up_prod:
    df_order, df_stops = load_production_data(up_prod)

    # =========================================================
    # VISÃO: PROGRAMADO X REALIZADO (ESTILO SEU DASHBOARD)
    # =========================================================
    if menu == "📊 PROGRAMADO X REAL":
        st.markdown("## 📊 Acompanhamento de Produção (Aderência ao Plano)")
        
        if not up_datas:
            st.warning("⚠️ Suba o arquivo de Programação (DATAS) no menu lateral.")
        else:
            df_plan = load_planner_data(up_datas)
            if df_plan is not None and not df_plan.empty:
                data_sel = st.date_input("Filtrar Data de Referência", df_order['Data'].max().date())
                
                # Filtrar Realizado e Programado
                real = df_order[df_order['Data'].dt.date == data_sel].groupby(['Máquina', 'Código'])['Peças Estoque - Ajuste'].sum().reset_index()
                real.columns = ['Máquina', 'Código', 'Realizado']
                prog = df_plan[df_plan['Data'] == data_sel]
                
                # Cruzamento
                df_comp = pd.merge(prog, real, on=['Máquina', 'Código'], how='outer').fillna(0)
                df_comp['Aderência %'] = (df_comp['Realizado'] / df_comp['Programado'] * 100).replace([float('inf')], 0).fillna(0)
                
                # KPIs Topo
                t_p, t_r = df_comp['Programado'].sum(), df_comp['Realizado'].sum()
                ad_g = (t_r / t_p * 100) if t_p > 0 else 0
                
                st.markdown(f"""
                    <div class="metric-container">
                        <div class="status-card"><div class="status-title">Total Programado</div><div class="status-value">{t_p:,.0f}</div></div>
                        <div class="status-card"><div class="status-title">Total Realizado</div><div class="status-value">{t_r:,.0f}</div></div>
                        <div class="status-card"><div class="status-title">Aderência Geral</div><div class="status-value">{ad_g:.1f}%</div></div>
                    </div>
                """, unsafe_allow_html=True)
                
                # Status por Máquina
                st.markdown("### 🖥️ Status por Máquina")
                maq_res = df_comp.groupby('Máquina').agg({'Programado':'sum', 'Realizado':'sum'}).reset_index()
                maq_res['Status %'] = (maq_res['Realizado'] / maq_res['Programado'] * 100).fillna(0)
                
                cols = st.columns(len(maq_res) if len(maq_res) > 0 else 1)
                for i, row in maq_res.iterrows():
                    cor = "#10b981" if row['Status %'] >= 90 else "#f59e0b" if row['Status %'] >= 50 else "#f43f5e"
                    cols[i].markdown(f"""<div style="background:#1e293b; padding:10px; border-radius:8px; border-left:5px solid {cor};">
                        <small>MÁQ {row['Máquina']}</small><br><b>{row['Status %']:.1f}%</b></div>""", unsafe_allow_html=True)
                
                st.markdown("<br>### 📋 Detalhamento SKU", unsafe_allow_html=True)
                st.dataframe(df_comp.sort_values('Máquina').style.format({'Programado': '{:,.0f}', 'Realizado': '{:,.0f}', 'Aderência %': '{:.1f}%'})
                             .background_gradient(subset=['Aderência %'], cmap='RdYlGn', vmin=0, vmax=100), use_container_width=True)
            else:
                st.error("Aba 'CALENDARIO MARÇO PEÇAS' não detectada no arquivo DATAS.")

    # =========================================================
    # VISÃO: PERFORMANCE GERAL
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()])
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq))]
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Machine Counter", f"{df_f['Machine Counter'].sum():,.0f}")
        c2.metric("Peças Estoque", f"{df_f['Peças Estoque - Ajuste'].sum():,.0f}")
        c3.metric("Run Time Total", f"{df_f['Run Time'].sum():,.0f}m")

        col_g1, col_g2 = st.columns(2)
        mov = (df_f['Run Time'].sum() / df_f['Horário Padrão'].sum() * 100) if df_f['Horário Padrão'].sum() > 0 else 0
        with col_g1:
            fig1 = go.Figure(go.Indicator(mode="gauge+number", value=mov, title={'text': "Movimentação %"}, gauge={'bar':{'color':"#10b981"}}))
            fig1.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig1, use_container_width=True)
        with col_g2:
            loss = ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100) if df_f['Machine Counter'].sum() > 0 else 0
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=loss, title={'text': "Loss %"}, gauge={'bar':{'color':"#f43f5e"}}))
            fig2.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig2, use_container_width=True)

    # =========================================================
    # VISÃO: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_d_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()])
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_d_s[0]) & (df_stops['Data'].dt.date <= f_d_s[1])]
        
        st.markdown("## 🛑 Top 10 Paradas por Minutos")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10), orientation='h', color_discrete_sequence=['#f43f5e']), use_container_width=True)

    # =========================================================
    # VISÃO: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        df_order['Mes'] = df_order['Data'].dt.month
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        m_idx = list(calendar.month_name).index(mes_sel) + 1
        
        df_c = df_order[df_order['Mes'] == m_idx]
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        
        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, m_idx))
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#64748b; font-weight:900; font-size:0.7rem;">{n}</div>'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_v = row['Run Time'].values[0]/row['Horário Padrão'].values[0]*100 if not row.empty and row['Horário Padrão'].values[0]>0 else 0
                cor = "#059669" if m_v > 85 else "#dc2626" if m_v > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">MOV: {m_v:.1f}%</div></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

    # =========================================================
    # VISÃO: ANÁLISE SEMANAL (5 PORQUÊS)
    # =========================================================
    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Configuração")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        df_sb = df_stops[df_stops['Máquina'] == maq_b].groupby('Problema')['Minutos'].sum().sort_values(ascending=False).head(1)
        pior_p = df_sb.index[0] if not df_sb.empty else "---"

        st.markdown(f"## 📋 Relatório Semanal - MÁQUINA {maq_b}")
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

else:
    st.info("💡 Carregue os arquivos Excel para iniciar a análise industrial.")
