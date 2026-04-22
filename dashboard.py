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
    
    /* Menu Lateral Botões Uniformes */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label > div:first-child { display: none !important; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 10px 15px !important; border-radius: 8px !important;
        margin-bottom: 5px !important; color: white !important;
        transition: all 0.3s ease; cursor: pointer; width: 100%;
        display: block !important; text-align: center; font-weight: 600;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
        border: none !important; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
    }

    /* Cartões Padronizados */
    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 15px; }
    .metric-card {
        background: #1e293b; padding: 15px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 90px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { color: #10b981; font-size: 1.6rem; font-weight: 900; line-height: 1; }

    /* Calendário */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .day-card { background: #0f172a; border-radius: 8px; padding: 10px; min-height: 80px; border: 1px solid rgba(255,255,255,0.05); }

    /* 5 Porquês (Impressão) */
    .five-why-box { border: 1px solid #000; padding: 15px; background: #fff; color: #000; margin-top: 15px; }
    .line-space { border-bottom: 1px solid #000; margin-bottom: 10px; height: 20px; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
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
    df_order = df_order.dropna(subset=['Data'])
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    
    def categorize(m): return "BABY" if m in ['2', '3', '4', '5', '6'] else "ADULTO"
    df_order['Categoria'] = df_order['Máquina'].apply(categorize)
    return df_order, df_stops

@st.cache_data
def load_planner_meta_row125(file, data_ref):
    try:
        xls = pd.ExcelFile(file)
        target = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        if not target: return 0
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        row_dates = df_raw.iloc[2, :].tolist()
        row_meta = df_raw.iloc[124, :].tolist() # Linha 125 é index 124
        
        meta_ate_hoje = 0
        for col_idx, d_val in enumerate(row_dates):
            if isinstance(d_val, (datetime, pd.Timestamp)):
                if d_val.date() <= data_ref:
                    meta_ate_hoje += pd.to_numeric(row_meta[col_idx], errors='coerce') or 0
        return meta_ate_hoje
    except: return 0

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    up_prod = st.file_uploader("1. Produção (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("2. DATAS (.xlsx)", type=["xlsx"])
    
    if up_prod:
        menu = st.radio("NAVEGAÇÃO", ["📋 REPORTE DIÁRIO", "📈 PERFORMANCE GERAL", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 SEMANAL"], label_visibility="collapsed")

if up_prod:
    df_order, df_stops = load_production_data(up_prod)

    # =========================================================
    # VISÃO: REPORTE DIÁRIO
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        data_atual = df_order['Data'].max().date()
        st.markdown(f"## 📋 Reporte Diário - Ref: {data_atual.strftime('%d/%m/%Y')}")
        
        # --- DESTAQUE TOPO ---
        df_mes = df_order[df_order['Data'].dt.month == data_atual.month]
        total_p_real = df_mes['Peças Estoque - Ajuste'].sum()
        total_mov = (df_mes['Run Time'].sum() / df_mes['Horário Padrão'].sum() * 100) if df_mes['Horário Padrão'].sum() > 0 else 0
        total_perda = ((df_mes['Machine Counter'].sum() - total_p_real) / df_mes['Machine Counter'].sum() * 100) if df_mes['Machine Counter'].sum() > 0 else 0
        meta_estoque = load_planner_meta_row125(up_datas, data_atual) if up_datas else 0

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Movimentação Mês</div><div class="metric-value">{total_mov:.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Perda Mês</div><div class="metric-value" style="color:#f43f5e">{total_perda:.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Estoque Total Mês</div><div class="metric-value">{total_p_real:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Meta Estoque (DATAS)</div><div class="metric-value" style="color:#3b82f6">{meta_estoque:,.0f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        # Filtro de Últimos 3 dias
        datas_disponiveis = sorted(df_order['Data'].dt.date.unique(), reverse=True)
        dias_selecionados = st.multiselect("Filtrar dias para visualização:", datas_disponiveis, default=datas_disponiveis[:3])

        for dia in dias_selecionados:
            st.subheader(f"📅 Produção {dia.strftime('%d/%m/%Y')}")
            df_dia = df_order[df_order['Data'].dt.date == dia]
            res = df_dia.groupby(['Categoria', 'Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
            res['Mov %'] = (res['Run Time']/res['Horário Padrão']*100).round(1)
            res['Perda %'] = ((res['Machine Counter']-res['Peças Estoque - Ajuste'])/res['Machine Counter']*100).round(1)
            st.dataframe(res[['Categoria','Máquina','Mov %','Perda %','Peças Estoque - Ajuste']].rename(columns={'Peças Estoque - Ajuste':'Pçs Estoque'}), use_container_width=True)

    # =========================================================
    # VISÃO: PERFORMANCE GERAL
    # =========================================================
    elif menu == "📈 PERFORMANCE GERAL":
        st.sidebar.subheader("Filtros")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()])
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f['Machine Counter'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f['Peças Estoque - Ajuste'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f['Run Time'].sum():,.0f}m</div></div>
            </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            val_mov = (df_f['Run Time'].sum()/df_f['Horário Padrão'].sum()*100) if df_f['Horário Padrão'].sum()>0 else 0
            fig1 = go.Figure(go.Indicator(mode="gauge+number", value=val_mov, title={'text':"Movimentação %"}, gauge={'bar':{'color':"#10b981"}, 'axis':{'range':[0,100]}, 'threshold':{'line':{'color':"white",'width':4},'value':90}}))
            st.plotly_chart(fig1, use_container_width=True)
        with col2:
            val_loss = ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100) if df_f['Machine Counter'].sum()>0 else 0
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=val_loss, title={'text':"Loss %"}, gauge={'bar':{'color':"#e74c3c"}, 'axis':{'range':[0,15]}, 'threshold':{'line':{'color':"white",'width':4},'value':2.5}}))
            st.plotly_chart(fig2, use_container_width=True)

    # =========================================================
    # VISÃO: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        f_data = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()])
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()))
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_stops['Turno'].unique()), default=sorted(df_stops['Turno'].unique()))
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data[0]) & (df_stops['Data'].dt.date <= f_data[1]) & (df_stops['Máquina'].isin(f_maq)) & (df_stops['Turno'].isin(f_turno))]
        
        st.markdown("### 🛑 Paradas por Minutos")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values().tail(10), orientation='h', color_discrete_sequence=['#f43f5e']), use_container_width=True)
        st.markdown("### 🛑 Paradas por Quantidade")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['QTD'].sum().sort_values().tail(10), orientation='h', color_discrete_sequence=['#3b82f6']), use_container_width=True)

    # =========================================================
    # VISÃO: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        m_idx = list(calendar.month_name).index(mes_sel) + 1
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='cal_maq')
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='cal_turno')
        
        df_c = df_order[(df_order['Data'].dt.month == m_idx) & (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        
        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, m_idx))
        html = '<div class="calendar-grid">'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_v = row['Run Time'].values[0]/row['Horário Padrão'].values[0]*100 if not row.empty and row['Horário Padrão'].values[0]>0 else 0
                cor = "#059669" if m_v > 85 else "#dc2626" if m_v > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><span style="color:#94a3b8; font-weight:bold;">{d}</span><br><small>MOV: {m_v:.1f}%</small></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

    # =========================================================
    # VISÃO: ANÁLISE SEMANAL (IMPRESSÃO)
    # =========================================================
    elif menu == "📋 SEMANAL":
        st.sidebar.subheader("Filtros Relatório")
        maq_sel = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_sel = st.sidebar.selectbox("Turno", sorted(df_order['Turno'].unique()))
        periodo = st.sidebar.date_input("Período Semana", [datetime.now()-timedelta(days=7), datetime.now()])

        st.markdown(f"""
            <div style="text-align:center; border:2px solid #10b981; padding:10px; border-radius:10px;">
                <h1 style="color:white; margin:0;">RELATÓRIO SEMANAL DE PERFORMANCE</h1>
                <h2 style="color:#10b981; margin:0;">MÁQUINA {maq_sel} - TURNO {turno_sel}</h2>
                <p style="color:#94a3b8;">Período: {periodo[0].strftime('%d/%m/%Y')} a {periodo[1].strftime('%d/%m/%Y')}</p>
            </div>
        """, unsafe_allow_html=True)

        df_b = df_order[(df_order['Data'].dt.date >= periodo[0]) & (df_order['Data'].dt.date <= periodo[1]) & (df_order['Máquina'] == maq_sel) & (df_order['Turno'] == turno_sel)]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo[0]) & (df_stops['Data'].dt.date <= periodo[1]) & (df_stops['Máquina'] == maq_sel) & (df_stops['Turno'] == turno_sel)]

        c1, c2, c3 = st.columns(3)
        with c1:
            mov_v = (df_b['Run Time'].sum()/df_b['Horário Padrão'].sum()*100) if df_b['Horário Padrão'].sum()>0 else 0
            st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=mov_v, title={'text':"Movimentação %"})).update_layout(height=250), use_container_width=True)
        with c2:
            loss_v = ((df_b['Machine Counter'].sum()-df_b['Peças Estoque - Ajuste'].sum())/df_b['Machine Counter'].sum()*100) if df_b['Machine Counter'].sum()>0 else 0
            st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=loss_v, title={'text':"Loss %"})).update_layout(height=250), use_container_width=True)
        with c3:
            st.markdown(f"""
                <div class="metric-card" style="height:100%"><div class="metric-title">Peças p/ Estoque</div><div class="metric-value">{df_b['Peças Estoque - Ajuste'].sum():,.0f}</div></div>
                <div class="metric-card" style="height:100%; margin-top:5px;"><div class="metric-title">Run Time / HP</div><div class="metric-value" style="font-size:1.1rem;">{df_b['Run Time'].sum():,.0f} / {df_b['Horário Padrão'].sum():,.0f}</div></div>
            """, unsafe_allow_html=True)

        st.subheader("🛑 Piores 5 Paradas da Semana")
        stop_data = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=False).head(5)
        st.plotly_chart(px.bar(stop_data, orientation='v', color_discrete_sequence=['#10b981']), use_container_width=True)
        
        pior_p = stop_data.index[0] if not stop_data.empty else "---"
        st.markdown(f"""
            <div class="five-why-box">
                <h3 style="margin-top:0; color:#059669;">ANÁLISE DE CAUSA RAIZ - 5 PORQUÊS</h3>
                <b>PROBLEMA FOCO:</b> {pior_p}<br><br>
                1. Por que? <div class="line-space"></div>
                2. Por que? <div class="line-space"></div>
                3. Por que? <div class="line-space"></div>
                4. Por que? <div class="line-space"></div>
                5. Por que? <div class="line-space"></div>
                <div style="display:flex; gap:20px;">
                    <div style="flex:1;"><b>CAUSA RAIZ:</b><div class="line-space"></div></div>
                    <div style="flex:1;"><b>PLANO DE AÇÃO:</b><div class="line-space"></div></div>
                </div>
            </div>
        """, unsafe_allow_html=True)

else:
    st.info("💡 Carregue os arquivos para começar.")
