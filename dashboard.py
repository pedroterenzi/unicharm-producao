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

    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 15px; }
    .metric-card {
        background: #1e293b; padding: 15px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 90px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { color: #10b981; font-size: 1.6rem; font-weight: 900; line-height: 1; }

    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .day-card { background: #0f172a; border-radius: 8px; padding: 10px; min-height: 80px; border: 1px solid rgba(255,255,255,0.05); }

    .five-why-box { border: 1px solid #000; padding: 15px; background: #fff; color: #000; margin-top: 15px; }
    .line-space { border-bottom: 1px solid #000; margin-bottom: 10px; height: 20px; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_production_data(file):
    # Produção (Result by order)
    df_order = pd.read_excel(file, sheet_name="Result by order")
    # Paradas (Stop machine item)
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Conversão Numérica
    nums_order = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed']
    for col in nums_order:
        if col in df_order.columns:
            df_order[col] = pd.to_numeric(df_order[col], errors='coerce').fillna(0)
            
    nums_stops = ['Minutos', 'QTD']
    for col in nums_stops:
        if col in df_stops.columns:
            df_stops[col] = pd.to_numeric(df_stops[col], errors='coerce').fillna(0)
    
    # Datas e Categorias
    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_order = df_order.dropna(subset=['Data'])
    df_stops = df_stops.dropna(subset=['Data'])
    
    # Padronização de strings para evitar erros no multiselect
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    df_stops['Turno'] = df_stops['Turno'].fillna(0).astype(int).astype(str)
    
    def categorize(m): return "BABY" if m in ['2', '3', '4', '5', '6'] else "ADULTO"
    df_order['Categoria'] = df_order['Máquina'].apply(categorize)
    return df_order, df_stops

@st.cache_data
def load_metas_datas(file, data_ref):
    try:
        xls = pd.ExcelFile(file)
        target = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        if not target: return 0, 0
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        
        row_dates = df_raw.iloc[2, :].tolist()
        row_meta = df_raw.iloc[124, :].tolist() # Linha 125
        
        meta_geral_mes = 0
        meta_dinamica_hoje = 0
        
        for col_idx, d_val in enumerate(row_dates):
            if isinstance(d_val, (datetime, pd.Timestamp)):
                valor = pd.to_numeric(row_meta[col_idx], errors='coerce') or 0
                meta_geral_mes += valor
                if d_val.date() <= data_ref:
                    meta_dinamica_hoje += valor
        return meta_geral_mes, meta_dinamica_hoje
    except: return 0, 0

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
    # REPORTE DIÁRIO
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        data_referencia = df_order['Data'].max().date()
        st.markdown(f"## 📋 Reporte Diário - Ref: {data_referencia.strftime('%d/%m/%Y')}")
        
        df_mes = df_order[df_order['Data'].dt.month == data_referencia.month]
        estoque_total_mes = df_mes['Peças Estoque - Ajuste'].sum()
        mov_mes = (df_mes['Run Time'].sum() / df_mes['Horário Padrão'].sum() * 100) if df_mes['Horário Padrão'].sum() > 0 else 0
        perda_mes = ((df_mes['Machine Counter'].sum() - estoque_total_mes) / df_mes['Machine Counter'].sum() * 100) if df_mes['Machine Counter'].sum() > 0 else 0
        
        meta_mes, meta_dinamica = load_metas_datas(up_datas, data_referencia) if up_datas else (0, 0)

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Movimentação Mês</div><div class="metric-value">{mov_mes:.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Perda Mês</div><div class="metric-value" style="color:#f43f5e">{perda_mes:.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Estoque Total Mês</div><div class="metric-value">{estoque_total_mes:,.0f}</div></div>
            </div>
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Meta Dinâmica (Até Hoje)</div><div class="metric-value" style="color:#3b82f6">{meta_dinamica:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Meta Geral Mês</div><div class="metric-value" style="color:#10b981">{meta_mes:,.0f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        datas_disp = sorted(df_order['Data'].dt.date.unique(), reverse=True)
        dias_sel = st.multiselect("Seletor de Histórico (Últimos dias):", datas_disp, default=datas_disp[:3])

        for dia in dias_sel:
            st.subheader(f"📅 Produção {dia.strftime('%d/%m/%Y')}")
            df_dia = df_order[df_order['Data'].dt.date == dia]
            res = df_dia.groupby(['Categoria', 'Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
            res['Mov %'] = (res['Run Time']/res['Horário Padrão']*100).round(1)
            res['Perda %'] = ((res['Machine Counter']-res['Peças Estoque - Ajuste'])/res['Machine Counter']*100).round(1)
            st.table(res[['Categoria','Máquina','Mov %','Perda %','Peças Estoque - Ajuste']])

    # =========================================================
    # PERFORMANCE GERAL
    # =========================================================
    elif menu == "📈 PERFORMANCE GERAL":
        st.sidebar.subheader("Filtros")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()])
        
        # Correção Multiselect: Garantir strings únicas e ordenadas
        maq_opts = sorted(df_order['Máquina'].unique())
        f_maq = st.sidebar.multiselect("Máquinas", maq_opts, default=maq_opts)
        
        tur_opts = sorted(df_order['Turno'].unique())
        f_turno = st.sidebar.multiselect("Turnos", tur_opts, default=tur_opts)
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & 
                        (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f['Machine Counter'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f['Peças Estoque - Ajuste'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f['Run Time'].sum():,.0f}m</div></div>
            </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        hp = df_f['Horário Padrão'].sum()
        with col1:
            val_m = (df_f['Run Time'].sum()/hp*100) if hp>0 else 0
            st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=val_m, title={'text':"Movimentação %"}, 
                gauge={'bar':{'color':"#10b981"}, 'axis':{'range':[0,100]}, 'threshold':{'line':{'color':"white",'width':4},'value':90}})).update_layout(height=350), use_container_width=True)
        with col2:
            mc = df_f['Machine Counter'].sum()
            val_l = ((mc - df_f['Peças Estoque - Ajuste'].sum())/mc*100) if mc>0 else 0
            st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=val_l, title={'text':"Loss %"}, 
                gauge={'bar':{'color':"#e74c3c"}, 'axis':{'range':[0,15]}, 'threshold':{'line':{'color':"white",'width':4},'value':2.5}})).update_layout(height=350), use_container_width=True)

    # =========================================================
    # TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros")
        f_data = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()])
        maq_opts = sorted(df_stops['Máquina'].unique())
        f_maq = st.sidebar.multiselect("Máquinas", maq_opts, default=maq_opts)
        tur_opts = sorted(df_stops['Turno'].unique())
        f_turno = st.sidebar.multiselect("Turnos", tur_opts, default=tur_opts)

        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data[0]) & (df_stops['Data'].dt.date <= f_data[1]) & 
                          (df_stops['Máquina'].isin(f_maq)) & (df_stops['Turno'].isin(f_turno))]
        
        st.markdown("### ⏱️ Paradas por Minutos")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values().tail(10), orientation='h', color_discrete_sequence=['#f43f5e']), use_container_width=True)
        st.markdown("### 🔢 Paradas por Quantidade")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['QTD'].sum().sort_values().tail(10), orientation='h', color_discrete_sequence=['#3b82f6']), use_container_width=True)

    # =========================================================
    # CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        m_idx = list(calendar.month_name).index(mes_sel) + 1
        maq_opts = sorted(df_order['Máquina'].unique())
        f_maq = st.sidebar.multiselect("Máquinas", maq_opts, default=maq_opts, key='cal_m')
        
        df_c = df_order[(df_order['Data'].dt.month == m_idx) & (df_order['Máquina'].isin(f_maq))]
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
    # SEMANAL (IMPRESSÃO)
    # =========================================================
    elif menu == "📋 SEMANAL":
        st.sidebar.subheader("Filtros")
        maq_sel = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_sel = st.sidebar.selectbox("Turno", sorted(df_order['Turno'].unique()))
        per = st.sidebar.date_input("Semana", [datetime.now()-timedelta(days=7), datetime.now()])
        
        df_b = df_order[(df_order['Data'].dt.date >= per[0]) & (df_order['Data'].dt.date <= per[1]) & (df_order['Máquina'] == maq_sel) & (df_order['Turno'] == turno_sel)]
        df_sb = df_stops[(df_stops['Data'].dt.date >= per[0]) & (df_stops['Data'].dt.date <= per[1]) & (df_stops['Máquina'] == maq_sel) & (df_stops['Turno'] == turno_sel)]

        st.markdown(f"""
            <div style="text-align:center; border:2px solid #10b981; padding:10px; border-radius:10px;">
                <h2 style="color:white; margin:0;">RELATÓRIO SEMANAL DE PERFORMANCE</h2>
                <h3 style="color:#10b981; margin:0;">MÁQUINA {maq_sel} - TURNO {turno_sel}</h3>
                <p style="color:#94a3b8;">Período: {per[0].strftime('%d/%m')} a {per[1].strftime('%d/%m/%Y')}</p>
            </div>
        """, unsafe_allow_html=True)

        v1, v2 = st.columns(2)
        with v1: st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=(df_b['Run Time'].sum()/df_b['Horário Padrão'].sum()*100 if df_b['Horário Padrão'].sum()>0 else 0), title={'text':"Movimentação %"})).update_layout(height=250), use_container_width=True)
        with v2: st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=((df_b['Machine Counter'].sum()-df_b['Peças Estoque - Ajuste'].sum())/df_b['Machine Counter'].sum()*100 if df_b['Machine Counter'].sum()>0 else 0), title={'text':"Loss %"})).update_layout(height=250), use_container_width=True)

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Peças Enviadas</div><div class="metric-value">{df_b['Peças Estoque - Ajuste'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Horário Padrão</div><div class="metric-value">{df_b['Horário Padrão'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time</div><div class="metric-value">{df_b['Run Time'].sum():,.0f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        st.subheader("🛑 Piores 5 Paradas")
        stop_data = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=False).head(5)
        st.plotly_chart(px.bar(stop_data, color_discrete_sequence=['#10b981']), use_container_width=True)
        
        pior = stop_data.index[0] if not stop_data.empty else "---"
        st.markdown(f"""
            <div class="five-why-box">
                <h3 style="color:#059669;">ANÁLISE 5 PORQUÊS: {pior}</h3>
                1. Por que? <div class="line-space"></div> 2. Por que? <div class="line-space"></div>
                3. Por que? <div class="line-space"></div> 4. Por que? <div class="line-space"></div>
                5. Por que? <div class="line-space"></div>
                <b>CAUSA RAIZ:</b> <div class="line-space"></div> <b>PLANO DE AÇÃO:</b> <div class="line-space"></div>
            </div>
        """, unsafe_allow_html=True)
else:
    st.info("💡 Carregue os arquivos para começar.")
