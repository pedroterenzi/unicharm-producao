import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM (MANTIDA INTEGRALMENTE) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Menu Lateral */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 10px 15px !important; border-radius: 8px !important;
        margin-bottom: 5px !important; color: white !important; cursor: pointer;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
    }

    /* Cards de Métricas Padronizados */
    .metric-container { display: flex; justify-content: space-between; gap: 8px; margin-bottom: 15px; }
    .metric-card {
        background: #1e293b; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 70px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    .metric-value { color: #10b981; font-size: 1.1rem; font-weight: 900; line-height: 1; }

    /* Calendário Estilizado */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .calendar-day-name { text-align: center; font-weight: 900; color: #10b981; font-size: 0.8rem; padding-bottom: 5px; }
    .day-card { 
        background: #0f172a; border-radius: 8px; padding: 10px; 
        min-height: 95px; border: 1px solid rgba(255,255,255,0.05); 
    }
    .day-number { font-size: 1rem; font-weight: 900; color: #f8fafc; }
    .day-status { font-size: 0.75rem; font-weight: 600; color: #ffffff; text-align: right; }

    /* Ranking e Feedback */
    .highlight-rank { background: #064e3b !important; color: #10b981 !important; font-weight: 900; border-radius: 5px; padding: 5px; }
    .feedback-box { padding: 12px; border-radius: 8px; margin-bottom: 10px; text-align: center; font-weight: 700; font-size: 0.9rem; }

    /* Área 5 Porquês (Clara para caneta) */
    .five-why-box {
        border: 1px solid #000; border-radius: 5px; padding: 15px; 
        background: #ffffff; color: #000; margin-top: 10px;
    }
    .five-why-line { border-bottom: 1px dotted #000; padding: 10px 0; font-size: 0.9rem; }

    .section-header {
        background: #1e293b; padding: 10px; border-radius: 5px;
        color: #10b981; font-weight: 800; text-transform: uppercase;
        margin-top: 20px; border-left: 5px solid #10b981; font-size: 0.9rem;
    }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file_obj):
    df_order = pd.read_excel(file_obj, sheet_name="Result by order")
    df_stops = pd.read_excel(file_obj, sheet_name="Stop machine item")
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
    
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_stops = df_stops.dropna(subset=['Data'])
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    df_stops['Turno'] = df_stops['Turno'].fillna(0).astype(int).astype(str)

    def categorize(m): return "BABY" if m in ['2', '3', '4', '5', '6'] else "ADULTO"
    df_order['Categoria'] = df_order['Máquina'].apply(categorize)
    
    return df_order, df_stops

@st.cache_data
def load_planner_metas(file, data_ref):
    try:
        xls = pd.ExcelFile(file)
        target = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        row_dates = df_raw.iloc[2, :].tolist()
        row_meta = df_raw.iloc[124, :].tolist() 
        
        meta_geral_mes = 0
        meta_ate_hoje = 0
        for col_idx, d_val in enumerate(row_dates):
            if isinstance(d_val, (datetime, pd.Timestamp)):
                valor = pd.to_numeric(row_meta[col_idx], errors='coerce') or 0
                meta_geral_mes += valor
                if d_val.date() <= data_ref:
                    meta_ate_hoje += valor
        return meta_geral_mes, meta_ate_hoje
    except: return 0, 0

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    up_prod = st.file_uploader("1. Produção (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("2. DATAS (.xlsx)", type=["xlsx"])
    st.markdown("---")
    if up_prod:
        menu = st.radio("NAVEGAÇÃO", ["📋 REPORTE DIÁRIO", "📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"])

if up_prod:
    df_order, df_stops = load_data(up_prod)

    def mini_gauge(label, value, color, target, height=150):
        fig = go.Figure(go.Indicator(
            mode="gauge+number", value=value,
            number={'suffix': "%", 'font': {'size': 18}},
            title={'text': label, 'font': {'size': 12}},
            gauge={'axis': {'range': [0, 100]}, 'bar': {'color': color},
                   'threshold': {'line': {'color': "white", 'width': 2}, 'value': target}}
        ))
        fig.update_layout(height=height, margin=dict(l=10, r=10, t=30, b=10), paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"})
        return fig

    # =========================================================
    # ABA: REPORTE DIÁRIO (KPIs em Cards Padronizados)
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        st.subheader("⚙️ Filtros da Página")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            data_ref_reporte = st.date_input("Data de Referência (Acumulado Mês)", df_order['Data'].max().date())
        with col_f2:
            datas_disp = sorted(df_order['Data'].dt.date.unique().tolist(), reverse=True)
            dias_selecionados = st.multiselect("Filtrar histórico (Tabelas):", datas_disp, default=datas_disp[:3] if len(datas_disp) >= 3 else datas_disp)

        st.markdown(f"## 📋 Reporte Diário de Produção - {data_ref_reporte.strftime('%d/%m/%Y')}")

        df_acumulado_mes = df_order[
            (df_order['Data'].dt.month == data_ref_reporte.month) & 
            (df_order['Data'].dt.year == data_ref_reporte.year) & 
            (df_order['Data'].dt.date <= data_ref_reporte)
        ]
        
        estoque_acum_mes = df_acumulado_mes['Peças Estoque - Ajuste'].sum()
        mc_acum_mes = df_acumulado_mes['Machine Counter'].sum()
        rt_acum_mes = df_acumulado_mes['Run Time'].sum()
        hp_acum_mes = df_acumulado_mes['Horário Padrão'].sum()
        
        mov_acum_mes = (rt_acum_mes / hp_acum_mes * 100) if hp_acum_mes > 0 else 0
        loss_acum_mes = ((mc_acum_mes - estoque_acum_mes) / mc_acum_mes * 100) if mc_acum_mes > 0 else 0
        
        # Metas e Gaps
        meta_mov, meta_loss = 90.0, 2.5
        gap_mov, gap_loss = mov_acum_mes - meta_mov, loss_acum_mes - meta_loss
        meta_geral_mes, meta_dinamica_hoje = load_planner_metas(up_datas, data_ref_reporte) if up_datas else (0, 0)

        # KPIs em Cards Padronizados (Enquadrados conforme solicitado)
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Movimentação Mês (Meta 90%)</div><div class="metric-value">{mov_acum_mes:.1f}%</div><div style="font-size:0.6rem; color:{'#10b981' if gap_mov>=0 else '#f43f5e'}">{gap_mov:+.1f}% vs meta</div></div>
                <div class="metric-card"><div class="metric-title">Loss Mês (Meta 2,5%)</div><div class="metric-value" style="color:#f43f5e">{loss_acum_mes:.1f}%</div><div style="font-size:0.6rem; color:{'#10b981' if gap_loss<=0 else '#f43f5e'}">{gap_loss:+.1f}% vs meta</div></div>
                <div class="metric-card"><div class="metric-title">Estoque Realizado Mês</div><div class="metric-value">{estoque_acum_mes:,.0f}</div></div>
            </div>
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Meta Acumulada (Até {data_ref_reporte.strftime('%d/%m')})</div><div class="metric-value" style="color:#3b82f6">{meta_dinamica_hoje:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Meta Geral do Mês (Datas)</div><div class="metric-value" style="color:#10b981">{meta_geral_mes:,.0f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        for dia in dias_selecionados:
            st.markdown(f"<div class='section-header'>DETALHAMENTO POR MÁQUINA - {dia.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
            df_dia = df_order[df_order['Data'].dt.date == dia]
            res = df_dia.groupby(['Categoria', 'Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
            res['Movimentação %'] = (res['Run Time'] / res['Horário Padrão'].replace(0,1) * 100).round(1)
            res['Perda %'] = ((res['Machine Counter'] - res['Peças Estoque - Ajuste']) / res['Machine Counter'].replace(0,1) * 100).round(1)
            st.table(res[['Categoria','Máquina','Movimentação %','Perda %','Peças Estoque - Ajuste']].rename(columns={'Peças Estoque - Ajuste':'Qtd Estoque'}))

    # =========================================================
    # ABA 1: PERFORMANCE (MANTIDA)
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t1')
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f["Machine Counter"].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f["Peças Estoque - Ajuste"].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f["Run Time"].sum():,.0f}m</div></div>
            </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        hp_sum = df_f['Horário Padrão'].sum()
        with col1: st.plotly_chart(mini_gauge("Movimentação (%)", (df_f['Run Time'].sum()/hp_sum*100 if hp_sum>0 else 0), "#10b981", 90, 300), use_container_width=True)
        with col2: st.plotly_chart(mini_gauge("Loss (%)", ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100 if df_f['Machine Counter'].sum()>0 else 0), "#e74c3c", 2.5, 300), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS (MANTIDA)
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1])]
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values().tail(10), orientation='h', title="Minutos Totais", color_discrete_sequence=['#f43f5e']), use_container_width=True)
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['QTD'].sum().sort_values().tail(10), orientation='h', title="Frequência (Qtd)", color_discrete_sequence=['#3b82f6']), use_container_width=True)

    # =========================================================
    # ABA 3: CALENDÁRIO (MANTIDA)
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        m_idx = list(calendar.month_name).index(mes_sel) + 1
        df_c = df_order[df_order['Data'].dt.month == m_idx]
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        cols = st.columns(7)
        for i, d in enumerate(['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']): cols[i].markdown(f"<div class='calendar-day-name'>{d}</div>", unsafe_allow_html=True)
        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, m_idx))
        html_grid = '<div class="calendar-grid">'
        for d in days:
            if d == 0: html_grid += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                mov = (row['Run Time'].values[0]/row['Horário Padrão'].values[0]*100) if not row.empty and row['Horário Padrão'].values[0]>0 else 0
                cor = "#059669" if mov > 85 else "#dc2626" if mov > 0 else "#1e293b"
                html_grid += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">M: {mov:.1f}%</div></div>'
        st.markdown(html_grid + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA 4: SEMANAL (MANTIDA)
    # =========================================================
    elif menu == "📋 ANÁLISE SEMANAL":
        maq_sel = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turn_sel = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        per = st.sidebar.date_input("Período", [datetime.now()-timedelta(days=7), datetime.now()])
        df_b = df_order[(df_order['Data'].dt.date >= per[0]) & (df_order['Data'].dt.date <= per[1]) & (df_order['Turno'].isin(turn_sel))]
        df_sb = df_stops[(df_stops['Data'].dt.date >= per[0]) & (df_stops['Data'].dt.date <= per[1]) & (df_stops['Máquina'] == maq_sel) & (df_stops['Turno'].isin(turn_sel))]
        
        st.markdown(f"""<div style="text-align:center; border:2px solid #10b981; padding:10px; border-radius:10px;"><h2 style="color:white; margin:0;">RELATÓRIO SEMANAL - MÁQUINA {maq_sel}</h2><p style="color:#94a3b8;">{per[0].strftime('%d/%m')} a {per[1].strftime('%d/%m/%Y')}</p></div>""", unsafe_allow_html=True)
        col_g, col_r = st.columns([2, 1])
        with col_g:
            stop_data = df_sb.groupby('Problema')['Minutos'].sum().sort_values().tail(5)
            st.plotly_chart(px.bar(stop_data, orientation='h', text_auto=True, color_discrete_sequence=['#10b981']).update_layout(height=300, paper_bgcolor='rgba(0,0,0,0)', font={'color':'white'}), use_container_width=True)
            pior_p = stop_data.index[-1] if not stop_data.empty else "Nenhuma Parada"
        with col_r:
            rk = df_b.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index(); rk['Mov %'] = (rk['Run Time']/rk['Horário Padrão']*100).round(1); rk = rk.sort_values('Mov %', ascending=False).reset_index(drop=True); rk.index += 1
            for i, r in rk.iterrows(): st.markdown(f"<div {'class=\"highlight-rank\"' if r['Máquina']==maq_sel else ''}>{i}º - MÁQ {r['Máquina']}: {r['Mov %']}%</div>", unsafe_allow_html=True)
            pos = rk[rk['Máquina']==maq_sel].index[0] if maq_sel in rk['Máquina'].values else 99
            msg, cor = ("🏆 Liderança semanal!", "#064e3b") if pos <= 2 else ("🚀 Foco para subir!", "#7f1d1d")
            st.markdown(f"<div class='feedback-box' style='background:{cor}; border-left:5px solid #10b981;'>{msg}</div>", unsafe_allow_html=True)

        st.markdown(f"""<div class="five-why-box"><h3 style="color:#059669; margin:0;">ANÁLISE 5 PORQUÊS: {pior_p}</h3>1. Por que? <div class="five-why-line"></div> 2. Por que? <div class="five-why-line"></div> 3. Por que? <div class="five-why-line"></div> 4. Por que? <div class="five-why-line"></div> 5. Por que? <div class="five-why-line"></div><b>CAUSA RAIZ / PLANO DE AÇÃO:</b> <div class="five-why-line"></div><div class="five-why-line"></div></div>""", unsafe_allow_html=True)

else:
    st.info("💡 Carregue os arquivos Excel para iniciar.")
