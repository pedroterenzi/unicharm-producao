import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta
import sqlite3

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# =========================================================
# BANCO DE DADOS LOCAL (SQLite)
# =========================================================
def init_db():
    conn = sqlite3.connect('reportes_turno.db')
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS reportes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_registro TEXT,
            turno TEXT,
            coordenador TEXT,
            ocorrencias TEXT,
            maq_analisada TEXT,
            problema TEXT,
            pq1 TEXT, pq2 TEXT, pq3 TEXT, pq4 TEXT, pq5 TEXT,
            oque TEXT,
            quem TEXT,
            quando TEXT,
            status TEXT
        )
    """)
    conn.commit()
    conn.close()

# Inicializa o banco de dados local
init_db()

# --- FUNÇÃO AUXILIAR DE FORMATAÇÃO (Milhares com ponto) ---
def fmt(valor):
    if pd.isna(valor) or valor is None:
        return "0"
    try:
        return f"{int(valor):,}".replace(",", ".")
    except:
        return str(valor)

# --- ESTILIZAÇÃO CSS PREMIUM (LIGHT MODE - FUNDO BRANCO) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    
    /* Fundo Branco */
    .stApp { background-color: #ffffff; color: #1e293b; }
    
    /* Menu Lateral Premium Light */
    [data-testid="stSidebar"] { background-color: #f8fafc; border-right: 1px solid #e2e8f0; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #ffffff; 
        border: 1px solid #e2e8f0;
        padding: 12px 18px !important; 
        border-radius: 12px !important;
        margin-bottom: 8px !important; 
        color: #334155 !important; 
        cursor: pointer;
        font-weight: 600;
        font-size: 0.9rem;
        transition: all 0.25s ease-in-out;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        display: block !important;
        text-align: center;
    }
    /* Efeito de Hover no menu */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label:hover {
        background-color: #f1f5f9 !important;
        border-color: #cbd5e1 !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }
    /* Botão Ativo Premium */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important; 
        color: white !important; 
        border: 1px solid #047857 !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25) !important;
    }

    /* Cards de Métricas Light */
    .metric-container { display: flex; justify-content: space-between; gap: 8px; margin-bottom: 15px; }
    .metric-card {
        background: #f8fafc; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid #e2e8f0;
        flex: 1; min-height: 80px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #64748b; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    .metric-value { color: #10b981; font-size: 1.3rem; font-weight: 900; line-height: 1; }

    /* Calendário Light */
    .calendar-day-name { text-align: center; font-weight: 900; color: #10b981; font-size: 0.8rem; padding-bottom: 5px; }
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .day-card { background: #f8fafc; border-radius: 8px; padding: 10px; min-height: 95px; border: 1px solid #e2e8f0; }
    .day-number { font-size: 1rem; font-weight: 900; color: #1e293b; }
    .day-status { font-size: 0.75rem; font-weight: 600; color: #64748b; text-align: right; }

    /* Ranking e Feedback */
    .highlight-rank { background: #dcfce7 !important; color: #166534 !important; font-weight: 900; border-radius: 5px; padding: 5px; }
    .feedback-box { padding: 12px; border-radius: 8px; margin-bottom: 10px; text-align: center; font-weight: 700; font-size: 0.9rem; border: 1px solid #e2e8f0; }

    /* 5 Porquês */
    .five-why-box { border: 2px solid #1e293b; padding: 15px; background: #ffffff; color: #000; margin-top: 15px; }
    .five-why-line { border-bottom: 1px solid #000; padding: 10px 0; font-size: 0.9rem; }

    /* Reporte Diário Header */
    .section-header {
        background: #f1f5f9; padding: 10px; border-radius: 5px;
        color: #0f172a; font-weight: 800; text-transform: uppercase;
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
def load_planner_metas_advanced(file, data_ref):
    try:
        xls = pd.ExcelFile(file)
        target = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        
        row_dates = df_raw.iloc[2, :].tolist()
        maq_lines = {'1': 6, '2': 28, '3': 47, '4': 58, '5': 77, '6': 96, '7': 113}
        
        plan_dia = {}
        plan_mes_acum = {}
        
        idx_col_ref = None
        for i, d in enumerate(row_dates):
            if isinstance(d, (datetime, pd.Timestamp)) and d.date() == data_ref:
                idx_col_ref = i
                break
        
        if idx_col_ref is not None:
            for maq, row_idx in maq_lines.items():
                plan_dia[maq] = pd.to_numeric(df_raw.iloc[row_idx, idx_col_ref], errors='coerce') or 0
                soma_mtd = 0
                for i, d in enumerate(row_dates):
                    if isinstance(d, (datetime, pd.Timestamp)):
                        if d.year == data_ref.year and d.month == data_ref.month and d.date() <= data_ref:
                            soma_mtd += pd.to_numeric(df_raw.iloc[row_idx, i], errors='coerce') or 0
                plan_mes_acum[maq] = soma_mtd
        
        row_meta_125 = df_raw.iloc[124, :].tolist()
        m_total_mes, m_mtd_total = 0, 0
        for i, d in enumerate(row_dates):
            if isinstance(d, (datetime, pd.Timestamp)):
                val = pd.to_numeric(row_meta_125[i], errors='coerce') or 0
                if d.year == data_ref.year and d.month == data_ref.month:
                    m_total_mes += val
                    if d.date() <= data_ref:
                        m_mtd_total += val
                        
        return plan_dia, plan_mes_acum, m_total_mes, m_mtd_total
    except:
        return {}, {}, 0, 0

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Excel Produção (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("📂 Carregar Excel DATAS (.xlsx)", type=["xlsx"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📋 REPORTE DIÁRIO", "📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL", "📝 LANÇAR REPORTE", "📊 ACOMPANHAMENTO"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    def mini_gauge(label, value, color, target, height=150):
        fig = go.Figure(go.Indicator(
            mode="gauge+number", value=value,
            number={'suffix': "%", 'font': {'size': 18, 'color': '#1e293b'}},
            title={'text': label, 'font': {'size': 12, 'color': '#64748b'}},
            gauge={'axis': {'range': [0, 100], 'tickcolor': '#1e293b'}, 'bar': {'color': color},
                   'threshold': {'line': {'color': "#1e293b", 'width': 2}, 'value': target}}
        ))
        fig.update_layout(height=height, margin=dict(l=10, r=10, t=30, b=10), paper_bgcolor='rgba(0,0,0,0)', font={'color': "#1e293b"})
        return fig

    # =========================================================
    # ABA: REPORTE DIÁRIO
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        st.subheader("⚙️ Filtros da Página")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            data_ref_reporte = st.date_input("Data de Referência", df_order['Data'].max().date())
        with col_f2:
            datas_disp = sorted(df_order['Data'].dt.date.unique().tolist(), reverse=True)
            dias_sel = st.multiselect("Filtrar histórico:", datas_disp, default=datas_disp[:3] if len(datas_disp) >= 3 else datas_disp)

        st.markdown(f"## 📋 Reporte Diário de Produção - {data_ref_reporte.strftime('%d/%m/%Y')}")

        plan_dia, plan_mes_acum, m_total_mes, m_mtd_total = load_planner_metas_advanced(up_datas, data_ref_reporte) if up_datas else ({}, {}, 0, 0)

        df_acumulado_mes = df_order[(df_order['Data'].dt.month == data_ref_reporte.month) & (df_order['Data'].dt.year == data_ref_reporte.year) & (df_order['Data'].dt.date <= data_ref_reporte)]
        estoque_acum_mes = df_acumulado_mes['Peças Estoque - Ajuste'].sum()
        total_mc_mes = df_acumulado_mes['Machine Counter'].sum()
        mov_acum_mes = (df_acumulado_mes['Run Time'].sum() / df_acumulado_mes['Horário Padrão'].sum() * 100) if df_acumulado_mes['Horário Padrão'].sum() > 0 else 0
        loss_acum_mes = ((total_mc_mes - estoque_acum_mes) / total_mc_mes * 100) if total_mc_mes > 0 else 0
        
        gap_mov = mov_acum_mes - 90.0
        gap_loss = loss_acum_mes - 2.5

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Movimentação Mês (Meta 90%)</div><div class="metric-value">{mov_acum_mes:.1f}%</div><div style="font-size:0.6rem; color:{'#10b981' if gap_mov>=0 else '#f43f5e'}">{gap_mov:+.1f}% vs meta</div></div>
                <div class="metric-card"><div class="metric-title">Loss Mês (Meta 2,5%)</div><div class="metric-value" style="color:#f43f5e">{loss_acum_mes:.1f}%</div><div style="font-size:0.6rem; color:{'#10b981' if gap_loss<=0 else '#f43f5e'}">{gap_loss:+.1f}% vs meta</div></div>
                <div class="metric-card"><div class="metric-title">Estoque Realizado MTD</div><div class="metric-value">{fmt(estoque_acum_mes)}</div></div>
            </div>
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Meta Acumulada MTD (Até {data_ref_reporte.strftime('%d/%m')})</div><div class="metric-value" style="color:#3b82f6">{fmt(m_mtd_total)}</div></div>
                <div class="metric-card"><div class="metric-title">Meta Geral do Mês Completo</div><div class="metric-value" style="color:#10b981">{fmt(m_total_mes)}</div></div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Gaps / Ganhos de Peças (Comparativo por Máquina)</div>", unsafe_allow_html=True)
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.markdown(f"**Comparativo do Dia {data_ref_reporte.strftime('%d/%m')}**")
            df_real_dia = df_order[df_order['Data'].dt.date == data_ref_reporte].groupby('Máquina')['Peças Estoque - Ajuste'].sum().to_dict()
            res_gap_dia = []
            for m in sorted(df_order['Máquina'].unique(), key=int):
                r, p = df_real_dia.get(m, 0), plan_dia.get(m, 0)
                res_gap_dia.append({'Máquina': m, 'Realizado': fmt(r), 'Planejado': fmt(p), 'Gap/Ganho': fmt(r-p)})
            st.table(pd.DataFrame(res_gap_dia))
        with col_t2:
            st.markdown(f"**Acumulado Mês (MTD) até {data_ref_reporte.strftime('%d/%m')}**")
            df_real_mes = df_acumulado_mes.groupby('Máquina')['Peças Estoque - Ajuste'].sum().to_dict()
            res_gap_mes = []
            for m in sorted(df_order['Máquina'].unique(), key=int):
                r, p = df_real_mes.get(m, 0), plan_mes_acum.get(m, 0)
                res_gap_mes.append({'Máquina': m, 'Realizado MTD': fmt(r), 'Planejado MTD': fmt(p), 'Gap/Ganho MTD': fmt(r-p)})
            st.table(pd.DataFrame(res_gap_mes))

        for dia in dias_sel:
            st.markdown(f"<div class='section-header'>DETALHAMENTO POR MÁQUINA - {dia.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
            df_dia = df_order[df_order['Data'].dt.date == dia]
            res = df_dia.groupby(['Categoria', 'Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
            
            res['Movimentação %'] = (res['Run Time'] / res['Horário Padrão'].replace(0,1) * 100).apply(lambda x: f"{x:.2f}%".replace('.', ','))
            res['Perda %'] = ((res['Machine Counter'] - res['Peças Estoque - Ajuste']) / res['Machine Counter'].replace(0,1) * 100).apply(lambda x: f"{x:.2f}%".replace('.', ','))
            
            res['Peças Estoque'] = res['Peças Estoque - Ajuste'].apply(fmt)
            st.table(res[['Categoria','Máquina','Movimentação %','Perda %','Peças Estoque']])

    # =========================================================
    # ABA: PERFORMANCE
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t1')
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        str_maquinas = ", ".join(f_maq) if f_maq else "Nenhuma"
        str_turnos_f = ", ".join(f_turno) if f_turno else "Nenhum"
        st.markdown(f"## 📈 Performance Industrial — Máquina(s): {str_maquinas} | Turno(s): {str_turnos_f}")
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{fmt(df_f["Machine Counter"].sum())}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{fmt(df_f["Peças Estoque - Ajuste"].sum())}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{fmt(df_f["Run Time"].sum())}m</div></div>
            </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        hp_sum = df_f['Horário Padrão'].sum()
        with col1: st.plotly_chart(mini_gauge("Movimentação (%)", (df_f['Run Time'].sum()/hp_sum*100 if hp_sum>0 else 0), "#10b981", 90, 280), use_container_width=True)
        with col2: st.plotly_chart(mini_gauge("Loss (%)", ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100 if df_f['Machine Counter'].sum()>0 else 0), "#e74c3c", 2.5, 280), use_container_width=True)

    # =========================================================
    # ABA: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        f_maq_s = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()), key='m2')
        f_turno_s = st.sidebar.multiselect("Turnos", sorted(df_stops['Turno'].unique()), default=sorted(df_stops['Turno'].unique()), key='ts2')
        
        df_s_f = df_stops[
            (df_stops['Data'].dt.date >= f_data_s[0]) & 
            (df_stops['Data'].dt.date <= f_data_s[1]) & 
            (df_stops['Máquina'].isin(f_maq_s)) & 
            (df_stops['Turno'].isin(f_turno_s))
        ]
        str_maquinas_s = ", ".join(f_maq_s) if f_maq_s else "Nenhuma"
        str_turnos_s = ", ".join(f_turno_s) if f_turno_s else "Nenhum"
        st.markdown(f"## 🛑 Análise de Paradas — Máquina(s): {str_maquinas_s} | Turno(s): {str_turnos_s}")
        
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values().tail(10), orientation='h', title="Minutos Totais", color_discrete_sequence=['#f43f5e']).update_layout(paper_bgcolor='white', plot_bgcolor='white', font={'color':'black'}), use_container_width=True)
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['QTD'].sum().sort_values().tail(10), orientation='h', title="Frequência (Qtd)", color_discrete_sequence=['#3b82f6']).update_layout(paper_bgcolor='white', plot_bgcolor='white', font={'color':'black'}), use_container_width=True)

    # =========================================================
    # ABA: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        m_idx = list(calendar.month_name).index(mes_sel) + 1
        df_c = df_order[(df_order['Data'].dt.month == m_idx)]
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        
        st.markdown(f"### 📅 Cronograma {mes_sel}")
        cols = st.columns(7)
        for i, d_name in enumerate(['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']): cols[i].markdown(f"<div class='calendar-day-name'>{d_name}</div>", unsafe_allow_html=True)
        
        ano_ref = df_order['Data'].max().year
        days = list(calendar.Calendar(0).itermonthdays(ano_ref, m_idx))
        html_grid = '<div class="calendar-grid">'
        for d in days:
            if d == 0: html_grid += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                mov = (row['Run Time'].values[0]/row['Horário Padrão'].replace(0,1).values[0]*100) if not row.empty else 0
                cor = "#059669" if mov > 85 else "#dc2626" if mov > 0 else "#f1f5f9"
                html_grid += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status" style="color:{"white" if mov > 0 else "#64748b"}">{mov:.1f}%</div></div>'
        st.markdown(html_grid + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA: ANÁLISE SEMANAL
    # =========================================================
    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Filtros Board")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_b = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='tb')
        periodo_b = st.sidebar.date_input("Período", [df_order['Data'].max() - timedelta(days=7), df_order['Data'].max()])
        
        df_b_all = df_order[(df_order['Data'].dt.date >= periodo_b[0]) & (df_order['Data'].dt.date <= periodo_b[1]) & (df_order['Turno'].isin(turno_b))]
        df_b = df_b_all[df_b_all['Máquina'] == maq_b]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo_b[0]) & (df_stops['Data'].dt.date <= periodo_b[1]) & (df_stops['Máquina'] == maq_b) & (df_stops['Turno'].isin(turno_b))]

        str_turnos = ", ".join(turno_b) if turno_b else "Nenhum"
        st.markdown(f"""<div style="text-align:center; border-bottom:3px solid #10b981; padding-bottom:10px; margin-bottom:15px;">
            <h1 style="color:#0f172a; margin:0;">RELATÓRIO SEMANAL DE PERFORMANCE - MÁQUINA {maq_b}</h1>
            <h3 style="color:#10b981; margin:0;">TURNO(S): {str_turnos}</h3>
            <p style="color:#64748b; font-size:1rem;">Período: {periodo_b[0].strftime('%d/%m')} a {periodo_b[1].strftime('%d/%m/%Y')}</p></div>""", unsafe_allow_html=True)

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].replace(0,1).sum()*100)
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].replace(0,1).sum()*100)
        pecas_v = df_b["Peças Estoque - Ajuste"].sum()

        v1, v2, v3 = st.columns([1, 1, 1])
        with v1: st.plotly_chart(mini_gauge("Movimentação", m_v, "#10b981", 85, 180), use_container_width=True)
        with v2: st.plotly_chart(mini_gauge("Loss", l_v, "#e74c3c", 5, 180), use_container_width=True)
        with v3: st.markdown(f'<div class="metric-card" style="height:150px;"><div class="metric-title">Peças Enviadas</div><div class="metric-value" style="font-size:1.8rem;">{fmt(pecas_v)}</div></div>', unsafe_allow_html=True)

        rank_df = df_b_all.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        rank_df['Mov %'] = (rank_df['Run Time'] / rank_df['Horário Padrão'].replace(0,1) * 100).round(1)
        rank_df = rank_df.sort_values('Mov %', ascending=False).reset_index(drop=True)
        rank_df.index += 1
        
        check_maq = rank_df[rank_df['Máquina'] == maq_b]
        if not check_maq.empty:
            posicao = check_maq.index[0]
            total_maqs = len(rank_df)
            if posicao <= 2:
                msg, col = ("🏆 Liderança semanal! Excelente performance.", "#dcfce7")
            else:
                msg, col = ("🚀 Foco na melhoria para subir o ranking semanal!", "#fee2e2")
            st.markdown(f'<div class="feedback-box" style="background:{col}; color:black; border-left:5px solid #10b981;">{msg}</div>', unsafe_allow_html=True)

        col_g, col_r = st.columns([2, 1])
        with col_g:
            st.markdown("🛑 **Impacto das Paradas (Piores 5)**")
            stop_data = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(5)
            st.plotly_chart(px.bar(stop_data, orientation='h', text_auto=True, color_discrete_sequence=['#10b981']).update_layout(height=300, paper_bgcolor='white', plot_bgcolor='white', font={'color':'black'}), use_container_width=True)
            pior_p = stop_data.index[-1] if not stop_data.empty else "Nenhuma Parada"

        with col_r:
            st.markdown("🏆 **Ranking Movimentação**")
            for i, r in rank_df.iterrows():
                style = "class='highlight-rank'" if r['Máquina'] == maq_b else ""
                st.markdown(f"<div {style}>{i}º - MÁQ {r['Máquina']}: {r['Mov %']}%</div>", unsafe_allow_html=True)

        st.markdown(f"""<div class="five-why-box"><h3 style="color:#059669; margin:0;">ANÁLISE 5 PORQUÊS: {pior_p}</h3>
            1. Por que? <div class="five-why-line"></div> 2. Por que? <div class="five-why-line"></div> 3. Por que? <div class="five-why-line"></div> 
            4. Por que? <div class="five-why-line"></div> 5. Por que? <div class="five-why-line"></div>
            <b>CAUSA RAIZ / PLANO DE AÇÃO:</b> <div class="five-why-line"></div><div class="five-why-line"></div></div>""", unsafe_allow_html=True)

    # =========================================================
    # ABA: LANÇAR REPORTE
    # =========================================================
    elif menu == "📝 LANÇAR REPORTE":
        st.markdown("## 📝 Formulário de Registro de Turno")
        with st.form("form_reporte", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1: data_rep = st.date_input("Data do Turno", datetime.now().date())
            with c2: turno_rep = st.selectbox("Turno", ["T1", "T2", "T3"])
            with c3: coord_rep = st.text_input("Coordenador Responsável").upper()
            st.markdown("<div class='section-header'>1. Principais Ocorrências / Paradas do Turno</div>", unsafe_allow_html=True)
            txt_ocorrencias = st.text_area("Descreva as ocorrências", height=120)
            st.markdown("<div class='section-header'>2. Análise de Causa Raiz (Top Ofensor)</div>", unsafe_allow_html=True)
            cc1, cc2 = st.columns(2)
            with cc1: maq_an = st.text_input("Máquina Analisada").upper()
            with cc2: prob_an = st.text_input("Problema Foco")
            p1 = st.text_input("Por que 1?")
            p2 = st.text_input("Por que 2?")
            p3 = st.text_input("Por que 3?")
            p4 = st.text_input("Por que 4?")
            p5 = st.text_input("Por que 5? (Causa Raiz)")
            st.markdown("<div class='section-header'>3. Plano de Ação Imediato</div>", unsafe_allow_html=True)
            ccc1, ccc2, ccc3 = st.columns([2, 1, 1])
            with ccc1: action_oque = st.text_area("O quê (Ações)")
            with ccc2: action_quem = st.text_input("Quem (Responsável)")
            with ccc3: action_quando = st.text_input("Quando (Prazo)")
            status_inicial = st.selectbox("Status Inicial do Plano", ["Pendente", "Em Andamento", "Resolvido"])
            submit = st.form_submit_button("💾 SALVAR REPORTE NO BANCO DE DADOS")
            
            if submit:
                if not coord_rep or not prob_an:
                    st.error("Por favor, preencha os campos essenciais como Coordenador e Problema Foco.")
                else:
                    conn = sqlite3.connect('reportes_turno.db')
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO reportes (data_registro, turno, coordenador, ocorrencias, maq_analisada, problema,
                        pq1, pq2, pq3, pq4, pq5, oque, quem, quando, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(data_rep), turno_rep, coord_rep, txt_ocorrencias, maq_an, prob_an, p1, p2, p3, p4, p5, action_oque, action_quem, action_quando, status_inicial))
                    conn.commit()
                    conn.close()
                    st.success("🎉 Reporte alocado e registrado com sucesso no banco de dados local!")

    # =========================================================
    # ABA: ACOMPANHAMENTO (EDICAO DINÂMICA COMPLETA + EXCLUSÃO)
    # =========================================================
    elif menu == "📊 ACOMPANHAMENTO":
        st.markdown("## 📊 Painel de Acompanhamento de Ações")
        conn = sqlite3.connect('reportes_turno.db')
        df_db = pd.read_sql_query("SELECT * FROM reportes ORDER BY data_registro DESC", conn)
        conn.close()
        
        if df_db.empty:
            st.info("Nenhum registro encontrado no banco de dados local.")
        else:
            f_col1, f_col2 = st.columns(2)
            with f_col1: filtro_status = st.multiselect("Filtrar por Status", df_db['status'].unique(), default=df_db['status'].unique())
            with f_col2: filtro_turno = st.multiselect("Filtrar por Turno", df_db['turno'].unique(), default=df_db['turno'].unique())
            df_filtrado_db = df_db[(df_db['status'].isin(filtro_status)) & (df_db['turno'].isin(filtro_turno))].copy()
            
            def colorir_linhas_por_status(row):
                if row['status'] == 'Pendente':
                    return ['background-color: #fee2e2; color: #b91c1c; font-weight: 600'] * len(row)
                elif row['status'] == 'Em Andamento':
                    return ['background-color: #fef3c7; color: #d97706; font-weight: 600'] * len(row)
                elif row['status'] == 'Resolvido':
                    return ['background-color: #dcfce7; color: #15803d; font-weight: 600'] * len(row)
                return [''] * len(row)

            cols_exibicao = ['id', 'data_registro', 'turno', 'coordenador', 'maq_analisada', 'problema', 'quem', 'quando', 'status']
            st.dataframe(df_filtrado_db[cols_exibicao].style.apply(colorir_linhas_por_status, axis=1), use_container_width=True)
            
            st.markdown("<div class='section-header'>✏️ Gerenciar / Editar Informações do Reporte</div>", unsafe_allow_html=True)
            id_selecionado = st.number_input("Digite o ID do reporte para carregar o painel de edição:", min_value=1, step=1)
            
            if id_selecionado in df_db['id'].values:
                # Recupera os dados da linha selecionada do banco
                row_sel = df_db[df_db['id'] == id_selecionado].iloc[0]
                
                st.markdown(f"#### Editando Dados do ID: `{id_selecionado}`")
                
                # Formula o painel de inputs preenchidos com os dados atuais para edição
                e_c1, e_c2, e_c3 = st.columns(3)
                with e_col1:
                    edit_coord = st.text_input("Editar Coordenador", value=str(row_sel['coordenador'])).upper()
                with e_col2:
                    edit_status = st.selectbox("Alterar Status", ["Pendente", "Em Andamento", "Resolvido"], index=["Pendente", "Em Andamento", "Resolvido"].index(row_sel['status']))
                with e_col3:
                    edit_maq = st.text_input("Editar Máquina Analisada", value=str(row_sel['maq_analisada'])).upper()
                
                edit_ocorrencias = st.text_area("Editar Ocorrências / Paradas do Turno", value=str(row_sel['ocorrencias']), height=100)
                edit_problema = st.text_input("Editar Problema Foco", value=str(row_sel['problema']))
                
                # Edição dos 5 porquês
                st.write("**Editar Análise dos 5 Porquês:**")
                epq1 = st.text_input("Por que 1?", value=str(row_sel['pq1']))
                epq2 = st.text_input("Por que 2?", value=str(row_sel['pq2']))
                epq3 = st.text_input("Por que 3?", value=str(row_sel['pq3']))
                epq4 = st.text_input("Por que 4?", value=str(row_sel['pq4']))
                epq5 = st.text_input("Por que 5? (Causa Raiz)", value=str(row_sel['pq5']))
                
                # Edição das ações
                st.write("**Editar Plano de Ação:**")
                ea_oque = st.text_area("O quê (Ação)", value=str(row_sel['oque']))
                ea_quem = st.text_input("Quem (Responsável)", value=str(row_sel['quem']))
                ea_quando = st.text_input("Quando (Prazo)", value=str(row_sel['quando']))
                
                st.markdown("---")
                col_actions1, col_actions2 = st.columns(2)
                
                # Ação 1: Salvar alterações de todos os campos editados
                with col_actions1:
                    if st.button("💾 SALVAR ALTERAÇÕES", use_container_width=True):
                        conn = sqlite3.connect('reportes_turno.db')
                        cursor = conn.cursor()
                        cursor.execute("""
                            UPDATE reportes 
                            SET coordenador = ?, status = ?, maq_analisada = ?, ocorrencias = ?, problema = ?, 
                                pq1 = ?, pq2 = ?, pq3 = ?, pq4 = ?, pq5 = ?, oque = ?, quem = ?, quando = ?
                            WHERE id = ?
                        """, (
                            edit_coord, edit_status, edit_maq, edit_ocorrencias, edit_problema,
                            epq1, epq2, epq3, epq4, epq5, ea_oque, ea_quem, ea_quando, int(id_selecionado)
                        ))
                        conn.commit()
                        conn.close()
                        st.success(f"🎉 Todas as informações do ID {id_selecionado} foram atualizadas com sucesso!")
                        st.rerun()
                
                # Ação 2: Exclusão definitiva do registro
                with col_actions2:
                    if st.button("❌ EXCLUIR REPORTE DEFINITIVAMENTE", type="primary", use_container_width=True):
                        conn = sqlite3.connect('reportes_turno.db')
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM reportes WHERE id = ?", (int(id_selecionado),))
                        conn.commit()
                        conn.close()
                        st.success(f"O reporte de ID {id_selecionado} foi excluído com sucesso do banco de dados!")
                        st.rerun()
            else:
                st.caption("Insira um ID válido presente na tabela acima para gerenciar ou editar.")
else:
    st.info("💡 Por favor, carregue os arquivos Excel para iniciar.")
