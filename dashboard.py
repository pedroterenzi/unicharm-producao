import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta, date
import sqlite3

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- INICIALIZAÇÃO DE ESTADOS DO STREAMLIT (Controle de Visibilidade e Resets) ---
if 'mostrar_edicao' not in st.session_state:
    st.session_state['mostrar_edicao'] = False
if 'id_atual' not in st.session_state:
    st.session_state['id_atual'] = 0

if 'mostrar_edicao_semanal' not in st.session_state:
    st.session_state['mostrar_edicao_semanal'] = False
if 'id_atual_semanal' not in st.session_state:
    st.session_state['id_atual_semanal'] = 0

# Contador para resetar os campos do Nippo Coordenadores após a gravação
if 'contador_nippo' not in st.session_state:
    st.session_state['contador_nippo'] = 0

# =========================================================
# BANCO DE DADOS LOCAL (SQLite)
# =========================================================
def init_db():
    conn = sqlite3.connect('reportes_turno.db')
    cursor = conn.cursor()
    # Tabela 1: Reportes Diários de Turno
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
    # Tabela 2: Análises Semanais dos Operadores
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS analises_semanais (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_registro TEXT,
            turno TEXT,
            maquina TEXT,
            pior_parada TEXT,
            pq1 TEXT, pq2 TEXT, pq3 TEXT, pq4 TEXT, pq5 TEXT,
            causa_raiz TEXT,
            plano_acao TEXT,
            prazo TEXT,
            responsavel TEXT,
            status TEXT
        )
    """)
    # Tabela 3: Tabela Nippo Coordenadores (Troca de Turno por Máquina)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS nippo_coordenadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            turno TEXT,
            coordenador TEXT,
            tecnico TEXT,
            maquina TEXT,
            itens_compartilhar TEXT,
            produtividade REAL,
            loss REAL,
            sku TEXT,
            palete_inicial TEXT,
            palete_final TEXT,
            total_ordem INTEGER,
            data_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

# Inicializa o banco de dados local com todas as tabelas estruturadas
init_db()

# --- FUNÇÃO AUXILIAR DE FORMATAÇÃO (Milhares com ponto) ---
def fmt(valor):
    if pd.isna(valor) or valor is None:
        return "0"
    try:
        return f"{int(valor):,}".replace(",", ".")
    except:
        return str(valor)

# --- ESTILIZAÇÃO CSS PREMIUM (LIGHT MODE - PADRONIZAÇÃO ESTÉTICA DO MENU) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    
    /* Fundo Branco */
    .stApp { background-color: #ffffff; color: #1e293b; }
    
    /* --- MENU LATERAL ULTRA MODERNO E PADRONIZADO --- */
    [data-testid="stSidebar"] { background-color: #f8fafc; border-right: 1px solid #e2e8f0; }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] {
        display: flex;
        flex-direction: column;
        gap: 2px;
        width: 100%;
    }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #ffffff !important; 
        border: 1px solid #e2e8f0 !important;
        padding: 12px 18px !important; 
        border-radius: 10px !important;
        margin-bottom: 5px !important; 
        color: #475569 !important; 
        cursor: pointer;
        font-weight: 500;
        font-size: 0.82rem;
        transition: all 0.2s ease-in-out;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important;
        display: flex !important;
        align-items: center;
        justify-content: flex-start;
        width: 100% !important;
        box-sizing: border-box !important;
    }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label:hover {
        background-color: #f1f5f9 !important;
        border-color: #cbd5e1 !important;
        color: #0f172a !important;
        transform: translateX(2px);
    }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important; 
        color: #ffffff !important; 
        border: 1px solid #047857 !important;
        font-weight: 600;
        box-shadow: 0 4px 10px rgba(16, 185, 129, 0.2) !important;
    }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label icon {
        display: none !important;
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
    st.markdown("<h1 style='font-size:1.4rem; color:#10b981; font-weight:900; margin-bottom:15px; text-align:center;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Excel Production (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("📂 Carregar Excel DATAS (.xlsx)", type=["xlsx"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", [
            "📋 REPORTE DIÁRIO", 
            "📈 PERFORMANCE", 
            "🛑 TOP 10 PARADAS", 
            "📅 CALENDÁRIO", 
            "📋 ANÁLISE SEMANAL", 
            "📝 LANÇAR REPORTE", 
            "📊 ACOMPANHAMENTO",
            "📝 LANÇAR ANÁLISE SEMANAL", 
            "📋 ACOMP. ANÁLISES SEMANAIS", 
            "📊 APRESENTAÇÃO SEMANAL",
            "📋 NIPPO COORDENADORES"
        ])

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
    # ABA: LANÇAR REPORTE DIÁRIO
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
                    st.error("Por favor, preencha os campos essenciais.")
                else:
                    conn = sqlite3.connect('reportes_turno.db')
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO reportes (data_registro, turno, coordenador, ocorrencias, maq_analisada, समस्या,
                        pq1, pq2, pq3, pq4, pq5, oque, quem, quando, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(data_rep), turno_rep, coord_rep, txt_ocorrencias, maq_an, prob_an, p1, p2, p3, p4, p5, action_oque, action_quem, action_quando, status_inicial))
                    conn.commit(); conn.close()
                    st.success("🎉 Reporte alocado com sucesso!")

    # =========================================================
    # ABA: ACOMPANHAMENTO DIÁRIO
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
            with f_col1: filtro_status = st.multiselect("Filtrar por Status", df_db['status'].unique(), default=df_db['status'].unique(), key='ds1')
            with f_col2: filtro_turno = st.multiselect("Filtrar por Turno", df_db['turno'].unique(), default=df_db['turno'].unique(), key='dt1')
            df_filtrado_db = df_db[(df_db['status'].isin(filtro_status)) & (df_db['turno'].isin(filtro_turno))].copy()
            
            def colorir_linhas_por_status(row):
                if row['status'] == 'Pendente': return ['background-color: #fee2e2; color: #b91c1c; font-weight: 600'] * len(row)
                elif row['status'] == 'Em Andamento': return ['background-color: #fef3c7; color: #d97706; font-weight: 600'] * len(row)
                elif row['status'] == 'Resolvido': return ['background-color: #dcfce7; color: #15803d; font-weight: 600'] * len(row)
                return [''] * len(row)

            st.dataframe(df_filtrado_db[['id', 'data_registro', 'turno', 'coordenador', 'maq_analisada', 'problema', 'quem', 'quando', 'status']].style.apply(colorir_linhas_por_status, axis=1), use_container_width=True)
            
            st.markdown("<div class='section-header'>✏️ Gerenciar / Editar Informações do Reporte</div>", unsafe_allow_html=True)
            id_selecionado = st.number_input("Digite o ID do reporte para gerenciar:", min_value=1, step=1)
            
            if id_selecionado in df_db['id'].values:
                if id_selecionado != st.session_state['id_atual']:
                    st.session_state['id_atual'] = id_selecionado
                    st.session_state['mostrar_edicao'] = False
                
                if not st.session_state['mostrar_edicao']:
                    if st.button("🔍 ABRIR PAINEL DE GERENCIAMENTO / EDICAO"):
                        st.session_state['mostrar_edicao'] = True
                        st.rerun()
                else:
                    if st.button("🔼 MINIMIZAR / FECHAR PAINEL DE EDIÇÃO"):
                        st.session_state['mostrar_edicao'] = False
                        st.rerun()
                
                if st.session_state['mostrar_edicao']:
                    row_sel = df_db[df_db['id'] == id_selecionado].iloc[0]
                    e_c1, e_c2, e_c3 = st.columns(3)
                    with e_c1: edit_coord = st.text_input("Editar Coordenador", value=str(row_sel['coordenador'])).upper()
                    with e_c2: edit_status = st.selectbox("Alterar Status", ["Pendente", "Em Andamento", "Resolvido"], index=["Pendente", "Em Andamento", "Resolvido"].index(row_sel['status']))
                    with e_c3: edit_maq = st.text_input("Editar Máquina Analisada", value=str(row_sel['maq_analisada'])).upper()
                    
                    edit_ocorrencias = st.text_area("Editar Ocorrências", value=str(row_sel['ocorrencias']), height=100)
                    edit_problema = st.text_input("Editar Problema Foco", value=str(row_sel['problema']))
                    epq1 = st.text_input("Por que 1?", value=str(row_sel['pq1']))
                    epq2 = st.text_input("Por que 2?", value=str(row_sel['pq2']))
                    epq3 = st.text_input("Por que 3?", value=str(row_sel['pq3']))
                    epq4 = st.text_input("Por que 4?", value=str(row_sel['pq4']))
                    epq5 = st.text_input("Por que 5? (Causa Raiz)", value=str(row_sel['pq5']))
                    ea_oque = st.text_area("O quê (Ação)", value=str(row_sel['oque']))
                    ea_quem = st.text_input("Quem", value=str(row_sel['quem']))
                    ea_quando = st.text_input("Quando", value=str(row_sel['quando']))
                    
                    st.markdown("---")
                    col_actions1, col_actions2 = st.columns(2)
                    with col_actions1:
                        if st.button("💾 SALVAR ALTERAÇÕES", use_container_width=True):
                            conn = sqlite3.connect('reportes_turno.db')
                            cursor = conn.cursor()
                            cursor.execute("""
                                UPDATE reportes SET coordenador=?, status=?, maq_analisada=?, ocorrencias=?, problema=?, 
                                pq1=?, pq2=?, pq3=?, pq4=?, pq5=?, oque=?, quem=?, quando=? WHERE id=?
                            """, (edit_coord, edit_status, edit_maq, edit_ocorrencias, edit_problema, epq1, epq2, epq3, epq4, epq5, ea_oque, ea_quem, ea_quando, int(id_selecionado)))
                            conn.commit(); conn.close()
                            st.session_state['mostrar_edicao'] = False
                            st.success("🎉 Atualizado!")
                            st.rerun()
                    with col_actions2:
                        if st.button("❌ EXCLUIR REPORTE DEFINITIVAMENTE", type="primary", use_container_width=True):
                            conn = sqlite3.connect('reportes_turno.db')
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM reportes WHERE id = ?", (int(id_selecionado),))
                            conn.commit(); conn.close()
                            st.session_state['mostrar_edicao'] = False
                            st.success("Excluído!")
                            st.rerun()

    # =========================================================
    # ABA: LANÇAR ANÁLISE SEMANAL (OPERADORES)
    # =========================================================
    elif menu == "📝 LANÇAR ANÁLISE SEMANAL":
        st.markdown("## 📝 Formulário de Lançamento — Análise Semanal (Operadores)")
        with st.form("form_analise_semanal", clear_on_submit=True):
            s1, s2, s3 = st.columns(3)
            with s1: semana_ref = st.date_input("Semana de Referência (Início/Data)", datetime.now().date())
            with s2: turno_sem = st.selectbox("Turno Analisado", ["T1", "T2", "T3"], key='ts_sem')
            with s3: maq_sem = st.selectbox("Máquina Alvo", sorted(df_order['Máquina'].unique()), key='mq_sem')
                
            pior_parada_sem = st.text_input("Pior Parada Detectada (Ofensor da Semana)")
            st.markdown("<div class='section-header'>Análise Causa Raiz — Método dos 5 Porquês</div>", unsafe_allow_html=True)
            spq1 = st.text_input("1º Por que?")
            spq2 = st.text_input("2º Por que?")
            spq3 = st.text_input("3º Por que?")
            spq4 = st.text_input("4º Por que?")
            spq5 = st.text_input("5º Por que? (Causa Raiz)")
            st.markdown("<div class='section-header'>Plano de Ação Semanal Bloqueante</div>", unsafe_allow_html=True)
            sa_oque = st.text_area("O quê (Plano de Ação)")
            sa_quem = st.text_input("Responsável (Quem)")
            sa_quando = st.text_input("Prazo Final (Quando)")
            sa_status = st.selectbox("Status Operacional", ["Pendente", "Em Andamento", "Resolvido"])
            submit_sem = st.form_submit_button("💾 REGISTRAR ANÁLISE SEMANAL NO BANCO")
            
            if submit_sem:
                if not pior_parada_sem or not spq5:
                    st.error("Campos essenciais como Pior Parada e Causa Raiz devem ser informados.")
                else:
                    conn = sqlite3.connect('reportes_turno.db')
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO analises_semanais (data_registro, turno, maquina, pior_parada, pq1, pq2, pq3, pq4, pq5, causa_raiz, plano_acao, prazo, responsavel, status)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (str(semana_ref), turno_sem, maq_sem, pior_parada_sem, spq1, spq2, spq3, spq4, spq5, spq5, sa_oque, sa_quando, sa_quem, sa_status))
                    conn.commit(); conn.close()
                    st.success("🎉 Análise semanal gravada de forma definitiva!")

    # =========================================================
    # ABA: ACOMPANHAMENTO ANÁLISES SEMANAIS
    # =========================================================
    elif menu == "📋 ACOMP. ANÁLISES SEMANAIS":
        st.markdown("## 📋 Acompanhamento Técnico — Análises dos Operadores")
        conn = sqlite3.connect('reportes_turno.db')
        df_db_sem = pd.read_sql_query("SELECT * FROM analises_semanais ORDER BY data_registro DESC", conn)
        conn.close()
        
        if df_db_sem.empty:
            st.info("Nenhuma análise semanal encontrada no banco de dados local.")
        else:
            col_fs1, col_fs2 = st.columns(2)
            with col_fs1: status_f = st.multiselect("Filtrar Status", df_db_sem['status'].unique(), default=df_db_sem['status'].unique(), key='sf1')
            with col_fs2: turno_f = st.multiselect("Filtrar Turno", df_db_sem['turno'].unique(), default=df_db_sem['turno'].unique(), key='tf1')
            df_f_sem = df_db_sem[(df_db_sem['status'].isin(status_f)) & (df_db_sem['turno'].isin(turno_f))]
            
            def colorir_linhas_por_status(row):
                if row['status'] == 'Pendente': return ['background-color: #fee2e2; color: #b91c1c; font-weight: 600'] * len(row)
                elif row['status'] == 'Em Andamento': return ['background-color: #fef3c7; color: #d97706; font-weight: 600'] * len(row)
                elif row['status'] == 'Resolvido': return ['background-color: #dcfce7; color: #15803d; font-weight: 600'] * len(row)
                return [''] * len(row)

            cols_ex = ['id', 'data_registro', 'turno', 'maquina', 'pior_parada', 'responsavel', 'prazo', 'status']
            st.dataframe(df_f_sem[cols_ex].style.apply(colorir_linhas_por_status, axis=1), use_container_width=True)
            
            st.markdown("<div class='section-header'>✏️ Central de Gerenciamento da Análise Semanal</div>", unsafe_allow_html=True)
            id_sel_sem = st.number_input("Digite o ID da Análise Semanal para gerenciar:", min_value=1, step=1, key='id_num_sem')
            
            if id_sel_sem in df_db_sem['id'].values:
                if id_sel_sem != st.session_state['id_atual_semanal']:
                    st.session_state['id_atual_semanal'] = id_sel_sem
                    st.session_state['mostrar_edicao_semanal'] = False
                    
                if not st.session_state['mostrar_edicao_semanal']:
                    if st.button("🔍 ABRIR PAINEL DE EDIÇÃO DA ANÁLISE"):
                        st.session_state['mostrar_edicao_semanal'] = True
                        st.rerun()
                else:
                    if st.button("🔼 MINIMIZAR / FECHAR PAINEL DE EDIÇÃO"):
                        st.session_state['mostrar_edicao_semanal'] = False
                        st.rerun()
                        
                if st.session_state['mostrar_edicao_semanal']:
                    row_s = df_db_sem[df_db_sem['id'] == id_sel_sem].iloc[0]
                    esc1, esc2, esc3 = st.columns(3)
                    with esc1: es_pior = st.text_input("Editar Pior Parada", value=str(row_s['pior_parada']))
                    with esc2: es_status = st.selectbox("Editar Status", ["Pendente", "Em Andamento", "Resolvido"], index=["Pendente", "Em Andamento", "Resolvido"].index(row_s['status']), key='status_ed_sem')
                    with esc3: es_maq = st.text_input("Editar Máquina", value=str(row_s['maquina'])).upper()
                    
                    st.write("**Editar os 5 Porquês:**")
                    ep1 = st.text_input("1º Por que?", value=str(row_s['pq1']), key='ep1')
                    ep2 = st.text_input("2º Por que?", value=str(row_s['pq2']), key='ep2')
                    ep3 = st.text_input("3º Por que?", value=str(row_s['pq3']), key='ep3')
                    ep4 = st.text_input("4º Por que?", value=str(row_s['pq4']), key='ep4')
                    ep5 = st.text_input("5º Por que?", value=str(row_s['pq5']), key='ep5')
                    
                    st.write("**Editar Plano:**")
                    e_oque = st.text_area("O quê (Plano)", value=str(row_s['plano_acao']))
                    e_quem = st.text_input("Quem (Responsável)", value=str(row_s['responsavel']))
                    e_quando = st.text_input("Quando (Prazo)", value=str(row_s['prazo']))
                    
                    st.markdown("---")
                    btn_col1, btn_col2 = st.columns(2)
                    with btn_col1:
                        if st.button("💾 SALVAR ATUALIZAÇÃO SEMANAL", use_container_width=True):
                            conn = sqlite3.connect('reportes_turno.db')
                            cursor = conn.cursor()
                            cursor.execute("""
                                UPDATE analises_semanais SET pior_parada=?, status=?, maquina=?, pq1=?, pq2=?, pq3=?, pq4=?, pq5=?, causa_raiz=?, plano_acao=?, responsavel=?, prazo=? WHERE id=?
                            """, (es_pior, es_status, es_maq, ep1, ep2, ep3, ep4, ep5, ep5, e_oque, e_quem, e_quando, int(id_sel_sem)))
                            conn.commit(); conn.close()
                            st.session_state['mostrar_edicao_semanal'] = False
                            st.success("🎉 Dados atualizados!")
                            st.rerun()
                    with btn_col2:
                        if st.button("❌ DELETAR ANÁLISE SEMANAL", type="primary", use_container_width=True):
                            conn = sqlite3.connect('reportes_turno.db')
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM analises_semanais WHERE id = ?", (int(id_sel_sem),))
                            conn.commit(); conn.close()
                            st.session_state['mostrar_edicao_semanal'] = False
                            st.success("Deletado do sistema!")
                            st.rerun()

    # =========================================================
    # ABA: APRESENTAÇÃO SEMANAL
    # =========================================================
    elif menu == "📊 APRESENTAÇÃO SEMANAL":
        st.markdown("<h2 style='text-align:center;'>📊 Reunião Geral de Fechamento & Apresentação Semanal</h2>", unsafe_allow_html=True)
        st.subheader("⚙️ Selecione os Parâmetros da Apresentação")
        ap_c1, ap_c2, ap_c3 = st.columns(3)
        with ap_c1: maq_ap = st.selectbox("Máquina em Análise", sorted(df_order['Máquina'].unique()), key='maq_ap')
        with ap_c2:
            turno_ap = st.selectbox("Turno", ["T1", "T2", "T3"], key='turno_ap')
            turno_lista = [turno_ap[-1]]
        with ap_c3: periodo_ap = st.date_input("Período Semana", [df_order['Data'].max() - timedelta(days=7), df_order['Data'].max()], key='per_ap')
            
        df_ap_bruto = df_order[(df_order['Data'].dt.date >= periodo_ap[0]) & (df_order['Data'].dt.date <= periodo_ap[1]) & (df_order['Turno'].isin(turno_lista))]
        df_ap_maq = df_ap_bruto[df_ap_bruto['Máquina'] == maq_ap]
        df_ap_stops = df_stops[(df_stops['Data'].dt.date >= periodo_ap[0]) & (df_stops['Data'].dt.date <= periodo_ap[1]) & (df_stops['Máquina'] == maq_ap) & (df_stops['Turno'].isin(turno_lista))]
        
        st.markdown(f"""<div style="background-color:#f1f5f9; padding:15px; border-radius:10px; border-left:6px solid #10b981; margin-bottom:15px;">
            <h3 style='margin:0; color:#0f172a;'>EXIBIÇÃO INTEGRADA — MÁQUINA {maq_ap} (TURNO {turno_ap})</h3>
            <p style='margin:0; color:#64748b;'>Análise estatística cruzada com o banco de dados vinculada estritamente ao período selecionado.</p></div>""", unsafe_allow_html=True)
            
        c_kpi1, c_kpi2, c_kpi3 = st.columns(3)
        mov_sem = (df_ap_maq["Run Time"].sum() / df_ap_maq["Horário Padrão"].replace(0,1).sum() * 100)
        loss_sem = ((df_ap_maq["Machine Counter"].sum() - df_ap_maq["Peças Estoque - Ajuste"].sum()) / df_ap_maq["Machine Counter"].replace(0,1).sum() * 100)
        pecas_sem = df_ap_maq["Peças Estoque - Ajuste"].sum()
        
        with c_kpi1: st.plotly_chart(mini_gauge("Movimentação Semanal", mov_sem, "#10b981", 85, 140), use_container_width=True)
        with c_kpi2: st.plotly_chart(mini_gauge("Loss Semanal", loss_sem, "#e74c3c", 5, 140), use_container_width=True)
        with c_kpi3: st.markdown(f'<div class="metric-card" style="height:110px;"><div class="metric-title">Volume Realizado Semanal</div><div class="metric-value" style="font-size:1.8rem; margin-top:10px;">{fmt(pecas_sem)}</div></div>', unsafe_allow_html=True)
        
        conn = sqlite3.connect('reportes_turno.db')
        df_query_db = pd.read_sql_query("""
            SELECT * FROM analises_semanais 
            WHERE maquina = ? AND turno = ? AND data_registro >= ? AND data_registro <= ?
            ORDER BY data_registro DESC LIMIT 1
        """, conn, params=(maq_ap, turno_ap, str(periodo_ap[0]), str(periodo_ap[1])))
        conn.close()
        
        col_la, col_lb = st.columns([2, 1])
        with col_la:
            st.markdown("🏆 **Ranking de Eficiência Semanal**")
            rk_sem = df_ap_bruto.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
            rk_sem['Mov %'] = (rk_sem['Run Time']/rk_sem['Horário Padrão'].replace(0,1)*100).round(1)
            rk_sem = rk_sem.sort_values('Mov %', ascending=False).reset_index(drop=True)
            rk_sem.index += 1
            st.dataframe(rk_sem, use_container_width=True)
            
        with col_lb:
            st.markdown("🛑 **Top Ofensores (Gráfico do Arquivo)**")
            stop_data_ap = df_ap_stops.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(5)
            if not stop_data_ap.empty:
                st.plotly_chart(px.bar(stop_data_ap, orientation='h', color_discrete_sequence=['#10b981']).update_layout(height=230, paper_bgcolor='white', plot_bgcolor='white', margin=dict(l=0,r=0,t=0,b=0)), use_container_width=True)
            else:
                st.write("Sem registros de falhas mecânicas no período.")

        st.markdown("<div class='section-header'>Análise Causa Raiz Realizada pelos Operadores (Puxada do Banco de Dados)</div>", unsafe_allow_html=True)
        if df_query_db.empty:
            st.warning(f"⚠️ Nenhuma análise técnica foi cadastrada no banco de dados para a Máquina {maq_ap} no turno {turno_ap} para o período de {periodo_ap[0].strftime('%d/%m')} até {periodo_ap[1].strftime('%d/%m/%Y')}.")
        else:
            dados_db_pior = df_query_db.iloc[0]
            st.markdown(f"""
                <div class="five-why-box">
                    <div style="font-size:1.1rem; font-weight:700; color:#059669; margin-bottom:10px;">
                        DIAGRAMA DE CAUSA RAIZ PREENCHIDO — OFENSOR: <span style="color:#e11d48;">{dados_db_pior['pior_parada']}</span>
                    </div>
                    <div class="five-why-line"><b>1º Por que?</b> {dados_db_pior['pq1']}</div>
                    <div class="five-why-line"><b>2º Por que?</b> {dados_db_pior['pq2']}</div>
                    <div class="five-why-line"><b>3º Por que?</b> {dados_db_pior['pq3']}</div>
                    <div class="five-why-line"><b>4º Por que?</b> {dados_db_pior['pq4']}</div>
                    <div class="five-why-line"><b>5º Por que? (Causa Raiz)</b> <span style="color:#b91c1c; font-weight:600;">{dados_db_pior['pq5']}</span></div>
                    <br>
                    <div style="display:flex; gap:10px; margin-top:5px;">
                        <div style="flex:1; border:1px solid #cbd5e1; background-color:#f8fafc; padding:12px; border-radius:5px;">
                            <b>CAUSA RAIZ CONSOLIDADA:</b><br><span style="color:#334155;">{dados_db_pior['causa_raiz']}</span>
                        </div>
                        <div style="flex:2; border:1px solid #cbd5e1; background-color:#f8fafc; padding:12px; border-radius:5px;">
                            <b>PLANO DE AÇÃO BLOQUEANTE / IMEDIATO:</b><br><span style="color:#334155;">{dados_db_pior['plano_acao']}</span><br>
                            <small style='color:#64748b;'><b>Quem:</b> {dados_db_pior['responsavel']} | <b>Quando:</b> {dados_db_pior['prazo']} | <b>Status:</b> {dados_db_pior['status']}</small>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

    # =========================================================
    # ABA: 📋 NIPPO COORDENADORES (COM IMPLEMENTAÇÃO DE AUTO-RESET)
    # =========================================================
    elif menu == "📋 NIPPO COORDENADORES":
        st.markdown("## 📋 Nippo Coordenadores — Troca de Turno Operacional")
        
        aba_lancar, aba_consultar = st.tabs(["📝 Lançar Fechamento", "🔍 Histórico por Máquina"])
        
        with aba_lancar:
            # FIX: Todos os campos abaixo recebem uma 'key' indexada pelo contador dinâmico. 
            # Quando salvamos com sucesso, o contador sobe e o Streamlit reconstrói os campos vazios.
            versao_chave = st.session_state['contador_nippo']
            
            st.subheader("Informações Gerais do Turno")
            col_n1, col_n2 = st.columns(2)
            with col_n1:
                data_nippo = st.date_input("Data do Nippo", date.today(), key=f"date_np_{versao_chave}")
                coordenador_nippo = st.text_input("Nome do Coordenador", placeholder="Ex: DANILO", key=f"coord_np_{versao_chave}").upper()
            with col_n2:
                turno_nippo = st.selectbox("Selecione o Turno do Nippo", ["1º Turno", "2º Turno", "3º Turno"], index=2, key=f"turno_np_{versao_chave}")
                tecnico_nippo = st.text_input("Nome do Técnico Responsável", placeholder="Ex: KANIGIA", key=f"tec_np_{versao_chave}").upper()
            
            st.markdown("<div class='section-header'>Lançamento Individual por Máquina (M1 a M7)</div>", unsafe_allow_html=True)
            
            maquinas_lista = [f"M{i}" for i in range(1, 8)]
            mapa_inputs_maquinas = {}
            
            for m_item in maquinas_lista:
                with st.expander(f"⚙️ Reporte de Campo — Máquina: {m_item}", expanded=True):
                    col_b1, col_b2, col_b3 = st.columns([2, 1, 1])
                    
                    with col_b1:
                        txt_compartilhar = st.text_area(f"Itens a compartilhar / Ocorrências ({m_item})", key=f"txt_nippo_{m_item}_{versao_chave}", height=90)
                    with col_b2:
                        sku_maq = st.text_input(f"SKU Atual ({m_item})", key=f"sku_nippo_{m_item}_{versao_chave}").upper()
                        prod_maq = st.number_input(f"Produtividade % ({m_item})", min_value=0.0, max_value=100.0, step=0.1, key=f"prod_nippo_{m_item}_{versao_chave}")
                        loss_maq = st.number_input(f"Loss % ({m_item})", min_value=0.0, max_value=100.0, step=0.1, key=f"loss_nippo_{m_item}_{versao_chave}")
                    with col_b3:
                        pal_ini_maq = st.text_input(f"Palete Inicial ({m_item})", key=f"pal_ini_nippo_{m_item}_{versao_chave}").upper()
                        pal_fim_maq = st.text_input(f"Palete Final ({m_item})", key=f"pal_fim_nippo_{m_item}_{versao_chave}").upper()
                        tot_ordem_maq = st.number_input(f"Total da Ordem ({m_item})", min_value=0, step=1, key=f"tot_nippo_{m_item}_{versao_chave}")
                        
                    mapa_inputs_maquinas[m_item] = {
                        "itens": txt_compartilhar, "sku": sku_maq, "prod": prod_maq,
                        "loss": loss_maq, "pal_ini": pal_ini_maq, "pal_fim": pal_fim_maq, "tot": tot_ordem_maq
                    }
            
            # Executa o salvamento e realiza o auto-reset via incremento de estado
            if st.button("💾 GRAVAR REPORTE NIPPO NO BANCO", type="primary", use_container_width=True):
                if not coordenador_nippo or not tecnico_nippo:
                    st.error("Não é possível salvar. Os campos Coordenador e Técnico são obrigatórios.")
                else:
                    conn = sqlite3.connect('reportes_turno.db')
                    cursor = conn.cursor()
                    try:
                        for m_item, dados in mapa_inputs_maquinas.items():
                            cursor.execute("""
                                INSERT INTO nippo_coordenadores 
                                (data, turno, coordenador, tecnico, maquina, itens_compartilhar, produtividade, loss, sku, palete_inicial, palete_final, total_ordem)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                str(data_nippo), turno_nippo, coordenador_nippo, tecnico_nippo, m_item,
                                dados["itens"], dados["prod"], dados["loss"], dados["sku"], dados["pal_ini"], dados["pal_fim"], int(dados["tot"])
                            ))
                        conn.commit()
                        
                        # MODIFICAÇÃO DE RESETS: Incrementa o contador para forçar o reset dos campos textuais
                        st.session_state['contador_nippo'] += 1
                        st.success(f"🎉 O Nippo completo do {turno_nippo} foi gravado e as caixas de texto foram limpas para o próximo uso!")
                        st.rerun()
                        
                    except Exception as error:
                        st.error(f"Falha operacional na gravação das linhas: {error}")
                    finally:
                        conn.close()
                        
        with aba_consultar:
            st.subheader("🔍 Filtros de Pesquisa Histórica")
            c_f1, c_f2, c_f3 = st.columns(3)
            with c_f1:
                query_data = st.date_input("Filtrar Data Específica", date.today(), key="q_data")
            with c_f2:
                query_turno = st.selectbox("Filtrar Turno Específico", ["Todos", "1º Turno", "2º Turno", "3º Turno"], index=0)
            with c_f3:
                query_maq = st.selectbox("Filtrar Máquina Foco", ["Todas", "M1", "M2", "M3", "M4", "M5", "M6", "M7"], index=0)
                
            conn = sqlite3.connect('reportes_turno.db')
            sql_txt = "SELECT data, turno, coordenador, tecnico, maquina, itens_compartilhar, sku, produtividade, loss, palete_inicial, palete_final, total_ordem FROM nippo_coordenadores WHERE data = ?"
            parametros_filtro = [str(query_data)]
            
            if query_turno != "Todos":
                sql_txt += " AND turno = ?"
                parametros_filtro.append(query_turno)
            if query_maq != "Todas":
                sql_txt += " AND maquina = ?"
                parametros_filtro.append(query_maq)
                
            df_nippo_res = pd.read_sql_query(sql_txt, conn, params=parametros_filtro)
            conn.close()
            
            if df_nippo_res.empty:
                st.warning(f"Nenhum diário Nippo encontrado para o dia {query_data.strftime('%d/%m/%Y')}.")
            else:
                st.dataframe(df_nippo_res, use_container_width=True)
                st.markdown("<div class='section-header'>¼ Detalhamento de Itens Compartilhados no Turno</div>", unsafe_allow_html=True)
                for _, linha in df_nippo_res.iterrows():
                    if str(linha['itens_compartilhar']).strip():
                        st.markdown(f"🔹 **{linha['maquina']} — SKU: {linha['sku']}** (Turno: {linha['turno']} | Coordenador: {linha['coordenador']} | Técnico: {linha['tecnico']})")
                        st.info(linha['itens_compartilhar'])
else:
    st.info("💡 Por favor, carregue os arquivos Excel para iniciar o Analytics Hub.")
