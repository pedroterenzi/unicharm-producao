import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta, date
from sqlalchemy import create_engine, text
import io
import hashlib
import re

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# =========================================================
# BANCO DE DADOS NA NUVEM (POSTGRESQL - NEON.TECH)
# =========================================================
CONNECTION_STRING = "postgresql://neondb_owner:npg_obg1nxhT6GdK@ep-bitter-dream-aierzna8.c-4.us-east-1.aws.neon.tech/neondb?sslmode=require"

@st.cache_resource
def obter_engine():
    return create_engine(CONNECTION_STRING, pool_pre_ping=True)

# Função auxiliar para criptografar senhas (Segurança Industrial)
def hash_senha(senha):
    return hashlib.sha256(str.encode(senha)).hexdigest()

# Função para validar a força da senha de forma robusta
def validar_forca_senha(senha):
    erros = []
    if len(senha) < 8: erros.append("Mínimo de 8 caracteres")
    if not re.search(r"[A-Z]", senha): erros.append("Pelo menos 1 letra MAIÚSCULA")
    if not re.search(r"[0-9]", senha): erros.append("Pelo menos 1 número")
    if not re.search(r"[@#\$%\^&\*!\+=\-\[\]\{\}\(\)\|\:\;\,\.\?\/\~\`\_\\]", senha):
        erros.append("Pelo menos 1 caratere especial (@, #, $, %, etc.)")
    return erros

def init_db():
    engine = obter_engine()
    
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id SERIAL PRIMARY KEY, login TEXT UNIQUE, senha TEXT, cargo TEXT
            )
        """))
        
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS reportes (
                id SERIAL PRIMARY KEY, data_registro TEXT, turno TEXT, coordenador TEXT,
                ocorrencias TEXT, maq_analisada TEXT, problema TEXT,
                pq1 TEXT, pq2 TEXT, pq3 TEXT, pq4 TEXT, pq5 TEXT
            )
        """))
        
    with engine.begin() as conn:
        try: conn.execute(text("ALTER TABLE reportes ADD COLUMN duracao TEXT;"))
        except: pass

    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS acoes_reportes (
                id SERIAL PRIMARY KEY, reporte_id INTEGER REFERENCES reportes(id) ON DELETE CASCADE,
                oque TEXT, quem TEXT, quando TEXT, status TEXT
            )
        """))
        
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS analises_semanais (
                id SERIAL PRIMARY KEY, data_registro TEXT, turno TEXT, maquina TEXT,
                pior_parada TEXT, pq1 TEXT, pq2 TEXT, pq3 TEXT, pq4 TEXT, pq5 TEXT, causa_raiz TEXT
            )
        """))
        
    with engine.begin() as conn:
        try: conn.execute(text("ALTER TABLE analises_semanais ADD COLUMN duracao TEXT;"))
        except: pass

    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS acoes_semanais (
                id SERIAL PRIMARY KEY, analise_id INTEGER REFERENCES analises_semanais(id) ON DELETE CASCADE,
                oque TEXT, quem TEXT, quando TEXT, status TEXT
            )
        """))
        
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS nippo_coordenadores (
                id SERIAL PRIMARY KEY, data TEXT, turno TEXT, coordenador TEXT, tecnico TEXT, maquina TEXT,
                itens_compartilhar TEXT, produtividade REAL, loss REAL, sku TEXT,
                palete_inicial TEXT, palete_final TEXT, total_ordem INTEGER,
                data_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))

try:
    init_db()
except Exception as e:
    st.error(f"⚠️ Erro de Autenticação na Nuvem: {e}.")

# --- CONTROLE DE SESSÃO DO USUÁRIO ---
if 'autenticado' not in st.session_state: st.session_state['autenticado'] = False
if 'usuario_logado' not in st.session_state: st.session_state['usuario_logado'] = None
if 'cargo_logado' not in st.session_state: st.session_state['cargo_logado'] = None
if 'contador_cadastro' not in st.session_state: st.session_state['contador_cadastro'] = 0

# Chaves de reprocessamento dinâmico para limpar os formulários após o salvamento bem-sucedido
if 'chave_form_reporte' not in st.session_state: st.session_state['chave_form_reporte'] = 0
if 'chave_form_semanal' not in st.session_state: st.session_state['chave_form_semanal'] = 0

# --- INICIALIZAÇÃO DE ESTADOS DO STREAMLIT ---
if 'mostrar_edicao' not in st.session_state: st.session_state['mostrar_edicao'] = False
if 'id_atual' not in st.session_state: st.session_state['id_atual'] = 0
if 'mostrar_edicao_semanal' not in st.session_state: st.session_state['mostrar_edicao_semanal'] = False
if 'id_atual_semanal' not in st.session_state: st.session_state['id_atual_semanal'] = 0
if 'contador_nippo' not in st.session_state: st.session_state['contador_nippo'] = 0
if 'mostrar_edicao_nippo' not in st.session_state: st.session_state['mostrar_edicao_nippo'] = False
if 'chave_nippo_edicao' not in st.session_state: st.session_state['chave_nippo_edicao'] = ""

def fmt(valor):
    if pd.isna(valor) or valor is None: return "0"
    try: return f"{int(valor):,}".replace(",", ".")
    except: return str(valor)

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

# --- ESTILIZAÇÃO CSS PREMIUM (LIGHT MODE) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #ffffff; color: #1e293b; }
    [data-testid="stSidebar"] { background-color: #f8fafc; border-right: 1px solid #e2e8f0; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] { display: flex; flex-direction: column; gap: 2px; width: 100%; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #ffffff !important; border: 1px solid #e2e8f0 !important;
        padding: 12px 18px !important; border-radius: 10px !important; margin-bottom: 5px !important; 
        color: #475569 !important; cursor: pointer; font-weight: 500; font-size: 0.82rem;
        transition: all 0.2s ease-in-out; box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important;
        display: flex !important; align-items: center; justify-content: flex-start; width: 100% !important; box-sizing: border-box !important;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label:hover { background-color: #f1f5f9 !important; border-color: #cbd5e1 !important; color: #0f172a !important; transform: translateX(2px); }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important; color: #ffffff !important; 
        border: 1px solid #047857 !important; font-weight: 600; box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25) !important;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label icon { display: none !important; }
    .metric-container { display: flex; justify-content: space-between; gap: 8px; margin-bottom: 15px; }
    .metric-card {
        background: #f8fafc; padding: 12px; border-radius: 10px; text-align: center; border: 1px solid #e2e8f0;
        flex: 1; min-height: 80px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #64748b; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    .metric-value { color: #10b981; font-size: 1.3rem; font-weight: 900; line-height: 1; }
    .calendar-day-name { text-align: center; font-weight: 900; color: #10b981; font-size: 0.8rem; padding-bottom: 5px; }
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .day-card { background: #f8fafc; border-radius: 8px; padding: 10px; min-height: 95px; border: 1px solid #e2e8f0; }
    .day-number { font-size: 1rem; font-weight: 900; color: #1e293b; }
    .day-status { font-size: 0.75rem; font-weight: 600; color: #64748b; text-align: right; }
    .highlight-rank { background: #dcfce7 !important; color: #166534 !important; font-weight: 900; border-radius: 5px; padding: 5px; }
    .feedback-box { padding: 12px; border-radius: 8px; margin-bottom: 10px; text-align: center; font-weight: 700; font-size: 0.9rem; border: 1px solid #e2e8f0; }
    .five-why-box { border: 2px solid #1e293b; padding: 15px; background: #ffffff; color: #000; margin-top: 15px; }
    .five-why-line { border-bottom: 1px solid #000; padding: 10px 0; font-size: 0.9rem; }
    .section-header { background: #f1f5f9; padding: 10px; border-radius: 5px; color: #0f172a; font-weight: 800; text-transform: uppercase; margin-top: 20px; border-left: 5px solid #10b981; font-size: 0.9rem; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA DE ARQUIVOS EXCEL
@st.cache_data
def load_data(file_obj):
    df_order = pd.read_excel(file_obj, sheet_name="Result by order")
    df_stops = pd.read_excel(file_obj, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    nums = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed', 'Minutos', 'QTD']
    for df in [df_order, df_stops]:
        for col in nums:
            if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order = df_order.dropna(subset=['Data'])
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_stops = df_stops.dropna(subset=['Data'])
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    df_stops['Turno'] = df_stops['Turno'].fillna(0).astype(int).astype(str)

    df_order['Categoria'] = df_order['Máquina'].apply(lambda m: "BABY" if m in ['2', '3', '4', '5', '6'] else "ADULTO")
    return df_order, df_stops

@st.cache_data
def load_planner_metas_advanced(file, data_ref):
    try:
        xls = pd.ExcelFile(file)
        target = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        row_dates = df_raw.iloc[2, :].tolist()
        maq_lines = {'1': 6, '2': 28, '3': 47, '4': 58, '5': 77, '6': 96, '7': 113}
        plan_dia, plan_mes_acum = {}, {}
        
        idx_col_ref = None
        for i, d in enumerate(row_dates):
            if isinstance(d, (datetime, pd.Timestamp)) and d.date() == data_ref:
                idx_col_ref = i; break
        
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
                    if d.date() <= data_ref: m_mtd_total += val
        return plan_dia, plan_mes_acum, m_total_mes, m_mtd_total
    except:
        return {}, {}, 0, 0

# =========================================================
# TELA DE LOGIN E CADASTRO VIA ABAS
# =========================================================
if not st.session_state['autenticado']:
    st.markdown("<h1 style='text-align:center; color:#10b981; font-weight:900; margin-top:40px;'>🏭 INDUSTRIAL ANALYTICS HUB</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; color:#64748b;'>Selecione a opção desejada para entrar no ecossistema da produção.</p>", unsafe_allow_html=True)
    
    col_centro = st.columns([1, 2, 1])[1]
    with col_centro:
        aba_login, aba_cadastro = st.tabs(["🔐 ACESSAR SISTEMA", "📝 CRIAR NOVA CONTA"])
        with aba_login:
            st.markdown("<div class='section-header'>📋 IDENTIFICAÇÃO DE USUÁRIO</div>", unsafe_allow_html=True)
            login_user = st.text_input("Usuário / Login", key="login_u").strip().lower()
            senha_user = st.text_input("Senha", type="password", key="senha_u")
            if st.button("🔓 ENTRAR NO HUB", use_container_width=True, type="primary"):
                engine = obter_engine()
                df_auth = pd.read_sql_query(
                    text("SELECT login, cargo FROM usuarios WHERE login = :login AND senha = :senha"),
                    engine, params={"login": login_user, "senha": hash_senha(senha_user)}
                )
                if not df_auth.empty:
                    st.session_state['autenticado'] = True
                    st.session_state['usuario_logado'] = df_auth.iloc[0]['login']
                    st.session_state['cargo_logado'] = df_auth.iloc[0]['cargo']
                    st.rerun()
                else: st.error("Usuário ou senha incorretos.")
        with aba_cadastro:
            v_cad = st.session_state['contador_cadastro']
            st.markdown("<div class='section-header'>📝 FORMULÁRIO DE AUTO CADASTRO</div>", unsafe_allow_html=True)
            cad_user = st.text_input("Defina seu Login", key=f"cad_u_{v_cad}").strip().lower()
            cad_senha = st.text_input("Defina sua Senha (Mínimo 8 caracteres)", type="password", key=f"cad_s_{v_cad}")
            cad_conf_senha = st.text_input("Confirme sua Senha", type="password", key=f"cad_cs_{v_cad}")
            cad_cargo = st.selectbox("Selecione seu Cargo", ["Gerente", "Coordenador", "Analista", "Técnico de Produção", "Operador", "Menor Aprendiz", "Assistente"], key=f"cad_c_{v_cad}")
            lista_erros_senha = validar_forca_senha(cad_senha) if cad_senha else []
            if cad_senha:
                if len(lista_erros_senha) > 0:
                    st.markdown("##### 🚨 Requisitos de Senha Pendentes:")
                    for erro in lista_erros_senha: st.markdown(f"<span style='color:#ef4444; font-size:0.9rem;'>✖ {erro}</span>", unsafe_allow_html=True)
                else: st.markdown("<span style='color:#10b981; font-weight:600;'>✔ Estrutura de Senha Forte Detectada!</span>", unsafe_allow_html=True)
            if st.button("💾 REGISTRAR MEU USUÁRIO", use_container_width=True, disabled=len(lista_erros_senha) > 0):
                if cad_senha != cad_conf_senha: st.error("As senhas não conferem.")
                else:
                    engine = obter_engine()
                    try:
                        with engine.begin() as conn:
                            conn.execute(text("INSERT INTO usuarios (login, senha, cargo) VALUES (:login, :senha, :cargo)"), {"login": cad_user, "senha": hash_senha(cad_senha), "cargo": cad_cargo})
                        st.session_state['contador_cadastro'] += 1
                        st.success("🎉 Cadastro realizado! Utilize o painel de login.")
                        st.rerun()
                    except: st.error("Usuário já existe.")

# =========================================================
# ECOSSISTEMA AUTENTICADO (MENU FILTRADO)
# =========================================================
else:
    cargo = st.session_state['cargo_logado']
    todas_abas = [
        "📋 REPORTE DIÁRIO", "📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", 
        "📋 ANÁLISE SEMANAL", "📝 LANÇAR REPORTE", "📊 ACOMPANHAMENTO", 
        "📝 LANÇAR ANÁLISE SEMANAL", "📋 ACOMP. ANÁLISES SEMANAIS", "📋 PAINEL UNIFICADO DE AÇÕES", 
        "📋 RELATÓRIO CONSOLIDADO", "📋 NIPPO COORDENADORES"
    ]
    if cargo == "Operador": abas_permitidas = ["📝 LANÇAR ANÁLISE SEMANAL", "📋 PAINEL UNIFICADO DE AÇÕES"]
    elif cargo in ["Menor Aprendiz", "Assistente"]: abas_permitidas = [a for a in todas_abas if a not in ["📊 ACOMPANHAMENTO", "📋 ACOMP. ANÁLISES SEMANAIS", "📋 NIPPO COORDENADORES", "📋 RELATÓRIO CONSOLIDADO"]]
    else: abas_permitidas = todas_abas

    with st.sidebar:
        st.markdown(f"👤 <b>{st.session_state['usuario_logado'].upper()}</b> ({cargo.upper()})", unsafe_allow_html=True)
        if st.button("🚪 Sair / Desconectar"):
            st.session_state['autenticado'] = False; st.rerun()
        st.markdown("---")
        uploaded_file = st.file_uploader("📂 Carregar Excel Produção (.xlsm)", type=["xlsm"])
        up_datas = st.file_uploader("📂 Carregar Excel DATAS (.xlsx)", type=["xlsx"])
        st.markdown("---")
        if uploaded_file: menu = st.radio("NAVEGAÇÃO", abas_permitidas)

    if uploaded_file:
        df_order, df_stops = load_data(file_obj=uploaded_file)

        # =========================================================
        # ABA: REPORTE DIÁRIO (MANTIDA ORIGINAL E COMPLETA)
        # =========================================================
        if menu == "📋 REPORTE DIÁRIO":
            st.subheader("⚙️ Filtros da Página")
            col_f1, col_f2 = st.columns(2)
            with col_f1: data_ref_reporte = st.date_input("Data de Referência", df_order['Data'].max().date())
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
        # ABA: PERFORMANCE (MANTIDA ORIGINAL E COMPLETA)
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
        # ABA: TOP 10 PARADAS (MANTIDA ORIGINAL E COMPLETA)
        # =========================================================
        elif menu == "🛑 TOP 10 PARADAS":
            st.sidebar.subheader("Filtros Paradas")
            f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
            f_maq_s = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()), key='m2')
            f_turno_s = st.sidebar.multiselect("Turnos", sorted(df_stops['Turno'].unique()), default=sorted(df_stops['Turno'].unique()), key='ts2')
            
            df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1]) & (df_stops['Máquina'].isin(f_maq_s)) & (df_stops['Turno'].isin(f_turno_s))]
            str_maquinas_s = ", ".join(f_maq_s) if f_maq_s else "Nenhuma"
            str_turnos_s = ", ".join(f_turno_s) if f_turno_s else "Nenhum"
            st.markdown(f"## 🛑 Análise de Paradas — Máquina(s): {str_maquinas_s} | Turno(s): {str_turnos_s}")
            
            st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values().tail(10), orientation='h', title="Minutos Totais", color_discrete_sequence=['#f43f5e']).update_layout(paper_bgcolor='white', plot_bgcolor='white', font={'color':'black'}), use_container_width=True)
            st.plotly_chart(px.bar(df_s_f.groupby('Problema')['QTD'].sum().sort_values().tail(10), orientation='h', title="Frequência (Qtd)", color_discrete_sequence=['#3b82f6']).update_layout(paper_bgcolor='white', plot_bgcolor='white', font={'color':'black'}), use_container_width=True)

        # =========================================================
        # ABA: CALENDÁRIO (MANTIDA ORIGINAL E COMPLETA)
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
        # ABA: ANÁLISE SEMANAL (MANTIDA ORIGINAL E COMPLETA)
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

            m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].replace(0,1).sum()*100) if not df_b.empty else 0
            l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].replace(0,1).sum()*100) if not df_b.empty else 0
            pecas_v = df_b["Peças Estoque - Ajuste"].sum() if not df_b.empty else 0

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
                if posicao <= 2: msg, col = ("🏆 Excelente performance.", "#dcfce7")
                else: msg, col = ("🚀 Foco na melhoria para subir o ranking semanal!", "#fee2e2")
                st.markdown(f'<div class="feedback-box" style="background:{col}; color:black; border-left:5px solid #10b981;">{msg}</div>', unsafe_allow_html=True)

            col_g, col_r = st.columns([2, 1])
            with col_g:
                st.markdown("🛑 *Impacto das Paradas (Piores 5)*")
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
        # 📝 LANÇAR REPORTE (SENHA E BLOQUEIO DE ENTER ACIDENTAL ATIVADOS)
        # =========================================================
        elif menu == "📝 LANÇAR REPORTE":
            v_rep = st.session_state['chave_form_reporte']
            st.markdown("## 📝 Lançar Reporte de Turno")
            with st.form("form_reporte", clear_on_submit=False):
                c1, c2, c3 = st.columns(3)
                with c1: data_rep = st.date_input("Data do Turno", datetime.now().date(), key=f"rep_data_{v_rep}")
                with c2: turno_rep = st.selectbox("Turno", ["T1", "T2", "T3"], key=f"rep_turno_{v_rep}")
                with c3: coord_rep = st.text_input("Coordenador Responsável", key=f"rep_coord_{v_rep}").upper()
                
                st.markdown("<div class='section-header'>1. Principais Ocorrências / Paradas do Turno</div>", unsafe_allow_html=True)
                txt_ocorrencias = st.text_area("Descreva as ocorrências", height=120, key=f"rep_oc_{v_rep}")
                
                st.markdown("<div class='section-header'>2. Análise de Causa Raiz (Top Ofensor)</div>", unsafe_allow_html=True)
                cc1, cc2, cc3 = st.columns([2, 2, 1])
                with cc1: maq_an = st.text_input("Máquina Analisada", key=f"rep_maq_{v_rep}").upper()
                with cc2: prob_an = st.text_input("Problema Foco", key=f"rep_prob_{v_rep}")
                with cc3: dur_an = st.text_input("Duração Parada (Ex: 45min)", key=f"rep_dur_{v_rep}")
                
                p1 = st.text_input("Por que 1?", key=f"rep_p1_{v_rep}")
                p2 = st.text_input("Por que 2?", key=f"rep_p2_{v_rep}")
                p3 = st.text_input("Por que 3?", key=f"rep_p3_{v_rep}")
                p4 = st.text_input("Por que 4?", key=f"rep_p4_{v_rep}")
                p5 = st.text_input("Por que 5? (Causa Raiz)", key=f"rep_p5_{v_rep}")
                
                st.markdown("<div class='section-header'>3. Plano de Ação Imediato Dinâmico</div>", unsafe_allow_html=True)
                num_acoes = st.number_input("Quantas ações esse problema gerou?", min_value=1, max_value=10, value=1, step=1, key=f"rep_num_{v_rep}")
                
                lista_reporte_inputs = []
                for idx in range(int(num_acoes)):
                    st.markdown(f"**Ação #{idx+1}**")
                    ac_c1, ac_c2, ac_c3, ac_c4 = st.columns([3, 2, 2, 2])
                    with ac_c1: inp_oque = st.text_input(f"O quê (Ação)", key=f"r_oq_{idx}_{v_rep}")
                    with ac_c2: inp_quem = st.text_input(f"Quem (Responsável)", key=f"r_qm_{idx}_{v_rep}").upper()
                    with ac_c3: inp_quando = st.text_input(f"Quando (Prazo)", key=f"r_qn_{idx}_{v_rep}")
                    with ac_c4: inp_status = st.selectbox(f"Status Inicial", ["Pendente", "Em Andamento", "Resolvido"], key=f"r_st_{idx}_{v_rep}")
                    lista_reporte_inputs.append({"oque": inp_oque, "quem": inp_quem, "quando": inp_quando, "status": inp_status})
                
                submit = st.form_submit_button("💾 SALVAR REPORTE NO BANCO DE DADOS")
                if submit:
                    vazios = [item for item in lista_reporte_inputs if item["oque"].strip() == "" or item["quem"].strip() == ""]
                    if not coord_rep or not prob_an or not dur_an:
                        st.error("⚠️ Preenchimento Obrigatório: Informe o Coordenador, Problema e Duração!")
                    elif len(vazios) > 0:
                        st.error("⚠️ Envio Bloqueado: Garanta que todas as caixas de Ação e Responsável geradas estejam preenchidas.")
                    else:
                        engine = obter_engine()
                        with engine.begin() as conn:
                            result = conn.execute(text("""
                                INSERT INTO reportes (data_registro, turno, coordenador, ocorrencias, maq_analisada, problema, duracao, pq1, pq2, pq3, pq4, pq5) 
                                VALUES (:data, :turno, :coord, :ocorrencias, :maq, :prob, :dur, :p1, :p2, :p3, :p4, :p5) RETURNING id
                            """), {"data": str(data_rep), "turno": turno_rep, "coord": coord_rep, "ocorrencias": txt_ocorrencias, "maq": maq_an, "prob": prob_an, "dur": dur_an, "p1": p1, "p2": p2, "p3": p3, "p4": p4, "p5": p5})
                            
                            reporte_id = result.fetchone()[0]
                            for item in lista_reporte_inputs:
                                conn.execute(text("INSERT INTO acoes_reportes (reporte_id, oque, quem, quando, status) VALUES (:reporte_id, :oque, :quem, :quando, :status)"), {"reporte_id": reporte_id, "oque": item["oque"], "quem": item["quem"], "quando": item["quando"], "status": item["status"]})
                        
                        st.session_state['chave_form_reporte'] += 1
                        st.success(f"🎉 Registro ID #{reporte_id} salvo com sucesso!"); st.rerun()

        # =========================================================
        # 📊 ABA: ACOMPANHAMENTO (MESA DE EDIÇÃO COMPLETA DE REPORTES DIÁRIOS)
        # =========================================================
        elif menu == "📊 ACOMPANHAMENTO":
            st.markdown("## 📊 Mesa de Edição — Reportes Diários")
            engine = obter_engine()
            df_db = pd.read_sql_query("SELECT id, data_registro, turno, coordenador, maq_analisada, problema, duracao FROM reportes ORDER BY id DESC", engine)
            
            if df_db.empty: st.info("Nenhum reporte estruturado encontrado na nuvem.")
            else:
                st.dataframe(df_db, use_container_width=True)
                id_selecionado = st.number_input("Digite o ID do reporte para editar dados de cabeçalho ou os 5 Porquês:", min_value=1, step=1)
                
                if id_selecionado in df_db['id'].values:
                    if id_selecionado != st.session_state['id_atual']:
                        st.session_state['id_atual'] = id_selecionado
                        st.session_state['mostrar_edicao'] = False
                        
                    if not st.session_state['mostrar_edicao']:
                        if st.button("🔍 CARREGAR DADOS COMPLETOS PARA EDIÇÃO"): st.session_state['mostrar_edicao'] = True; st.rerun()
                    else:
                        if st.button("🔼 MINIMIZAR FORMULÁRIO"): st.session_state['mostrar_edicao'] = False; st.rerun()
                            
                    if st.session_state['mostrar_edicao']:
                        row_sel = pd.read_sql_query(text("SELECT * FROM reportes WHERE id = :id"), engine, params={"id": int(id_selecionado)}).iloc[0]
                        
                        ec1, ec2, ec3 = st.columns(3)
                        with ec1: ed_coord = st.text_input("Coordenador", value=str(row_sel['coordenador'])).upper()
                        with ec2: ed_maq = st.text_input("Máquina Analisada", value=str(row_sel['maq_analisada'])).upper()
                        with ec3: ed_dur = st.text_input("Duração Parada", value=str(row_sel['duracao']))
                        
                        ed_oc = st.text_area("Principais Ocorrências do Turno", value=str(row_sel['ocorrencias']))
                        ed_prob = st.text_input("Problema Foco", value=str(row_sel['problema']))
                        
                        st.write("**Fluxo de Diagnóstico (5 Porquês):**")
                        ep1 = st.text_input("Por que 1?", value=str(row_sel['pq1']))
                        ep2 = st.text_input("Por que 2?", value=str(row_sel['pq2']))
                        ep3 = st.text_input("Por que 3?", value=str(row_sel['pq3']))
                        ep4 = st.text_input("Por que 4?", value=str(row_sel['pq4']))
                        ep5 = st.text_input("Por que 5 (Causa Raiz)?", value=str(row_sel['pq5']))
                        
                        st.markdown("#### Planos de Ação Relacionados a esse ID")
                        df_ac_edit = pd.read_sql_query(text("SELECT id, oque as \"Ação\", quem as \"Responsável\", quando as \"Prazo\", status as \"Status\" FROM acoes_reportes WHERE reporte_id = :id"), engine, params={"id": int(id_selecionado)})
                        tabela_ed_flow = st.data_editor(df_ac_edit, num_rows="dynamic", column_config={"Status": st.column_config.SelectboxColumn("Status", options=["Pendente", "Em Andamento", "Resolvido"], required=True)}, use_container_width=True, key=f"dt_rep_{id_selecionado}")
                        
                        b_c1, b_c2 = st.columns(2)
                        with b_c1:
                            if st.button("💾 SALVAR ALTERAÇÕES COMPLETAS"):
                                with engine.begin() as conn:
                                    conn.execute(text("UPDATE reportes SET coordenador=:c, ocorrencias=:oc, maq_analisada=:m, problema=:p, duracao=:d, pq1=:p1, pq2=:p2, pq3=:p3, pq4=:p4, pq5=:p5 WHERE id=:id"), {"c": ed_coord, "oc": ed_oc, "m": ed_maq, "p": ed_prob, "d": ed_dur, "p1": ep1, "p2": ep2, "p3": ep3, "p4": ep4, "p5": ep5, "id": int(id_selecionado)})
                                    conn.execute(text("DELETE FROM acoes_reportes WHERE reporte_id = :id"), {"id": int(id_selecionado)})
                                    for _, r in tabela_ed_flow.iterrows():
                                        if str(r["Ação"]).strip() != "":
                                            conn.execute(text("INSERT INTO acoes_reportes (reporte_id, oque, quem, quando, status) VALUES (:id, :oq, :qm, :qn, :st)"), {"id": int(id_selecionado), "oq": r["Ação"], "qm": str(r["Responsável"]).upper(), "qn": r["Prazo"], "st": r["Status"]})
                                st.session_state['mostrar_edicao'] = False; st.success("Sincronizado!"); st.rerun()
                        with b_c2:
                            if st.button("❌ APAGAR DEFINITIVAMENTE", type="primary"):
                                with engine.begin() as conn: conn.execute(text("DELETE FROM reportes WHERE id = :id"), {"id": int(id_selecionado)})
                                st.session_state['mostrar_edicao'] = False; st.rerun()

        # =========================================================
        # 📝 LANÇAR ANÁLISE SEMANAL (VALIDAÇÃO E LIMPEZA CONTRA ENTER)
        # =========================================================
        elif menu == "📝 LANÇAR ANÁLISE SEMANAL":
            v_sem = st.session_state['chave_form_semanal']
            st.markdown("## 📝 Formulário de Lançamento — Análise Semanal")
            with st.form("form_analise_semanal", clear_on_submit=False):
                s1, s2, s3 = st.columns(3)
                with s1: semana_ref = st.date_input("Semana de Referência", datetime.now().date(), key=f"s_dt_{v_sem}")
                with s2: turno_sem = st.selectbox("Turno Analisado", ["T1", "T2", "T3"], key=f"s_tr_{v_sem}")
                with s3: maq_sem = st.selectbox("Máquina Alvo", sorted(df_order['Máquina'].unique()), key=f"s_mq_{v_sem}")
                pior_parada_sem = st.text_input("Pior Parada Detectada", key=f"s_pr_{v_sem}")
                dur_sem = st.text_input("Duração Total Parada (Ex: 3h)", key=f"s_dr_{v_sem}")
                
                st.markdown("<div class='section-header'>Análise Causa Raiz — Método dos 5 Porquês</div>", unsafe_allow_html=True)
                spq1 = st.text_input("1º Por que?", key=f"s_p1_{v_sem}")
                spq2 = st.text_input("2º Por que?", key=f"s_p2_{v_sem}")
                spq3 = st.text_input("3º Por que?", key=f"s_p3_{v_sem}")
                spq4 = st.text_input("4º Por que?", key=f"s_p4_{v_sem}")
                spq5 = st.text_input("5º Por que? (Causa Raiz)", key=f"s_p5_{v_sem}")
                
                st.markdown("<div class='section-header'>Plano de Ações Semanais Combinadas</div>", unsafe_allow_html=True)
                num_acoes_sem = st.number_input("Quantas ações esse ofensor gerou para a semana?", min_value=1, max_value=10, value=1, step=1, key=f"s_nm_{v_sem}")
                
                lista_sem_inputs = []
                for idx in range(int(num_acoes_sem)):
                    st.markdown(f"**Ação Semanal #{idx+1}**")
                    sem_c1, sem_c2, sem_c3, sem_c4 = st.columns([3, 2, 2, 2])
                    with sem_c1: sem_oque = st.text_input(f"O quê (Ação Semanal)", key=f"s_oq_{idx}_{v_sem}")
                    with sem_c2: sem_quem = st.text_input(f"Quem (Responsável)", key=f"s_qm_{idx}_{v_sem}").upper()
                    with sem_c3: sem_quando = st.text_input(f"Quando (Prazo)", key=f"s_qn_{idx}_{v_sem}")
                    with sem_c4: sem_status = st.selectbox(f"Status", ["Pendente", "Em Andamento", "Resolvido"], key=f"s_st_{idx}_{v_sem}")
                    lista_sem_inputs.append({"oque": sem_oque, "quem": sem_quem, "quando": sem_quando, "status": sem_status})
                
                submit_sem = st.form_submit_button("💾 REGISTRAR ANÁLISE SEMANAL NO BANCO")
                if submit_sem:
                    vazios_sem = [item for item in lista_sem_inputs if item["oque"].strip() == "" or item["quem"].strip() == ""]
                    if not pior_parada_sem or not spq5 or not dur_sem:
                        st.error("⚠️ Preenchimento Obrigatório: Preencha o Ofensor, Causa Raiz e Duração!")
                    elif len(vazios_sem) > 0:
                        st.error("⚠️ Envio Bloqueado: Certifique-se de que todas as caixas de Ação Semanal geradas estejam preenchidas.")
                    else:
                        engine = obter_engine()
                        with engine.begin() as conn:
                            result_sem = conn.execute(text("INSERT INTO analises_semanais (data_registro, turno, maquina, pior_parada, duracao, pq1, pq2, pq3, pq4, pq5, causa_raiz) VALUES (:data, :turn, :maq, :pior, :dur, :p1, :p2, :p3, :p4, :p5, :causa) RETURNING id"), {"data": str(semana_ref), "turn": turno_sem, "maq": maq_sem, "pior": pior_parada_sem, "dur": dur_sem, "p1": spq1, "p2": spq2, "p3": spq3, "p4": spq4, "p5": spq5, "causa": spq5})
                            analise_id = result_sem.fetchone()[0]
                            for item in lista_sem_inputs:
                                conn.execute(text("INSERT INTO acoes_semanais (analise_id, oque, quem, quando, status) VALUES (:analise_id, :oque, :quem, :quando, :status)"), {"analise_id": analise_id, "oque": item["oque"], "quem": item["quem"], "quando": item["quando"], "status": item["status"]})
                        
                        st.session_state['chave_form_semanal'] += 1
                        st.success(f"🎉 Análise Semanal ID #{analise_id} gravada!"); st.rerun()

        # =========================================================
        # 📋 ACOMP. ANÁLISES SEMANAIS (MESA DE EDIÇÃO DE TÍTULOS E 5PQ)
        # =========================================================
        elif menu == "📋 ACOMP. ANÁLISES SEMANAIS":
            st.markdown("## 📋 Mesa de Edição — Análises Semanais")
            engine = obter_engine()
            df_as = pd.read_sql_query("SELECT id, data_registro, turno, maquina, pior_parada, duracao FROM analises_semanais ORDER BY id DESC", engine)
            if df_as.empty: st.info("Nenhuma análise semanal registrada no banco remoto.")
            else:
                st.dataframe(df_as, use_container_width=True)
                id_sel_sem = st.number_input("Digite o ID da Análise Semanal para editar dados ou os 5 Porquês:", min_value=1, step=1, key='id_num_sem')
                
                if id_sel_sem in df_as['id'].values:
                    if id_sel_sem != st.session_state['id_atual_semanal']:
                        st.session_state['id_atual_semanal'] = id_sel_sem
                        st.session_state['mostrar_edicao_semanal'] = False
                        
                    if not st.session_state['mostrar_edicao_semanal']:
                        if st.button("🔍 CARREGAR DADOS DA ANÁLISE"): st.session_state['mostrar_edicao_semanal'] = True; st.rerun()
                    else:
                        if st.button("🔼 MINIMIZAR FORMULÁRIO Técnico"): st.session_state['mostrar_edicao_semanal'] = False; st.rerun()
                            
                    if st.session_state['mostrar_edicao_semanal']:
                        row_s = pd.read_sql_query(text("SELECT * FROM analises_semanais WHERE id = :id"), engine, params={"id": int(id_sel_sem)}).iloc[0]
                        
                        esc1, esc2, esc3 = st.columns(3)
                        with esc1: es_pior = st.text_input("Ofensor Semanal", value=str(row_s['pior_parada']))
                        with esc2: es_maq = st.text_input("Máquina Alvo", value=str(row_s['maquina'])).upper()
                        with esc3: es_dur = st.text_input("Corrigir Duração Mapeada", value=str(row_s['duracao']))
                        
                        sep1 = st.text_input("1º Por que?", value=str(row_s['pq1']))
                        sep2 = st.text_input("2º Por que?", value=str(row_s['pq2']))
                        sep3 = st.text_input("3º Por que?", value=str(row_s['pq3']))
                        sep4 = st.text_input("4º Por que?", value=str(row_s['pq4']))
                        sep5 = st.text_input("5º Por que? (Causa Raiz)", value=str(row_s['pq5']))
                        
                        st.write("**Planos de Ação Relacionados**")
                        df_ac_sem_edit = pd.read_sql_query(text("SELECT id, oque as \"Ação Semanal\", quem as \"Responsável\", quando as \"Prazo\", status as \"Status\" FROM acoes_semanais WHERE analise_id = :id"), engine, params={"id": int(id_sel_sem)})
                        tabela_ed_sem_flow = st.data_editor(df_ac_sem_edit, num_rows="dynamic", column_config={"Status": st.column_config.SelectboxColumn("Status", options=["Pendente", "Em Andamento", "Resolvido"], required=True)}, use_container_width=True, key=f"dt_sem_{id_sel_sem}")
                        
                        btn_col1, btn_col2 = st.columns(2)
                        with btn_col1:
                            if st.button("💾 GRAVAR ALTERAÇÕES DA ANÁLISE"):
                                with engine.begin() as conn:
                                    conn.execute(text("UPDATE analises_semanais SET pior_parada=:p, maquina=:m, duracao=:d, pq1=:p1, pq2=:p2, pq3=:p3, pq4=:p4, pq5=:p5, causa_raiz=:c WHERE id=:id"), {"p": es_pior, "m": es_maq, "d": es_dur, "p1": sep1, "p2": sep2, "p3": sep3, "p4": sep4, "p5": sep5, "c": sep5, "id": int(id_sel_sem)})
                                    conn.execute(text("DELETE FROM acoes_semanais WHERE analise_id = :id"), {"id": int(id_sel_sem)})
                                    for _, r in tabela_ed_sem_flow.iterrows():
                                        if str(r["Ação Semanal"]).strip() != "":
                                            conn.execute(text("INSERT INTO acoes_semanais (analise_id, oque, quem, quando, status) VALUES (:id, :oq, :qm, :qn, :st)"), {"id": int(id_sel_sem), "oq": r["Ação Semanal"], "qm": str(r["Responsável"]).upper(), "qn": r["Prazo"], "st": r["Status"]})
                                st.session_state['mostrar_edicao_semanal'] = False; st.success("Atualizado!"); st.rerun()
                        with btn_col2:
                            if st.button("❌ APAGAR BLOCO SEMANAL", type="primary"):
                                with engine.begin() as conn: conn.execute(text("DELETE FROM analises_semanais WHERE id = :id"), {"id": int(id_sel_sem)})
                                st.session_state['mostrar_edicao_semanal'] = False; st.rerun()

        # =========================================================
        # 📋 ABA: CENTRAL E PAINEL UNIFICADO DE AÇÕES
        # =========================================================
        elif menu == "📋 PAINEL UNIFICADO DE AÇÕES":
            st.markdown("## 📋 Painel Unificado de Ações Industriais (Central de Cobrança)")
            st.caption("Esta tela compila em tempo real todas as ações (Reportes Diários + Análises Semanais) divididas por Status.")
            
            engine = obter_engine()
            df_ac_rep = pd.read_sql_query("""
                SELECT 'DIÁRIO' as "Origem", r.maq_analisada as "Máquina", r.problema as "Problema / Ofensor", ar.oque as "O que Fazer", ar.quem as "Responsável", ar.quando as "Prazo", ar.status 
                FROM acoes_reportes ar JOIN reportes r ON ar.reporte_id = r.id
            """, engine)
            df_ac_sem = pd.read_sql_query("""
                SELECT 'SEMANAL' as "Origem", ans.maquina as "Máquina", ans.pior_parada as "Problema / Ofensor", asm.oque as "O que Fazer", asm.quem as "Responsável", asm.quando as "Prazo", asm.status 
                FROM acoes_semanais asm JOIN analises_semanais ans ON asm.analise_id = ans.id
            """, engine)
            
            df_unificado = pd.concat([df_ac_rep, df_ac_sem], ignore_index=True)
            
            if df_unificado.empty: st.info("Nenhuma ação cadastrada no sistema.")
            else:
                status_cards = ["Pendente", "Em Andamento", "Resolvido"]
                cores_cards = ["#ef4444", "#f59e0b", "#10b981"]
                
                cols_cards = st.columns(3)
                for i, st_name in enumerate(status_cards):
                    total_st = len(df_unificado[df_unificado['status'] == st_name])
                    cols_cards[i].markdown(f"""
                    <div style="background-color:#f8fafc; border:1px solid #e2e8f0; border-top:5px solid {cores_cards[i]}; border-radius:10px; padding:15px; text-align:center; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <h4 style="margin:0; color:#475569; font-size:0.9rem; text-transform:uppercase;">AÇÕES {st_name.upper()}</h4>
                        <h2 style="margin:5px 0 0 0; color:{cores_cards[i]}; font-weight:900; font-size:2rem;">{total_st}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                aba_p, aba_and, aba_ok = st.tabs(["🔴 AÇÕES PENDENTES", "🟡 AÇÕES EM ANDAMENTO", "🟢 AÇÕES REALIZADAS"])
                colunas_exibicao = ["Origem", "Máquina", "Problema / Ofensor", "O que Fazer", "Responsável", "Prazo"]
                
                with aba_p:
                    df_p = df_unificado[df_unificado['status'] == "Pendente"]
                    if df_p.empty: st.success("🎉 Nenhuma ação pendente! Tudo em andamento ou resolvido.")
                    else: st.dataframe(df_p[colunas_exibicao], use_container_width=True)
                    
                with aba_and:
                    df_and = df_unificado[df_unificado['status'] == "Em Andamento"]
                    if df_and.empty: st.info("Nenhuma ação em andamento no momento.")
                    else: st.dataframe(df_and[colunas_exibicao], use_container_width=True)
                    
                with aba_ok:
                    df_ok = df_unificado[df_unificado['status'] == "Resolvido"]
                    if df_ok.empty: st.warning("Ainda não temos ações concluídas gravadas na nuvem.")
                    else: st.dataframe(df_ok[colunas_exibicao], use_container_width=True)

        # =========================================================
        # 📋 ABA: RELATÓRIO CONSOLIDADO DE OCORRÊNCIAS OPERACIONAIS
        # =========================================================
        elif menu == "📋 RELATÓRIO CONSOLIDADO":
            st.markdown("## 📋 Relatório Consolidado de Ocorrências Operacionais (Sumário Matinal)")
            data_matinal = st.date_input("Data de Análise dos Turnos", date.today() - timedelta(days=1))
            st.markdown(f"### 📑 Painel Geral Consolidado — Turnos de {data_matinal.strftime('%d/%m/%Y')}")
            
            engine = obter_engine()
            df_reportes_dia = pd.read_sql_query(text("SELECT * FROM reportes WHERE data_registro = :data ORDER BY turno ASC"), engine, params={"data": str(data_matinal)})
            
            if df_reportes_dia.empty:
                st.warning(f"⚠️ Nenhum reporte operacional foi lançado na nuvem para a data {data_matinal.strftime('%d/%m/%Y')}.")
            else:
                st.markdown("<div class='section-header'>📋 SUMÁRIO EXECUTIVO — PRINCIPAIS OCORRÊNCIAS DO DIA</div>", unsafe_allow_html=True)
                for _, r in df_reportes_dia.iterrows():
                    st.markdown(f"🔹 **Turno {r['turno']}** (Coordenador: {r['coordenador']})")
                    st.info(r['ocorrencias'] if r['ocorrencias'].strip() else "Nenhuma grande intercorrência anotada no livro de turno.")
                
                st.markdown("<div class='section-header'>🛑 ANÁLISE DOS TOP OFENSORES E DIAGNÓSTICO 5 PORQUÊS</div>", unsafe_allow_html=True)
                for _, r in df_reportes_dia.iterrows():
                    st.markdown(f"""
                    <div class="five-why-box">
                        <h4 style="color:#059669; margin:0;">⚙️ Turno {r['turno']} — Máquina: {r['maq_analisada']} | Ofensor: {r['problema']}</h4>
                        <h5 style="color:#e11d48; margin-top:5px; margin-bottom:15px;">⏱ DURAÇÃO DA PARADA: {r['duracao']}</h5>
                        <div class="five-why-line"><b>1º Por que?</b> {r['pq1']}</div>
                        <div class="five-why-line"><b>2º Por que?</b> {r['pq2']}</div>
                        <div class="five-why-line"><b>3º Por que?</b> {r['pq3']}</div>
                        <div class="five-why-line"><b>4º Por que?</b> {r['pq4']}</div>
                        <div class="five-why-line"><b>5º Por que? (Causa Raiz)</b> <span style="color:#b91c1c; font-weight:600;">{r['pq5']}</span></div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    df_ac_sem = pd.read_sql_query(text("SELECT oque as \"Ação Programada\", quem as \"Responsável\", quando as \"Prazo\", status as \"Status\" FROM acoes_reportes WHERE reporte_id = :id"), engine, params={"id": int(r['id'])})
                    if not df_ac_sem.empty:
                        st.caption("📋 **Matriz de Planos Bloqueantes Alocados:**")
                        st.dataframe(df_ac_sem, use_container_width=True)
                    else:
                        st.caption("ℹ️ Nenhum plano de ação cadastrado para este ofensor.")

        # =========================================================
        # 📋 NIPPO COORDENADORES
        # =========================================================
        elif menu == "📋 NIPPO COORDENADORES":
            st.markdown("## 📋 Nippo Coordenadores — Troca de Turno Operacional")
            aba_lancar, aba_consultar = st.tabs(["📝 Lançar Fechamento", "🔍 Histórico / Gerenciamento"])
            with aba_lancar:
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
                        with col_b1: txt_compartilhar = st.text_area(f"Itens a compartilhar / Ocorrências ({m_item})", key=f"txt_nippo_{m_item}_{versao_chave}", height=90)
                        with col_b2:
                            sku_maq = st.text_input(f"SKU Atual ({m_item})", key=f"sku_nippo_{m_item}_{versao_chave}").upper()
                            prod_maq = st.number_input(f"Produtividade % ({m_item})", min_value=0.0, max_value=100.0, step=0.1, key=f"prod_nippo_{m_item}_{versao_chave}")
                            loss_maq = st.number_input(f"Loss % ({m_item})", min_value=0.0, max_value=100.0, step=0.1, key=f"loss_nippo_{m_item}_{versao_chave}")
                        with col_b3:
                            pal_ini_maq = st.text_input(f"Palete Inicial ({m_item})", key=f"pal_ini_nippo_{m_item}_{versao_chave}").upper()
                            pal_fim_maq = st.text_input(f"Palete Final ({m_item})", key=f"pal_fim_nippo_{m_item}_{versao_chave}").upper()
                            tot_ordem_maq = st.number_input(f"Total da Ordem ({m_item})", min_value=0, step=1, key=f"tot_nippo_{m_item}_{versao_chave}")
                        mapa_inputs_maquinas[m_item] = {"itens": txt_compartilhar, "sku": sku_maq, "prod": prod_maq, "loss": loss_maq, "pal_ini": pal_ini_maq, "pal_fim": pal_fim_maq, "tot": tot_ordem_maq}
                
                if st.button("💾 GRAVAR REPORTE NIPPO NO BANCO", type="primary", use_container_width=True):
                    if not coordenador_nippo or not tecnico_nippo: st.error("Os campos Coordenador e Técnico são obrigatórios.")
                    else:
                        engine = obter_engine()
                        with engine.begin() as conn:
                            for m_item, dados in mapa_inputs_maquinas.items():
                                conn.execute(text("INSERT INTO nippo_coordenadores (data, turno, coordenador, tecnico, maquina, itens_compartilhar, produtividade, loss, sku, palete_inicial, palete_final, total_ordem) VALUES (:data, :turno, :coord, :tec, :maq, :itens, :prod, :loss, :sku, :p_ini, :p_fim, :tot)"), {"data": str(data_nippo), "turno": turno_nippo, "coord": coordenador_nippo, "tec": tecnico_nippo, "maq": m_item, "itens": dados["itens"], "prod": dados["prod"], "loss": dados["loss"], "sku": dados["sku"], "p_ini": dados["pal_ini"], "p_fim": dados["pal_fim"], "tot": int(dados["tot"])})
                        st.session_state['contador_nippo'] += 1; st.success("🎉 O Nippo completo foi gravado!"); st.rerun()
            with aba_consultar:
                query_data = st.date_input("Filtrar Data", date.today(), key="q_data")
                engine = obter_engine()
                df_nippo_res = pd.read_sql_query(text("SELECT id, data, turno, coordenador, tecnico, maquina, itens_compartilhar, sku, produtividade, loss, palete_inicial, palete_final, total_ordem FROM nippo_coordenadores WHERE data = :data"), engine, params={"data": str(query_data)})
                if df_nippo_res.empty: st.warning(f"Nenhum diário Nippo encontrado.")
                else: st.dataframe(df_nippo_res, use_container_width=True)

    else:
        st.info("💡 Por favor, carregue os arquivos Excel para iniciar o Analytics Hub.")
