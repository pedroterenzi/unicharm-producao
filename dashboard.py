import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM REFINADA ---
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

    /* Cards Padronizados */
    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 15px; }
    .metric-card {
        background: #1e293b; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 80px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    .metric-value { color: #10b981; font-size: 1.3rem; font-weight: 900; line-height: 1; }

    /* Calendário Estilizado */
    .calendar-header { text-align: center; font-weight: 900; color: #10b981; padding: 5px; font-size: 0.8rem; text-transform: uppercase; }
    .day-card { background: #0f172a; border-radius: 8px; padding: 10px; min-height: 95px; border: 1px solid rgba(255,255,255,0.05); transition: 0.3s; }
    .day-card:hover { border: 1px solid #10b981; }
    .day-number { font-size: 1.1rem; font-weight: 900; color: #f8fafc; }
    .day-status { font-size: 0.7rem; font-weight: 600; color: #94a3b8; margin-top: 5px; }

    /* Ranking e Mensagens */
    .ranking-box { background: #1e293b; padding: 15px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.1); }
    .highlight-rank { background: #064e3b !important; color: #10b981 !important; font-weight: 900; }
    .motivation-msg { padding: 15px; border-radius: 10px; margin-top: 10px; font-weight: 600; text-align: center; border-left: 5px solid; }

    /* 5 Porquês */
    .five-why-box { border: 1px solid #000; padding: 15px; background: #fff; color: #000; margin-top: 15px; }
    .line-space { border-bottom: 1px solid #000; margin-bottom: 8px; height: 18px; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_production_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Forçar conversão da coluna AN (Estoque) e outras
    df_order['Peças Estoque - Ajuste'] = pd.to_numeric(df_order['Peças Estoque - Ajuste'], errors='coerce').fillna(0)
    nums = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Average Speed']
    for col in nums: df_order[col] = pd.to_numeric(df_order[col], errors='coerce').fillna(0)
    
    df_stops['Minutos'] = pd.to_numeric(df_stops['Minutos'], errors='coerce').fillna(0)
    df_stops['QTD'] = pd.to_numeric(df_stops['QTD'], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order = df_order.dropna(subset=['Data'])
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
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
        row_meta = df_raw.iloc[124, :].tolist() # Linha 125
        
        m_geral, m_hoje = 0, 0
        for idx, d in enumerate(row_dates):
            if isinstance(d, (datetime, pd.Timestamp)):
                val = pd.to_numeric(row_meta[idx], errors='coerce') or 0
                m_geral += val
                if d.date() <= data_ref: m_hoje += val
        return m_geral, m_hoje
    except: return 0, 0

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    up_p = st.file_uploader("1. Produção (.xlsm)", type=["xlsm"])
    up_d = st.file_uploader("2. DATAS (.xlsx)", type=["xlsx"])
    if up_p:
        menu = st.radio("NAVEGAÇÃO", ["📋 REPORTE DIÁRIO", "📈 PERFORMANCE GERAL", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 SEMANAL"], label_visibility="collapsed")

if up_p:
    df_order, df_stops = load_production_data(up_p)

    # =========================================================
    # VISÃO: REPORTE DIÁRIO
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        data_ref = df_order['Data'].max().date()
        st.markdown(f"## 📋 Reporte Diário - {data_ref.strftime('%d/%m/%Y')}")
        
        # Cálculo Ajustado do Estoque (Coluna AN filtrada pelo mês atual)
        df_mes = df_order[df_order['Data'].dt.month == data_ref.month]
        estoque_correto = df_mes['Peças Estoque - Ajuste'].sum()
        mov_mes = (df_mes['Run Time'].sum() / df_mes['Horário Padrão'].sum() * 100) if df_mes['Horário Padrão'].sum() > 0 else 0
        loss_mes = ((df_mes['Machine Counter'].sum() - estoque_correto) / df_mes['Machine Counter'].sum() * 100) if df_mes['Machine Counter'].sum() > 0 else 0
        
        m_geral, m_hoje = load_planner_metas(up_d, data_ref) if up_d else (0,0)

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Movimentação Mês</div><div class="metric-value">{mov_mes:.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Loss Mês</div><div class="metric-value" style="color:#f43f5e">{loss_mes:.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Estoque Total Mês</div><div class="metric-value">{estoque_correto:,.0f}</div></div>
            </div>
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Meta Acumulada Dia</div><div class="metric-value" style="color:#3b82f6">{m_hoje:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Meta Geral Mês</div><div class="metric-value" style="color:#10b981">{m_geral:,.0f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        datas_disp = sorted(df_order['Data'].dt.date.unique(), reverse=True)
        dias_sel = st.multiselect("Histórico Diário:", datas_disp, default=datas_disp[:3])
        for d in dias_sel:
            st.markdown(f"### 📅 Produção {d.strftime('%d/%m')}")
            df_d = df_order[df_order['Data'].dt.date == d]
            res = df_d.groupby(['Categoria','Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
            res['Mov %'] = (res['Run Time']/res['Horário Padrão']*100).round(1)
            res['Perda %'] = ((res['Machine Counter']-res['Peças Estoque - Ajuste'])/res['Machine Counter']*100).round(1)
            st.table(res[['Categoria','Máquina','Mov %','Perda %','Peças Estoque - Ajuste']])

    # =========================================================
    # VISÃO: PERFORMANCE GERAL
    # =========================================================
    elif menu == "📈 PERFORMANCE GERAL":
        st.sidebar.subheader("Filtros")
        f_d = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()])
        f_m = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        f_t = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        
        df_f = df_order[(df_order['Data'].dt.date >= f_d[0]) & (df_order['Data'].dt.date <= f_d[1]) & (df_order['Máquina'].isin(f_m)) & (df_order['Turno'].isin(f_t))]
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f['Machine Counter'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f['Peças Estoque - Ajuste'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f['Run Time'].sum():,.0f}m</div></div>
            </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            val_m = (df_f['Run Time'].sum()/df_f['Horário Padrão'].sum()*100) if df_f['Horário Padrão'].sum()>0 else 0
            fig1 = go.Figure(go.Indicator(mode="gauge+number", value=val_m, title={'text':"Movimentação %", 'font':{'size':18}}, gauge={'bar':{'color':"#10b981"}, 'axis':{'range':[0,100]}}))
            fig1.update_layout(height=280, margin=dict(t=50, b=20), paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig1, use_container_width=True)
        with col2:
            val_l = ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100) if df_f['Machine Counter'].sum()>0 else 0
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=val_l, title={'text':"Loss %", 'font':{'size':18}}, gauge={'bar':{'color':"#e74c3c"}, 'axis':{'range':[0,15]}}))
            fig2.update_layout(height=280, margin=dict(t=50, b=20), paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig2, use_container_width=True)

    # =========================================================
    # VISÃO: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        m_idx = list(calendar.month_name).index(mes_sel) + 1
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='cal_m')
        
        df_c = df_order[(df_order['Data'].dt.month == m_idx) & (df_order['Máquina'].isin(f_maq))]
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        
        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, m_idx))
        
        # Grid dos dias da semana
        cols_name = st.columns(7)
        for i, name in enumerate(['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']):
            cols_name[i].markdown(f"<div class='calendar-header'>{name}</div>", unsafe_allow_html=True)

        html_grid = '<div class="calendar-grid">'
        for d in days:
            if d == 0: html_grid += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_v = row['Run Time'].values[0]/row['Horário Padrão'].values[0]*100 if not row.empty and row['Horário Padrão'].values[0]>0 else 0
                cor = "#059669" if m_v > 85 else "#dc2626" if m_v > 0 else "#1e293b"
                html_grid += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">MOV: {m_v:.1f}%</div></div>'
        st.markdown(html_grid + '</div>', unsafe_allow_html=True)

    # =========================================================
    # VISÃO: SEMANAL (IMPRESSÃO)
    # =========================================================
    elif menu == "📋 SEMANAL":
        st.sidebar.subheader("Filtros do Relatório")
        maq_sel = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_sel = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        per = st.sidebar.date_input("Semana", [datetime.now()-timedelta(days=7), datetime.now()])
        
        st.markdown(f"""
            <div style="text-align:center; border:2px solid #10b981; padding:10px; border-radius:10px;">
                <h2 style="color:white; margin:0;">RELATÓRIO SEMANAL DE PERFORMANCE - MÁQUINA {maq_sel}</h2>
                <h4 style="color:#94a3b8; margin:5px 0;">Período: {per[0].strftime('%d/%m')} a {per[1].strftime('%d/%m/%Y')}</h4>
            </div>
        """, unsafe_allow_html=True)

        df_b = df_order[(df_order['Data'].dt.date >= per[0]) & (df_order['Data'].dt.date <= per[1]) & (df_order['Turno'].isin(turno_sel))]
        df_sb = df_stops[(df_stops['Data'].dt.date >= per[0]) & (df_stops['Data'].dt.date <= per[1]) & (df_stops['Máquina'] == maq_sel) & (df_stops['Turno'].isin(turno_sel))]
        
        # Gráfico Horizontal e Ranking
        col_g, col_r = st.columns([2, 1])
        
        with col_g:
            stop_data = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(5)
            fig_bar = px.bar(stop_data, orientation='h', text_auto=True, title="Piores 5 Paradas (Minutos)", color_discrete_sequence=['#10b981'])
            fig_bar.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig_bar, use_container_width=True)
            pior_p = stop_data.index[-1] if not stop_data.empty else "---"

        with col_r:
            st.markdown("<div class='ranking-box'>", unsafe_allow_html=True)
            st.markdown("🏆 **RANKING MOVIMENTAÇÃO**")
            rk = df_b.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
            rk['Mov %'] = (rk['Run Time']/rk['Horário Padrão']*100).round(1)
            rk = rk.sort_values('Mov %', ascending=False).reset_index(drop=True)
            rk.index += 1
            
            for i, r in rk.iterrows():
                is_this = "highlight-rank" if r['Máquina'] == maq_sel else ""
                st.markdown(f"<div class='{is_this}' style='padding:5px; border-radius:5px;'>{i}º - MÁQ {r['Máquina']}: {r['Mov %']}%</div>", unsafe_allow_html=True)
            
            # IA de Mensagens
            pos = rk[rk['Máquina']==maq_sel].index[0] if maq_sel in rk['Máquina'].values else 99
            if pos <= 2:
                msg, cor = f"🏆 PARABÉNS! A Máquina {maq_sel} está voando! Liderança absoluta na semana.", "#064e3b"
            elif pos <= 4:
                msg, cor = f"💪 ÓTIMO TRABALHO! A Máquina {maq_sel} está no pelotão de elite. Continue assim!", "#1e293b"
            else:
                msg, cor = f"🚀 MOTIVAÇÃO! Hora de analisar os 5 porquês e ajustar os detalhes. Vamos subir esse ranking!", "#7f1d1d"
            st.markdown(f"<div class='motivation-msg' style='background:{cor}; border-color:#10b981;'>{msg}</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown(f"""
            <div class="five-why-box">
                <h3 style="color:#059669; margin:0;">ANÁLISE 5 PORQUÊS: {pior_p}</h3>
                1. Por que? <div class="line-space"></div> 2. Por que? <div class="line-space"></div>
                3. Por que? <div class="line-space"></div> 4. Por que? <div class="line-space"></div>
                5. Por que? <div class="line-space"></div>
                <b>CAUSA RAIZ / PLANO DE AÇÃO:</b> <div class="line-space"></div> <div class="line-space"></div>
            </div>
        """, unsafe_allow_html=True)

else:
    st.info("💡 Carregue os arquivos para começar.")
