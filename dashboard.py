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
    
    /* Menu Lateral */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 10px 15px !important; border-radius: 8px !important;
        margin-bottom: 5px !important; color: white !important; cursor: pointer;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
    }

    /* Cards de Métricas Reduzidos */
    .metric-container { display: flex; justify-content: space-between; gap: 8px; margin-bottom: 15px; }
    .metric-card {
        background: #1e293b; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 70px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 2px; }
    .metric-value { color: #10b981; font-size: 1.1rem; font-weight: 900; }

    /* Calendário */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .day-card { 
        background: #0f172a; border-radius: 8px; padding: 10px; 
        min-height: 90px; border: 1px solid rgba(255,255,255,0.05); 
    }
    .day-status { font-size: 0.75rem; font-weight: 600; color: #ffffff; text-align: right; }

    /* Área 5 Porquês (Clara para caneta) */
    .five-why-box {
        border: 1px solid #000; border-radius: 5px; padding: 15px; 
        background: #ffffff; color: #000; margin-top: 10px;
    }
    .five-why-line { border-bottom: 1px dotted #000; padding: 10px 0; font-size: 0.9rem; }
    .feedback-box { padding: 12px; border-radius: 8px; margin-bottom: 10px; text-align: center; font-weight: 700; font-size: 0.9rem; }
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
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    df_stops['Turno'] = df_stops['Turno'].fillna(0).astype(int).astype(str)
    return df_order, df_stops

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Excel (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

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
    # ABA 1: PERFORMANCE GERAL
    # =========================================================
    if menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t1')
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & 
                        (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        df_f = df_f.copy()
        df_f['T_Dispo'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        oee_df = df_f.groupby('Máquina').agg({'Run Time':'sum', 'Peças Estoque - Ajuste':'sum', 'Machine Counter':'sum', 'Average Speed':'mean', 'Horário Padrão':'sum', 'T_Dispo':'sum'}).reset_index()
        oee_df['OEE'] = ((oee_df['Run Time']/oee_df['T_Dispo']) * (oee_df['Machine Counter']/(oee_df['Average Speed']*oee_df['Run Time'].replace(0,1))) * (oee_df['Peças Estoque - Ajuste']/oee_df['Machine Counter']) * 100).fillna(0).round(1)

        st.markdown("## 📈 Performance Industrial")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f["Machine Counter"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f["Peças Estoque - Ajuste"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-card"><div class="metric-title">OEE Médio</div><div class="metric-value">{oee_df["OEE"].mean():.1f}%</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="metric-card"><div class="metric-title">Run Time</div><div class="metric-value">{df_f["Run Time"].sum():,.0f}m</div></div>', unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        hp_sum = df_f['Horário Padrão'].sum()
        with col1:
            val = (df_f['Run Time'].sum()/hp_sum*100) if hp_sum > 0 else 0
            st.plotly_chart(mini_gauge("Movimentação (%)", val, "#10b981", 85, 300), use_container_width=True)
        with col2:
            val_l = ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100) if df_f['Machine Counter'].sum() > 0 else 0
            st.plotly_chart(mini_gauge("Loss (%)", val_l, "#e74c3c", 5, 300), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        f_maq_s = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()), key='m2')
        f_turno_s = st.sidebar.multiselect("Turnos", sorted(df_stops['Turno'].unique()), default=sorted(df_stops['Turno'].unique()), key='ts2')
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1]) & 
                          (df_stops['Máquina'].isin(f_maq_s)) & (df_stops['Turno'].isin(f_turno_s))]
        
        st.markdown("## 🛑 Análise de Paradas")
        top_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
        st.plotly_chart(px.bar(top_min, orientation='h', title="Minutos Totais", color_discrete_sequence=['#f43f5e']), use_container_width=True)
        
        top_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
        st.plotly_chart(px.bar(top_qtd, orientation='h', title="Frequência (Quantidade)", color_discrete_sequence=['#3b82f6']), use_container_width=True)

    # =========================================================
    # ABA 3: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        meses_lista = list(calendar.month_name)[1:]
        mes_sel = st.sidebar.selectbox("Mês", meses_lista, index=datetime.now().month-1)
        mes_idx = meses_lista.index(mes_sel) + 1
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m3')
        f_turno_c = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t3')
        
        df_c = df_order[(df_order['Data'].dt.month == mes_idx) & (df_order['Máquina'].isin(f_maq_c)) & (df_order['Turno'].isin(f_turno_c))]
        
        st.markdown(f"## 📅 Operação - {mes_sel}")
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'].replace(0,1) * 100).fillna(0)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'].replace(0,1) * 100).fillna(0)

        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, mes_idx))
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#64748b; font-weight:900; font-size:0.7rem;">{n}</div>'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m_v, l_v = (row['Mov'].values[0], row['Loss'].values[0]) if not row.empty else (0,0)
                cor = "#059669" if m_v > 85 else "#dc2626" if m_v > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">M: {m_v:.1f}%<br>L: {l_v:.1f}%</div></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA 4: ANÁLISE SEMANAL
    # =========================================================
    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Filtros Board")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_b = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='tb')
        periodo_b = st.sidebar.date_input("Período", [df_order['Data'].max() - timedelta(days=7), df_order['Data'].max()])
        
        df_b_all = df_order[(df_order['Data'].dt.date >= periodo_b[0]) & (df_order['Data'].dt.date <= periodo_b[1]) & (df_order['Turno'].isin(turno_b))]
        df_b = df_b_all[df_b_all['Máquina'] == maq_b]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo_b[0]) & (df_stops['Data'].dt.date <= periodo_b[1]) & 
                         (df_stops['Máquina'] == maq_b) & (df_stops['Turno'].isin(turno_b))]

        # Título Dinâmico com Turnos
        str_turnos = ", ".join(turno_b) if turno_b else "Nenhum"
        st.markdown(f"""<div style="text-align:center; border-bottom:3px solid #10b981; padding-bottom:10px; margin-bottom:15px;">
            <h1 style="color:white; margin:0;">RELATÓRIO SEMANAL DE PERFORMANCE - MÁQUINA {maq_b}</h1>
            <h3 style="color:#10b981; margin:0;">TURNO(S): {str_turnos}</h3>
            <p style="color:#94a3b8; font-size:1rem;">Período: {periodo_b[0].strftime('%d/%m')} a {periodo_b[1].strftime('%d/%m/%Y')}</p></div>""", unsafe_allow_html=True)

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].replace(0,1).sum()*100)
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].replace(0,1).sum()*100)
        pecas_v = df_b["Peças Estoque - Ajuste"].sum()

        v1, v2, v3 = st.columns([1, 1, 1])
        with v1: st.plotly_chart(mini_gauge("Movimentação", m_v, "#10b981", 85, 180), use_container_width=True)
        with v2: st.plotly_chart(mini_gauge("Loss", l_v, "#e74c3c", 5, 180), use_container_width=True)
        with v3: st.markdown(f'<div class="metric-card" style="height:150px;"><div class="metric-title">Peças Enviadas</div><div class="metric-value" style="font-size:1.8rem;">{pecas_v:,.0f}</div></div>', unsafe_allow_html=True)

        # Ranking e Mensagem
        rank_df = df_b_all.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        rank_df['Mov %'] = (rank_df['Run Time'] / rank_df['Horário Padrão'].replace(0,1) * 100).round(1)
        rank_df = rank_df.sort_values('Mov %', ascending=False).reset_index(drop=True)
        rank_df.index += 1
        
        check_maq = rank_df[rank_df['Máquina'] == maq_b]
        if not check_maq.empty:
            posicao = check_maq.index[0]
            total_maqs = len(rank_df)
            if posicao <= 3:
                msg, cor_msg = f"🌟 EXCELENTE! Máquina {maq_b} no TOP 3 ({posicao}º).", "#064e3b"
            elif posicao > (total_maqs - 3):
                msg, cor_msg = f"💪 VAMOS LÁ! Máquina {maq_b} na posição {posicao}º. Foco total!", "#7f1d1d"
            else:
                msg, cor_msg = f"📈 BOM TRABALHO! Máquina {maq_b} na posição {posicao}º.", "#1e293b"
            st.markdown(f'<div class="feedback-box" style="background:{cor_msg}; color:white;">{msg}</div>', unsafe_allow_html=True)

        col_rank, col_stops = st.columns([1, 2])
        with col_rank:
            st.markdown("🏆 **Ranking Mov. (%)**")
            st.dataframe(rank_df[['Máquina', 'Mov %']].style.apply(lambda s: ['background-color: #10b981' if s.Máquina == maq_b else '' for _ in s], axis=1), use_container_width=True)

        with col_stops:
            st.markdown("🛑 **Impacto das Paradas (%)**")
            stop_imp = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(5)
            if not stop_imp.empty:
                pior_parada = stop_imp.index[-1]
                total_min_p = df_sb['Minutos'].sum()
                df_p_plot = stop_imp.reset_index()
                df_p_plot['%'] = (df_p_plot['Minutos'] / total_min_p * 100).round(1)
                df_p_plot['Label'] = df_p_plot.apply(lambda r: f"{r['Minutos']} min ({r['%']}%)", axis=1)
                
                fig_b = px.bar(df_p_plot, x='Minutos', y='Problema', orientation='h', text='Label', color_discrete_sequence=['#10b981'])
                fig_b.update_layout(height=250, margin=dict(l=0,r=0,t=0,b=0), paper_bgcolor='rgba(0,0,0,0)', font={'color':'white'})
                st.plotly_chart(fig_b, use_container_width=True)
            else:
                pior_parada = "Nenhuma parada registrada"

        st.markdown(f"""
            <div class="five-why-box">
                <div class="five-why-title">ANÁLISE DE CAUSA RAIZ - 5 PORQUÊS</div>
                <div style="font-weight:bold; margin-bottom:10px;">PROBLEMA FOCO: <span style="color:#e74c3c">{pior_parada}</span></div>
                <div class="five-why-line"><b>1º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>2º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>3º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>4º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>5º Por que?</b> _________________________________________________________________</div>
                <br>
                <div style="display:flex; gap:10px;">
                    <div style="flex:1; border:1px solid #000; padding:10px; min-height:80px;"><b>CAUSA RAIZ:</b></div>
                    <div style="flex:2; border:1px solid #000; padding:10px; min-height:80px;"><b>AÇÃO CORRETIVA:</b></div>
                </div>
            </div>
        """, unsafe_allow_html=True)
else:
    st.info("💡 Carregue o arquivo Excel para iniciar.")
