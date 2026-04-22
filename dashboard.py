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
    
    /* Navegação Lateral */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #1e293b; border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 12px 20px !important; border-radius: 10px !important;
        margin-bottom: 8px !important; color: white !important;
        transition: all 0.3s ease; cursor: pointer;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #10b981 0%, #064e3b 100%) !important;
        border: none !important;
    }

    /* Cards de Métricas */
    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 20px; }
    .metric-card {
        background: #1e293b; padding: 15px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 100px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { color: #10b981; font-size: 1.8rem; font-weight: 900; line-height: 1; }

    /* Área 5 Porquês (Clara para escrita à caneta) */
    .five-why-box {
        border: 2px solid #cbd5e1; border-radius: 10px; padding: 25px; 
        background: #ffffff; margin-top: 20px; color: #1e293b;
    }
    .five-why-line { border-bottom: 1px dotted #94a3b8; padding: 12px 0; font-size: 1rem; color: #334155; }
    .five-why-title { color: #059669; font-weight: 900; font-size: 1.2rem; margin-bottom: 15px; border-bottom: 2px solid #059669; padding-bottom: 5px; }

    /* Ajustes para Impressão */
    @media print {
        .stApp { background-color: white !important; color: black !important; }
        [data-testid="stSidebar"], header { display: none !important; }
        .metric-card { border: 2px solid black !important; background: white !important; }
        .metric-value, .metric-title { color: black !important; }
        .five-why-box { border: 2px solid black !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO DOS DADOS
@st.cache_data
def load_data(file_obj):
    df_order = pd.read_excel(file_obj, sheet_name="Result by order")
    df_stops = pd.read_excel(file_obj, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    cols_num = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed', 'Minutos', 'QTD']
    for df in [df_order, df_stops]:
        for col in cols_num:
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
    st.markdown("<h1 style='color:#10b981;'>🏭 INDUSTRIAL HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Relatório (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📝 ANÁLISE SEMANAL"])

# --- LÓGICA POR ABA ---
if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

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
        
        df_f['T_Dispo'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        oee_df = df_f.groupby('Máquina').agg({'Run Time':'sum', 'Tempo_Disponivel':'sum', 'Peças Estoque - Ajuste':'sum', 'Machine Counter':'sum', 'Average Speed':'mean', 'Horário Padrão':'sum', 'T_Dispo':'sum'}).reset_index()
        oee_df['Disp'] = (oee_df['Run Time'] / oee_df['T_Dispo']).clip(0,1)
        oee_df['Qual'] = (oee_df['Peças Estoque - Ajuste'] / oee_df['Machine Counter']).fillna(0).clip(0,1)
        oee_df['Perf'] = (oee_df['Machine Counter'] / (oee_df['Average Speed'] * oee_df['Run Time'])).fillna(0).clip(0,1)
        oee_df['OEE'] = (oee_df['Disp'] * oee_df['Perf'] * oee_df['Qual'] * 100).round(1)

        st.markdown("## 📈 Performance e Eficiência")
        
        total_mc = df_f["Machine Counter"].sum()
        total_est = df_f["Peças Estoque - Ajuste"].sum()
        total_hp = df_f["Horário Padrão"].sum()
        total_rt = df_f["Run Time"].sum()

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{total_mc:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{total_est:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Horário Padrão (min)</div><div class="metric-value">{total_hp:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{total_rt:,.0f}</div></div>
            </div>
            """, unsafe_allow_html=True)

        col_v1, col_v2 = st.columns(2)
        def create_gauge(label, value, color, target):
            fig = go.Figure(go.Indicator(mode="gauge+number", value=value, title={'text': label, 'font': {'size': 24}},
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': color}, 'threshold': {'line': {'color': "white", 'width': 4}, 'value': target}}))
            fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=450, margin=dict(l=50, r=50, t=80, b=50))
            return fig

        col_v1.plotly_chart(create_gauge("Movimentação (%)", (total_rt/total_hp*100 if total_hp>0 else 0), "#10b981", 85), use_container_width=True)
        col_v2.plotly_chart(create_gauge("Loss (%)", ((total_mc-total_est)/total_mc*100 if total_mc>0 else 0), "#e74c3c", 5), use_container_width=True)

        st.markdown("### 🏆 Ranking OEE por Máquina")
        st.dataframe(oee_df[['Máquina', 'Disp', 'Perf', 'Qual', 'OEE']].sort_values('OEE', ascending=False), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS (Filtros Máquina e Turno incluídos)
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        f_maq_s = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()), key='m2')
        f_turno_s = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t2')
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1]) & 
                          (df_stops['Máquina'].isin(f_maq_s)) & (df_stops['Turno'].isin(f_turno_s))]
        
        st.markdown("## 🛑 Análise de Paradas")
        
        top_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
        fig_min = px.bar(top_min, orientation='h', title="Top 10 por Minutos Totais", color_discrete_sequence=['#f43f5e'])
        st.plotly_chart(fig_min, use_container_width=True)
        
        st.markdown("---")
        
        top_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
        fig_qtd = px.bar(top_qtd, orientation='h', title="Top 10 por Quantidade", color_discrete_sequence=['#3b82f6'])
        st.plotly_chart(fig_qtd, use_container_width=True)

    # =========================================================
    # ABA 3: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        mes_idx = list(calendar.month_name).index(mes_sel)
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m3')
        f_turno_c = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t3')
        
        df_c = df_order[(df_order['Data'].dt.month == mes_idx) & (df_order['Máquina'].isin(f_maq_c)) & (df_order['Turno'].isin(f_turno_c))]
        
        st.markdown(f"## 📅 Calendário Operacional - {mes_sel}")
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time'] / cal_data['Horário Padrão'] * 100).fillna(0)
        cal_data['Loss'] = ((cal_data['Machine Counter'] - cal_data['Peças Estoque - Ajuste']) / cal_data['Machine Counter'] * 100).fillna(0)

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
    # ABA 4: ANÁLISE SEMANAL (Board de Impressão)
    # =========================================================
    elif menu == "📝 ANÁLISE SEMANAL":
        st.sidebar.subheader("Filtros Análise")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        # Período customizado para o Board
        periodo_b = st.sidebar.date_input("Escolha o Período", [datetime.now() - timedelta(days=7), datetime.now()])
        
        df_b = df_order[(df_order['Data'].dt.date >= periodo_b[0]) & (df_order['Data'].dt.date <= periodo_b[1]) & (df_order['Máquina'] == maq_b)]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo_b[0]) & (df_stops['Data'].dt.date <= periodo_b[1]) & (df_stops['Máquina'] == maq_b)]

        st.markdown(f"""<div style="text-align:center; border-bottom:3px solid #10b981; padding-bottom:10px; margin-bottom:20px;">
            <h1 style="color:white; margin:0;">ANÁLISE DE PERFORMANCE - MÁQUINA {maq_b}</h1>
            <p style="color:#94a3b8; font-size:1.2rem;">Período: {periodo_b[0].strftime('%d/%m/%Y')} a {periodo_b[1].strftime('%d/%m/%Y')}</p></div>""", unsafe_allow_html=True)

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].sum()*100) if df_b["Horário Padrão"].sum()>0 else 0
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].sum()*100) if df_b["Machine Counter"].sum()>0 else 0
        
        st.markdown(f"""<div class="metric-container">
            <div class="metric-card"><div class="metric-title">Movimentação</div><div class="metric-value">{m_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Loss (Perda)</div><div class="metric-value" style="color:#f43f5e">{l_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Peças Enviadas</div><div class="metric-value">{df_b["Peças Estoque - Ajuste"].sum():,.0f}</div></div>
            <div class="metric-card"><div class="metric-title">Total Minutos Produzidos</div><div class="metric-value">{df_b["Run Time"].sum():,.0f}m</div></div></div>""", unsafe_allow_html=True)

        st.markdown("### 🛑 Gráfico de Paradas (Impacto %)")
        stop_imp = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True)
        if not stop_imp.empty:
            total_p = stop_imp.sum()
            fig = px.bar(stop_imp/total_p*100, orientation='h', labels={'value':'% do Tempo Total Parado','Problema':''}, color_discrete_sequence=['#10b981'])
            fig.update_layout(height=400, font=dict(size=14))
            st.plotly_chart(fig, use_container_width=True)

        st.markdown(f"""
            <div class="five-why-box">
                <div class="five-why-title">ANÁLISE DE CAUSA RAIZ - FERRAMENTA 5 PORQUÊS</div>
                <div style="font-weight:bold; margin-bottom:10px;">PROBLEMA FOCO: __________________________________________________________________________</div>
                <div class="five-why-line"><b>1º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>2º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>3º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>4º Por que?</b> ________________________________________________________________________________________</div>
                <div class="five-why-line"><b>5º Por que?</b> ________________________________________________________________________________________</div>
                <br>
                <div style="display:flex; gap:20px;">
                    <div style="flex:1; border:1px solid #94a3b8; padding:10px; min-height:100px;"><b>CAUSA RAIZ:</b></div>
                    <div style="flex:2; border:1px solid #94a3b8; padding:10px; min-height:100px;"><b>AÇÃO CORRETIVA / PLANO DE AÇÃO:</b></div>
                </div>
            </div>
        """, unsafe_allow_html=True)

else:
    st.info("💡 Carregue o arquivo Excel para iniciar.")
