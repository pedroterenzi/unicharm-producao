import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta
from fpdf import FPDF

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #020617; }
    
    /* Botões do Menu Lateral */
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

    /* Cards de Métricas Reduzidos */
    .metric-container { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 20px; }
    .metric-card {
        background: #1e293b; padding: 12px; border-radius: 10px;
        text-align: center; border: 1px solid rgba(255,255,255,0.1);
        flex: 1; min-height: 80px; display: flex; flex-direction: column; justify-content: center;
    }
    .metric-title { color: #94a3b8; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; margin-bottom: 4px; }
    .metric-value { color: #10b981; font-size: 1.3rem; font-weight: 900; line-height: 1; }

    /* Calendário */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .day-card { background: #0f172a; border-radius: 8px; padding: 10px; min-height: 90px; border: 1px solid rgba(255,255,255,0.05); }
    .day-number { font-size: 0.9rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.65rem; font-weight: 600; color: #ffffff; text-align: right; line-height: 1.2; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNÇÃO GERADORA DE PDF (Frente e Verso A4)
def create_pdf(dados, pior_parada, df_stops_week):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    
    # PÁGINA 1: INDICADORES E GRÁFICO
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, f"RELATÓRIO SEMANAL DE PERFORMANCE - MÁQUINA {dados['maq']}", ln=True, align='C')
    pdf.set_font('Arial', '', 12)
    pdf.cell(0, 8, f"Período: {dados['periodo']}", ln=True, align='C')
    pdf.ln(10)
    
    # Grid de KPIs
    pdf.set_fill_color(230, 230, 230)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(47, 10, "MOVIMENTAÇÃO", 1, 0, 'C', True)
    pdf.cell(47, 10, "LOSS (PERDA)", 1, 0, 'C', True)
    pdf.cell(47, 10, "PEÇAS ENVIADAS", 1, 0, 'C', True)
    pdf.cell(47, 10, "RUN TIME TOTAL", 1, 1, 'C', True)
    
    pdf.set_font('Arial', '', 14)
    pdf.cell(47, 15, f"{dados['mov']:.1f}%", 1, 0, 'C')
    pdf.cell(47, 15, f"{dados['loss']:.1f}%", 1, 0, 'C')
    pdf.cell(47, 15, f"{dados['pecas']:,.0f}", 1, 0, 'C')
    pdf.cell(47, 15, f"{dados['runtime']:,.0f}m", 1, 1, 'C')
    
    pdf.ln(15)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, "RESUMO DE PARADAS (TOP 5)", ln=True)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(100, 8, "Problema", 1, 0, 'C', True)
    pdf.cell(44, 8, "Minutos", 1, 0, 'C', True)
    pdf.cell(44, 8, "Impacto %", 1, 1, 'C', True)
    
    pdf.set_font('Arial', '', 9)
    total_m = df_stops_week['Minutos'].sum()
    for _, row in df_stops_week.iterrows():
        p = (row['Minutos']/total_m*100) if total_m > 0 else 0
        pdf.cell(100, 8, str(row['Problema'])[:55], 1)
        pdf.cell(44, 8, f"{row['Minutos']}", 1, 0, 'C')
        pdf.cell(44, 8, f"{p:.1f}%", 1, 1, 'C')

    # PÁGINA 2: 5 PORQUÊS (VERSO)
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.set_text_color(16, 185, 129)
    pdf.cell(0, 10, "ANÁLISE DE CAUSA RAIZ - 5 PORQUÊS", ln=True, align='C')
    pdf.set_text_color(0, 0, 0)
    pdf.ln(10)
    
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, f"PROBLEMA FOCO: {pior_parada}", ln=True)
    pdf.ln(5)
    
    pdf.set_font('Arial', '', 11)
    for i in range(1, 6):
        pdf.cell(0, 12, f"{i}º Por que? __________________________________________________________________________", ln=True)
        pdf.ln(5)
        
    pdf.ln(10)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 25, "CAUSA RAIZ: ___________________________________________________________________________", 1, 1)
    pdf.ln(5)
    pdf.cell(0, 35, "PLANO DE AÇÃO: ________________________________________________________________________", 1, 1)
    
    return pdf.output(dest='S').encode('latin-1')

# 3. CARREGAMENTO DE DADOS
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

# --- UI PRINCIPAL ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Relatório (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    if menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq))]
        
        df_f = df_f.copy()
        df_f['T_Dispo'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        oee_df = df_f.groupby('Máquina').agg({'Run Time':'sum', 'Peças Estoque - Ajuste':'sum', 'Machine Counter':'sum', 'Average Speed':'mean', 'T_Dispo':'sum', 'Horário Padrão':'sum'}).reset_index()
        oee_df['OEE'] = ((oee_df['Run Time']/oee_df['T_Dispo']) * (oee_df['Machine Counter']/(oee_df['Average Speed']*oee_df['Run Time'].replace(0,1))) * (oee_df['Peças Estoque - Ajuste']/oee_df['Machine Counter']) * 100).fillna(0).round(1)

        st.markdown("## 📈 Performance e Eficiência")
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f['Machine Counter'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{df_f['Peças Estoque - Ajuste'].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">OEE Médio</div><div class="metric-value">{oee_df['OEE'].mean():.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Run Time</div><div class="metric-value">{df_f['Run Time'].sum():,.0f}m</div></div>
            </div>
            """, unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            m_v = (df_f['Run Time'].sum()/df_f['Horário Padrão'].sum()*100) if df_f['Horário Padrão'].sum()>0 else 0
            fig = go.Figure(go.Indicator(mode="gauge+number", value=m_v, title={'text':"Movimentação %"}, gauge={'bar':{'color':"#10b981"}, 'axis':{'range':[0,100]}, 'threshold':{'line':{'color':"white",'width':4},'value':85}}))
            fig.update_layout(height=380, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            l_v = ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100) if df_f['Machine Counter'].sum()>0 else 0
            fig2 = go.Figure(go.Indicator(mode="gauge+number", value=l_v, title={'text':"Loss %"}, gauge={'bar':{'color':"#e74c3c"}, 'axis':{'range':[0,15]}, 'threshold':{'line':{'color':"white",'width':4},'value':5}}))
            fig2.update_layout(height=380, paper_bgcolor='rgba(0,0,0,0)', font={'color':"white"})
            st.plotly_chart(fig2, use_container_width=True)

    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        f_maq_s = st.sidebar.multiselect("Máquinas", sorted(df_stops['Máquina'].unique()), default=sorted(df_stops['Máquina'].unique()), key='m2')
        
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1]) & (df_stops['Máquina'].isin(f_maq_s))]
        
        st.markdown("## 🛑 Top 10 Paradas")
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10), orientation='h', title="Minutos Totais", color_discrete_sequence=['#f43f5e']), use_container_width=True)
        st.plotly_chart(px.bar(df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10), orientation='h', title="Frequência (Quantidade)", color_discrete_sequence=['#3b82f6']), use_container_width=True)

    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        mes_sel = st.sidebar.selectbox("Mês", list(calendar.month_name)[1:], index=datetime.now().month-1)
        mes_idx = list(calendar.month_name).index(mes_sel) + 1
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m3')
        f_turno_c = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t3')
        
        df_c = df_order[(df_order['Data'].dt.month == mes_idx) & (df_order['Máquina'].isin(f_maq_c)) & (df_order['Turno'].isin(f_turno_c))]
        st.markdown(f"## 📅 Operação - {mes_sel}")
        
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
        cal_data['Mov'] = (cal_data['Run Time']/cal_data['Horário Padrão'].replace(0,1)*100).fillna(0)
        cal_data['Loss'] = ((cal_data['Machine Counter']-cal_data['Peças Estoque - Ajuste'])/cal_data['Machine Counter'].replace(0,1)*100).fillna(0)

        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, mes_idx))
        html = '<div class="calendar-grid">'
        for n in ['SEG','TER','QUA','QUI','SEX','SAB','DOM']: html += f'<div style="text-align:center; color:#64748b; font-weight:900; font-size:0.7rem;">{n}</div>'
        for d in days:
            if d == 0: html += '<div></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                m, l = (row['Mov'].values[0], row['Loss'].values[0]) if not row.empty else (0,0)
                cor = "#059669" if m > 85 else "#dc2626" if m > 0 else "#1e293b"
                html += f'<div class="day-card" style="background:{cor}"><span class="day-number">{d}</span><div class="day-status">M: {m:.1f}%<br>L: {l:.1f}%</div></div>'
        st.markdown(html + '</div>', unsafe_allow_html=True)

    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Configuração do Relatório")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_b = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        periodo_b = st.sidebar.date_input("Período", [df_order['Data'].max() - timedelta(days=7), df_order['Data'].max()])
        
        df_b = df_order[(df_order['Data'].dt.date >= periodo_b[0]) & (df_order['Data'].dt.date <= periodo_b[1]) & (df_order['Máquina'] == maq_b) & (df_order['Turno'].isin(turno_b))]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo_b[0]) & (df_stops['Data'].dt.date <= periodo_b[1]) & (df_stops['Máquina'] == maq_b) & (df_stops['Turno'].isin(turno_b))]

        # Cálculos para o PDF
        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].replace(0,1).sum()*100)
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].replace(0,1).sum()*100)
        p_stop = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=False).head(5).reset_index()
        pior_p = p_stop['Problema'].iloc[0] if not p_stop.empty else "Nenhuma Parada"

        dados_pdf = {'maq': maq_b, 'periodo': f"{periodo_b[0].strftime('%d/%m')} a {periodo_b[1].strftime('%d/%m/%Y')}", 'mov': m_v, 'loss': l_v, 'pecas': df_b["Peças Estoque - Ajuste"].sum(), 'runtime': df_b["Run Time"].sum()}
        
        pdf_bytes = create_pdf(dados_pdf, pior_p, p_stop)
        st.download_button(label="📥 BAIXAR PDF (FRENTE E VERSO)", data=pdf_bytes, file_name=f"Relatorio_Semanal_MQ{maq_b}.pdf", mime="application/pdf")

        st.markdown(f"### RELATÓRIO SEMANAL - MÁQUINA {maq_b}")
        st.markdown(f"""<div class="metric-container">
            <div class="metric-card"><div class="metric-title">Movimentação</div><div class="metric-value">{m_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Loss</div><div class="metric-value" style="color:#f43f5e">{l_v:.1f}%</div></div>
            <div class="metric-card"><div class="metric-title">Peças Enviadas</div><div class="metric-value">{dados_pdf['pecas']:,.0f}</div></div>
            <div class="metric-card"><div class="metric-title">Run Time</div><div class="metric-value">{dados_pdf['runtime']:,.0f}m</div></div></div>""", unsafe_allow_html=True)
        
        if not p_stop.empty:
            st.plotly_chart(px.bar(p_stop.sort_values('Minutos'), x='Minutos', y='Problema', orientation='h', title="TOP 5 PARADAS", color_discrete_sequence=['#10b981']), use_container_width=True)
else:
    st.info("💡 Carregue o arquivo Excel para iniciar.")
