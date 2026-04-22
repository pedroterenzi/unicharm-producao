import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS PREMIUM REFINADA ---
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

    /* Cartões Padronizados - Performance Geral */
    .metric-container {
        display: flex;
        justify-content: space-between;
        gap: 10px;
        margin-bottom: 20px;
    }
    .metric-card {
        background: #1e293b; 
        padding: 12px; 
        border-radius: 10px;
        text-align: center; 
        border: 1px solid rgba(255,255,255,0.1);
        flex: 1;
        min-height: 100px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .metric-title { 
        color: #94a3b8; 
        font-size: 0.65rem; 
        font-weight: 700; 
        text-transform: uppercase; 
        margin-bottom: 4px;
        white-space: nowrap;
        overflow: hidden;
    }
    .metric-value { 
        color: #10b981; 
        font-size: 1.2rem; 
        font-weight: 900; 
        line-height: 1;
    }

    /* Calendário */
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; }
    .day-card { 
        background: #0f172a; border-radius: 8px; padding: 10px; 
        min-height: 90px; border: 1px solid rgba(255,255,255,0.05); 
    }
    .day-number { font-size: 0.9rem; font-weight: 900; color: #94a3b8; }
    .day-status { font-size: 0.65rem; font-weight: 600; color: #ffffff; text-align: right; line-height: 1.2; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file):
    df_order = pd.read_excel(file, sheet_name="Result by order")
    df_stops = pd.read_excel(file, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # Tratamento numérico robusto
    for df in [df_order, df_stops]:
        for col in df.columns:
            if col not in ['Data', 'Máquina', 'Turno', 'Problema']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    return df_order, df_stops

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#10b981;'>💎 INDUSTRIAL HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Carregar Relatório (.xlsm)", type=["xlsm"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO"])

# --- LÓGICA POR ABA ---
if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    # =========================================================
    # ABA 1: PERFORMANCE GERAL (Cartões Padronizados e Velocímetros)
    # =========================================================
    if menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros Performance")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()])
        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq))]
        
        # Cálculos OEE
        df_f['Tempo_Disponivel'] = df_f['Turno'].map({'1': 455, '2': 440, '3': 415}).astype(float)
        oee_data = df_f.groupby('Máquina').agg({
            'Run Time':'sum', 
            'Tempo_Disponivel':'sum', 
            'Peças Estoque - Ajuste':'sum', 
            'Machine Counter':'sum', 
            'Average Speed':'mean',
            'Horário Padrão':'sum'
        }).reset_index()
        
        oee_data['Disp'] = (oee_data['Run Time'] / oee_data['Tempo_Disponivel']).clip(0,1)
        oee_data['Qual'] = (oee_data['Peças Estoque - Ajuste'] / oee_data['Machine Counter']).fillna(0).clip(0,1)
        oee_data['Perf'] = (oee_data['Machine Counter'] / (oee_data['Average Speed'] * oee_data['Run Time'])).fillna(0).clip(0,1)
        oee_data['OEE'] = (oee_data['Disp'] * oee_data['Perf'] * oee_data['Qual'] * 100).round(1)

        st.markdown("## 📈 Performance e Eficiência")

        # Cartões Padronizados
        total_mc = df_f["Machine Counter"].sum()
        total_est = df_f["Peças Estoque - Ajuste"].sum()
        avg_oee = oee_data["OEE"].mean()
        total_rt = df_f["Run Time"].sum()
        
        # Movimentação e Loss para os velocímetros
        hp_sum = df_f['Horário Padrão'].sum()
        mov_p = (total_rt / hp_sum * 100) if hp_sum > 0 else 0
        loss_p = ((total_mc - total_est) / total_mc * 100) if total_mc > 0 else 0

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{total_mc:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value">{total_est:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">OEE Médio</div><div class="metric-value">{avg_oee:.1f}%</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{total_rt:,.0f}m</div></div>
            </div>
            """, unsafe_allow_html=True)

        # Velocímetros
        col_v1, col_v2 = st.columns(2)
        
        def create_gauge(label, value, color, target):
            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=value,
                title={'text': label, 'font': {'size': 18, 'color': 'white'}},
                gauge={
                    'axis': {'range': [0, 100], 'tickcolor': "white"},
                    'bar': {'color': color},
                    'bgcolor': "rgba(0,0,0,0)",
                    'threshold': {'line': {'color': "white", 'width': 4}, 'thickness': 0.75, 'value': target}
                }
            ))
            fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font={'color': "white"}, height=250, margin=dict(t=30, b=0, l=30, r=30))
            return fig

        with col_v1:
            st.plotly_chart(create_gauge("Movimentação (%)", mov_p, "#10b981", 85), use_container_width=True)
        with col_v2:
            st.plotly_chart(create_gauge("Loss (%)", loss_p, "#f43f5e", 5), use_container_width=True)

        st.markdown("### 🏆 Ranking por Máquina")
        st.dataframe(oee_data[['Máquina', 'Disp', 'Perf', 'Qual', 'OEE']].sort_values('OEE', ascending=False), use_container_width=True)

    # =========================================================
    # ABA 2: TOP 10 PARADAS (Um por linha)
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        st.sidebar.subheader("Filtros Paradas")
        f_data_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()])
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_data_s[0]) & (df_stops['Data'].dt.date <= f_data_s[1])]
        
        st.markdown("## 🛑 Top 10 Paradas")
        
        top_min = df_s_f.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(10)
        fig_min = px.bar(top_min, orientation='h', title="Minutos Totais por Motivo", color_discrete_sequence=['#f43f5e'])
        fig_min.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color': 'white'})
        st.plotly_chart(fig_min, use_container_width=True)
        
        st.markdown("---")
        
        top_qtd = df_s_f.groupby('Problema')['QTD'].sum().sort_values(ascending=True).tail(10)
        fig_qtd = px.bar(top_qtd, orientation='h', title="Frequência (Quantidade) por Motivo", color_discrete_sequence=['#3b82f6'])
        fig_qtd.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color': 'white'})
        st.plotly_chart(fig_qtd, use_container_width=True)

    # =========================================================
    # ABA 3: CALENDÁRIO (Mês, Máquina, Turno)
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        st.sidebar.subheader("Filtros Calendário")
        mes_lista = list(calendar.month_name)[1:]
        mes_sel = st.sidebar.selectbox("Mês", mes_lista, index=datetime.now().month-1)
        mes_idx = mes_lista.index(mes_sel) + 1
        
        f_maq_c = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        f_turno_c = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        
        df_c = df_order[(df_order['Data'].dt.month == mes_idx) & (df_order['Máquina'].isin(f_maq_c)) & (df_order['Turno'].isin(f_turno_c))]
        
        st.markdown(f"## 📅 Operação - {mes_sel}")
        
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
                m_val = row['Mov'].values[0] if not row.empty else 0
                l_val = row['Loss'].values[0] if not row.empty else 0
                cor = "#059669" if m_val > 85 else "#dc2626" if m_val > 0 else "#1e293b"
                html += f'''<div class="day-card" style="background:{cor}">
                            <span class="day-number">{d}</span>
                            <div class="day-status">MOV: {m_val:.1f}%<br>LOSS: {l_val:.1f}%</div>
                          </div>'''
        st.markdown(html + '</div>', unsafe_allow_html=True)
else:
    st.info("💡 Por favor, carregue o arquivo Excel (.xlsm) no menu lateral.")
