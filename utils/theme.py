"""CSS escuro profissional para o dashboard - estilo Power BI dark."""

import streamlit as st


def aplicar_tema():
    """Aplica o tema CSS escuro profissional."""
    st.markdown("""
    <style>
        /* Esconder menu e rodape padrao */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Fundo geral - gradiente escuro */
        .stApp {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            color: #e2e8f0;
        }

        /* Titulo e texto geral */
        .stApp h1, .stApp h2, .stApp h3 {
            color: #ffffff !important;
        }
        .stApp p, .stApp label, .stApp span {
            color: #cbd5e1;
        }

        /* Cards KPI */
        .kpi-card {
            background: rgba(255,255,255,0.08);
            border-radius: 15px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.1);
            text-align: center;
            backdrop-filter: blur(10px);
            transition: all 0.3s ease;
        }
        .kpi-card:hover {
            transform: translateY(-3px);
            border-color: rgba(102, 126, 234, 0.5);
            box-shadow: 0 4px 20px rgba(102, 126, 234, 0.2);
        }
        .kpi-label {
            font-size: 0.75rem;
            color: #aaa;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 8px;
        }
        .kpi-value {
            font-size: 2rem;
            font-weight: 700;
            color: #fbbf24;
            margin-bottom: 4px;
        }
        .kpi-detail {
            font-size: 0.85rem;
            color: #94a3b8;
        }

        /* Cores positivas/negativas */
        .positive { color: #4ade80; }
        .negative { color: #f87171; }
        .neutral { color: #60a5fa; }

        /* Status badges */
        .status-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .status-critica { background: rgba(248, 113, 113, 0.2); color: #f87171; }
        .status-atencao { background: rgba(251, 191, 36, 0.2); color: #fbbf24; }
        .status-normal { background: rgba(74, 222, 128, 0.2); color: #4ade80; }
        .status-lotada { background: rgba(96, 165, 250, 0.2); color: #60a5fa; }
        .status-super { background: rgba(167, 139, 250, 0.2); color: #a78bfa; }

        /* Barra de ocupacao */
        .ocupacao-bar {
            width: 100%;
            height: 8px;
            background: rgba(255,255,255,0.1);
            border-radius: 4px;
            overflow: hidden;
        }
        .ocupacao-fill {
            height: 100%;
            border-radius: 4px;
            transition: width 0.3s ease;
        }

        /* Header customizado */
        .dashboard-header {
            background: rgba(255,255,255,0.1);
            border-radius: 15px;
            padding: 20px 25px;
            margin-bottom: 25px;
            border: none;
            display: flex;
            justify-content: space-between;
            align-items: center;
            backdrop-filter: blur(10px);
        }

        /* Secao */
        .section-title {
            font-size: 1.3rem;
            font-weight: 600;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            color: #ffffff;
        }

        /* Tabelas Streamlit */
        .stDataFrame {
            border-radius: 10px;
            overflow: hidden;
        }

        /* Sidebar */
        [data-testid="stSidebar"] {
            background: #0f1629;
        }
        [data-testid="stSidebar"] .block-container {
            padding-top: 2rem;
        }
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] label {
            color: #cbd5e1 !important;
        }
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #ffffff !important;
        }

        /* Metricas do Streamlit */
        [data-testid="stMetric"] {
            background: rgba(255,255,255,0.08);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 15px;
            padding: 15px 20px;
            backdrop-filter: blur(10px);
        }
        [data-testid="stMetricLabel"] {
            color: #aaa !important;
            font-size: 0.8rem !important;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.8rem !important;
            font-weight: 700 !important;
            color: #fbbf24 !important;
        }
        [data-testid="stMetricDelta"] {
            color: #94a3b8 !important;
        }
        [data-testid="stMetricDelta"] svg {
            display: none;
        }

        /* Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }
        .stTabs [data-baseweb="tab"] {
            background: rgba(255,255,255,0.05);
            border-radius: 8px;
            padding: 8px 16px;
            color: #94a3b8;
        }
        .stTabs [aria-selected="true"] {
            background: rgba(102, 126, 234, 0.3) !important;
            color: #ffffff !important;
            font-weight: 600;
        }

        /* Dividers */
        hr {
            border-color: rgba(255,255,255,0.1);
        }

        /* Radio buttons */
        .stRadio label {
            color: #cbd5e1 !important;
        }
        .stRadio [data-testid="stMarkdownContainer"] p {
            color: #e2e8f0 !important;
        }

        /* Captions */
        .stCaption, small {
            color: #64748b !important;
        }

        /* ====== Tabelas HTML customizadas ====== */
        .dark-table-container {
            background: rgba(0,0,0,0.2);
            border-radius: 12px;
            border: none;
            overflow: hidden;
            margin-bottom: 15px;
        }
        .dark-table-header {
            background: rgba(102, 126, 234, 0.3);
            padding: 12px 15px;
            font-weight: 600;
            font-size: 0.95rem;
            color: #ffffff;
            border-bottom: none;
        }
        .dark-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.85rem;
        }
        .dark-table th {
            padding: 8px 12px;
            text-align: left;
            font-weight: 600;
            color: #aaa;
            text-transform: uppercase;
            font-size: 0.7rem;
            letter-spacing: 0.5px;
            background: rgba(255,255,255,0.05);
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .dark-table th.num { text-align: right; }
        .dark-table td {
            padding: 7px 12px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #e2e8f0;
        }
        .dark-table td.num { text-align: right; font-variant-numeric: tabular-nums; }
        .dark-table tr:hover { background: rgba(255,255,255,0.03); }
        .dark-table tr.total-row {
            background: rgba(102, 126, 234, 0.2);
            font-weight: 700;
        }
        .dark-table tr.total-row td { color: #ffffff; border-top: 1px solid rgba(255,255,255,0.1); }
        .dark-table tr.subtotal-row {
            background: rgba(255,255,255,0.05);
            font-weight: 600;
        }
        .dark-table tr.subtotal-row td { color: #e2e8f0; }
        .dark-table tr.segment-header td {
            color: #fbbf24;
            font-weight: 600;
            font-size: 0.8rem;
            padding-top: 12px;
            border-bottom: none;
        }
        .val-pos { color: #4ade80 !important; }
        .val-neg { color: #f87171 !important; }
        .val-warn { color: #fbbf24 !important; }

        /* Progress bar metas */
        .meta-card {
            background: rgba(255,255,255,0.08);
            border-radius: 12px;
            padding: 15px 18px;
            border: 1px solid rgba(255,255,255,0.1);
            min-width: 0;
            backdrop-filter: blur(10px);
        }
        .meta-card-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }
        .meta-card-name { font-weight: 600; font-size: 0.95rem; color: #ffffff; }
        .meta-card-pct { font-size: 1.3rem; font-weight: 700; }
        .meta-bar {
            width: 100%;
            height: 10px;
            background: rgba(255,255,255,0.1);
            border-radius: 5px;
            overflow: hidden;
            margin-bottom: 6px;
        }
        .meta-bar-fill {
            height: 100%;
            border-radius: 5px;
            transition: width 0.5s ease;
        }
        .meta-card-bottom {
            display: flex;
            justify-content: space-between;
            font-size: 0.78rem;
            color: #94a3b8;
        }

        /* Total rede barra */
        .meta-total {
            background: rgba(255,255,255,0.08);
            border-radius: 12px;
            padding: 18px 20px;
            border: 1px solid rgba(255,255,255,0.15);
            margin-top: 12px;
            backdrop-filter: blur(10px);
        }
        .meta-total-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }
        .meta-total-name { font-weight: 700; font-size: 1.05rem; color: #ffffff; }
        .meta-total-pct { font-size: 1.5rem; font-weight: 700; }
        .meta-total-bar {
            width: 100%;
            height: 14px;
            background: rgba(255,255,255,0.1);
            border-radius: 7px;
            overflow: hidden;
            margin-bottom: 6px;
        }
        .meta-total-bar-fill {
            height: 100%;
            border-radius: 7px;
        }
        .meta-total-bottom {
            display: flex;
            justify-content: space-between;
            font-size: 0.82rem;
            color: #94a3b8;
        }

        /* Integral cards */
        .integral-card {
            background: rgba(0,0,0,0.2);
            border-radius: 10px;
            padding: 15px;
            text-align: center;
        }
    </style>
    """, unsafe_allow_html=True)


def kpi_card(label, value, detail="", color="#fbbf24"):
    """Renderiza um card KPI customizado."""
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" style="color: {color};">{value}</div>
        <div class="kpi-detail">{detail}</div>
    </div>
    """
