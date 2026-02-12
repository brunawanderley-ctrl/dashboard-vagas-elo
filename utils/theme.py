"""CSS claro profissional para o dashboard."""

import streamlit as st


def aplicar_tema():
    """Aplica o tema CSS claro profissional."""
    st.markdown("""
    <style>
        /* Esconder menu e rodape padrao */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Fundo geral */
        .stApp {
            background: #f8f9fa;
        }

        /* Cards KPI */
        .kpi-card {
            background: #ffffff;
            border-radius: 16px;
            padding: 20px;
            border: 1px solid #e2e8f0;
            text-align: center;
            transition: all 0.3s ease;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        .kpi-card:hover {
            transform: translateY(-3px);
            border-color: rgba(102, 126, 234, 0.4);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.15);
        }
        .kpi-label {
            font-size: 0.75rem;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 8px;
        }
        .kpi-value {
            font-size: 2rem;
            font-weight: 700;
            margin-bottom: 4px;
        }
        .kpi-detail {
            font-size: 0.85rem;
            color: #94a3b8;
        }

        /* Cores positivas/negativas */
        .positive { color: #16a34a; }
        .negative { color: #dc2626; }
        .neutral { color: #2563eb; }

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
        .status-critica { background: rgba(220, 38, 38, 0.1); color: #dc2626; }
        .status-atencao { background: rgba(202, 138, 4, 0.1); color: #b45309; }
        .status-normal { background: rgba(22, 163, 74, 0.1); color: #16a34a; }
        .status-lotada { background: rgba(37, 99, 235, 0.1); color: #2563eb; }
        .status-super { background: rgba(124, 58, 237, 0.1); color: #7c3aed; }

        /* Barra de ocupacao */
        .ocupacao-bar {
            width: 100%;
            height: 8px;
            background: #e2e8f0;
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
            background: #ffffff;
            border-radius: 16px;
            padding: 20px 25px;
            margin-bottom: 25px;
            border: 1px solid #e2e8f0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }

        /* Secao */
        .section-title {
            font-size: 1.3rem;
            font-weight: 600;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 1px solid #e2e8f0;
            color: #1e293b;
        }

        /* Tabelas Streamlit */
        .stDataFrame {
            border-radius: 10px;
            overflow: hidden;
        }

        /* Sidebar */
        [data-testid="stSidebar"] {
            background: #f0f2f6;
        }
        [data-testid="stSidebar"] .block-container {
            padding-top: 2rem;
        }

        /* Metricas do Streamlit */
        [data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            padding: 15px 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        [data-testid="stMetricLabel"] {
            color: #64748b !important;
            font-size: 0.8rem !important;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.8rem !important;
            font-weight: 700 !important;
            color: #1e293b !important;
        }

        /* Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }
        .stTabs [data-baseweb="tab"] {
            background: #f1f5f9;
            border-radius: 8px;
            padding: 8px 16px;
            color: #64748b;
        }
        .stTabs [aria-selected="true"] {
            background: rgba(102, 126, 234, 0.15) !important;
            color: #667eea !important;
            font-weight: 600;
        }

        /* Dividers */
        hr {
            border-color: #e2e8f0;
        }

        /* ====== Tabelas HTML customizadas ====== */
        .dark-table-container {
            background: #ffffff;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            overflow: hidden;
            margin-bottom: 15px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        .dark-table-header {
            background: rgba(102, 126, 234, 0.1);
            padding: 10px 15px;
            font-weight: 600;
            font-size: 0.95rem;
            color: #1e293b;
            border-bottom: 1px solid #e2e8f0;
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
            color: #64748b;
            text-transform: uppercase;
            font-size: 0.7rem;
            letter-spacing: 0.5px;
            background: #f8fafc;
            border-bottom: 2px solid #e2e8f0;
        }
        .dark-table th.num { text-align: right; }
        .dark-table td {
            padding: 7px 12px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
        }
        .dark-table td.num { text-align: right; font-variant-numeric: tabular-nums; }
        .dark-table tr:hover { background: #f8fafc; }
        .dark-table tr.total-row {
            background: rgba(102, 126, 234, 0.08);
            font-weight: 700;
        }
        .dark-table tr.total-row td { color: #1e293b; border-top: 2px solid #e2e8f0; }
        .dark-table tr.subtotal-row {
            background: #f8fafc;
            font-weight: 600;
        }
        .dark-table tr.subtotal-row td { color: #334155; }
        .dark-table tr.segment-header td {
            color: #b45309;
            font-weight: 600;
            font-size: 0.8rem;
            padding-top: 12px;
            border-bottom: none;
        }
        .val-pos { color: #16a34a !important; }
        .val-neg { color: #dc2626 !important; }
        .val-warn { color: #b45309 !important; }

        /* Progress bar metas */
        .meta-card {
            background: #ffffff;
            border-radius: 12px;
            padding: 15px 18px;
            border: 1px solid #e2e8f0;
            min-width: 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        .meta-card-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }
        .meta-card-name { font-weight: 600; font-size: 0.95rem; color: #1e293b; }
        .meta-card-pct { font-size: 1.3rem; font-weight: 700; }
        .meta-bar {
            width: 100%;
            height: 10px;
            background: #e2e8f0;
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
            background: #ffffff;
            border-radius: 12px;
            padding: 18px 20px;
            border: 1px solid #cbd5e1;
            margin-top: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        .meta-total-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }
        .meta-total-name { font-weight: 700; font-size: 1.05rem; color: #1e293b; }
        .meta-total-pct { font-size: 1.5rem; font-weight: 700; }
        .meta-total-bar {
            width: 100%;
            height: 14px;
            background: #e2e8f0;
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
    </style>
    """, unsafe_allow_html=True)


def kpi_card(label, value, detail="", color="#1e293b"):
    """Renderiza um card KPI customizado."""
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" style="color: {color};">{value}</div>
        <div class="kpi-detail">{detail}</div>
    </div>
    """
