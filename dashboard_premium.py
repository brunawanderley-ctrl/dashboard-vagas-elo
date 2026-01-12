import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import sqlite3
import subprocess
import os
import base64
import io
from pathlib import Path
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Vagas Colégio Elo",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = Path(__file__).parent

# Toggle de tema (precisa vir antes do CSS)
if 'tema_escuro' not in st.session_state:
    st.session_state.tema_escuro = True

# CSS Dinâmico baseado no tema
if st.session_state.get('tema_toggle', True):
    # CSS Premium Dark Mode
    st.markdown("""
<style>
    /* Dark theme base */
    .stApp {
        background: linear-gradient(135deg, #0f0f1a 0%, #1a1a2e 50%, #16213e 100%);
    }

    /* Main container */
    .main .block-container {
        padding: 2rem 3rem;
        max-width: 1400px;
    }

    /* Headers */
    h1, h2, h3 {
        color: #ffffff !important;
        font-weight: 600 !important;
    }

    h2, h3 {
        color: #e0e0ff !important;
    }

    h1 {
        font-size: 2.5rem !important;
        background: linear-gradient(90deg, #8fa4f3 0%, #a78bda 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    /* Markdown headers */
    .stMarkdown h3 {
        color: #c4d0ff !important;
        font-size: 1.4rem !important;
    }

    /* Metric cards */
    [data-testid="stMetric"] {
        background: linear-gradient(145deg, #1e1e30 0%, #252540 100%);
        border: 1px solid rgba(102, 126, 234, 0.2);
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
    }

    [data-testid="stMetric"] label {
        color: #a0a0b0 !important;
        font-size: 0.9rem !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #ffffff !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
    }

    [data-testid="stMetric"] [data-testid="stMetricDelta"] {
        color: #4ade80 !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: rgba(30, 30, 48, 0.8);
        border-radius: 12px;
        padding: 0.5rem;
        gap: 0.5rem;
    }

    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #a0a0b0;
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 500;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white !important;
    }

    /* Expander */
    .streamlit-expanderHeader {
        background: rgba(30, 30, 48, 0.8);
        border-radius: 12px;
        color: #ffffff;
    }

    /* Dataframe */
    .stDataFrame {
        background: rgba(30, 30, 48, 0.8);
        border-radius: 12px;
    }

    /* Divider */
    hr {
        border-color: rgba(102, 126, 234, 0.2);
    }

    /* Caption */
    .stCaption {
        color: #606080 !important;
    }

    /* Button */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
    }

    /* Info box */
    .stAlert {
        background: rgba(102, 126, 234, 0.1);
        border: 1px solid rgba(102, 126, 234, 0.3);
        border-radius: 12px;
    }

    /* Plotly charts dark theme */
    .js-plotly-plot {
        border-radius: 16px;
    }

    /* Premium card class */
    .premium-card {
        background: linear-gradient(145deg, #1e1e30 0%, #252540 100%);
        border: 1px solid rgba(102, 126, 234, 0.2);
        border-radius: 20px;
        padding: 2rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
    }

    /* Glowing effect for important metrics */
    .glow {
        box-shadow: 0 0 20px rgba(102, 126, 234, 0.5);
    }
</style>
""", unsafe_allow_html=True)
else:
    # CSS Light Mode
    st.markdown("""
<style>
    /* Light theme base */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4e8f0 50%, #dce2ed 100%);
    }

    /* Main container */
    .main .block-container {
        padding: 2rem 3rem;
        max-width: 1400px;
    }

    /* Headers */
    h1, h2, h3 {
        color: #1e3a5f !important;
        font-weight: 600 !important;
    }

    h1 {
        font-size: 2.5rem !important;
        background: linear-gradient(90deg, #4a6fa5 0%, #6b5b95 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    /* Metric cards */
    [data-testid="stMetric"] {
        background: linear-gradient(145deg, #ffffff 0%, #f0f2f6 100%);
        border: 1px solid rgba(74, 111, 165, 0.2);
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    }

    [data-testid="stMetric"] label {
        color: #5a6a7a !important;
        font-size: 0.9rem !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #1e3a5f !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: rgba(255, 255, 255, 0.8);
        border-radius: 12px;
        padding: 0.5rem;
        gap: 0.5rem;
    }

    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #5a6a7a;
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 500;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #4a6fa5 0%, #6b5b95 100%);
        color: white !important;
    }

    /* Button */
    .stButton > button {
        background: linear-gradient(135deg, #4a6fa5 0%, #6b5b95 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(74, 111, 165, 0.3);
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(74, 111, 165, 0.4);
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #ffffff 0%, #f5f7fa 100%);
    }

    /* Divider */
    hr {
        border-color: rgba(74, 111, 165, 0.2);
    }
</style>
""", unsafe_allow_html=True)

# Layout do tema para gráficos Plotly (dinâmico)
PLOTLY_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#a0a0b0', family='Inter, sans-serif'),
    title=dict(font=dict(color='#ffffff', size=18)),
    xaxis=dict(
        gridcolor='rgba(102, 126, 234, 0.1)',
        linecolor='rgba(102, 126, 234, 0.2)',
        tickfont=dict(color='#a0a0b0')
    ),
    yaxis=dict(
        gridcolor='rgba(102, 126, 234, 0.1)',
        linecolor='rgba(102, 126, 234, 0.2)',
        tickfont=dict(color='#a0a0b0')
    ),
    legend=dict(
        bgcolor='rgba(0,0,0,0)',
        font=dict(color='#a0a0b0')
    ),
    margin=dict(t=60, b=40, l=40, r=40)
)

# Cores premium
COLORS = {
    'primary': '#667eea',
    'secondary': '#764ba2',
    'success': '#4ade80',
    'warning': '#fbbf24',
    'danger': '#f87171',
    'info': '#60a5fa',
    'gradient': ['#667eea', '#764ba2', '#a855f7', '#ec4899']
}

BASE_PATH = Path(__file__).parent / "output"

# Carrega dados atuais
@st.cache_data(ttl=60)
def carregar_dados():
    with open(BASE_PATH / "resumo_ultimo.json") as f:
        resumo = json.load(f)
    with open(BASE_PATH / "vagas_ultimo.json") as f:
        vagas = json.load(f)
    return resumo, vagas

# Carrega histórico do banco
@st.cache_data(ttl=60)
def carregar_historico():
    db_path = BASE_PATH / "vagas.db"
    if not db_path.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0

    conn = sqlite3.connect(db_path)

    query_unidades = """
    SELECT e.data_extracao, v.unidade_codigo, v.unidade_nome,
           SUM(v.vagas) as vagas, SUM(v.matriculados) as matriculados,
           SUM(v.novatos) as novatos, SUM(v.veteranos) as veteranos,
           SUM(v.disponiveis) as disponiveis
    FROM vagas v JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id, v.unidade_codigo ORDER BY e.data_extracao
    """
    df_unidades = pd.read_sql_query(query_unidades, conn)

    query_total = """
    SELECT e.data_extracao, SUM(v.vagas) as vagas, SUM(v.matriculados) as matriculados,
           SUM(v.novatos) as novatos, SUM(v.veteranos) as veteranos, SUM(v.disponiveis) as disponiveis
    FROM vagas v JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id ORDER BY e.data_extracao
    """
    df_total = pd.read_sql_query(query_total, conn)

    query_segmento = """
    SELECT e.data_extracao, v.segmento, SUM(v.vagas) as vagas,
           SUM(v.matriculados) as matriculados, SUM(v.disponiveis) as disponiveis
    FROM vagas v JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id, v.segmento ORDER BY e.data_extracao
    """
    df_segmento = pd.read_sql_query(query_segmento, conn)

    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM 'extrações'")
    num_extracoes = cursor.fetchone()[0]
    conn.close()

    for df in [df_unidades, df_total, df_segmento]:
        if not df.empty:
            df['data_extracao'] = pd.to_datetime(df['data_extracao'])
            df['data_formatada'] = df['data_extracao'].dt.strftime('%d/%m %H:%M')

    return df_unidades, df_total, df_segmento, num_extracoes

def criar_df_turmas(vagas_data):
    """Cria DataFrame com todas as turmas"""
    rows = []
    for unidade in vagas_data["unidades"]:
        for turma in unidade.get("turmas", []):
            rows.append({
                "Unidade": unidade["nome"],
                "Segmento": turma["segmento"],
                "Turma": turma["turma"],
                "Vagas": turma["vagas"],
                "Matriculados": turma["matriculados"],
                "Novatos": turma["novatos"],
                "Veteranos": turma["veteranos"],
                "Pre-matriculados": turma["pre_matriculados"],
                "Disponiveis": turma["disponiveis"],
            })
    return pd.DataFrame(rows)

def criar_df_resumo(resumo_data):
    """Cria DataFrame com resumo por unidade/segmento"""
    rows = []
    for unidade in resumo_data["unidades"]:
        for segmento, dados in unidade["segmentos"].items():
            rows.append({
                "Unidade": unidade["nome"],
                "Segmento": segmento,
                "Vagas": dados["vagas"],
                "Novatos": dados["novatos"],
                "Veteranos": dados["veteranos"],
                "Matriculados": dados["matriculados"],
                "Disponiveis": dados["disponiveis"],
            })
    return pd.DataFrame(rows)

def gerar_relatorio_pdf(resumo, df_perf, df_turmas, total):
    """Gera relatório PDF executivo em formato HTML para impressão"""
    data_hoje = datetime.now().strftime('%d/%m/%Y às %H:%M')
    data_extracao = resumo['data_extracao'][:16].replace('T', ' ')
    ocupacao_geral = round(total['matriculados'] / total['vagas'] * 100, 1) if total['vagas'] > 0 else 0
    ating_meta = round(total['matriculados'] / 4100 * 100, 1)
    ating_novatos = round(total['novatos'] / 1000 * 100, 1)

    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Relatório Executivo - Colégio Elo</title>
    <style>@page{{size:A4;margin:1.5cm}}body{{font-family:'Segoe UI',Arial,sans-serif;color:#1e3a5f;line-height:1.4;font-size:11px}}
    .header{{text-align:center;border-bottom:3px solid #667eea;padding-bottom:15px;margin-bottom:20px}}
    .header h1{{color:#667eea;margin:0;font-size:24px}}.header p{{color:#64748b;margin:5px 0 0 0;font-size:12px}}
    .kpi-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:20px}}
    .kpi-box{{background:linear-gradient(135deg,#f8fafc 0%,#e2e8f0 100%);padding:12px;border-radius:8px;text-align:center;border-left:4px solid #667eea}}
    .kpi-box.green{{border-left-color:#10b981}}.kpi-box.yellow{{border-left-color:#f59e0b}}.kpi-box.red{{border-left-color:#ef4444}}
    .kpi-label{{color:#64748b;font-size:9px;text-transform:uppercase}}.kpi-value{{color:#1e3a5f;font-size:22px;font-weight:700;margin:5px 0}}
    .kpi-detail{{color:#94a3b8;font-size:9px}}.section{{margin-bottom:20px}}.section h2{{color:#667eea;font-size:14px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;margin-bottom:10px}}
    table{{width:100%;border-collapse:collapse;font-size:10px}}th{{background:#667eea;color:white;padding:8px 5px;text-align:left}}
    td{{padding:6px 5px;border-bottom:1px solid #e2e8f0}}tr:nth-child(even){{background:#f8fafc}}
    .status-ok{{color:#10b981;font-weight:600}}.status-warn{{color:#f59e0b;font-weight:600}}.status-crit{{color:#ef4444;font-weight:600}}
    .unidade-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:15px}}
    .unidade-box{{background:#f8fafc;padding:10px;border-radius:8px;text-align:center}}
    .unidade-nome{{font-weight:600;color:#1e3a5f;font-size:11px}}.unidade-ating{{font-size:18px;font-weight:700;margin:5px 0}}
    .footer{{text-align:center;color:#94a3b8;font-size:9px;margin-top:20px;padding-top:10px;border-top:1px solid #e2e8f0}}
    @media print{{body{{-webkit-print-color-adjust:exact;print-color-adjust:exact}}}}</style></head><body>
    <div class="header"><h1>Relatório Executivo - Colégio Elo</h1><p>Período: {resumo['periodo']} | Gerado em: {data_hoje} | Dados de: {data_extracao}</p></div>
    <div class="kpi-grid">
    <div class="kpi-box {'green' if ocupacao_geral >= 80 else 'yellow' if ocupacao_geral >= 60 else 'red'}"><div class="kpi-label">Ocupação Geral</div><div class="kpi-value">{ocupacao_geral}%</div><div class="kpi-detail">{total['matriculados']:,} / {total['vagas']:,} vagas</div></div>
    <div class="kpi-box {'green' if ating_meta >= 100 else 'yellow' if ating_meta >= 80 else 'red'}"><div class="kpi-label">Meta Matrículas (4.100)</div><div class="kpi-value">{ating_meta}%</div><div class="kpi-detail">{total['matriculados']:,} alunos ({total['matriculados'] - 4100:+,})</div></div>
    <div class="kpi-box {'green' if ating_novatos >= 100 else 'yellow' if ating_novatos >= 80 else 'red'}"><div class="kpi-label">Meta Novatos (1.000)</div><div class="kpi-value">{ating_novatos}%</div><div class="kpi-detail">{total['novatos']:,} novatos ({total['novatos'] - 1000:+,})</div></div>
    </div>
    <div class="kpi-grid">
    <div class="kpi-box"><div class="kpi-label">Total Matriculados</div><div class="kpi-value">{total['matriculados']:,}</div></div>
    <div class="kpi-box"><div class="kpi-label">Veteranos</div><div class="kpi-value">{total['veteranos']:,}</div></div>
    <div class="kpi-box"><div class="kpi-label">Vagas Disponíveis</div><div class="kpi-value">{total['disponiveis']:,}</div></div>
    </div>
    <div class="section"><h2>Atingimento por Unidade</h2><div class="unidade-grid">"""

    for _, row in df_perf.iterrows():
        cor = '#10b981' if row['Gap'] >= 0 else '#f59e0b' if row['Atingimento'] >= 80 else '#ef4444'
        sinal = '+' if row['Gap'] >= 0 else ''
        sinal_nov = '+' if row['Gap_Novatos'] >= 0 else ''
        html += f"""<div class="unidade-box" style="border-left:4px solid {cor};"><div class="unidade-nome">{row['Nome_curto']}</div><div class="unidade-ating" style="color:{cor};">{row['Atingimento']:.1f}%</div><div style="font-size:9px;color:#64748b;">Matr: {int(row['Matriculados'])} / {int(row['Meta'])} ({sinal}{int(row['Gap'])})<br>Nov: {int(row['Novatos'])} / {int(row['Meta_Novatos'])} ({sinal_nov}{int(row['Gap_Novatos'])})</div></div>"""

    html += """</div></div><div class="section"><h2>Performance por Unidade</h2><table><thead><tr><th>Unidade</th><th>Vagas</th><th>Matr.</th><th>Ocupação</th><th>Ating.</th><th>Novatos</th><th>Meta Nov.</th><th>Vet.</th><th>Disp.</th></tr></thead><tbody>"""

    for _, row in df_perf.iterrows():
        status_class = 'status-ok' if row['Atingimento'] >= 100 else 'status-warn' if row['Atingimento'] >= 80 else 'status-crit'
        html += f"""<tr><td><strong>{row['Nome_curto']}</strong></td><td>{int(row['Vagas'])}</td><td>{int(row['Matriculados'])}</td><td>{row['Ocupacao']:.1f}%</td><td class="{status_class}">{row['Atingimento']:.1f}%</td><td>{int(row['Novatos'])}</td><td>{int(row['Meta_Novatos'])}</td><td>{int(row['Veteranos'])}</td><td>{int(row['Vagas']) - int(row['Matriculados'])}</td></tr>"""

    html += """</tbody></table></div>"""

    turmas_lotadas = df_turmas[df_turmas['Ocupação %'] >= 95].head(10)
    if len(turmas_lotadas) > 0:
        html += """<div class="section"><h2>Turmas Lotadas (≥95%)</h2><table><thead><tr><th>Unidade</th><th>Turma</th><th>Ocupação</th><th>Matr.</th><th>Vagas</th></tr></thead><tbody>"""
        for _, t in turmas_lotadas.iterrows():
            unidade_curta = t['Unidade'].split('(')[1].replace(')', '') if '(' in t['Unidade'] else t['Unidade']
            html += f"""<tr><td>{unidade_curta}</td><td>{t['Turma']}</td><td class="status-crit">{t['Ocupação %']:.0f}%</td><td>{int(t['Matriculados'])}</td><td>{int(t['Vagas'])}</td></tr>"""
        html += """</tbody></table></div>"""

    turmas_vazias = df_turmas[df_turmas['Ocupação %'] < 50].head(10)
    if len(turmas_vazias) > 0:
        html += """<div class="section"><h2>Turmas com Oportunidade (<50%)</h2><table><thead><tr><th>Unidade</th><th>Turma</th><th>Ocupação</th><th>Disponíveis</th></tr></thead><tbody>"""
        for _, t in turmas_vazias.iterrows():
            unidade_curta = t['Unidade'].split('(')[1].replace(')', '') if '(' in t['Unidade'] else t['Unidade']
            html += f"""<tr><td>{unidade_curta}</td><td>{t['Turma']}</td><td class="status-warn">{t['Ocupação %']:.0f}%</td><td>{int(t['Disponiveis'])}</td></tr>"""
        html += """</tbody></table></div>"""

    html += """<div class="footer"><p>Colégio Elo - Relatório Executivo Confidencial</p><p>Para imprimir: Ctrl+P (ou Cmd+P) → Salvar como PDF</p></div></body></html>"""
    return html.replace(",", ".")

try:
    resumo, vagas = carregar_dados()
    df_hist_unidades, df_hist_total, df_hist_segmento, num_extracoes = carregar_historico()
    df_turmas_all = criar_df_turmas(vagas)
    df_resumo_all = criar_df_resumo(resumo)
except FileNotFoundError:
    st.error("Arquivos de dados não encontrados. Execute a extração primeiro.")
    st.stop()

# ===== SIDEBAR - TEMA =====
tema_escuro = st.sidebar.toggle("Tema Escuro", value=True, key="tema_toggle")

st.sidebar.divider()

# ===== SIDEBAR - FILTROS =====
st.sidebar.header("Filtros")

# Filtro de Unidade
unidades_lista = ["Todas"] + list(df_resumo_all["Unidade"].unique())
unidade_selecionada = st.sidebar.selectbox("Unidade", unidades_lista)

# Filtro de Segmento
segmentos_lista = ["Todos"] + list(df_resumo_all["Segmento"].unique())
segmento_selecionado = st.sidebar.selectbox("Segmento", segmentos_lista)

# Filtro de Turno
def extrair_turno(turma_nome):
    turma_lower = turma_nome.lower()
    if "manhã" in turma_lower or "manha" in turma_lower:
        return "Manhã"
    elif "tarde" in turma_lower:
        return "Tarde"
    elif "integral" in turma_lower:
        return "Integral"
    else:
        return "Outro"

df_turmas_all["Turno"] = df_turmas_all["Turma"].apply(extrair_turno)
turnos_lista = ["Todos"] + list(df_turmas_all["Turno"].unique())
turno_selecionado = st.sidebar.selectbox("Turno", turnos_lista)

# Aplica filtros
df_resumo_filtrado = df_resumo_all.copy()
df_turmas_filtrado = df_turmas_all.copy()

if unidade_selecionada != "Todas":
    df_resumo_filtrado = df_resumo_filtrado[df_resumo_filtrado["Unidade"] == unidade_selecionada]
    df_turmas_filtrado = df_turmas_filtrado[df_turmas_filtrado["Unidade"] == unidade_selecionada]

if segmento_selecionado != "Todos":
    df_resumo_filtrado = df_resumo_filtrado[df_resumo_filtrado["Segmento"] == segmento_selecionado]
    df_turmas_filtrado = df_turmas_filtrado[df_turmas_filtrado["Segmento"] == segmento_selecionado]

if turno_selecionado != "Todos":
    df_turmas_filtrado = df_turmas_filtrado[df_turmas_filtrado["Turno"] == turno_selecionado]

# Seletor de Turma específica
st.sidebar.divider()
st.sidebar.header("Buscar Turma")
turmas_opcoes = ["Todas"] + sorted(df_turmas_filtrado["Turma"].unique().tolist())
turma_selecionada = st.sidebar.selectbox("Selecione a turma", turmas_opcoes)

st.sidebar.divider()

# ===== SIDEBAR - EXPORTAR =====
st.sidebar.header("Exportar")

csv = df_turmas_all.to_csv(index=False).encode('utf-8')
st.sidebar.download_button(
    label="Baixar CSV completo",
    data=csv,
    file_name=f"vagas_colegio_elo_{resumo['data_extracao'][:10]}.csv",
    mime="text/csv",
)

st.sidebar.divider()

# Info
st.sidebar.info(
    f"**Período:** {resumo['periodo']}\n\n"
    f"**Unidades:** {len(resumo['unidades'])}\n\n"
    f"**Total de turmas:** {len(df_turmas_all)}"
)

# Header Premium
col_title, col_btn = st.columns([5, 1])

with col_title:
    st.markdown("""
        <h1 style='margin-bottom: 0;'>Dashboard de Vagas</h1>
        <p style='color: #667eea; font-size: 1.2rem; margin-top: 0.5rem;'>Colégio Elo - Visão Executiva</p>
    """, unsafe_allow_html=True)

with col_btn:
    st.write("")
    if st.button("🔄 Atualizar", use_container_width=True):
        status_container = st.empty()
        status_container.info("⏳ Iniciando extração do SIGA...")

        try:
            # Tenta encontrar o Python do venv
            venv_python = BASE_DIR / "venv" / "bin" / "python"
            if not venv_python.exists():
                venv_python = "python3"  # Fallback para python do sistema

            extrator_script = BASE_DIR / "extrair_vagas.py"

            if not extrator_script.exists():
                status_container.error(f"Script não encontrado: {extrator_script}")
            else:
                status_container.info("⏳ Extraindo dados do SIGA... (pode levar alguns minutos)")

                result = subprocess.run(
                    [str(venv_python), str(extrator_script)],
                    capture_output=True,
                    text=True,
                    timeout=600,
                    cwd=str(BASE_DIR),
                    env={**os.environ, "PYTHONUNBUFFERED": "1"}
                )

                if result.returncode == 0:
                    status_container.success("✅ Dados atualizados com sucesso!")
                    st.cache_data.clear()
                    import time
                    time.sleep(1)
                    st.rerun()
                else:
                    status_container.error(f"❌ Erro na extração")
                    with st.expander("Ver detalhes do erro"):
                        st.code(result.stderr or result.stdout or "Sem detalhes")

        except subprocess.TimeoutExpired:
            status_container.error("⏰ Timeout: extração demorou mais de 10 minutos")
        except Exception as e:
            status_container.error(f"❌ Erro: {str(e)}")
            with st.expander("Ver detalhes"):
                import traceback
                st.code(traceback.format_exc())

# Info bar
st.markdown(f"""
    <div style='display: flex; gap: 2rem; color: #606080; font-size: 0.85rem; margin-bottom: 2rem;'>
        <span>📅 Última atualização: <strong style='color: #a0a0b0;'>{resumo['data_extracao'][:16].replace('T', ' ')}</strong></span>
        <span>📊 Período: <strong style='color: #a0a0b0;'>{resumo['periodo']}</strong></span>
        <span>🔢 Extrações: <strong style='color: #a0a0b0;'>{num_extracoes}</strong></span>
    </div>
""", unsafe_allow_html=True)

# Métricas principais (com filtro aplicado)
if unidade_selecionada != "Todas" or segmento_selecionado != "Todos":
    total = {
        "vagas": df_resumo_filtrado["Vagas"].sum(),
        "novatos": df_resumo_filtrado["Novatos"].sum(),
        "veteranos": df_resumo_filtrado["Veteranos"].sum(),
        "matriculados": df_resumo_filtrado["Matriculados"].sum(),
        "disponiveis": df_resumo_filtrado["Disponiveis"].sum(),
    }
else:
    total = resumo['total_geral']

ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1) if total['vagas'] > 0 else 0

col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    st.metric("OCUPAÇÃO", f"{ocupacao}%")
with col2:
    st.metric("MATRICULADOS", f"{total['matriculados']:,}".replace(",", "."))
with col3:
    st.metric("VETERANOS", f"{total['veteranos']:,}".replace(",", "."))
with col4:
    st.metric("NOVATOS", f"{total['novatos']:,}".replace(",", "."))
with col5:
    st.metric("VAGAS TOTAIS", f"{total['vagas']:,}".replace(",", "."))
with col6:
    st.metric("DISPONÍVEIS", f"{total['disponiveis']:,}".replace(",", "."))

st.markdown("<br>", unsafe_allow_html=True)

# Gráficos principais
col_left, col_right = st.columns(2)

with col_left:
    st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>Ocupação por Unidade</h3>", unsafe_allow_html=True)

    df_unidades = pd.DataFrame([
        {
            'Unidade': u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome'],
            'Ocupação': round(u['total']['matriculados'] / u['total']['vagas'] * 100, 1),
            'Matriculados': u['total']['matriculados'],
            'Vagas': u['total']['vagas']
        }
        for u in resumo['unidades']
    ])

    fig1 = go.Figure()

    # Barra de fundo (vagas totais)
    fig1.add_trace(go.Bar(
        name='Capacidade',
        x=df_unidades['Unidade'],
        y=[100] * len(df_unidades),
        marker_color='rgba(102, 126, 234, 0.15)',
        hoverinfo='skip'
    ))

    # Barra de ocupação - escala 6 cores
    def cor_ocupacao(o):
        if o >= 90: return '#065f46'    # Excelente (verde escuro) - 90-100%
        elif o >= 80: return '#22c55e'  # Boa (verde) - 80-89%
        elif o >= 70: return '#a3e635'  # Atenção (verde-amarelo) - 70-79%
        elif o >= 50: return '#facc15'  # Risco (amarelo) - 50-69%
        elif o >= 38: return '#f97316'  # Crítica (laranja) - 38-49%
        else: return '#dc2626'          # Congelada (vermelho) - 0-37%
    colors = [cor_ocupacao(o) for o in df_unidades['Ocupação']]

    fig1.add_trace(go.Bar(
        name='Ocupação',
        x=df_unidades['Unidade'],
        y=df_unidades['Ocupação'],
        marker_color=colors,
        text=df_unidades.apply(lambda r: f"{r['Ocupação']}%<br>({int(r['Matriculados'])})", axis=1),
        textposition='outside',
        textfont=dict(color='#ffffff', size=12, family='Inter')
    ))

    fig1.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#a0a0b0', family='Inter, sans-serif'),
        barmode='overlay',
        showlegend=False,
        height=380,
        yaxis=dict(gridcolor='rgba(102, 126, 234, 0.1)', range=[0, 120], title=''),
        xaxis=dict(gridcolor='rgba(102, 126, 234, 0.1)', title='')
    )

    st.plotly_chart(fig1, use_container_width=True)

with col_right:
    st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>Distribuição por Segmento</h3>", unsafe_allow_html=True)

    segmentos_total = {}
    for unidade in resumo['unidades']:
        for seg, vals in unidade['segmentos'].items():
            if seg not in segmentos_total:
                segmentos_total[seg] = {'matriculados': 0, 'vagas': 0}
            segmentos_total[seg]['matriculados'] += vals['matriculados']
            segmentos_total[seg]['vagas'] += vals['vagas']

    df_seg = pd.DataFrame([
        {'Segmento': seg, 'Matriculados': v['matriculados'], 'Vagas': v['vagas']}
        for seg, v in segmentos_total.items()
    ])

    ordem = ['Ed. Infantil', 'Fund. I', 'Fund. II', 'Ens. Médio']
    df_seg['ordem'] = df_seg['Segmento'].map({s: i for i, s in enumerate(ordem)})
    df_seg = df_seg.sort_values('ordem')

    fig2 = go.Figure()

    fig2.add_trace(go.Bar(
        name='Vagas',
        x=df_seg['Segmento'],
        y=df_seg['Vagas'],
        marker_color='rgba(102, 126, 234, 0.3)',
        text=df_seg['Vagas'],
        textposition='outside',
        textfont=dict(color='#667eea')
    ))

    fig2.add_trace(go.Bar(
        name='Matriculados',
        x=df_seg['Segmento'],
        y=df_seg['Matriculados'],
        marker=dict(
            color=df_seg['Matriculados'],
            colorscale=[[0, '#667eea'], [1, '#764ba2']]
        ),
        text=df_seg['Matriculados'],
        textposition='outside',
        textfont=dict(color='#ffffff')
    ))

    fig2.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#a0a0b0', family='Inter, sans-serif'),
        barmode='group',
        height=350,
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1,
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='#a0a0b0')
        )
    )

    st.plotly_chart(fig2, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== METAS POR UNIDADE =====
# Mapeamento por código da unidade (mais preciso)
METAS_POR_CODIGO = {
    "01-BV": 1280,   # Boa Viagem
    "02-CD": 1270,   # Candeias (Jaboatão)
    "03-JG": 750,    # Janga (Paulista)
    "04-CDR": 800,   # Cordeiro
}
# Metas de NOVATOS por unidade
METAS_NOVATOS = {
    "01-BV": 280,    # Boa Viagem (-25)
    "02-CD": 284,    # Candeias (-30)
    "03-JG": 204,    # Janga (+25)
    "04-CDR": 232,   # Cordeiro (+30)
}
META_NOVATOS = 1000
META_TOTAL = 1280 + 1270 + 750 + 800  # 4100

# ===== INSIGHTS EXECUTIVOS - CEO =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>💡 Insights Executivos</h3>", unsafe_allow_html=True)

# Calcula métricas por unidade com metas
df_perf_unidade = df_resumo_all.groupby('Unidade').agg({
    'Vagas': 'sum', 'Matriculados': 'sum', 'Novatos': 'sum', 'Veteranos': 'sum'
}).reset_index()

# Adiciona metas e calcula gaps
def get_meta(unidade_nome):
    # Usa código da unidade para precisão
    if "01-BV" in unidade_nome or "Boa Viagem" in unidade_nome:
        return 1280
    elif "02-CD" in unidade_nome or "Jaboatão" in unidade_nome or "Candeias" in unidade_nome:
        return 1270
    elif "03-JG" in unidade_nome or "Paulista" in unidade_nome or "Janga" in unidade_nome:
        return 750
    elif "04-CDR" in unidade_nome or "Cordeiro" in unidade_nome:
        return 800
    return 0

def get_meta_novatos(unidade_nome):
    # Meta de novatos por unidade
    if "01-BV" in unidade_nome or "Boa Viagem" in unidade_nome:
        return 305
    elif "02-CD" in unidade_nome or "Jaboatão" in unidade_nome or "Candeias" in unidade_nome:
        return 314
    elif "03-JG" in unidade_nome or "Paulista" in unidade_nome or "Janga" in unidade_nome:
        return 179
    elif "04-CDR" in unidade_nome or "Cordeiro" in unidade_nome:
        return 202
    return 0

df_perf_unidade['Meta'] = df_perf_unidade['Unidade'].apply(get_meta)
df_perf_unidade['Gap'] = df_perf_unidade['Matriculados'] - df_perf_unidade['Meta']
df_perf_unidade['Atingimento'] = (df_perf_unidade['Matriculados'] / df_perf_unidade['Meta'] * 100).round(1)
df_perf_unidade['Ocupacao'] = (df_perf_unidade['Matriculados'] / df_perf_unidade['Vagas'] * 100).round(1)

# Metas de novatos por unidade
df_perf_unidade['Meta_Novatos'] = df_perf_unidade['Unidade'].apply(get_meta_novatos)
df_perf_unidade['Gap_Novatos'] = df_perf_unidade['Novatos'] - df_perf_unidade['Meta_Novatos']
df_perf_unidade['Ating_Novatos'] = (df_perf_unidade['Novatos'] / df_perf_unidade['Meta_Novatos'] * 100).round(1)

# Extrai nome curto
df_perf_unidade['Nome_curto'] = df_perf_unidade['Unidade'].apply(
    lambda x: x.split('(')[1].replace(')', '') if '(' in x else x
)

# Calcula totais
total_meta = META_TOTAL
gap_total = total['matriculados'] - total_meta
atingimento_total = (total['matriculados'] / total_meta * 100) if total_meta > 0 else 0

gap_novatos = total['novatos'] - META_NOVATOS
atingimento_novatos = (total['novatos'] / META_NOVATOS * 100) if META_NOVATOS > 0 else 0

# Linha 1 - Metas gerais
col_meta1, col_meta2, col_meta3 = st.columns(3)

with col_meta1:
    cor_meta = '#10b981' if gap_total >= 0 else '#f59e0b' if atingimento_total >= 80 else '#ef4444'
    sinal = '+' if gap_total >= 0 else ''
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid {cor_meta};'>
        <p style='color: #94a3b8; font-size: 0.75rem; margin: 0; text-transform: uppercase;'>Meta Matrículas ({total_meta:,})</p>
        <p style='color: {cor_meta}; font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0;'>{atingimento_total:.1f}%</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>{sinal}{gap_total:,} alunos ({total['matriculados']:,}/{total_meta:,})</p>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

with col_meta2:
    cor_novatos = '#10b981' if gap_novatos >= 0 else '#f59e0b' if atingimento_novatos >= 80 else '#ef4444'
    sinal_nov = '+' if gap_novatos >= 0 else ''
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid {cor_novatos};'>
        <p style='color: #94a3b8; font-size: 0.75rem; margin: 0; text-transform: uppercase;'>Meta Novatos ({META_NOVATOS:,})</p>
        <p style='color: {cor_novatos}; font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0;'>{atingimento_novatos:.1f}%</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>{sinal_nov}{gap_novatos:,} novatos ({total['novatos']:,}/{META_NOVATOS:,})</p>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

with col_meta3:
    taxa_retencao_geral = (total['veteranos'] / total['matriculados'] * 100) if total['matriculados'] > 0 else 0
    cor_retencao = '#10b981' if taxa_retencao_geral >= 70 else '#f59e0b' if taxa_retencao_geral >= 50 else '#ef4444'
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid {cor_retencao};'>
        <p style='color: #94a3b8; font-size: 0.75rem; margin: 0; text-transform: uppercase;'>Retenção de Veteranos</p>
        <p style='color: {cor_retencao}; font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0;'>{taxa_retencao_geral:.1f}%</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>{total['veteranos']:,} veteranos renovaram</p>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

# Linha 2 - Metas por unidade (cards elegantes)
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div style='display: flex; align-items: center; margin-bottom: 1rem;'>
    <h4 style='margin: 0; color: #e0e0ff;'>🎯 Atingimento de Metas por Unidade</h4>
</div>
""", unsafe_allow_html=True)

cols_unidades = st.columns(4)

for i, (_, row) in enumerate(df_perf_unidade.iterrows()):
    with cols_unidades[i % 4]:
        cor = '#10b981' if row['Gap'] >= 0 else '#f59e0b' if row['Atingimento'] >= 80 else '#ef4444'
        sinal = '+' if row['Gap'] >= 0 else ''
        cor_nov = '#10b981' if row['Gap_Novatos'] >= 0 else '#f59e0b' if row['Ating_Novatos'] >= 80 else '#ef4444'
        sinal_nov = '+' if row['Gap_Novatos'] >= 0 else ''

        # Ícone baseado no status
        icone = '✅' if row['Gap'] >= 0 else '⚠️' if row['Atingimento'] >= 80 else '🔴'

        # Barra de progresso visual
        progresso = min(row['Atingimento'], 100)

        st.markdown(f"""
        <div style='background: linear-gradient(145deg, #0f172a 0%, #1e293b 100%); padding: 1.2rem; border-radius: 16px; border: 1px solid rgba(102, 126, 234, 0.2); box-shadow: 0 4px 20px rgba(0,0,0,0.3); margin-bottom: 0.5rem;'>
            <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;'>
                <span style='color: #ffffff; font-size: 1.1rem; font-weight: 700;'>{row['Nome_curto']}</span>
                <span style='font-size: 1.2rem;'>{icone}</span>
            </div>
            <p style='color: {cor}; font-size: 2.2rem; font-weight: 800; margin: 0.3rem 0; text-align: center;'>{row['Atingimento']:.1f}%</p>
            <div style='background: rgba(255,255,255,0.1); border-radius: 10px; height: 8px; margin: 0.5rem 0; overflow: hidden;'>
                <div style='background: linear-gradient(90deg, {cor} 0%, {cor}aa 100%); width: {progresso}%; height: 100%; border-radius: 10px;'></div>
            </div>
            <div style='display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem; margin-top: 0.8rem;'>
                <div style='background: rgba(255,255,255,0.05); padding: 0.4rem; border-radius: 8px; text-align: center;'>
                    <p style='color: #94a3b8; font-size: 0.6rem; margin: 0; text-transform: uppercase;'>Matrículas</p>
                    <p style='color: #ffffff; font-size: 0.9rem; font-weight: 600; margin: 0;'>{int(row['Matriculados'])}<span style='color: #64748b; font-size: 0.7rem;'>/{int(row['Meta'])}</span></p>
                    <p style='color: {cor}; font-size: 0.7rem; margin: 0;'>{sinal}{int(row['Gap'])}</p>
                </div>
                <div style='background: rgba(255,255,255,0.05); padding: 0.4rem; border-radius: 8px; text-align: center;'>
                    <p style='color: #94a3b8; font-size: 0.6rem; margin: 0; text-transform: uppercase;'>Novatos</p>
                    <p style='color: #ffffff; font-size: 0.9rem; font-weight: 600; margin: 0;'>{int(row['Novatos'])}<span style='color: #64748b; font-size: 0.7rem;'>/{int(row['Meta_Novatos'])}</span></p>
                    <p style='color: {cor_nov}; font-size: 0.7rem; margin: 0;'>{sinal_nov}{int(row['Gap_Novatos'])}</p>
                </div>
            </div>
            <p style='color: #64748b; font-size: 0.65rem; margin: 0.5rem 0 0 0; text-align: center;'>Veteranos: {int(row['Veteranos'])} | Ocupação: {row['Ocupacao']:.0f}%</p>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== TERMÔMETRO DE METAS POR UNIDADE =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>🎯 Termômetro de Metas por Unidade</h3>", unsafe_allow_html=True)

# Função para cor do termômetro de metas (escala 6 cores)
def cor_meta(atingimento):
    if atingimento >= 100: return '#065f46'   # Meta atingida (verde escuro)
    elif atingimento >= 90: return '#22c55e'  # Quase lá (verde)
    elif atingimento >= 80: return '#a3e635'  # Bom progresso (verde-amarelo)
    elif atingimento >= 60: return '#facc15'  # Atenção (amarelo)
    elif atingimento >= 40: return '#f97316'  # Risco (laranja)
    else: return '#dc2626'                    # Crítico (vermelho)

def status_meta(atingimento):
    if atingimento >= 100: return 'Atingida'
    elif atingimento >= 90: return 'Quase lá'
    elif atingimento >= 80: return 'Bom'
    elif atingimento >= 60: return 'Atenção'
    elif atingimento >= 40: return 'Risco'
    else: return 'Crítico'

# Termômetros de Matrículas
st.markdown("<h4 style='color: #e2e8f0; font-weight: 500; margin-top: 10px;'>Meta de Matrículas</h4>", unsafe_allow_html=True)
cols_meta_mat = st.columns(len(df_perf_unidade))

for idx, (_, row) in enumerate(df_perf_unidade.iterrows()):
    atingimento = row['Atingimento']
    cor = cor_meta(atingimento)
    status = status_meta(atingimento)
    gap = int(row['Gap'])
    sinal_gap = '+' if gap >= 0 else ''

    with cols_meta_mat[idx]:
        st.markdown(f"""
        <div style='text-align: center; padding: 10px;'>
            <div style='font-size: 14px; color: #e2e8f0; font-weight: 600; margin-bottom: 8px;'>{row['Nome_curto']}</div>
            <div style='position: relative; width: 50px; height: 160px; margin: 0 auto; background: linear-gradient(to top, #1a1a2e 0%, #2d2d44 100%); border-radius: 25px; border: 2px solid #3d3d5c; overflow: hidden;'>
                <div style='position: absolute; bottom: 0; width: 100%; height: {min(atingimento, 100)}%; background: linear-gradient(to top, {cor}, {cor}dd); border-radius: 0 0 23px 23px; transition: height 0.5s;'></div>
                <div style='position: absolute; width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;'>
                    <span style='font-size: 16px; font-weight: bold; color: white; text-shadow: 1px 1px 2px rgba(0,0,0,0.8);'>{atingimento:.0f}%</span>
                </div>
            </div>
            <div style='margin-top: 8px;'>
                <span style='background: {cor}; color: {"white" if atingimento < 80 or atingimento >= 90 else "#1a1a2e"}; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 500;'>{status}</span>
            </div>
            <div style='font-size: 11px; color: #a0a0b0; margin-top: 5px;'>{int(row['Matriculados'])} / {int(row['Meta'])}</div>
            <div style='font-size: 10px; color: {cor}; font-weight: 600;'>Gap: {sinal_gap}{gap}</div>
        </div>
        """, unsafe_allow_html=True)

# Termômetros de Novatos
st.markdown("<h4 style='color: #e2e8f0; font-weight: 500; margin-top: 20px;'>Meta de Novatos</h4>", unsafe_allow_html=True)
cols_meta_nov = st.columns(len(df_perf_unidade))

for idx, (_, row) in enumerate(df_perf_unidade.iterrows()):
    atingimento_nov = (row['Novatos'] / row['Meta_Novatos'] * 100) if row['Meta_Novatos'] > 0 else 0
    cor = cor_meta(atingimento_nov)
    status = status_meta(atingimento_nov)
    gap_nov = int(row['Novatos'] - row['Meta_Novatos'])
    sinal_gap = '+' if gap_nov >= 0 else ''

    with cols_meta_nov[idx]:
        st.markdown(f"""
        <div style='text-align: center; padding: 10px;'>
            <div style='font-size: 14px; color: #e2e8f0; font-weight: 600; margin-bottom: 8px;'>{row['Nome_curto']}</div>
            <div style='position: relative; width: 50px; height: 160px; margin: 0 auto; background: linear-gradient(to top, #1a1a2e 0%, #2d2d44 100%); border-radius: 25px; border: 2px solid #3d3d5c; overflow: hidden;'>
                <div style='position: absolute; bottom: 0; width: 100%; height: {min(atingimento_nov, 100)}%; background: linear-gradient(to top, {cor}, {cor}dd); border-radius: 0 0 23px 23px; transition: height 0.5s;'></div>
                <div style='position: absolute; width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;'>
                    <span style='font-size: 16px; font-weight: bold; color: white; text-shadow: 1px 1px 2px rgba(0,0,0,0.8);'>{atingimento_nov:.0f}%</span>
                </div>
            </div>
            <div style='margin-top: 8px;'>
                <span style='background: {cor}; color: {"white" if atingimento_nov < 80 or atingimento_nov >= 90 else "#1a1a2e"}; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 500;'>{status}</span>
            </div>
            <div style='font-size: 11px; color: #a0a0b0; margin-top: 5px;'>{int(row['Novatos'])} / {int(row['Meta_Novatos'])}</div>
            <div style='font-size: 10px; color: {cor}; font-weight: 600;'>Gap: {sinal_gap}{gap_nov}</div>
        </div>
        """, unsafe_allow_html=True)

# Legenda das metas
st.markdown("""
<div style='display: flex; flex-wrap: wrap; gap: 8px; justify-content: center; margin-top: 15px;'>
    <span style='background: #065f46; color: white; padding: 4px 10px; border-radius: 4px; font-size: 11px;'>≥100% Atingida</span>
    <span style='background: #22c55e; color: white; padding: 4px 10px; border-radius: 4px; font-size: 11px;'>90-99% Quase lá</span>
    <span style='background: #a3e635; color: #1a1a2e; padding: 4px 10px; border-radius: 4px; font-size: 11px;'>80-89% Bom</span>
    <span style='background: #facc15; color: #1a1a2e; padding: 4px 10px; border-radius: 4px; font-size: 11px;'>60-79% Atenção</span>
    <span style='background: #f97316; color: white; padding: 4px 10px; border-radius: 4px; font-size: 11px;'>40-59% Risco</span>
    <span style='background: #dc2626; color: white; padding: 4px 10px; border-radius: 4px; font-size: 11px;'>&lt;40% Crítico</span>
</div>
""", unsafe_allow_html=True)

# ===== ALERTAS DE AÇÃO IMEDIATA POR UNIDADE =====
st.markdown("""
<h3 style='color: #ffffff; font-weight: 700;'>⚠️ Alertas de Ação por Unidade</h3>
""", unsafe_allow_html=True)

# Seletor de quantidade
col_config1, col_config2 = st.columns([3, 1])
with col_config2:
    qtd_alertas = st.selectbox("Exibir", [5, 10, 15, 20, "Todos"], index=0, key="qtd_alertas")

# Turmas com ocupação calculada
turmas_criticas = df_turmas_all.copy()
turmas_criticas['Ocupacao'] = (turmas_criticas['Matriculados'] / turmas_criticas['Vagas'] * 100).round(1)
turmas_criticas['Unidade_curta'] = turmas_criticas['Unidade'].apply(
    lambda x: x.split('(')[1].replace(')', '') if '(' in x else x
)

# Tabs por unidade
unidades_unicas = sorted(turmas_criticas['Unidade_curta'].unique())
tabs_alertas = st.tabs(unidades_unicas)

for i, tab in enumerate(tabs_alertas):
    with tab:
        unidade = unidades_unicas[i]
        df_unidade = turmas_criticas[turmas_criticas['Unidade_curta'] == unidade]

        # Turmas lotadas e vazias desta unidade
        lotadas = df_unidade[df_unidade['Ocupacao'] >= 95].sort_values('Ocupacao', ascending=False)
        vazias = df_unidade[df_unidade['Ocupacao'] < 50].sort_values('Ocupacao')

        # Limita quantidade
        limite = None if qtd_alertas == "Todos" else qtd_alertas
        lotadas_exibir = lotadas if limite is None else lotadas.head(limite)
        vazias_exibir = vazias if limite is None else vazias.head(limite)

        col_a1, col_a2 = st.columns(2)

        with col_a1:
            st.markdown(f"""
            <div style='background: rgba(239, 68, 68, 0.15); padding: 1rem; border-radius: 12px; border-left: 5px solid #ef4444; margin-bottom: 0.8rem;'>
                <p style='color: #ffffff; font-weight: 700; font-size: 1rem; margin: 0;'>🔴 TURMAS LOTADAS (≥95%)</p>
                <p style='color: #fca5a5; font-size: 0.85rem; margin: 0.3rem 0 0 0;'>{len(lotadas)} turmas encontradas</p>
            </div>
            """, unsafe_allow_html=True)

            if len(lotadas_exibir) > 0:
                for _, t in lotadas_exibir.iterrows():
                    st.markdown(f"""
                    <div style='background: rgba(30, 41, 59, 0.9); padding: 0.7rem 1rem; border-radius: 10px; margin-bottom: 0.4rem; border: 1px solid rgba(239, 68, 68, 0.3);'>
                        <div style='display: flex; justify-content: space-between; align-items: center;'>
                            <span style='color: #ffffff; font-weight: 600;'>{t['Turma']}</span>
                            <span style='color: #ef4444; font-weight: 800; font-size: 1.1rem;'>{t['Ocupacao']:.0f}%</span>
                        </div>
                        <p style='color: #94a3b8; font-size: 0.8rem; margin: 0.3rem 0 0 0;'>{t['Segmento']} • {int(t['Matriculados'])}/{int(t['Vagas'])} alunos</p>
                    </div>
                    """, unsafe_allow_html=True)
                if limite and len(lotadas) > limite:
                    st.caption(f"... e mais {len(lotadas) - limite} turmas")
            else:
                st.markdown("<p style='color: #10b981; font-size: 0.9rem; padding: 0.5rem;'>✅ Nenhuma turma lotada nesta unidade</p>", unsafe_allow_html=True)

        with col_a2:
            st.markdown(f"""
            <div style='background: rgba(251, 191, 36, 0.15); padding: 1rem; border-radius: 12px; border-left: 5px solid #f59e0b; margin-bottom: 0.8rem;'>
                <p style='color: #ffffff; font-weight: 700; font-size: 1rem; margin: 0;'>🟡 BAIXA OCUPAÇÃO (<50%)</p>
                <p style='color: #fcd34d; font-size: 0.85rem; margin: 0.3rem 0 0 0;'>{len(vazias)} turmas - foco em captação</p>
            </div>
            """, unsafe_allow_html=True)

            if len(vazias_exibir) > 0:
                for _, t in vazias_exibir.iterrows():
                    vagas_disp = int(t['Vagas'] - t['Matriculados'])
                    st.markdown(f"""
                    <div style='background: rgba(30, 41, 59, 0.9); padding: 0.7rem 1rem; border-radius: 10px; margin-bottom: 0.4rem; border: 1px solid rgba(251, 191, 36, 0.3);'>
                        <div style='display: flex; justify-content: space-between; align-items: center;'>
                            <span style='color: #ffffff; font-weight: 600;'>{t['Turma']}</span>
                            <span style='color: #f59e0b; font-weight: 800; font-size: 1.1rem;'>{vagas_disp} vagas</span>
                        </div>
                        <p style='color: #94a3b8; font-size: 0.8rem; margin: 0.3rem 0 0 0;'>{t['Segmento']} • {t['Ocupacao']:.0f}% ocupação</p>
                    </div>
                    """, unsafe_allow_html=True)
                if limite and len(vazias) > limite:
                    st.caption(f"... e mais {len(vazias) - limite} turmas")
            else:
                st.markdown("<p style='color: #10b981; font-size: 0.9rem; padding: 0.5rem;'>✅ Todas as turmas acima de 50%</p>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== GRÁFICO DE OCUPAÇÃO POR UNIDADE/SEGMENTO =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>Taxa de Ocupação por Unidade e Segmento</h3>", unsafe_allow_html=True)

df_ocupacao = df_resumo_filtrado.copy()
df_ocupacao["Ocupacao"] = (df_ocupacao["Matriculados"] / df_ocupacao["Vagas"] * 100).round(1)

fig_ocup = px.bar(
    df_ocupacao,
    x="Unidade" if unidade_selecionada == "Todas" else "Segmento",
    y="Ocupacao",
    color="Segmento" if unidade_selecionada == "Todas" else "Unidade",
    barmode="group",
    color_discrete_sequence=[COLORS['primary'], COLORS['success'], COLORS['warning'], COLORS['danger']],
    labels={"Ocupacao": "Taxa de Ocupação (%)"},
    text="Ocupacao"
)
fig_ocup.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
fig_ocup.update_layout(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#a0a0b0', family='Inter, sans-serif'),
    height=400,
    yaxis=dict(gridcolor='rgba(102, 126, 234, 0.1)', range=[0, 120]),
    xaxis=dict(gridcolor='rgba(102, 126, 234, 0.1)'),
    legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#a0a0b0'))
)
st.plotly_chart(fig_ocup, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== NOVATOS vs VETERANOS =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>Composição: Novatos vs Veteranos</h3>", unsafe_allow_html=True)

col_nv1, col_nv2 = st.columns(2)

with col_nv1:
    # Pizza geral
    fig_pizza_nv = go.Figure(data=[go.Pie(
        labels=["Novatos", "Veteranos"],
        values=[total["novatos"], total["veteranos"]],
        hole=0.4,
        marker_colors=[COLORS['warning'], COLORS['primary']]
    )])
    fig_pizza_nv.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#a0a0b0', family='Inter, sans-serif'),
        title_text="Distribuição Geral",
        title_font_color='#ffffff',
        height=350,
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#a0a0b0'))
    )
    st.plotly_chart(fig_pizza_nv, use_container_width=True)

with col_nv2:
    # Barra por unidade/segmento
    df_nv = df_resumo_filtrado.groupby("Unidade" if unidade_selecionada == "Todas" else "Segmento").agg({
        "Novatos": "sum",
        "Veteranos": "sum"
    }).reset_index()

    fig_nv_bar = px.bar(
        df_nv,
        x="Unidade" if unidade_selecionada == "Todas" else "Segmento",
        y=["Novatos", "Veteranos"],
        barmode="stack",
        color_discrete_map={"Novatos": COLORS['warning'], "Veteranos": COLORS['primary']},
        labels={"value": "Quantidade", "variable": "Tipo"}
    )
    fig_nv_bar.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#a0a0b0', family='Inter, sans-serif'),
        height=350,
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#a0a0b0'), title='')
    )
    st.plotly_chart(fig_nv_bar, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== TERMÔMETRO DE OCUPAÇÃO POR UNIDADE =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>🌡️ Termômetro de Ocupação por Unidade</h3>", unsafe_allow_html=True)

# Função para determinar cor do termômetro - escala 6 cores
def cor_termometro(ocupacao):
    if ocupacao >= 90: return '#065f46'    # Excelente (verde escuro) - 90-100%
    elif ocupacao >= 80: return '#22c55e'  # Boa (verde) - 80-89%
    elif ocupacao >= 70: return '#a3e635'  # Atenção (verde-amarelo) - 70-79%
    elif ocupacao >= 50: return '#facc15'  # Risco (amarelo) - 50-69%
    elif ocupacao >= 38: return '#f97316'  # Crítica (laranja) - 38-49%
    else: return '#dc2626'                 # Congelada (vermelho) - 0-37%

def classificacao_termometro(ocupacao):
    if ocupacao >= 90: return 'Excelente'
    elif ocupacao >= 80: return 'Boa'
    elif ocupacao >= 70: return 'Atenção'
    elif ocupacao >= 50: return 'Risco'
    elif ocupacao >= 38: return 'Crítica'
    else: return 'Congelada'

# Calcula ocupação por unidade
df_termo = df_unidades.copy()
cols_termo = st.columns(len(df_termo))

for idx, (_, row) in enumerate(df_termo.iterrows()):
    ocupacao = row['Ocupação']
    cor = cor_termometro(ocupacao)
    classif = classificacao_termometro(ocupacao)
    unidade_nome = row['Unidade'].split('(')[1].replace(')', '') if '(' in row['Unidade'] else row['Unidade']

    with cols_termo[idx]:
        # Termômetro visual
        st.markdown(f"""
        <div style='text-align: center; padding: 10px;'>
            <div style='font-size: 14px; color: #e2e8f0; font-weight: 600; margin-bottom: 8px;'>{unidade_nome}</div>
            <div style='position: relative; width: 50px; height: 180px; margin: 0 auto; background: linear-gradient(to top, #1a1a2e 0%, #2d2d44 100%); border-radius: 25px; border: 2px solid #3d3d5c; overflow: hidden;'>
                <div style='position: absolute; bottom: 0; width: 100%; height: {min(ocupacao, 100)}%; background: linear-gradient(to top, {cor}, {cor}dd); border-radius: 0 0 23px 23px; transition: height 0.5s;'></div>
                <div style='position: absolute; width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;'>
                    <span style='font-size: 18px; font-weight: bold; color: white; text-shadow: 1px 1px 2px rgba(0,0,0,0.8);'>{ocupacao:.0f}%</span>
                </div>
            </div>
            <div style='margin-top: 8px;'>
                <span style='background: {cor}; color: {"white" if ocupacao < 70 or ocupacao >= 90 else "#1a1a2e"}; padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 500;'>{classif}</span>
            </div>
            <div style='font-size: 11px; color: #a0a0b0; margin-top: 5px;'>{int(row['Matriculados'])} / {int(row['Vagas'])}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== MAPA DE CALOR DE OCUPAÇÃO =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>📊 Mapa de Calor - Ocupação por Unidade e Segmento</h3>", unsafe_allow_html=True)

# Prepara dados para heatmap
df_heatmap = df_resumo_all.copy()
df_heatmap['Ocupacao'] = (df_heatmap['Matriculados'] / df_heatmap['Vagas'] * 100).round(1)

# Extrai nome curto da unidade
df_heatmap['Unidade_curta'] = df_heatmap['Unidade'].apply(
    lambda x: x.split('(')[1].replace(')', '') if '(' in x else x
)

# Cria pivot table para heatmap
pivot_ocupacao = df_heatmap.pivot_table(
    values='Ocupacao',
    index='Segmento',
    columns='Unidade_curta',
    aggfunc='mean'
).round(1)

# Ordena segmentos
ordem_seg = ['Ed. Infantil', 'Fund. I', 'Fund. II', 'Ens. Médio']
pivot_ocupacao = pivot_ocupacao.reindex([s for s in ordem_seg if s in pivot_ocupacao.index])

fig_heatmap = go.Figure(data=go.Heatmap(
    z=pivot_ocupacao.values,
    x=pivot_ocupacao.columns.tolist(),
    y=pivot_ocupacao.index.tolist(),
    zmin=0,    # Força escala a começar em 0%
    zmax=100,  # Força escala a terminar em 100%
    colorscale=[
        [0, '#dc2626'],      # 0-37%: Congelada (vermelho)
        [0.37, '#dc2626'],   # 37%: fim Congelada
        [0.38, '#f97316'],   # 38-49%: Crítica (laranja)
        [0.49, '#f97316'],   # 49%: fim Crítica
        [0.50, '#facc15'],   # 50-69%: Risco (amarelo)
        [0.69, '#facc15'],   # 69%: fim Risco
        [0.70, '#a3e635'],   # 70-79%: Atenção (verde-amarelo)
        [0.79, '#a3e635'],   # 79%: fim Atenção
        [0.80, '#22c55e'],   # 80-89%: Boa (verde)
        [0.89, '#22c55e'],   # 89%: fim Boa
        [0.90, '#065f46'],   # 90-100%: Excelente (verde escuro)
        [1.0, '#065f46']     # 100%: Excelente
    ],
    text=pivot_ocupacao.values,
    texttemplate='%{text:.1f}%',
    textfont={"size": 14, "color": "white"},
    hovertemplate='Unidade: %{x}<br>Segmento: %{y}<br>Ocupação: %{z:.1f}%<extra></extra>',
    colorbar=dict(
        title=dict(text='Ocupação %', font=dict(color='#a0a0b0')),
        tickfont=dict(color='#a0a0b0'),
        tickvals=[0, 38, 50, 70, 80, 90, 100],
        ticktext=['0%', '38%', '50%', '70%', '80%', '90%', '100%']
    )
))

fig_heatmap.update_layout(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#a0a0b0', family='Inter, sans-serif'),
    height=350,
    xaxis=dict(tickfont=dict(color='#e0e0ff', size=12)),
    yaxis=dict(tickfont=dict(color='#e0e0ff', size=12))
)

st.plotly_chart(fig_heatmap, use_container_width=True)

# Legenda das faixas de ocupação
st.markdown("""
<div style='display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; margin-top: 10px;'>
    <span style='background: #065f46; color: white; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 500;'>90-100% Excelente</span>
    <span style='background: #22c55e; color: white; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 500;'>80-89% Boa</span>
    <span style='background: #a3e635; color: #1a1a2e; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 500;'>70-79% Atenção</span>
    <span style='background: #facc15; color: #1a1a2e; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 500;'>50-69% Risco</span>
    <span style='background: #f97316; color: white; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 500;'>38-49% Crítica</span>
    <span style='background: #dc2626; color: white; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: 500;'>0-37% Congelada</span>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Seção de histórico
if num_extracoes >= 2:
    st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>📈 Evolução Histórica</h3>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Visão Geral", "Por Unidade"])

    with tab1:
        df_hist_total['ocupacao'] = round(df_hist_total['matriculados'] / df_hist_total['vagas'] * 100, 1)

        fig_hist = go.Figure()

        fig_hist.add_trace(go.Scatter(
            x=df_hist_total['data_formatada'],
            y=df_hist_total['ocupacao'],
            mode='lines+markers',
            name='Ocupação',
            line=dict(color=COLORS['primary'], width=3),
            marker=dict(size=10, color=COLORS['primary']),
            fill='tozeroy',
            fillcolor='rgba(102, 126, 234, 0.1)'
        ))

        fig_hist.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#a0a0b0', family='Inter, sans-serif'),
            height=300,
            yaxis=dict(gridcolor='rgba(102, 126, 234, 0.1)', title='Ocupação %', range=[0, 100])
        )

        st.plotly_chart(fig_hist, use_container_width=True)

    with tab2:
        fig_unid = go.Figure()
        cores_unid = [COLORS['primary'], COLORS['success'], COLORS['warning'], '#ec4899']

        for i, unidade in enumerate(df_hist_unidades['unidade_nome'].unique()):
            df_u = df_hist_unidades[df_hist_unidades['unidade_nome'] == unidade]
            nome = unidade.split('(')[1].replace(')', '') if '(' in unidade else unidade

            fig_unid.add_trace(go.Scatter(
                x=df_u['data_formatada'],
                y=df_u['matriculados'],
                mode='lines+markers',
                name=nome,
                line=dict(color=cores_unid[i % len(cores_unid)], width=2),
                marker=dict(size=8)
            ))

        fig_unid.update_layout(**PLOTLY_LAYOUT, height=300, hovermode='x unified')
        st.plotly_chart(fig_unid, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# Detalhamento por unidade
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>🏫 Detalhamento por Unidade</h3>", unsafe_allow_html=True)

tabs = st.tabs([u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome']
                for u in resumo['unidades']])

for i, tab in enumerate(tabs):
    with tab:
        unidade = resumo['unidades'][i]
        unidade_vagas = vagas['unidades'][i]

        t = unidade['total']
        ocup = round(t['matriculados'] / t['vagas'] * 100, 1)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ocupação", f"{ocup}%")
        c2.metric("Matriculados", t['matriculados'])
        c3.metric("Disponíveis", t['disponiveis'])
        c4.metric("Novatos / Veteranos", f"{t['novatos']} / {t['veteranos']}")

        col_a, col_b = st.columns([2, 1])

        with col_a:
            df_seg_u = pd.DataFrame([
                {'Segmento': seg, **vals}
                for seg, vals in unidade['segmentos'].items()
            ])

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=df_seg_u['Segmento'],
                y=df_seg_u['vagas'],
                name='Vagas',
                marker_color='rgba(102, 126, 234, 0.3)'
            ))
            fig.add_trace(go.Bar(
                x=df_seg_u['Segmento'],
                y=df_seg_u['matriculados'],
                name='Matriculados',
                marker_color=COLORS['primary']
            ))
            fig.update_layout(**PLOTLY_LAYOUT, height=280, barmode='group')
            st.plotly_chart(fig, use_container_width=True)

        with col_b:
            fig_pie = go.Figure(data=[go.Pie(
                labels=['Novatos', 'Veteranos'],
                values=[t['novatos'], t['veteranos']],
                hole=.6,
                marker_colors=[COLORS['warning'], COLORS['primary']]
            )])
            fig_pie.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#a0a0b0', family='Inter, sans-serif'),
                height=280,
                showlegend=True,
                legend=dict(orientation='h', yanchor='bottom', y=-0.2, bgcolor='rgba(0,0,0,0)')
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with st.expander("📋 Ver todas as turmas"):
            df_turmas = pd.DataFrame(unidade_vagas['turmas'])
            # Calcula ocupação
            df_turmas['ocupacao'] = (df_turmas['matriculados'] / df_turmas['vagas'] * 100).round(1)
            df_turmas = df_turmas[['segmento', 'turma', 'vagas', 'matriculados', 'ocupacao', 'disponiveis', 'pre_matriculados']]
            df_turmas.columns = ['Segmento', 'Turma', 'Vagas', 'Matr.', 'Ocup.%', 'Disp.', 'Pré-Matr.']
            st.dataframe(df_turmas, use_container_width=True, hide_index=True)

# ===== PAINEL EXECUTIVO - CEO =====
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>📊 Painel Executivo</h3>", unsafe_allow_html=True)

# Prepara DataFrame com todas as informações
df_relatorio = df_turmas_filtrado.copy()
df_relatorio['Ocupação %'] = (df_relatorio['Matriculados'] / df_relatorio['Vagas'] * 100).round(1)

# Filtra se uma turma específica foi selecionada
if turma_selecionada != "Todas":
    df_relatorio = df_relatorio[df_relatorio['Turma'] == turma_selecionada]

# ===== KPIs ESTRATÉGICOS =====
st.markdown("<h4 style='color: #e2e8f0; font-weight: 600;'>Indicadores Estratégicos</h4>", unsafe_allow_html=True)

# Calcula KPIs
total_vagas = df_relatorio['Vagas'].sum()
total_matriculados = df_relatorio['Matriculados'].sum()
total_novatos = df_relatorio['Novatos'].sum()
total_veteranos = df_relatorio['Veteranos'].sum()
total_disponiveis = df_relatorio['Disponiveis'].sum()
total_pre_matr = df_relatorio['Pre-matriculados'].sum()

taxa_ocupacao = (total_matriculados / total_vagas * 100) if total_vagas > 0 else 0
taxa_retencao = (total_veteranos / total_matriculados * 100) if total_matriculados > 0 else 0
taxa_captacao = (total_novatos / total_vagas * 100) if total_vagas > 0 else 0
taxa_conversao_pre = (total_pre_matr / total_disponiveis * 100) if total_disponiveis > 0 else 0

col_kpi1, col_kpi2, col_kpi3, col_kpi4, col_kpi5 = st.columns(5)

with col_kpi1:
    st.metric(
        "Taxa de Ocupação",
        f"{taxa_ocupacao:.1f}%",
        help="Matriculados / Vagas totais"
    )
with col_kpi2:
    st.metric(
        "Taxa de Retenção",
        f"{taxa_retencao:.1f}%",
        help="Veteranos / Total matriculados"
    )
with col_kpi3:
    st.metric(
        "Taxa de Captação",
        f"{taxa_captacao:.1f}%",
        help="Novatos / Vagas totais"
    )
with col_kpi4:
    st.metric(
        "Pré-matrículas",
        f"{total_pre_matr}",
        delta=f"{taxa_conversao_pre:.0f}% das vagas",
        help="Potenciais novos alunos"
    )
with col_kpi5:
    receita_potencial = total_disponiveis  # Cada vaga = potencial receita
    st.metric(
        "Vagas Disponíveis",
        f"{total_disponiveis}",
        help="Oportunidade de crescimento"
    )

st.markdown("<br>", unsafe_allow_html=True)

# ===== ALERTAS EXECUTIVOS =====
turmas_criticas = df_relatorio[df_relatorio['Ocupação %'] >= 95]
turmas_atencao = df_relatorio[(df_relatorio['Ocupação %'] >= 85) & (df_relatorio['Ocupação %'] < 95)]
turmas_oportunidade = df_relatorio[df_relatorio['Ocupação %'] < 60]

col_alerta1, col_alerta2, col_alerta3 = st.columns(3)

with col_alerta1:
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid #3b82f6;'>
        <p style='color: #94a3b8; font-size: 0.8rem; margin: 0; text-transform: uppercase;'>Turmas Lotadas (≥95%)</p>
        <p style='color: #ffffff; font-size: 2rem; font-weight: 700; margin: 0.5rem 0;'>{len(turmas_criticas)}</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>Considerar abertura de novas turmas</p>
    </div>
    """, unsafe_allow_html=True)

with col_alerta2:
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid #f59e0b;'>
        <p style='color: #94a3b8; font-size: 0.8rem; margin: 0; text-transform: uppercase;'>Turmas em Atenção (85-95%)</p>
        <p style='color: #ffffff; font-size: 2rem; font-weight: 700; margin: 0.5rem 0;'>{len(turmas_atencao)}</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>Monitorar próximas matrículas</p>
    </div>
    """, unsafe_allow_html=True)

with col_alerta3:
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid #10b981;'>
        <p style='color: #94a3b8; font-size: 0.8rem; margin: 0; text-transform: uppercase;'>Oportunidade (<60%)</p>
        <p style='color: #ffffff; font-size: 2rem; font-weight: 700; margin: 0.5rem 0;'>{len(turmas_oportunidade)}</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>Potencial para campanhas de captação</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== RANKING DE PERFORMANCE POR UNIDADE =====
st.markdown("<h4 style='color: #e2e8f0; font-weight: 600;'>Ranking de Unidades por Performance</h4>", unsafe_allow_html=True)

df_ranking = df_relatorio.groupby('Unidade').agg({
    'Vagas': 'sum',
    'Matriculados': 'sum',
    'Novatos': 'sum',
    'Veteranos': 'sum',
    'Disponiveis': 'sum'
}).reset_index()

df_ranking['Ocupação'] = (df_ranking['Matriculados'] / df_ranking['Vagas'] * 100).round(1)
df_ranking['Retenção'] = (df_ranking['Veteranos'] / df_ranking['Matriculados'] * 100).round(1)
df_ranking['Captação'] = (df_ranking['Novatos'] / df_ranking['Vagas'] * 100).round(1)
df_ranking = df_ranking.sort_values('Ocupação', ascending=False)

# Extrai nome curto
df_ranking['Unidade'] = df_ranking['Unidade'].apply(
    lambda x: x.split('(')[1].replace(')', '') if '(' in x else x
)

st.dataframe(
    df_ranking[['Unidade', 'Vagas', 'Matriculados', 'Ocupação', 'Retenção', 'Captação', 'Disponiveis']],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Ocupação": st.column_config.ProgressColumn(
            "Ocupação %",
            format="%.1f%%",
            min_value=0,
            max_value=100,
        ),
        "Retenção": st.column_config.ProgressColumn(
            "Retenção %",
            format="%.1f%%",
            min_value=0,
            max_value=100,
        ),
        "Captação": st.column_config.ProgressColumn(
            "Captação %",
            format="%.1f%%",
            min_value=0,
            max_value=100,
        ),
    }
)

st.markdown("<br>", unsafe_allow_html=True)

# ===== RELATÓRIO DETALHADO DAS TURMAS =====
st.markdown("<h4 style='color: #e2e8f0; font-weight: 600;'>Detalhamento por Turma</h4>", unsafe_allow_html=True)

# Filtros inline para o detalhamento
col_filtro1, col_filtro2, col_filtro3, col_filtro4 = st.columns(4)

with col_filtro1:
    unidades_det = ["Todas"] + sorted(df_relatorio['Unidade'].apply(
        lambda x: x.split('(')[1].replace(')', '') if '(' in x else x
    ).unique().tolist())
    filtro_unidade_det = st.selectbox("Filtrar Unidade", unidades_det, key="filtro_unidade_det")

with col_filtro2:
    segmentos_det = ["Todos"] + sorted(df_relatorio['Segmento'].unique().tolist())
    filtro_segmento_det = st.selectbox("Filtrar Segmento", segmentos_det, key="filtro_segmento_det")

with col_filtro3:
    turnos_det = ["Todos"] + sorted(df_relatorio['Turno'].unique().tolist())
    filtro_turno_det = st.selectbox("Filtrar Turno", turnos_det, key="filtro_turno_det")

with col_filtro4:
    ordenacao = st.selectbox("Ordenar por", ["Ocupação (maior)", "Ocupação (menor)", "Vagas (maior)", "Disponíveis (maior)"], key="ordenacao_det")

# Aplica filtros do detalhamento
df_det = df_relatorio.copy()

if filtro_unidade_det != "Todas":
    df_det = df_det[df_det['Unidade'].str.contains(filtro_unidade_det, case=False)]

if filtro_segmento_det != "Todos":
    df_det = df_det[df_det['Segmento'] == filtro_segmento_det]

if filtro_turno_det != "Todos":
    df_det = df_det[df_det['Turno'] == filtro_turno_det]

# Aplica ordenação
if ordenacao == "Ocupação (maior)":
    df_det = df_det.sort_values('Ocupação %', ascending=False)
elif ordenacao == "Ocupação (menor)":
    df_det = df_det.sort_values('Ocupação %', ascending=True)
elif ordenacao == "Vagas (maior)":
    df_det = df_det.sort_values('Vagas', ascending=False)
elif ordenacao == "Disponíveis (maior)":
    df_det = df_det.sort_values('Disponiveis', ascending=False)

# Reorganiza colunas para exibição
colunas_exibir = ['Unidade', 'Segmento', 'Turma', 'Turno', 'Vagas', 'Matriculados', 'Ocupação %', 'Novatos', 'Veteranos', 'Disponiveis', 'Pre-matriculados']
df_exibir = df_det[colunas_exibir].copy()

# Extrai nome curto da unidade
df_exibir['Unidade'] = df_exibir['Unidade'].apply(lambda x: x.split('(')[1].replace(')', '') if '(' in x else x)
df_exibir.columns = ['Unidade', 'Segmento', 'Turma', 'Turno', 'Vagas', 'Matr.', 'Ocup.', 'Nov.', 'Vet.', 'Disp.', 'Pré']

# Função para cor da barra de ocupação (mesma escala do termômetro)
def cor_barra_ocupacao(ocupacao):
    try:
        ocupacao = float(ocupacao) if ocupacao else 0
    except:
        ocupacao = 0
    if ocupacao >= 90: return '#065f46'    # Excelente (verde escuro) - 90-100%
    elif ocupacao >= 80: return '#22c55e'  # Boa (verde) - 80-89%
    elif ocupacao >= 70: return '#a3e635'  # Atenção (verde-amarelo) - 70-79%
    elif ocupacao >= 50: return '#facc15'  # Risco (amarelo) - 50-69%
    elif ocupacao >= 38: return '#f97316'  # Crítica (laranja) - 38-49%
    else: return '#dc2626'                 # Congelada (vermelho) - 0-37%

# Cria HTML da tabela com barras coloridas
def criar_barra_html(ocupacao):
    try:
        ocupacao = float(ocupacao) if ocupacao else 0
    except:
        ocupacao = 0
    cor = cor_barra_ocupacao(ocupacao)
    largura = min(ocupacao, 100)
    return f'''<div style="display: flex; align-items: center; gap: 8px;">
        <div style="flex: 1; background: #2d2d44; border-radius: 4px; height: 18px; overflow: hidden;">
            <div style="width: {largura}%; height: 100%; background: linear-gradient(90deg, {cor}, {cor}cc); border-radius: 4px;"></div>
        </div>
        <span style="min-width: 45px; text-align: right; font-size: 12px;">{ocupacao:.1f}%</span>
    </div>'''

# Gera HTML da tabela
html_rows = []
for _, row in df_exibir.iterrows():
    try:
        ocupacao_val = float(row['Ocup.']) if row['Ocup.'] else 0
    except:
        ocupacao_val = 0
    barra_html = criar_barra_html(ocupacao_val)
    html_rows.append(f'''
    <tr>
        <td>{row['Unidade']}</td>
        <td>{row['Segmento']}</td>
        <td>{row['Turma']}</td>
        <td>{row['Turno']}</td>
        <td style="text-align: center;">{int(row['Vagas'])}</td>
        <td style="text-align: center;">{int(row['Matr.'])}</td>
        <td style="min-width: 150px;">{barra_html}</td>
        <td style="text-align: center;">{int(row['Nov.'])}</td>
        <td style="text-align: center;">{int(row['Vet.'])}</td>
        <td style="text-align: center;">{int(row['Disp.'])}</td>
        <td style="text-align: center;">{int(row['Pré'])}</td>
    </tr>
    ''')

html_table = f'''
<div style="max-height: 450px; overflow-y: auto; border-radius: 8px; border: 1px solid #3d3d5c;">
<table style="width: 100%; border-collapse: collapse; font-size: 12px;">
    <thead>
        <tr style="background: linear-gradient(90deg, #667eea, #764ba2); color: white; position: sticky; top: 0;">
            <th style="padding: 10px 8px; text-align: left;">Unidade</th>
            <th style="padding: 10px 8px; text-align: left;">Segmento</th>
            <th style="padding: 10px 8px; text-align: left;">Turma</th>
            <th style="padding: 10px 8px; text-align: left;">Turno</th>
            <th style="padding: 10px 8px; text-align: center;">Vagas</th>
            <th style="padding: 10px 8px; text-align: center;">Matr.</th>
            <th style="padding: 10px 8px; text-align: left; min-width: 150px;">Ocupação</th>
            <th style="padding: 10px 8px; text-align: center;">Nov.</th>
            <th style="padding: 10px 8px; text-align: center;">Vet.</th>
            <th style="padding: 10px 8px; text-align: center;">Disp.</th>
            <th style="padding: 10px 8px; text-align: center;">Pré</th>
        </tr>
    </thead>
    <tbody>
        {"".join(html_rows)}
    </tbody>
</table>
</div>
<style>
    table tbody tr:nth-child(odd) {{ background: #1a1a2e; }}
    table tbody tr:nth-child(even) {{ background: #16162a; }}
    table tbody tr:hover {{ background: #2d2d44; }}
    table td {{ padding: 8px; color: #e0e0ff; border-bottom: 1px solid #2d2d44; }}
</style>
'''

st.markdown(html_table, unsafe_allow_html=True)

st.caption(f"Exibindo {len(df_exibir)} turmas • Filtros aplicados: {filtro_unidade_det} | {filtro_segmento_det} | {filtro_turno_det}")

st.markdown("<br>", unsafe_allow_html=True)

# Download do relatório filtrado
csv_filtrado = df_relatorio.to_csv(index=False).encode('utf-8')
col_dl1, col_dl2, col_dl3 = st.columns([1, 1, 1])
with col_dl1:
    st.download_button(
        label="📥 Exportar CSV",
        data=csv_filtrado,
        file_name=f"relatorio_executivo_{resumo['data_extracao'][:10]}.csv",
        mime="text/csv",
    )
with col_dl2:
    excel_data = df_relatorio.to_csv(index=False, sep=';').encode('utf-8')
    st.download_button(
        label="📊 Exportar Excel",
        data=excel_data,
        file_name=f"relatorio_executivo_{resumo['data_extracao'][:10]}.csv",
        mime="text/csv",
    )
with col_dl3:
    # Gera relatório PDF (HTML para impressão)
    html_pdf = gerar_relatorio_pdf(resumo, df_perf_unidade, df_relatorio, total)
    st.download_button(
        label="📄 Relatório PDF",
        data=html_pdf.encode('utf-8'),
        file_name=f"relatorio_executivo_{resumo['data_extracao'][:10]}.html",
        mime="text/html",
        help="Baixe e abra no navegador. Use Ctrl+P para salvar como PDF."
    )

# Footer
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(f"""
    <div style='text-align: center; color: #404060; font-size: 0.8rem; padding: 2rem 0;'>
        <p>Dashboard atualizado automaticamente às 6h - Última extração: {resumo['data_extracao'][:16].replace('T', ' ')}</p>
        <p style='color: #303050;'>Colégio Elo - 2026</p>
    </div>
""", unsafe_allow_html=True)
