import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import sqlite3
import subprocess
import os
import time
from pathlib import Path
from datetime import datetime

# ===== CONSTANTES =====
BASE_DIR = Path(__file__).parent
BASE_PATH = BASE_DIR / "output"

# Metas por unidade
METAS_MATRICULAS = {
    "01-BV": 1250, "02-CD": 1200, "03-JG": 850, "04-CDR": 800
}
METAS_NOVATOS = {
    "01-BV": 285, "02-CD": 273, "03-JG": 227, "04-CDR": 215
}
META_NOVATOS_TOTAL = 1000
META_MATRICULAS_TOTAL = sum(METAS_MATRICULAS.values())  # 4100

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

# Layout padrão Plotly
PLOTLY_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#a0a0b0', family='Inter, sans-serif'),
    xaxis=dict(gridcolor='rgba(102, 126, 234, 0.1)', tickfont=dict(color='#a0a0b0')),
    yaxis=dict(gridcolor='rgba(102, 126, 234, 0.1)', tickfont=dict(color='#a0a0b0')),
    legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#a0a0b0')),
    margin=dict(t=40, b=40, l=40, r=40)
)

# ===== FUNÇÕES AUXILIARES =====
def extrair_nome_curto(nome_completo):
    """Extrai nome curto da unidade: '1 - BV (Boa Viagem)' -> 'Boa Viagem'"""
    if '(' in str(nome_completo):
        return nome_completo.split('(')[1].replace(')', '')
    return str(nome_completo)

def calcular_ocupacao(matriculados, vagas):
    """Calcula taxa de ocupação em porcentagem"""
    return round((matriculados / vagas * 100), 1) if vagas > 0 else 0.0

def cor_por_porcentagem(valor, limites=(80, 60)):
    """Retorna cor baseada em porcentagem: verde/amarelo/vermelho"""
    alto, medio = limites
    if valor >= alto:
        return '#10b981'  # verde
    elif valor >= medio:
        return '#f59e0b'  # amarelo
    return '#ef4444'  # vermelho

def cor_ocupacao_6_niveis(ocupacao):
    """Escala de 6 cores para ocupação"""
    if ocupacao >= 90: return '#065f46'    # Excelente (verde escuro)
    elif ocupacao >= 80: return '#22c55e'  # Boa (verde)
    elif ocupacao >= 70: return '#a3e635'  # Atenção (verde-amarelo)
    elif ocupacao >= 50: return '#facc15'  # Risco (amarelo)
    elif ocupacao >= 38: return '#f97316'  # Crítica (laranja)
    return '#dc2626'                        # Congelada (vermelho)

def status_meta(atingimento):
    """Retorna status textual da meta"""
    if atingimento >= 100: return 'Atingida'
    elif atingimento >= 90: return 'Quase lá'
    elif atingimento >= 80: return 'Bom'
    elif atingimento >= 60: return 'Atenção'
    elif atingimento >= 40: return 'Risco'
    return 'Crítico'

def get_meta_unidade(unidade_nome, tipo='matriculas'):
    """Obtém meta de matrículas ou novatos para uma unidade"""
    metas = METAS_MATRICULAS if tipo == 'matriculas' else METAS_NOVATOS
    for codigo, valor in metas.items():
        if codigo in unidade_nome or any(nome in unidade_nome for nome in
            ['Boa Viagem', 'Jaboatão', 'Candeias', 'Paulista', 'Janga', 'Cordeiro']):
            if '01-BV' in codigo and ('01-BV' in unidade_nome or 'Boa Viagem' in unidade_nome):
                return valor
            elif '02-CD' in codigo and ('02-CD' in unidade_nome or 'Jaboatão' in unidade_nome or 'Candeias' in unidade_nome):
                return valor
            elif '03-JG' in codigo and ('03-JG' in unidade_nome or 'Paulista' in unidade_nome or 'Janga' in unidade_nome):
                return valor
            elif '04-CDR' in codigo and ('04-CDR' in unidade_nome or 'Cordeiro' in unidade_nome):
                return valor
    return 0

def extrair_turno(turma_nome):
    """Extrai turno do nome da turma"""
    turma_lower = turma_nome.lower()
    if "manhã" in turma_lower or "manha" in turma_lower:
        return "Manhã"
    elif "tarde" in turma_lower:
        return "Tarde"
    elif "integral" in turma_lower:
        return "Integral"
    return "Outro"

def extrair_serie(turma_nome):
    """Extrai série do nome da turma para cálculo de retenção"""
    turma_lower = turma_nome.lower()
    # Normaliza variações: remove espaço antes de º e garante espaço depois
    import re
    turma_normalizada = re.sub(r'\s*º\s*', 'º ', turma_lower)
    # Educação Infantil (ordem V->IV->III->II para evitar substring match)
    if "infantil v" in turma_lower or "infantil 5" in turma_lower:
        return "Infantil V"
    elif "infantil iv" in turma_lower or "infantil 4" in turma_lower:
        return "Infantil IV"
    elif "infantil iii" in turma_lower or "infantil 3" in turma_lower:
        return "Infantil III"
    elif "infantil ii" in turma_lower or "infantil 2" in turma_lower:
        return "Infantil II"
    # Fundamental I (usa turma_normalizada para variações)
    elif "1º ano" in turma_normalizada or "1° ano" in turma_lower:
        return "1º ano"
    elif "2º ano" in turma_normalizada or "2° ano" in turma_lower:
        return "2º ano"
    elif "3º ano" in turma_normalizada or "3° ano" in turma_lower:
        return "3º ano"
    elif "4º ano" in turma_normalizada or "4° ano" in turma_lower:
        return "4º ano"
    elif "5º ano" in turma_normalizada or "5° ano" in turma_lower:
        return "5º ano"
    # Fundamental II
    elif "6º ano" in turma_normalizada or "6° ano" in turma_lower:
        return "6º ano"
    elif "7º ano" in turma_normalizada or "7° ano" in turma_lower:
        return "7º ano"
    elif "8º ano" in turma_normalizada or "8° ano" in turma_lower:
        return "8º ano"
    elif "9º ano" in turma_normalizada or "9° ano" in turma_lower:
        return "9º ano"
    # Ensino Médio
    elif "1ª série" in turma_lower or "1a série" in turma_lower:
        return "1ª série EM"
    elif "2ª série" in turma_lower or "2a série" in turma_lower:
        return "2ª série EM"
    elif "3ª série" in turma_lower or "3a série" in turma_lower:
        return "3ª série EM"
    return None

def gerar_termometro_html(nome, valor_atual, meta, tipo='matriculas'):
    """Gera HTML de termômetro para metas"""
    atingimento = (valor_atual / meta * 100) if meta > 0 else 0
    cor = cor_ocupacao_6_niveis(atingimento)
    status = status_meta(atingimento)
    gap = int(valor_atual - meta)
    sinal = '+' if gap >= 0 else ''
    cor_texto = 'white' if atingimento < 80 or atingimento >= 90 else '#1a1a2e'

    return f"""
    <div style='text-align: center; padding: 10px;'>
        <div style='font-size: 14px; color: #e2e8f0; font-weight: 600; margin-bottom: 8px;'>{nome}</div>
        <div style='position: relative; width: 50px; height: 160px; margin: 0 auto; background: linear-gradient(to top, #1a1a2e 0%, #2d2d44 100%); border-radius: 25px; border: 2px solid #3d3d5c; overflow: hidden;'>
            <div style='position: absolute; bottom: 0; width: 100%; height: {min(atingimento, 100)}%; background: linear-gradient(to top, {cor}, {cor}dd); border-radius: 0 0 23px 23px;'></div>
            <div style='position: absolute; width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;'>
                <span style='font-size: 16px; font-weight: bold; color: white; text-shadow: 1px 1px 2px rgba(0,0,0,0.8);'>{atingimento:.0f}%</span>
            </div>
        </div>
        <div style='margin-top: 8px;'>
            <span style='background: {cor}; color: {cor_texto}; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 500;'>{status}</span>
        </div>
        <div style='font-size: 11px; color: #a0a0b0; margin-top: 5px;'>{int(valor_atual)} / {int(meta)}</div>
        <div style='font-size: 10px; color: {cor}; font-weight: 600;'>Gap: {sinal}{gap}</div>
    </div>
    """

# ===== CONFIGURAÇÃO DA PÁGINA =====
st.set_page_config(
    page_title="Vagas Colégio Elo",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    h1, h2, h3, h4, h5, h6 {
        color: #f1f5f9 !important;
        font-weight: 600 !important;
    }

    h2, h3 {
        color: #e0e0ff !important;
    }

    h4, h5, h6 {
        color: #e2e8f0 !important;
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

    .stMarkdown h4 {
        color: #e2e8f0 !important;
        font-size: 1.2rem !important;
    }

    /* Default text color */
    .stMarkdown p, .stMarkdown span, .stMarkdown div {
        color: #cbd5e1;
    }

    /* Selectbox and input labels */
    .stSelectbox label, .stMultiSelect label, .stTextInput label {
        color: #e2e8f0 !important;
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
    <div class="kpi-box {'green' if ocupacao_geral >= 80 else 'yellow' if ocupacao_geral >= 60 else 'red'}"><div class="kpi-label">Ocupação Geral</div><div class="kpi-value">{ocupacao_geral:.1f}%</div><div class="kpi-detail">{total['matriculados']:,} / {total['vagas']:,} vagas</div></div>
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

# Filtro de Turno (usa função global extrair_turno)
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
        <h1 style='margin-bottom: 0; color: #f1f5f9; font-size: 3rem;'>Matrículas</h1>
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
    st.metric("OCUPAÇÃO GERAL", f"{ocupacao:.1f}%")
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

# Quadro de Quantidade de Turmas
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>📚 Quantidade de Turmas por Unidade</h3>", unsafe_allow_html=True)

# Calcula quantidade de turmas por unidade
df_turmas_count = df_turmas_all.groupby('Unidade').agg({
    'Turma': 'count',
    'Vagas': 'sum',
    'Matriculados': 'sum'
}).reset_index()
df_turmas_count.columns = ['Unidade', 'Total Turmas', 'Vagas', 'Matriculados']
df_turmas_count['Nome_curto'] = df_turmas_count['Unidade'].apply(extrair_nome_curto)

# Cards com totais por unidade
cols_turmas = st.columns(len(df_turmas_count))
for idx, (_, row) in enumerate(df_turmas_count.iterrows()):
    with cols_turmas[idx]:
        st.markdown(f"""
        <div style='background: linear-gradient(145deg, #1e1e30 0%, #252540 100%);
                    border: 1px solid rgba(102, 126, 234, 0.3);
                    border-radius: 12px; padding: 1rem; text-align: center;'>
            <p style='color: #a0a0b0; font-size: 0.8rem; margin: 0;'>{row['Nome_curto']}</p>
            <p style='color: #667eea; font-size: 2rem; font-weight: 700; margin: 0.3rem 0;'>{int(row['Total Turmas'])}</p>
            <p style='color: #64748b; font-size: 0.7rem; margin: 0;'>turmas</p>
        </div>
        """, unsafe_allow_html=True)

# Total geral de turmas
total_turmas = df_turmas_count['Total Turmas'].sum()
st.markdown(f"<p style='color: #94a3b8; text-align: center; margin-top: 0.5rem;'>Total: <strong style='color: #f1f5f9;'>{int(total_turmas)} turmas</strong></p>", unsafe_allow_html=True)

# Detalhamento expandível
with st.expander("📋 Ver detalhamento por segmento e turno"):
    # Agrupa por unidade, segmento e turno
    df_detail = df_turmas_all.groupby(['Unidade', 'Segmento', 'Turno']).agg({
        'Turma': 'count',
        'Vagas': 'sum',
        'Matriculados': 'sum'
    }).reset_index()
    df_detail.columns = ['Unidade', 'Segmento', 'Turno', 'Qtd Turmas', 'Vagas', 'Matriculados']
    df_detail['Unidade'] = df_detail['Unidade'].apply(extrair_nome_curto)
    df_detail['Ocupação %'] = (df_detail['Matriculados'] / df_detail['Vagas'] * 100).round(1)

    # Ordena por unidade e segmento
    ordem_seg = {'Ed. Infantil': 1, 'Fund. 1': 2, 'Fund. 2': 3, 'Ens. Médio': 4}
    df_detail['ordem_seg'] = df_detail['Segmento'].map(ordem_seg).fillna(5)
    df_detail = df_detail.sort_values(['Unidade', 'ordem_seg', 'Turno'])
    df_detail = df_detail.drop('ordem_seg', axis=1)

    st.dataframe(
        df_detail,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Ocupação %": st.column_config.ProgressColumn(
                "Ocupação %",
                format="%.1f%%",
                min_value=0,
                max_value=100,
            ),
        }
    )

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

    # Barra de ocupação - escala 6 cores (usa função global)
    colors = [cor_ocupacao_6_niveis(o) for o in df_unidades['Ocupação']]

    fig1.add_trace(go.Bar(
        name='Ocupação',
        x=df_unidades['Unidade'],
        y=df_unidades['Ocupação'],
        marker_color=colors,
        text=df_unidades.apply(lambda r: f"{r['Ocupação']:.1f}%<br>({int(r['Matriculados'])})", axis=1),
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

    ordem = ['Ed. Infantil', 'Fund. 1', 'Fund. 2', 'Ens. Médio']
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

# ===== INSIGHTS EXECUTIVOS - CEO =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>💡 Insights Executivos</h3>", unsafe_allow_html=True)

# Calcula métricas por unidade com metas (usa funções globais)
df_perf_unidade = df_resumo_all.groupby('Unidade').agg({
    'Vagas': 'sum', 'Matriculados': 'sum', 'Novatos': 'sum', 'Veteranos': 'sum'
}).reset_index()

# Adiciona metas usando funções globais
df_perf_unidade['Meta'] = df_perf_unidade['Unidade'].apply(lambda x: get_meta_unidade(x, 'matriculas'))
df_perf_unidade['Gap'] = df_perf_unidade['Matriculados'] - df_perf_unidade['Meta']
df_perf_unidade['Atingimento'] = (df_perf_unidade['Matriculados'] / df_perf_unidade['Meta'] * 100).round(1)
df_perf_unidade['Ocupacao'] = df_perf_unidade.apply(lambda r: calcular_ocupacao(r['Matriculados'], r['Vagas']), axis=1)
df_perf_unidade['Meta_Novatos'] = df_perf_unidade['Unidade'].apply(lambda x: get_meta_unidade(x, 'novatos'))
df_perf_unidade['Gap_Novatos'] = df_perf_unidade['Novatos'] - df_perf_unidade['Meta_Novatos']
df_perf_unidade['Ating_Novatos'] = (df_perf_unidade['Novatos'] / df_perf_unidade['Meta_Novatos'] * 100).round(1)
df_perf_unidade['Nome_curto'] = df_perf_unidade['Unidade'].apply(extrair_nome_curto)

# Calcula totais (usa constantes globais)
gap_total = total['matriculados'] - META_MATRICULAS_TOTAL
atingimento_total = (total['matriculados'] / META_MATRICULAS_TOTAL * 100) if META_MATRICULAS_TOTAL > 0 else 0
gap_novatos = total['novatos'] - META_NOVATOS_TOTAL
atingimento_novatos = (total['novatos'] / META_NOVATOS_TOTAL * 100) if META_NOVATOS_TOTAL > 0 else 0

# Linha 1 - Metas gerais
col_meta1, col_meta2, col_meta3 = st.columns(3)

with col_meta1:
    cor_meta = cor_por_porcentagem(atingimento_total if gap_total >= 0 else atingimento_total, (100, 80))
    sinal = '+' if gap_total >= 0 else ''
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid {cor_meta};'>
        <p style='color: #94a3b8; font-size: 0.75rem; margin: 0; text-transform: uppercase;'>Meta Matrículas ({META_MATRICULAS_TOTAL:,})</p>
        <p style='color: {cor_meta}; font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0;'>{atingimento_total:.1f}%</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>{sinal}{gap_total:,} alunos ({total['matriculados']:,}/{META_MATRICULAS_TOTAL:,})</p>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

with col_meta2:
    cor_novatos = cor_por_porcentagem(atingimento_novatos if gap_novatos >= 0 else atingimento_novatos, (100, 80))
    sinal_nov = '+' if gap_novatos >= 0 else ''
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid {cor_novatos};'>
        <p style='color: #94a3b8; font-size: 0.75rem; margin: 0; text-transform: uppercase;'>Meta Novatos ({META_NOVATOS_TOTAL:,})</p>
        <p style='color: {cor_novatos}; font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0;'>{atingimento_novatos:.1f}%</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>{sinal_nov}{gap_novatos:,} novatos ({total['novatos']:,}/{META_NOVATOS_TOTAL:,})</p>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

with col_meta3:
    taxa_retencao_geral = (total['veteranos'] / total['matriculados'] * 100) if total['matriculados'] > 0 else 0
    cor_retencao = cor_por_porcentagem(taxa_retencao_geral, (70, 50))
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1e3a5f 0%, #2d4a6f 100%); padding: 1.2rem; border-radius: 12px; border-left: 4px solid {cor_retencao};'>
        <p style='color: #94a3b8; font-size: 0.75rem; margin: 0; text-transform: uppercase;'>% Veteranos</p>
        <p style='color: {cor_retencao}; font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0;'>{taxa_retencao_geral:.1f}%</p>
        <p style='color: #64748b; font-size: 0.75rem; margin: 0;'>{total['veteranos']:,} rematriculados</p>
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

# Termômetros de Matrículas (usa função global gerar_termometro_html)
st.markdown("<h4 style='color: #e2e8f0; font-weight: 500; margin-top: 10px;'>Meta de Matrículas</h4>", unsafe_allow_html=True)
cols_meta_mat = st.columns(len(df_perf_unidade))
for idx, (_, row) in enumerate(df_perf_unidade.iterrows()):
    with cols_meta_mat[idx]:
        st.markdown(gerar_termometro_html(row['Nome_curto'], row['Matriculados'], row['Meta']), unsafe_allow_html=True)

# Termômetros de Novatos
st.markdown("<h4 style='color: #e2e8f0; font-weight: 500; margin-top: 20px;'>Meta de Novatos</h4>", unsafe_allow_html=True)
cols_meta_nov = st.columns(len(df_perf_unidade))
for idx, (_, row) in enumerate(df_perf_unidade.iterrows()):
    with cols_meta_nov[idx]:
        st.markdown(gerar_termometro_html(row['Nome_curto'], row['Novatos'], row['Meta_Novatos']), unsafe_allow_html=True)

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

# Turmas com ocupação calculada (usa funções globais)
turmas_criticas = df_turmas_all.copy()
turmas_criticas['Ocupacao'] = turmas_criticas.apply(lambda r: calcular_ocupacao(r['Matriculados'], r['Vagas']), axis=1)
turmas_criticas['Unidade_curta'] = turmas_criticas['Unidade'].apply(extrair_nome_curto)

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
                    seg_nome = t['Segmento'] if len(str(t['Segmento'])) > 4 else t['Segmento']
                    st.markdown(f"""
                    <div style='background: rgba(30, 41, 59, 0.9); padding: 0.7rem 1rem; border-radius: 10px; margin-bottom: 0.4rem; border: 1px solid rgba(239, 68, 68, 0.3);'>
                        <div style='display: flex; justify-content: space-between; align-items: center;'>
                            <span style='color: #ffffff; font-weight: 600;'>{t['Turma']}</span>
                            <span style='color: #ef4444; font-weight: 800; font-size: 1.1rem;'>{t['Ocupacao']:.0f}% <span style='font-size: 0.75rem; color: #fca5a5;'>({int(t['Matriculados'])} alunos)</span></span>
                        </div>
                        <p style='color: #94a3b8; font-size: 0.8rem; margin: 0.3rem 0 0 0;'>{seg_nome} • {int(t['Matriculados'])}/{int(t['Vagas'])} vagas preenchidas</p>
                    </div>
                    """, unsafe_allow_html=True)
                if limite and len(lotadas) > limite:
                    st.caption(f"... e mais {len(lotadas) - limite} turmas")
            else:
                st.markdown("<p style='color: #10b981; font-size: 0.9rem; padding: 0.5rem;'>✅ Nenhuma turma lotada nesta unidade</p>", unsafe_allow_html=True)

        with col_a2:
            st.markdown(f"""
            <div style='background: rgba(251, 191, 36, 0.15); padding: 1rem; border-radius: 12px; border-left: 5px solid #f59e0b; margin-bottom: 0.8rem;'>
                <p style='color: #ffffff; font-weight: 700; font-size: 1rem; margin: 0;'>🟡 BAIXA OCUPAÇÃO GERAL (<50%)</p>
                <p style='color: #fcd34d; font-size: 0.85rem; margin: 0.3rem 0 0 0;'>{len(vazias)} turmas - foco em captação</p>
            </div>
            """, unsafe_allow_html=True)

            if len(vazias_exibir) > 0:
                for _, t in vazias_exibir.iterrows():
                    vagas_disp = int(t['Vagas'] - t['Matriculados'])
                    seg_nome = t['Segmento'] if len(str(t['Segmento'])) > 4 else t['Segmento']
                    st.markdown(f"""
                    <div style='background: rgba(30, 41, 59, 0.9); padding: 0.7rem 1rem; border-radius: 10px; margin-bottom: 0.4rem; border: 1px solid rgba(251, 191, 36, 0.3);'>
                        <div style='display: flex; justify-content: space-between; align-items: center;'>
                            <span style='color: #ffffff; font-weight: 600;'>{t['Turma']}</span>
                            <span style='color: #f59e0b; font-weight: 800; font-size: 1.1rem;'>{t['Ocupacao']:.0f}% <span style='font-size: 0.75rem; color: #fcd34d;'>({int(t['Matriculados'])} alunos)</span></span>
                        </div>
                        <p style='color: #94a3b8; font-size: 0.8rem; margin: 0.3rem 0 0 0;'>{seg_nome} • {vagas_disp} vagas disponíveis</p>
                    </div>
                    """, unsafe_allow_html=True)
                if limite and len(vazias) > limite:
                    st.caption(f"... e mais {len(vazias) - limite} turmas")
            else:
                st.markdown("<p style='color: #10b981; font-size: 0.9rem; padding: 0.5rem;'>✅ Todas as turmas acima de 50%</p>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===== COMPARATIVO 2025 vs 2026 =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>📊 Comparativo 2025 vs 2026 - Novatos e Veteranos</h3>", unsafe_allow_html=True)

# Tenta carregar dados de 2025
resumo_2025_path = Path(__file__).parent / "output" / "resumo_2025.json"
if resumo_2025_path.exists():
    with open(resumo_2025_path, "r", encoding="utf-8") as f:
        resumo_2025 = json.load(f)

    # Prepara dados para comparação
    dados_comp = []
    segmentos_validos = ["Ed. Infantil", "Fund. 1", "Fund. 2", "Ens. Médio"]

    for unidade_2026 in resumo.get("unidades", []):
        codigo = unidade_2026["codigo"]
        nome_curto = codigo.split("-")[1] if "-" in codigo else codigo

        # Encontra mesma unidade em 2025
        unidade_2025 = next((u for u in resumo_2025.get("unidades", []) if u["codigo"] == codigo), None)

        for seg in segmentos_validos:
            dados_2026 = unidade_2026.get("segmentos", {}).get(seg, {})
            dados_2025_seg = unidade_2025.get("segmentos", {}).get(seg, {}) if unidade_2025 else {}

            nov_2026 = dados_2026.get("novatos", 0)
            vet_2026 = dados_2026.get("veteranos", 0)
            matr_2026 = dados_2026.get("matriculados", 0)

            nov_2025 = dados_2025_seg.get("novatos", 0)
            vet_2025 = dados_2025_seg.get("veteranos", 0)
            matr_2025 = dados_2025_seg.get("matriculados", 0)

            dados_comp.append({
                "Unidade": nome_curto,
                "Segmento": seg,
                "Novatos_2025": nov_2025,
                "Novatos_2026": nov_2026,
                "Var_Nov": nov_2026 - nov_2025,
                "Var_Nov_Pct": ((nov_2026 - nov_2025) / nov_2025 * 100) if nov_2025 > 0 else 0,
                "Veteranos_2025": vet_2025,
                "Veteranos_2026": vet_2026,
                "Var_Vet": vet_2026 - vet_2025,
                "Var_Vet_Pct": ((vet_2026 - vet_2025) / vet_2025 * 100) if vet_2025 > 0 else 0,
                "Total_2025": matr_2025,
                "Total_2026": matr_2026,
                "Var_Total": matr_2026 - matr_2025,
            })

    df_comp = pd.DataFrame(dados_comp)

    # Tabela comparativa por unidade
    col_comp1, col_comp2 = st.columns(2)

    with col_comp1:
        st.markdown("<h4 style='color: #e2e8f0;'>Comparativo de Novatos</h4>", unsafe_allow_html=True)

        # Agrupa por unidade
        df_nov_unidade = df_comp.groupby("Unidade").agg({
            "Novatos_2025": "sum",
            "Novatos_2026": "sum",
            "Var_Nov": "sum"
        }).reset_index()

        html_nov = "<table style='width:100%; font-size:12px; border-collapse:collapse;'>"
        html_nov += "<tr style='background:#667eea; color:white;'><th style='padding:8px;'>Unidade</th><th>2025</th><th>2026</th><th>Variação</th></tr>"
        for _, r in df_nov_unidade.iterrows():
            cor_var = "#10b981" if r["Var_Nov"] >= 0 else "#ef4444"
            sinal = "+" if r["Var_Nov"] >= 0 else ""
            html_nov += f"<tr style='background:#1a1a2e;'><td style='padding:6px; color:#e0e0ff;'>{r['Unidade']}</td>"
            html_nov += f"<td style='text-align:center; color:#94a3b8;'>{int(r['Novatos_2025'])}</td>"
            html_nov += f"<td style='text-align:center; color:#e0e0ff; font-weight:600;'>{int(r['Novatos_2026'])}</td>"
            html_nov += f"<td style='text-align:center; color:{cor_var}; font-weight:600;'>{sinal}{int(r['Var_Nov'])}</td></tr>"
        # Total
        total_2025 = df_nov_unidade["Novatos_2025"].sum()
        total_2026 = df_nov_unidade["Novatos_2026"].sum()
        var_total = total_2026 - total_2025
        cor_total = "#10b981" if var_total >= 0 else "#ef4444"
        sinal_t = "+" if var_total >= 0 else ""
        html_nov += f"<tr style='background:#2d2d44; font-weight:700;'><td style='padding:8px; color:#ffffff;'>TOTAL</td>"
        html_nov += f"<td style='text-align:center; color:#ffffff;'>{int(total_2025)}</td>"
        html_nov += f"<td style='text-align:center; color:#ffffff;'>{int(total_2026)}</td>"
        html_nov += f"<td style='text-align:center; color:{cor_total};'>{sinal_t}{int(var_total)}</td></tr>"
        html_nov += "</table>"
        st.markdown(html_nov, unsafe_allow_html=True)

    with col_comp2:
        st.markdown("<h4 style='color: #e2e8f0;'>Comparativo de Veteranos</h4>", unsafe_allow_html=True)

        df_vet_unidade = df_comp.groupby("Unidade").agg({
            "Veteranos_2025": "sum",
            "Veteranos_2026": "sum",
            "Var_Vet": "sum"
        }).reset_index()

        html_vet = "<table style='width:100%; font-size:12px; border-collapse:collapse;'>"
        html_vet += "<tr style='background:#764ba2; color:white;'><th style='padding:8px;'>Unidade</th><th>2025</th><th>2026</th><th>Variação</th></tr>"
        for _, r in df_vet_unidade.iterrows():
            cor_var = "#10b981" if r["Var_Vet"] >= 0 else "#ef4444"
            sinal = "+" if r["Var_Vet"] >= 0 else ""
            html_vet += f"<tr style='background:#1a1a2e;'><td style='padding:6px; color:#e0e0ff;'>{r['Unidade']}</td>"
            html_vet += f"<td style='text-align:center; color:#94a3b8;'>{int(r['Veteranos_2025'])}</td>"
            html_vet += f"<td style='text-align:center; color:#e0e0ff; font-weight:600;'>{int(r['Veteranos_2026'])}</td>"
            html_vet += f"<td style='text-align:center; color:{cor_var}; font-weight:600;'>{sinal}{int(r['Var_Vet'])}</td></tr>"
        # Total
        total_2025_v = df_vet_unidade["Veteranos_2025"].sum()
        total_2026_v = df_vet_unidade["Veteranos_2026"].sum()
        var_total_v = total_2026_v - total_2025_v
        cor_total_v = "#10b981" if var_total_v >= 0 else "#ef4444"
        sinal_tv = "+" if var_total_v >= 0 else ""
        html_vet += f"<tr style='background:#2d2d44; font-weight:700;'><td style='padding:8px; color:#ffffff;'>TOTAL</td>"
        html_vet += f"<td style='text-align:center; color:#ffffff;'>{int(total_2025_v)}</td>"
        html_vet += f"<td style='text-align:center; color:#ffffff;'>{int(total_2026_v)}</td>"
        html_vet += f"<td style='text-align:center; color:{cor_total_v};'>{sinal_tv}{int(var_total_v)}</td></tr>"
        html_vet += "</table>"
        st.markdown(html_vet, unsafe_allow_html=True)

    # Tabela detalhada por segmento
    st.markdown("<h4 style='color: #e2e8f0; margin-top: 20px;'>Detalhamento por Segmento</h4>", unsafe_allow_html=True)

    html_seg = "<div style='max-height:350px; overflow-y:auto;'>"
    html_seg += "<table style='width:100%; font-size:11px; border-collapse:collapse;'>"
    html_seg += "<tr style='background:linear-gradient(90deg, #667eea, #764ba2); color:white; position:sticky; top:0;'>"
    html_seg += "<th style='padding:8px;'>Unidade</th><th>Segmento</th>"
    html_seg += "<th>Novatos 25</th><th>Novatos 26</th><th>Var.</th>"
    html_seg += "<th>Veteranos 25</th><th>Veteranos 26</th><th>Var.</th>"
    html_seg += "<th>Total 25</th><th>Total 26</th><th>Var.</th></tr>"

    for _, r in df_comp.iterrows():
        cor_nov = "#10b981" if r["Var_Nov"] >= 0 else "#ef4444"
        cor_vet = "#10b981" if r["Var_Vet"] >= 0 else "#ef4444"
        cor_tot = "#10b981" if r["Var_Total"] >= 0 else "#ef4444"
        s_nov = "+" if r["Var_Nov"] >= 0 else ""
        s_vet = "+" if r["Var_Vet"] >= 0 else ""
        s_tot = "+" if r["Var_Total"] >= 0 else ""

        html_seg += f"<tr style='background:#1a1a2e; border-bottom:1px solid #2d2d44;'>"
        html_seg += f"<td style='padding:6px; color:#e0e0ff;'>{r['Unidade']}</td>"
        html_seg += f"<td style='color:#94a3b8;'>{r['Segmento']}</td>"
        html_seg += f"<td style='text-align:center; color:#94a3b8;'>{int(r['Novatos_2025'])}</td>"
        html_seg += f"<td style='text-align:center; color:#e0e0ff;'>{int(r['Novatos_2026'])}</td>"
        html_seg += f"<td style='text-align:center; color:{cor_nov}; font-weight:600;'>{s_nov}{int(r['Var_Nov'])}</td>"
        html_seg += f"<td style='text-align:center; color:#94a3b8;'>{int(r['Veteranos_2025'])}</td>"
        html_seg += f"<td style='text-align:center; color:#e0e0ff;'>{int(r['Veteranos_2026'])}</td>"
        html_seg += f"<td style='text-align:center; color:{cor_vet}; font-weight:600;'>{s_vet}{int(r['Var_Vet'])}</td>"
        html_seg += f"<td style='text-align:center; color:#94a3b8;'>{int(r['Total_2025'])}</td>"
        html_seg += f"<td style='text-align:center; color:#e0e0ff;'>{int(r['Total_2026'])}</td>"
        html_seg += f"<td style='text-align:center; color:{cor_tot}; font-weight:600;'>{s_tot}{int(r['Var_Total'])}</td></tr>"

    html_seg += "</table></div>"
    st.markdown(html_seg, unsafe_allow_html=True)

    # Análise comparativa
    st.markdown("<h4 style='color: #e2e8f0; margin-top: 20px;'>📈 Análise Comparativa</h4>", unsafe_allow_html=True)

    total_nov_2025 = df_comp["Novatos_2025"].sum()
    total_nov_2026 = df_comp["Novatos_2026"].sum()
    total_vet_2025 = df_comp["Veteranos_2025"].sum()
    total_vet_2026 = df_comp["Veteranos_2026"].sum()
    total_matr_2025 = df_comp["Total_2025"].sum()
    total_matr_2026 = df_comp["Total_2026"].sum()

    var_nov_pct = ((total_nov_2026 - total_nov_2025) / total_nov_2025 * 100) if total_nov_2025 > 0 else 0
    var_vet_pct = ((total_vet_2026 - total_vet_2025) / total_vet_2025 * 100) if total_vet_2025 > 0 else 0
    var_matr_pct = ((total_matr_2026 - total_matr_2025) / total_matr_2025 * 100) if total_matr_2025 > 0 else 0

    analise_html = f"""
    <div style='background: rgba(102, 126, 234, 0.1); padding: 15px; border-radius: 10px; border-left: 4px solid #667eea;'>
        <p style='color: #e0e0ff; margin: 5px 0;'><strong>Novatos:</strong> {int(total_nov_2025)} (2025) → {int(total_nov_2026)} (2026) |
        <span style='color: {"#10b981" if var_nov_pct >= 0 else "#ef4444"};'>{"+" if var_nov_pct >= 0 else ""}{var_nov_pct:.1f}%</span></p>
        <p style='color: #e0e0ff; margin: 5px 0;'><strong>Veteranos:</strong> {int(total_vet_2025)} (2025) → {int(total_vet_2026)} (2026) |
        <span style='color: {"#10b981" if var_vet_pct >= 0 else "#ef4444"};'>{"+" if var_vet_pct >= 0 else ""}{var_vet_pct:.1f}%</span></p>
        <p style='color: #e0e0ff; margin: 5px 0;'><strong>Total Matrículas:</strong> {int(total_matr_2025)} (2025) → {int(total_matr_2026)} (2026) |
        <span style='color: {"#10b981" if var_matr_pct >= 0 else "#ef4444"};'>{"+" if var_matr_pct >= 0 else ""}{var_matr_pct:.1f}%</span></p>
    </div>
    """
    st.markdown(analise_html, unsafe_allow_html=True)

else:
    st.info("📋 Para visualizar o comparativo 2025 vs 2026, execute o script `extrair_2025.py` para extrair os dados de 2025 do ActiveSoft.")

st.markdown("<br>", unsafe_allow_html=True)

# ===== RETENÇÃO REAL POR SÉRIE =====
st.markdown("<h3 style='color: #f1f5f9; font-weight: 600;'>📈 Retenção Real por Série (2025 → 2026)</h3>", unsafe_allow_html=True)
st.caption("Retenção = alunos da série anterior (2025) que avançaram e permaneceram na escola (2026)")

# Carrega dados detalhados de 2025 e 2026
dados_2025_path = Path(__file__).parent / "output" / "dados_2025.json"
dados_2026_path = Path(__file__).parent / "output" / "vagas_ultimo.json"

if dados_2025_path.exists() and dados_2026_path.exists():
    with open(dados_2025_path, "r", encoding="utf-8") as f:
        dados_2025_full = json.load(f)
    with open(dados_2026_path, "r", encoding="utf-8") as f:
        dados_2026_full = json.load(f)

    # Mapeamento de progressão: série atual -> série anterior (usa extrair_serie global)
    PROGRESSAO = {
        "Infantil III": "Infantil II",
        "Infantil IV": "Infantil III",
        "Infantil V": "Infantil IV",
        "1º ano": "Infantil V",
        "2º ano": "1º ano",
        "3º ano": "2º ano",
        "4º ano": "3º ano",
        "5º ano": "4º ano",
        "6º ano": "5º ano",
        "7º ano": "6º ano",
        "8º ano": "7º ano",
        "9º ano": "8º ano",
        "1ª série EM": "9º ano",
        "2ª série EM": "1ª série EM",
        "3ª série EM": "2ª série EM",
    }

    # Agrupa dados por unidade e série
    def agrupar_por_serie(dados, ano):
        resultado = {}
        for unidade in dados.get("unidades", []):
            codigo = unidade["codigo"]
            if codigo not in resultado:
                resultado[codigo] = {}
            for turma in unidade.get("turmas", []):
                serie = extrair_serie(turma.get("turma", ""))
                if serie:
                    if serie not in resultado[codigo]:
                        resultado[codigo][serie] = {"matriculados": 0, "veteranos": 0, "novatos": 0}
                    resultado[codigo][serie]["matriculados"] += turma.get("matriculados", 0)
                    resultado[codigo][serie]["veteranos"] += turma.get("veteranos", 0)
                    resultado[codigo][serie]["novatos"] += turma.get("novatos", 0)
        return resultado

    dados_2025_serie = agrupar_por_serie(dados_2025_full, "2025")
    dados_2026_serie = agrupar_por_serie(dados_2026_full, "2026")

    # Calcula retenção real por série
    retencao_data = []
    for codigo_unidade in dados_2026_serie.keys():
        nome_unidade = codigo_unidade.split("-")[1] if "-" in codigo_unidade else codigo_unidade
        for serie_atual, serie_anterior in PROGRESSAO.items():
            # Veteranos na série atual (2026) vieram da série anterior (2025)
            vet_2026 = dados_2026_serie.get(codigo_unidade, {}).get(serie_atual, {}).get("veteranos", 0)
            # Total de alunos na série anterior em 2025
            total_2025 = dados_2025_serie.get(codigo_unidade, {}).get(serie_anterior, {}).get("matriculados", 0)

            if total_2025 > 0:
                retencao = (vet_2026 / total_2025) * 100
                retencao_data.append({
                    "Unidade": nome_unidade,
                    "Série 2026": serie_atual,
                    "Base 2025": serie_anterior,
                    "Alunos 2025": total_2025,
                    "Veteranos 2026": vet_2026,
                    "Retenção %": round(retencao, 1)
                })

    if retencao_data:
        df_retencao = pd.DataFrame(retencao_data)

        # Tabela resumo por unidade
        col_ret1, col_ret2 = st.columns(2)

        with col_ret1:
            st.markdown("<h4 style='color: #e2e8f0;'>Retenção por Unidade</h4>", unsafe_allow_html=True)
            df_ret_unidade = df_retencao.groupby("Unidade").agg({
                "Alunos 2025": "sum",
                "Veteranos 2026": "sum"
            }).reset_index()
            df_ret_unidade["Retenção %"] = (df_ret_unidade["Veteranos 2026"] / df_ret_unidade["Alunos 2025"] * 100).round(1)

            html_ret = "<table style='width:100%; font-size:12px; border-collapse:collapse;'>"
            html_ret += "<tr style='background:#10b981; color:white;'><th style='padding:8px;'>Unidade</th><th>Base 2025</th><th>Rematriculados 2026</th><th>Retenção</th></tr>"
            for _, r in df_ret_unidade.iterrows():
                cor = "#10b981" if r["Retenção %"] >= 80 else "#f59e0b" if r["Retenção %"] >= 60 else "#ef4444"
                html_ret += f"<tr style='background:#1a1a2e;'><td style='padding:6px; color:#e0e0ff;'>{r['Unidade']}</td>"
                html_ret += f"<td style='text-align:center; color:#94a3b8;'>{int(r['Alunos 2025'])}</td>"
                html_ret += f"<td style='text-align:center; color:#e0e0ff;'>{int(r['Veteranos 2026'])}</td>"
                html_ret += f"<td style='text-align:center; color:{cor}; font-weight:600;'>{r['Retenção %']:.1f}%</td></tr>"
            # Total
            total_base = df_ret_unidade["Alunos 2025"].sum()
            total_ret = df_ret_unidade["Veteranos 2026"].sum()
            ret_total = (total_ret / total_base * 100) if total_base > 0 else 0
            cor_total = "#10b981" if ret_total >= 80 else "#f59e0b" if ret_total >= 60 else "#ef4444"
            html_ret += f"<tr style='background:#2d2d44; font-weight:700;'><td style='padding:8px; color:#ffffff;'>TOTAL</td>"
            html_ret += f"<td style='text-align:center; color:#ffffff;'>{int(total_base)}</td>"
            html_ret += f"<td style='text-align:center; color:#ffffff;'>{int(total_ret)}</td>"
            html_ret += f"<td style='text-align:center; color:{cor_total};'>{ret_total:.1f}%</td></tr>"
            html_ret += "</table>"
            st.markdown(html_ret, unsafe_allow_html=True)

        with col_ret2:
            st.markdown("<h4 style='color: #e2e8f0;'>Séries com Maior Evasão</h4>", unsafe_allow_html=True)
            # Séries com menor retenção (maior evasão)
            df_evasao = df_retencao[df_retencao["Alunos 2025"] >= 5].sort_values("Retenção %").head(8)
            if len(df_evasao) > 0:
                html_eva = "<table style='width:100%; font-size:11px; border-collapse:collapse;'>"
                html_eva += "<tr style='background:#ef4444; color:white;'><th style='padding:6px;'>Unidade</th><th>Série</th><th>Base</th><th>Rematriculados</th><th>Evasão</th></tr>"
                for _, r in df_evasao.iterrows():
                    evasao = 100 - r["Retenção %"]
                    cor = "#ef4444" if evasao >= 40 else "#f59e0b" if evasao >= 20 else "#10b981"
                    html_eva += f"<tr style='background:#1a1a2e;'><td style='padding:5px; color:#e0e0ff;'>{r['Unidade']}</td>"
                    html_eva += f"<td style='color:#94a3b8;'>{r['Série 2026']}</td>"
                    html_eva += f"<td style='text-align:center; color:#94a3b8;'>{int(r['Alunos 2025'])}</td>"
                    html_eva += f"<td style='text-align:center; color:#e0e0ff;'>{int(r['Veteranos 2026'])}</td>"
                    html_eva += f"<td style='text-align:center; color:{cor}; font-weight:600;'>{evasao:.1f}%</td></tr>"
                html_eva += "</table>"
                st.markdown(html_eva, unsafe_allow_html=True)
                st.caption("Mostrando séries com base mínima de 5 alunos")

        # Detalhamento expandível
        with st.expander("📋 Ver detalhamento completo por série"):
            st.dataframe(
                df_retencao.sort_values(["Unidade", "Série 2026"]),
                use_container_width=True,
                hide_index=True
            )
    else:
        st.warning("Não foi possível calcular a retenção. Verifique os dados de 2025 e 2026.")
else:
    st.info("📋 Para calcular a retenção real, são necessários os dados de 2025 (`extrair_2025.py`) e 2026.")

st.markdown("<br>", unsafe_allow_html=True)

# ===== GRÁFICO DE OCUPAÇÃO GERAL POR UNIDADE/SEGMENTO =====
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

# ===== TERMÔMETRO DE OCUPAÇÃO GERAL POR UNIDADE =====
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

# ===== MAPA DE CALOR DE OCUPAÇÃO GERAL =====
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
ordem_seg = ['Ed. Infantil', 'Fund. 1', 'Fund. 2', 'Ens. Médio']
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
        c1.metric("Ocupação", f"{ocup:.1f}%")
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

# Calcula ocupação com tratamento para divisão por zero
df_relatorio['Ocupação %'] = df_relatorio.apply(
    lambda row: round((row['Matriculados'] / row['Vagas'] * 100), 1) if row['Vagas'] > 0 else 0.0,
    axis=1
)

# Garante que não há valores NaN ou infinitos
df_relatorio['Ocupação %'] = df_relatorio['Ocupação %'].fillna(0).replace([float('inf'), float('-inf')], 0)

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
        "% Veteranos",
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
df_ranking['% Veteranos'] = (df_ranking['Veteranos'] / df_ranking['Matriculados'] * 100).round(1)
df_ranking['Captação'] = (df_ranking['Novatos'] / df_ranking['Vagas'] * 100).round(1)
df_ranking = df_ranking.sort_values('Ocupação', ascending=False)

# Extrai nome curto
df_ranking['Unidade'] = df_ranking['Unidade'].apply(
    lambda x: x.split('(')[1].replace(')', '') if '(' in x else x
)

st.dataframe(
    df_ranking[['Unidade', 'Vagas', 'Matriculados', 'Ocupação', '% Veteranos', 'Captação', 'Disponiveis']],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Ocupação": st.column_config.ProgressColumn(
            "Ocupação %",
            format="%.1f%%",
            min_value=0,
            max_value=100,
        ),
        "% Veteranos": st.column_config.ProgressColumn(
            "% Veteranos",
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

# Adiciona coluna Série para filtro
df_relatorio['Série'] = df_relatorio['Turma'].apply(extrair_serie)

# Filtros inline para o detalhamento
col_filtro1, col_filtro2, col_filtro3, col_filtro4, col_filtro5 = st.columns(5)

with col_filtro1:
    unidades_det = ["Todas"] + sorted(df_relatorio['Unidade'].apply(
        lambda x: x.split('(')[1].replace(')', '') if '(' in x else x
    ).unique().tolist())
    filtro_unidade_det = st.selectbox("Filtrar Unidade", unidades_det, key="filtro_unidade_det")

with col_filtro2:
    segmentos_det = ["Todos"] + sorted(df_relatorio['Segmento'].unique().tolist())
    filtro_segmento_det = st.selectbox("Filtrar Segmento", segmentos_det, key="filtro_segmento_det")

with col_filtro3:
    series_det = ["Todas"] + sorted(df_relatorio['Série'].unique().tolist())
    filtro_serie_det = st.selectbox("Filtrar Série", series_det, key="filtro_serie_det")

with col_filtro4:
    turnos_det = ["Todos"] + sorted(df_relatorio['Turno'].unique().tolist())
    filtro_turno_det = st.selectbox("Filtrar Turno", turnos_det, key="filtro_turno_det")

with col_filtro5:
    ordenacao = st.selectbox("Ordenar por", ["Ocupação (maior)", "Ocupação (menor)", "Vagas (maior)", "Disponíveis (maior)"], key="ordenacao_det")

# Aplica filtros do detalhamento
df_det = df_relatorio.copy()

# Garante que todas as colunas numéricas não têm NaN
colunas_numericas = ['Vagas', 'Matriculados', 'Novatos', 'Veteranos', 'Disponiveis', 'Pre-matriculados', 'Ocupação %']
for col in colunas_numericas:
    if col in df_det.columns:
        df_det[col] = pd.to_numeric(df_det[col], errors='coerce').fillna(0)

if filtro_unidade_det != "Todas":
    df_det = df_det[df_det['Unidade'].str.contains(filtro_unidade_det, case=False)]

if filtro_segmento_det != "Todos":
    df_det = df_det[df_det['Segmento'] == filtro_segmento_det]

if filtro_serie_det != "Todas":
    df_det = df_det[df_det['Série'] == filtro_serie_det]

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

# Verifica se todas as colunas existem
colunas_disponiveis = [col for col in colunas_exibir if col in df_det.columns]
if len(colunas_disponiveis) != len(colunas_exibir):
    # Adiciona colunas faltantes com valor 0
    for col in colunas_exibir:
        if col not in df_det.columns:
            df_det[col] = 0

df_exibir = df_det[colunas_exibir].copy()

# Extrai nome curto da unidade
df_exibir['Unidade'] = df_exibir['Unidade'].apply(lambda x: x.split('(')[1].replace(')', '') if '(' in str(x) else str(x))
df_exibir.columns = ['Unidade', 'Segmento', 'Turma', 'Turno', 'Vagas', 'Matr.', 'Ocup.', 'Nov.', 'Vet.', 'Disp.', 'Pré']

# Prepara DataFrame para exibição
df_exibir['Ocup.'] = df_exibir['Ocup.'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "0%")

# Converte colunas numéricas para int
for col in ['Vagas', 'Matr.', 'Nov.', 'Vet.', 'Disp.', 'Pré']:
    df_exibir[col] = df_exibir[col].apply(lambda x: int(float(x)) if pd.notna(x) else 0)

# Exibe tabela usando Streamlit nativo
if len(df_exibir) > 0:
    st.dataframe(
        df_exibir,
        use_container_width=True,
        height=450,
        hide_index=True
    )
    st.caption(f"Exibindo {len(df_exibir)} turmas • Filtros: {filtro_unidade_det} | {filtro_segmento_det} | {filtro_turno_det}")
else:
    st.info("Nenhuma turma encontrada com os filtros selecionados.")

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
