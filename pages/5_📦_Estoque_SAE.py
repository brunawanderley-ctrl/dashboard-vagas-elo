#!/usr/bin/env python3
"""
Dashboard de Controle de Estoque - SAE + Socioemocional + Elo Tech (Livros)
Colegio Elo - Comparativo Pedido x Enviado x Vendido

Organizado por secoes: SAE | Socioemocional | Elo Tech | Visao Geral
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import json
import re
import io
from pathlib import Path
from datetime import datetime
import sys
sys.path.insert(0, str(Path(__file__).parent.parent / "utils"))
from dados_estoque import (
    PEDIDO_SAE, ESTOQUE_ENVIADO, UNIDADES, ORDEM_SEGMENTOS,
    BALANCO_FISICO, AJUSTE_ANO_PASSADO, ELOTECH_EM_ABERTO,
    ELOTECH_SERIE_PDF, AJUSTE_SOCIO_2025,
)
from siga_api import atualizar_via_api

# CSS
st.markdown("""
<style>
    .main { background: #f8fafc !important; }
    .stApp { background: #f8fafc !important; }
    .stMarkdown, .stText, p, span, label { color: #1e293b !important; }
    h1, h2, h3, h4, h5, h6 { color: #0f172a !important; }
    [data-testid="stSidebar"] { background: #ffffff !important; }
    [data-testid="stSidebar"] * { color: #1e293b !important; }
    [data-testid="stMetricValue"] { color: #667eea !important; font-weight: 700 !important; }
    [data-testid="stMetricLabel"] { color: #475569 !important; }
    .stDataFrame { color: #1e293b !important; }
    .stButton > button {
        background: linear-gradient(135deg, #667eea, #764ba2) !important;
        color: white !important; border: none !important;
    }
    .positive { color: #22c55e !important; font-weight: bold; }
    .negative { color: #ef4444 !important; font-weight: bold; }
    .warning { color: #f59e0b !important; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# Constantes
BASE_DIR = Path(__file__).parent.parent
OUTPUT_DIR = BASE_DIR / "output"

CODIGOS_SERVICOS = {
    "901", "902", "903", "904",
    "913", "914", "915", "920", "921",
    "916", "917", "918", "919",
    "912", "991", "992",
    "933", "934", "941", "942", "943", "944", "945", "946",
    "995",
}

# Cordeiro NAO tem Elo Tech
CODIGOS_EXCLUIR_POR_UNIDADE = {
    "CDR": {"995"},
}

SERVICOS_MAP = {
    "901": ("Infantil", "Infantil II", "SAE"),
    "902": ("Infantil", "Infantil III", "SAE"),
    "903": ("Infantil", "Infantil IV", "SAE"),
    "904": ("Infantil", "Infantil V", "SAE"),
    "913": ("Fund1", "3º Ano", "SAE"),
    "914": ("Fund1", "4º Ano", "SAE"),
    "915": ("Fund1", "5º Ano", "SAE"),
    "920": ("Fund1", "2º Ano", "SAE"),
    "921": ("Fund1", "1º Ano", "SAE"),
    "916": ("Fund2", "6º Ano", "SAE"),
    "917": ("Fund2", "7º Ano", "SAE"),
    "918": ("Fund2", "8º Ano", "SAE"),
    "919": ("Fund2", "9º Ano", "SAE"),
    "912": ("Médio", "1º Ano", "SAE"),
    "991": ("Médio", "2º Ano", "SAE"),
    "992": ("Médio", "3º Ano", "SAE"),
    "933": ("Infantil", "Infantil IV", "Socioemocional"),
    "934": ("Infantil", "Infantil V", "Socioemocional"),
    "941": ("Fund1", "1º Ano", "Socioemocional"),
    "942": ("Fund1", "2º Ano", "Socioemocional"),
    "943": ("Fund1", "3º Ano", "Socioemocional"),
    "944": ("Fund1", "4º Ano", "Socioemocional"),
    "945": ("Fund1", "5º Ano", "Socioemocional"),
    "946": ("Fund2", "6º Ano", "Socioemocional"),
    "995": ("Geral", "Todas", "Elo Tech"),
}

RE_SERVICO = re.compile(r'^(\d{3})\s*-\s*(.+?)(?:\s*\(|$)')
RE_TURMA = re.compile(r'Turma\s+(\w+)')

ORDEM_SERIES_COMPLETA = [
    ('Infantil', 'Infantil II'), ('Infantil', 'Infantil III'),
    ('Infantil', 'Infantil IV'), ('Infantil', 'Infantil V'),
    ('Fund1', '1º Ano'), ('Fund1', '2º Ano'), ('Fund1', '3º Ano'),
    ('Fund1', '4º Ano'), ('Fund1', '5º Ano'),
    ('Fund2', '6º Ano'), ('Fund2', '7º Ano'),
    ('Fund2', '8º Ano'), ('Fund2', '9º Ano'),
    ('Médio', '1º Ano'), ('Médio', '2º Ano'), ('Médio', '3º Ano'),
]

ORDEM_SERIES_SOCIO = [
    ('Infantil', 'Infantil IV'), ('Infantil', 'Infantil V'),
    ('Fund1', '1º Ano'), ('Fund1', '2º Ano'), ('Fund1', '3º Ano'),
    ('Fund1', '4º Ano'), ('Fund1', '5º Ano'),
    ('Fund2', '6º Ano'),
]

ORDEM_SERIES_ELOTECH = [
    ('Fund1', '2º Ano'), ('Fund1', '3º Ano'),
    ('Fund1', '4º Ano'), ('Fund1', '5º Ano'),
]

# Mapa inverso para converter nome -> codigo
UNIDADES_INV = {v: k for k, v in UNIDADES.items()}


# ===========================================
# FUNCOES DE ATUALIZACAO
# ===========================================

def parse_tsv(tsv_content, unidade_codigo):
    """Parser de TSV dos dados do SIGA"""
    registros = []
    servico_codigo = ""
    turma_atual = ""
    excluir = CODIGOS_EXCLUIR_POR_UNIDADE.get(unidade_codigo, set())

    for linha in tsv_content.split('\n'):
        linha = linha.strip()
        if not linha or linha.startswith('Subtotal') or linha.startswith('Matrícula'):
            continue

        if ' - ' in linha and '(' in linha:
            match = RE_SERVICO.match(linha)
            if match:
                servico_codigo = match.group(1)
                if servico_codigo not in CODIGOS_SERVICOS or servico_codigo in excluir:
                    servico_codigo = ""
                    continue
                turma_match = RE_TURMA.search(linha)
                turma_atual = turma_match.group(1) if turma_match else ""
            continue

        if not servico_codigo:
            continue

        campos = linha.split('\t')
        if len(campos) >= 6:
            matricula = campos[0].strip()
            if matricula and (matricula[0].isdigit() or '-' in matricula):
                info = SERVICOS_MAP.get(servico_codigo, ("", "", ""))
                registros.append({
                    "unidade": unidade_codigo,
                    "servico_codigo": servico_codigo,
                    "turma": turma_atual,
                    "matricula": matricula,
                    "nome": campos[1].strip(),
                    "titulo": campos[2].strip(),
                    "parcela": campos[3].strip(),
                    "dt_baixa": campos[4].strip(),
                    "valor": campos[5].strip(),
                    "recebido": campos[-1].strip() if len(campos) > 6 else "",
                    "segmento": info[0],
                    "serie": info[1],
                    "tipo": info[2],
                })

    return registros


def atualizar_dados():
    """Atualiza dados via API SIGA (requests, sem Playwright).

    Returns:
        -2 se credenciais nao configuradas
        -1 se erro na API
        (total_registros, datetime) se sucesso
    """
    try:
        creds = st.secrets["siga"]
    except Exception:
        return -2

    registros, erro = atualizar_via_api(
        creds["instituicao"], creds["login"], creds["senha"]
    )
    if erro:
        return -1

    agora = datetime.now()
    json_path = OUTPUT_DIR / "recebimento_final.json"
    dados = {
        "data_extracao": agora.isoformat(),
        "ultima_atualizacao": agora.strftime("%d/%m/%Y %H:%M:%S"),
        "periodo": {
            "data_inicial": "01/08/2025",
            "data_final": agora.strftime("%d/%m/%Y"),
        },
        "total_registros": len(registros),
        "registros": registros,
    }
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    return len(registros), agora


# ===========================================
# CARREGAR DADOS
# ===========================================

@st.cache_data(ttl=300)
def carregar_vendas():
    """Carrega dados de vendas do JSON"""
    json_path = OUTPUT_DIR / "recebimento_final.json"
    if not json_path.exists():
        return None, None

    with open(json_path, 'r', encoding='utf-8') as f:
        dados = json.load(f)

    df = pd.DataFrame(dados.get('registros', []))
    ultima_att = dados.get('ultima_atualizacao', 'N/A')
    return df, ultima_att


def calcular_vendas_por_serie_unidade(df, tipo="SAE"):
    """Calcula vendas (alunos unicos) por serie e unidade"""
    df_filtrado = df[df['tipo'] == tipo].copy()
    vendas = df_filtrado.groupby(['segmento', 'serie', 'unidade'])['matricula'].nunique().reset_index()
    vendas.columns = ['segmento', 'serie', 'unidade', 'vendido']
    return vendas


def calcular_vendas_elotech_por_serie(df):
    """Cruza matriculas Elo Tech com SAE/Socioemocional para descobrir a serie."""
    df_tech = df[df['tipo'] == 'Elo Tech'].copy()
    df_ref = df[df['tipo'].isin(['SAE', 'Socioemocional'])].copy()
    if df_tech.empty:
        return pd.DataFrame(columns=['segmento', 'serie', 'unidade', 'vendido'])
    mat_serie = df_ref.drop_duplicates(subset=['matricula', 'unidade'])\
        .set_index(['matricula', 'unidade'])['serie'].to_dict()
    df_tech['serie_real'] = df_tech.apply(
        lambda r: ELOTECH_SERIE_PDF.get((r['matricula'], r['unidade']),
                  mat_serie.get((r['matricula'], r['unidade']),
                                r['serie'] if r['serie'] != 'Todas' else 'Sem serie')), axis=1
    )
    df_tech['segmento_real'] = df_tech['serie_real'].apply(
        lambda s: 'Fund1' if s in ['1º Ano', '2º Ano', '3º Ano', '4º Ano', '5º Ano']
        else 'Fund2' if s in ['6º Ano', '7º Ano', '8º Ano', '9º Ano']
        else 'Outros'
    )
    vendas = df_tech.groupby(['segmento_real', 'serie_real', 'unidade'])['matricula']\
        .nunique().reset_index()
    vendas.columns = ['segmento', 'serie', 'unidade', 'vendido']
    return vendas


def criar_tabela_estoque(df_vendas):
    """Cria tabela completa de controle de estoque"""
    registros = []

    for cod, (serie, segmento, pedido_inicial, pedido_compl) in PEDIDO_SAE.items():
        pedido_total = pedido_inicial + pedido_compl
        estoque_enviado = ESTOQUE_ENVIADO.get(cod, {})
        ajuste_ap = AJUSTE_ANO_PASSADO.get(cod, {})

        for unidade_cod, unidade_nome in UNIDADES.items():
            enviado = estoque_enviado.get(unidade_cod, 0)
            ajuste = ajuste_ap.get(unidade_cod, 0)

            vendido = 0
            if df_vendas is not None:
                filtro = (
                    (df_vendas['segmento'] == segmento) &
                    (df_vendas['serie'] == serie) &
                    (df_vendas['unidade'] == unidade_cod)
                )
                if filtro.any():
                    vendido = df_vendas.loc[filtro, 'vendido'].values[0]

            estoque_restante = enviado - vendido + ajuste

            registros.append({
                'codigo': cod,
                'segmento': segmento,
                'serie': serie,
                'unidade_cod': unidade_cod,
                'unidade': unidade_nome,
                'pedido_inicial': pedido_inicial,
                'pedido_compl': pedido_compl,
                'pedido_total': pedido_total,
                'enviado': enviado,
                'ajuste': ajuste,
                'vendido': vendido,
                'estoque': estoque_restante
            })

    df = pd.DataFrame(registros)
    df['segmento_ordem'] = df['segmento'].map({s: i for i, s in enumerate(ORDEM_SEGMENTOS)})
    df = df.sort_values(['segmento_ordem', 'serie', 'unidade_cod'])
    df = df.drop(columns=['segmento_ordem'])
    return df


def criar_tabela_completa(df_dados, coluna_valor, unidade_selecionada=None):
    """Cria tabela com TODAS as series"""
    if unidade_selecionada and unidade_selecionada != "Todas":
        unidades_mostrar = [unidade_selecionada]
    else:
        unidades_mostrar = list(UNIDADES.values())

    registros = []
    for segmento, serie in ORDEM_SERIES_COMPLETA:
        registro = {'segmento': segmento, 'serie': serie, 'Descrição': f"{segmento} - {serie}"}
        for unidade_nome in unidades_mostrar:
            filtro = (
                (df_dados['segmento'] == segmento) &
                (df_dados['serie'] == serie) &
                (df_dados['unidade'] == unidade_nome)
            )
            registro[unidade_nome] = int(df_dados.loc[filtro, coluna_valor].sum()) if filtro.any() else 0
        registros.append(registro)

    df_resultado = pd.DataFrame(registros)
    colunas_valores = [c for c in df_resultado.columns if c not in ['segmento', 'serie', 'Descrição']]
    df_resultado['TOTAL'] = df_resultado[colunas_valores].sum(axis=1)

    totais = {'segmento': '', 'serie': '', 'Descrição': 'TOTAL'}
    for col in colunas_valores:
        totais[col] = df_resultado[col].sum()
    totais['TOTAL'] = df_resultado['TOTAL'].sum()
    df_resultado = pd.concat([df_resultado, pd.DataFrame([totais])], ignore_index=True)

    df_resultado = df_resultado.drop(columns=['segmento', 'serie'])
    cols = ['Descrição'] + colunas_valores + ['TOTAL']
    return df_resultado[cols]


def colorir_estoque(val):
    if isinstance(val, str):
        return ''
    if val < 0:
        return 'background-color: #fecaca; color: #dc2626; font-weight: bold'
    elif val <= 5:
        return 'background-color: #fef3c7; color: #d97706; font-weight: bold'
    else:
        return 'background-color: #d1fae5; color: #059669'


def colorir_diferenca(val):
    if isinstance(val, str):
        return ''
    if val < 0:
        return 'background-color: #fecaca; color: #dc2626; font-weight: bold'
    elif val > 0:
        return 'background-color: #fef3c7; color: #d97706; font-weight: bold'
    return 'background-color: #d1fae5; color: #059669'


# ===========================================
# HELPER: GERAR EXCEL EM MEMORIA
# ===========================================

def _to_excel_bytes(df, sheet_name="Dados"):
    """Converte um DataFrame em bytes Excel (openpyxl)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


def _to_excel_multi_sheets(sheets_dict):
    """Converte multiplos DataFrames em um unico Excel com varias abas.
    sheets_dict: {nome_aba: dataframe}
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


# ===========================================
# HELPER: HTML DE IMPRESSAO POR SECAO
# ===========================================

def df_to_html_colorido(df, titulo, colunas_colorir=None):
    """Converte DataFrame para HTML com celulas coloridas para impressao."""
    html = f"<h3>{titulo}</h3>\n" if titulo else ""
    html += "<table>\n<thead><tr>"
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead>\n<tbody>\n"
    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            val = row[col]
            style = ""
            if colunas_colorir and col in colunas_colorir and not isinstance(val, str):
                try:
                    v = float(val)
                    if v < 0:
                        style = ' class="cell-neg"'
                    elif v <= 5:
                        style = ' class="cell-warn"'
                    else:
                        style = ' class="cell-ok"'
                except (ValueError, TypeError):
                    pass
            elif colunas_colorir is None and col not in ("Descricao", "Descrição", "Unidade") and not isinstance(val, str):
                try:
                    v = float(val)
                    if v < 0:
                        style = ' class="cell-neg"'
                    elif v <= 5:
                        style = ' class="cell-warn"'
                    else:
                        style = ' class="cell-ok"'
                except (ValueError, TypeError):
                    pass
            html += f"<td{style}>{val}</td>"
        html += "</tr>\n"
    html += "</tbody></table>\n"
    return html


def gerar_html_impressao_secao(titulo, kpis, tabelas_html, data_str, filtro_unidade="Todas"):
    """Gera HTML de impressao para uma secao individual."""
    unidade_label = f" - {filtro_unidade}" if filtro_unidade != "Todas" else ""
    kpi_boxes = ""
    for lbl, val in kpis:
        kpi_boxes += f'<div class="kpi-box"><div class="val">{val}</div><div class="lbl">{lbl}</div></div>\n'

    tabelas_combined = "\n".join(tabelas_html)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Colegio Elo - {titulo}</title>
<style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 11px; color: #1e293b; padding: 20px;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
    }}
    .header {{
        text-align: center; border-bottom: 3px solid #667eea;
        padding-bottom: 10px; margin-bottom: 20px;
    }}
    .header h1 {{ font-size: 18px; color: #0f172a; }}
    .header p {{ font-size: 12px; color: #667eea; margin-top: 4px; }}
    .header .date {{ font-size: 10px; color: #64748b; margin-top: 4px; }}
    h3 {{ font-size: 13px; color: #334155; margin: 16px 0 6px 0; }}
    table {{
        width: 100%; border-collapse: collapse; margin-bottom: 16px;
        page-break-inside: auto;
    }}
    th {{
        background-color: #667eea !important; color: #ffffff !important;
        font-weight: 600; padding: 5px 8px; text-align: center;
        font-size: 10px; border: 1px solid #94a3b8;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }}
    td {{
        padding: 4px 8px; text-align: center; border: 1px solid #cbd5e1;
        font-size: 10px;
    }}
    td:first-child {{ text-align: left; font-weight: 500; }}
    tr:nth-child(even) {{ background-color: #f1f5f9 !important; }}
    tr:last-child {{ font-weight: bold; background-color: #e2e8f0 !important; }}
    .cell-neg {{
        background-color: #fecaca !important; color: #dc2626 !important;
        font-weight: bold;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }}
    .cell-warn {{
        background-color: #fef3c7 !important; color: #d97706 !important;
        font-weight: bold;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }}
    .cell-ok {{
        background-color: #d1fae5 !important; color: #059669 !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }}
    .legend {{
        font-size: 9px; color: #64748b; margin-top: 4px; margin-bottom: 12px;
    }}
    .kpi-row {{
        display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap;
    }}
    .kpi-box {{
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px;
        padding: 8px 16px; text-align: center; flex: 1; min-width: 100px;
    }}
    .kpi-box .val {{ font-size: 16px; font-weight: 700; color: #667eea; }}
    .kpi-box .lbl {{ font-size: 9px; color: #64748b; }}
    @media print {{
        body {{ padding: 10px; }}
        .no-print {{ display: none !important; }}
        table {{ page-break-inside: auto; }}
        tr {{ page-break-inside: avoid; page-break-after: auto; }}
    }}
</style>
</head>
<body>
<div class="header">
    <h1>Colegio Elo - {titulo}{unidade_label}</h1>
    <p>Controle de Estoque</p>
    <div class="date">Gerado em: {data_str}</div>
</div>

<div class="kpi-row">
{kpi_boxes}
</div>

{tabelas_combined}

<div class="legend">
    <strong style="color:#dc2626">Vermelho</strong>: Negativo (falta) &nbsp;
    <strong style="color:#d97706">Amarelo</strong>: Baixo (0-5) &nbsp;
    <strong style="color:#059669">Verde</strong>: OK (>5)
</div>

<script>
    window.onload = function() {{ window.print(); }};
</script>
</body>
</html>"""
    return html


def gerar_html_impressao_completo(kpis, tabelas_html, data_str, filtro_unidade="Todas"):
    """Gera HTML completo de impressao com todas as secoes."""
    return gerar_html_impressao_secao(
        "Controle de Estoque Completo", kpis, tabelas_html, data_str, filtro_unidade
    )


# ===========================================
# ATAS DE ENTREGA (FUNCOES)
# ===========================================

def gerar_ata_html(df_ata_src, titulo, col_serie='serie'):
    """Gera HTML de ata de entrega agrupada por unidade e turma (serie+turma)."""
    data_atual = datetime.now().strftime("%d/%m/%Y")
    unidades_ata = sorted(df_ata_src['unidade'].unique())
    total_alunos = len(df_ata_src)
    # Montar label de turma: "2o Ano A" ou "2o Ano" se turma vazia
    df_ata_src = df_ata_src.copy()
    df_ata_src['_turma_label'] = df_ata_src.apply(
        lambda r: f"{r[col_serie]} {r['turma']}".strip() if r['turma'].strip() else str(r[col_serie]),
        axis=1
    )
    html = f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<title>Ata de Entrega - {titulo}</title><style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',Tahoma,sans-serif; font-size:11px; color:#1e293b; padding:20px;
  -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }}
.header {{ text-align:center; border-bottom:3px solid #667eea; padding-bottom:10px; margin-bottom:15px; }}
.header h1 {{ font-size:16px; color:#0f172a; }} .header p {{ font-size:11px; color:#64748b; margin-top:4px; }}
h2 {{ font-size:14px; color:#0f172a; margin:20px 0 4px; background:#e0e7ff; padding:6px 10px; }}
h3 {{ font-size:12px; color:#334155; margin:12px 0 4px; border-bottom:1px solid #e2e8f0; padding-bottom:3px; }}
table {{ width:100%; border-collapse:collapse; margin-bottom:8px; page-break-inside:auto; }}
th {{ background:#667eea !important; color:#fff !important; font-weight:600; padding:5px 8px;
  text-align:center; font-size:10px; border:1px solid #94a3b8;
  -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }}
td {{ padding:4px 8px; text-align:center; border:1px solid #cbd5e1; font-size:10px; }}
td:nth-child(2) {{ text-align:left; }} td:last-child {{ min-width:120px; }}
tr:nth-child(even) {{ background:#f8fafc !important; }}
.assinatura {{ margin-top:30px; text-align:center; page-break-inside:avoid; margin-bottom:20px; }}
.assinatura .linha {{ border-top:1px solid #1e293b; width:300px; margin:30px auto 4px; }}
.assinatura p {{ font-size:10px; color:#64748b; }}
.no-print {{ margin-bottom:12px; }}
@media print {{ .no-print {{ display:none !important; }} tr {{ page-break-inside:avoid; }} h2 {{ page-break-before:auto; }} }}
</style></head><body>
<div class="header"><h1>Colegio Elo - Ata de Entrega: {titulo}</h1>
<p>Data: {data_atual} | Total: {total_alunos} alunos</p></div>
<button class="no-print" onclick="window.print()" style="padding:8px 20px;background:#667eea;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;">Imprimir</button>
"""
    for u_cod in unidades_ata:
        u_nome = UNIDADES.get(u_cod, u_cod)
        df_u = df_ata_src[df_ata_src['unidade'] == u_cod]
        html += f'<h2>{u_nome} ({len(df_u)} alunos)</h2>\n'
        for turma_label in sorted(df_u['_turma_label'].unique()):
            df_t = df_u[df_u['_turma_label'] == turma_label].sort_values('nome')
            html += f'<h3>{turma_label} ({len(df_t)} alunos)</h3>\n'
            html += '<table><thead><tr><th>N</th><th>Nome do Aluno</th><th>Matricula</th><th>Assinatura</th></tr></thead><tbody>\n'
            for i, (_, row) in enumerate(df_t.iterrows(), 1):
                html += f'<tr><td>{i}</td><td style="text-align:left">{row["nome"]}</td><td>{row["matricula"]}</td><td></td></tr>\n'
            html += '</tbody></table>\n'
        html += '<div class="assinatura"><div class="linha"></div><p>Responsavel pela Entrega - ' + u_nome + '</p></div>\n'
    html += '</body></html>'
    return html


def gerar_ata_elotech_html(df_ata_src):
    """Gera HTML de ata Elo Tech por unidade e serie, com colunas turno e turma."""
    data_atual = datetime.now().strftime("%d/%m/%Y")
    unidades_ata = sorted(df_ata_src['unidade'].unique())
    total_alunos = len(df_ata_src)
    html = f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<title>Ata de Entrega - Elo Tech</title><style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',Tahoma,sans-serif; font-size:11px; color:#1e293b; padding:20px;
  -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }}
.header {{ text-align:center; border-bottom:3px solid #667eea; padding-bottom:10px; margin-bottom:15px; }}
.header h1 {{ font-size:16px; color:#0f172a; }} .header p {{ font-size:11px; color:#64748b; margin-top:4px; }}
h2 {{ font-size:14px; color:#0f172a; margin:20px 0 4px; background:#e0e7ff; padding:6px 10px; }}
h3 {{ font-size:12px; color:#334155; margin:12px 0 4px; border-bottom:1px solid #e2e8f0; padding-bottom:3px; }}
table {{ width:100%; border-collapse:collapse; margin-bottom:8px; page-break-inside:auto; }}
th {{ background:#667eea !important; color:#fff !important; font-weight:600; padding:5px 8px;
  text-align:center; font-size:10px; border:1px solid #94a3b8;
  -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }}
td {{ padding:4px 8px; text-align:center; border:1px solid #cbd5e1; font-size:10px; }}
td:nth-child(2) {{ text-align:left; }} td:last-child {{ min-width:120px; }}
tr:nth-child(even) {{ background:#f8fafc !important; }}
.assinatura {{ margin-top:30px; text-align:center; page-break-inside:avoid; margin-bottom:20px; }}
.assinatura .linha {{ border-top:1px solid #1e293b; width:300px; margin:30px auto 4px; }}
.assinatura p {{ font-size:10px; color:#64748b; }}
.no-print {{ margin-bottom:12px; }}
@media print {{ .no-print {{ display:none !important; }} tr {{ page-break-inside:avoid; }} h2 {{ page-break-before:auto; }} }}
</style></head><body>
<div class="header"><h1>Colegio Elo - Ata de Entrega: Livros Elo Tech</h1>
<p>Data: {data_atual} | Total: {total_alunos} alunos</p></div>
<button class="no-print" onclick="window.print()" style="padding:8px 20px;background:#667eea;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;">Imprimir</button>
"""
    for u_cod in unidades_ata:
        u_nome = UNIDADES.get(u_cod, u_cod)
        df_u = df_ata_src[df_ata_src['unidade'] == u_cod]
        html += f'<h2>{u_nome} ({len(df_u)} alunos)</h2>\n'
        for serie in sorted(df_u['serie_real'].unique()):
            df_s = df_u[df_u['serie_real'] == serie].sort_values('nome')
            html += f'<h3>{serie} ({len(df_s)} alunos)</h3>\n'
            html += '<table><thead><tr><th>N</th><th>Nome do Aluno</th><th>Matricula</th><th>Turno</th><th>Turma</th><th>Assinatura</th></tr></thead><tbody>\n'
            for i, (_, row) in enumerate(df_s.iterrows(), 1):
                turno = row.get('turno', '')
                turma = row.get('turma_real', '')
                html += f'<tr><td>{i}</td><td style="text-align:left">{row["nome"]}</td><td>{row["matricula"]}</td><td>{turno}</td><td>{turma}</td><td></td></tr>\n'
            html += '</tbody></table>\n'
        html += '<div class="assinatura"><div class="linha"></div><p>Responsavel pela Entrega - ' + u_nome + '</p></div>\n'
    html += '</body></html>'
    return html


def _filtrar_unidade_ata(df_src, filtro_un):
    """Filtra DataFrame de ata pela unidade selecionada."""
    if filtro_un and filtro_un != "Todas":
        u_cod = UNIDADES_INV.get(filtro_un, "")
        return df_src[df_src['unidade'] == u_cod]
    return df_src


def _abrir_impressao(html_content):
    """Renderiza script para abrir janela de impressao."""
    components.html(
        f"""
        <script>
            var printWindow = window.open('', '_blank');
            printWindow.document.write({json.dumps(html_content)});
            printWindow.document.close();
        </script>
        <p style="color: #059669; font-size: 14px;">Janela de impressao aberta. Se nao aparecer, verifique o bloqueador de pop-ups.</p>
        """,
        height=40,
    )


# ===========================================
# AUTO-ATUALIZACAO PROGRAMADA
# ===========================================

HORARIOS_ATUALIZACAO = [7, 9, 12, 14, 17, 19]


def _precisa_atualizar():
    """Verifica se dados estao defasados em relacao ao horario programado."""
    json_path = OUTPUT_DIR / "recebimento_final.json"
    if not json_path.exists():
        return True

    agora = datetime.now()
    horarios_passados = [h for h in HORARIOS_ATUALIZACAO if h <= agora.hour]
    if not horarios_passados:
        return False

    ultimo_horario = max(horarios_passados)
    limite = agora.replace(hour=ultimo_horario, minute=0, second=0, microsecond=0)
    mtime = datetime.fromtimestamp(json_path.stat().st_mtime)
    return mtime < limite


if '_auto_att_feita' not in st.session_state:
    st.session_state['_auto_att_feita'] = False

if not st.session_state['_auto_att_feita'] and _precisa_atualizar():
    _resultado_auto = atualizar_dados()
    st.session_state['_auto_att_feita'] = True
    if isinstance(_resultado_auto, tuple):
        _n_auto, _ = _resultado_auto
        st.cache_data.clear()
        st.toast(f"Dados atualizados automaticamente: {_n_auto} registros ({datetime.now().strftime('%H:%M')})")
        st.rerun()


# ===========================================
# HEADER + BOTAO ATUALIZAR
# ===========================================

df_vendas_raw, ultima_atualizacao = carregar_vendas()

# Injetar alunos Elo Tech com pagamento "Em aberto" (faturados, nao pagos)
if df_vendas_raw is not None and ELOTECH_EM_ABERTO:
    _mats_existentes = set(df_vendas_raw[df_vendas_raw['tipo'] == 'Elo Tech']['matricula'].unique())
    _novos = []
    for mat, nome, un, serie_pdf in ELOTECH_EM_ABERTO:
        if mat not in _mats_existentes:
            _novos.append({
                'unidade': un, 'servico_codigo': '995', 'turma': '',
                'matricula': mat, 'nome': nome, 'titulo': '', 'parcela': 'TAXA',
                'dt_baixa': '', 'valor': '299,00', 'recebido': '0,00',
                'segmento': 'Geral', 'serie': serie_pdf, 'tipo': 'Elo Tech',
            })
    if _novos:
        df_vendas_raw = pd.concat([df_vendas_raw, pd.DataFrame(_novos)], ignore_index=True)

col_title, col_btn = st.columns([4, 1])
with col_title:
    st.markdown("""
        <h1 style='margin-bottom: 0; color: #1e293b;'>📦 Controle de Estoque SAE</h1>
        <p style='color: #667eea; font-size: 1.1rem; margin-top: 0.5rem;'>Colegio Elo - Pedido x Enviado x Vendido</p>
    """, unsafe_allow_html=True)
with col_btn:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Atualizar Dados", type="primary", use_container_width=True):
        # Capturar contagem anterior para comparacao
        _json_antes = OUTPUT_DIR / "recebimento_final.json"
        _regs_antes = 0
        if _json_antes.exists():
            try:
                with open(_json_antes, 'r', encoding='utf-8') as _f:
                    _regs_antes = len(json.load(_f).get('registros', []))
            except Exception:
                pass

        with st.spinner("Conectando ao SIGA (4 unidades em paralelo)..."):
            resultado = atualizar_dados()

        if resultado == -2:
            st.error("Credenciais SIGA nao configuradas em Secrets. Configure em .streamlit/secrets.toml")
        elif resultado == -1:
            st.error("Erro ao conectar com SIGA. Verifique internet/credenciais e tente novamente.")
        else:
            total_reg, dt_att = resultado
            _diff = total_reg - _regs_antes
            _diff_txt = f" ({_diff:+d} vs anterior)" if _regs_antes > 0 else ""
            st.cache_data.clear()
            st.success(f"Atualizado! {total_reg} registros{_diff_txt} em {dt_att.strftime('%H:%M:%S')}")
            if _diff < 0 and _regs_antes > 0:
                st.warning(f"Atencao: {abs(_diff)} registros a menos que a extracao anterior ({_regs_antes}). Verifique se ha titulos removidos no SIGA.")
            st.rerun()

if ultima_atualizacao:
    st.caption(f"Dados atualizados em: {ultima_atualizacao}")

if df_vendas_raw is None:
    st.error("Dados de vendas nao encontrados. Clique em 'Atualizar Dados' ou execute a extracao.")
    st.stop()


# ===========================================
# COMPUTACAO GLOBAL (antes de qualquer render)
# ===========================================

# --- Vendas e estoque SAE ---
df_vendas_sae = calcular_vendas_por_serie_unidade(df_vendas_raw, "SAE")
df_vendas_socio = calcular_vendas_por_serie_unidade(df_vendas_raw, "Socioemocional")
df_estoque_completo = criar_tabela_estoque(df_vendas_sae)

# --- Vendas Elo Tech (computado antes de render) ---
df_vendas_elotech = calcular_vendas_elotech_por_serie(df_vendas_raw)

# --- Detail Elo Tech (para expanders e ata) ---
df_vendas_elotech_detail = df_vendas_raw[df_vendas_raw['tipo'] == 'Elo Tech'].copy()
df_ref_detail = df_vendas_raw[df_vendas_raw['tipo'].isin(['SAE', 'Socioemocional'])].copy()
_mat_serie = df_ref_detail.drop_duplicates(subset=['matricula', 'unidade'])\
    .set_index(['matricula', 'unidade'])['serie'].to_dict()
_mat_turma = df_ref_detail[df_ref_detail['turma'].str.strip() != '']\
    .drop_duplicates(subset=['matricula', 'unidade'])\
    .set_index(['matricula', 'unidade'])['turma'].to_dict()
df_vendas_elotech_detail['serie_real'] = df_vendas_elotech_detail.apply(
    lambda r: ELOTECH_SERIE_PDF.get((r['matricula'], r['unidade']),
              _mat_serie.get((r['matricula'], r['unidade']),
                             r['serie'] if r['serie'] != 'Todas' else 'Sem serie')), axis=1
)
df_vendas_elotech_detail['turma_real'] = df_vendas_elotech_detail.apply(
    lambda r: _mat_turma.get((r['matricula'], r['unidade']),
                             r['turma'].strip() if r['turma'].strip() else ''), axis=1
)
_TURNO_MAP = {'A': 'Manha', 'B': 'Tarde', 'C': 'Integral'}
df_vendas_elotech_detail['turno'] = df_vendas_elotech_detail['turma_real'].map(
    lambda t: _TURNO_MAP.get(t.strip().upper(), '') if t.strip() else ''
)

# --- Excel export data (SAE) ---
df_estoque_export = df_estoque_completo[['segmento', 'serie', 'unidade', 'pedido_total', 'enviado', 'vendido', 'ajuste', 'estoque']].rename(
    columns={'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade',
             'pedido_total': 'Pedido', 'enviado': 'Enviado', 'vendido': 'Vendido',
             'ajuste': 'Ajuste 2025', 'estoque': 'Estoque'}
)

# ===========================================
# FILTROS (TOPO)
# ===========================================
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    filtro_unidade = st.selectbox("Unidade", ["Todas"] + list(UNIDADES.values()), key="filtro_unidade")
with col_f2:
    segmento_filtro = st.selectbox("Segmento", ["Todos"] + ORDEM_SEGMENTOS, key="filtro_segmento")
with col_f3:
    if segmento_filtro != "Todos":
        series_disponiveis = sorted(df_estoque_completo[df_estoque_completo['segmento'] == segmento_filtro]['serie'].unique().tolist())
    else:
        series_disponiveis = sorted(df_estoque_completo['serie'].unique().tolist())
    serie_filtro = st.selectbox("Serie", ["Todas"] + series_disponiveis, key="filtro_serie")

# Aplica filtros
df_estoque = df_estoque_completo.copy()
if filtro_unidade != "Todas":
    df_estoque = df_estoque[df_estoque['unidade'] == filtro_unidade]
if segmento_filtro != "Todos":
    df_estoque = df_estoque[df_estoque['segmento'] == segmento_filtro]
if serie_filtro != "Todas":
    df_estoque = df_estoque[df_estoque['serie'] == serie_filtro]

# Filtros ativos
filtros_ativos = []
if filtro_unidade != "Todas":
    filtros_ativos.append(f"Unidade: {filtro_unidade}")
if segmento_filtro != "Todos":
    filtros_ativos.append(f"Segmento: {segmento_filtro}")
if serie_filtro != "Todas":
    filtros_ativos.append(f"Serie: {serie_filtro}")

if filtros_ativos:
    st.info(f"Filtros ativos: {' | '.join(filtros_ativos)}")

st.divider()

# ===========================================
# KPIs GLOBAIS
# ===========================================
pedido_total = df_estoque.drop_duplicates(subset=['codigo'])['pedido_total'].sum()
enviado_total = df_estoque['enviado'].sum()
vendido_total = df_estoque['vendido'].sum()
estoque_total = df_estoque['estoque'].sum()
ajuste_total = df_estoque['ajuste'].sum()
percentual_venda = (vendido_total / pedido_total * 100) if pedido_total > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Pedido Total", f"{pedido_total:,}")
with col2:
    st.metric("Enviado Total", f"{enviado_total:,}")
with col3:
    st.metric("Vendido SAE", f"{vendido_total:,}")
with col4:
    st.metric("Estoque Restante", f"{estoque_total:,}")
with col5:
    st.metric("% Vendido", f"{percentual_venda:.1f}%")


# ===========================================
# BRIEFING KAROL — Analise de Inconsistencias por Unidade
# ===========================================

st.markdown("### Briefing Karol — Pendencias por Unidade")
st.caption(f"Analise automatica gerada em {datetime.now().strftime('%d/%m/%Y %H:%M')} — "
           f"dados do SIGA de {ultima_atualizacao or 'N/A'}")

_pedido_socio = {
    'Infantil IV': 100, 'Infantil V': 125,
    '1\u00ba Ano': 147, '2\u00ba Ano': 140,
    '3\u00ba Ano': 135, '4\u00ba Ano': 130,
    '5\u00ba Ano': 155, '6\u00ba Ano': 160,
}
_pedido_elotech = {'2\u00ba Ano': 150, '3\u00ba Ano': 155, '4\u00ba Ano': 135, '5\u00ba Ano': 150}

for _u_cod, _u_nome in [("BV", "Boa Viagem"), ("CD", "Candeias"), ("JG", "Janga"), ("CDR", "Cordeiro")]:
    _alertas_sae = []
    _alertas_socio = []
    _alertas_tech = []
    _totais_u = {'vendido_sae': 0, 'enviado_sae': 0, 'estoque_sae': 0, 'vendido_socio': 0, 'vendido_tech': 0}

    # --- SAE ---
    for _cod, (_serie, _seg, _, _) in PEDIDO_SAE.items():
        _aj = AJUSTE_ANO_PASSADO.get(_cod, {}).get(_u_cod, 0)
        _env = ESTOQUE_ENVIADO.get(_cod, {}).get(_u_cod, 0)
        _f = (df_vendas_sae['segmento'] == _seg) & (df_vendas_sae['serie'] == _serie) & (df_vendas_sae['unidade'] == _u_cod)
        _vend = int(df_vendas_sae.loc[_f, 'vendido'].values[0]) if _f.any() else 0
        _est = _env - _vend + _aj
        _liq = _vend - _aj
        _totais_u['vendido_sae'] += _vend
        _totais_u['enviado_sae'] += _env
        _totais_u['estoque_sae'] += _est

        if _est < 0:
            _alertas_sae.append(f"**{_serie}**: estoque NEGATIVO ({_est}). Vendido={_vend}, Enviado={_env}")
        elif _est <= 3 and _env > 0:
            _alertas_sae.append(f"**{_serie}**: estoque critico ({_est} restantes)")
        if _aj > 0 and _aj >= _vend:
            _alertas_sae.append(f"**{_serie}**: Ajuste 2025 ({_aj}) >= Vendido ({_vend}) — venda liquida zero ou negativa, checar contratos")

    # --- Socioemocional ---
    for _seg_s, _serie_s in ORDEM_SERIES_SOCIO:
        _f_s = (df_vendas_socio['segmento'] == _seg_s) & (df_vendas_socio['serie'] == _serie_s) & (df_vendas_socio['unidade'] == _u_cod)
        _vend_s = int(df_vendas_socio.loc[_f_s, 'vendido'].sum()) if _f_s.any() else 0
        _aj_s = AJUSTE_SOCIO_2025.get(_serie_s, {}).get(_u_cod, 0)
        _liq_s = _vend_s - _aj_s
        _totais_u['vendido_socio'] += _vend_s
        if _aj_s > 0 and _aj_s >= _vend_s:
            _alertas_socio.append(f"**{_serie_s}**: Ajuste ({_aj_s}) >= Vendido ({_vend_s}) — conferir contratos novos")
        if _vend_s == 0:
            _alertas_socio.append(f"**{_serie_s}**: zero vendas registradas no SIGA")

    # --- Elo Tech ---
    if _u_cod != 'CDR':
        for _seg_e, _serie_e in ORDEM_SERIES_ELOTECH:
            _f_e = (df_vendas_elotech['segmento'] == _seg_e) & (df_vendas_elotech['serie'] == _serie_e) & (df_vendas_elotech['unidade'] == _u_cod)
            _vend_e = int(df_vendas_elotech.loc[_f_e, 'vendido'].sum()) if _f_e.any() else 0
            _totais_u['vendido_tech'] += _vend_e
        # Alunos sem serie identificada
        _tech_detail = df_vendas_raw[(df_vendas_raw['tipo'] == 'Elo Tech') & (df_vendas_raw['unidade'] == _u_cod)]
        _tech_unicos = _tech_detail.drop_duplicates(subset=['matricula'])
        _sem_serie = _tech_unicos[~_tech_unicos['matricula'].isin(
            [k[0] for k in ELOTECH_SERIE_PDF if k[1] == _u_cod]
        )]
        _sem_serie_sae = _sem_serie[~_sem_serie['matricula'].isin(
            df_vendas_raw[df_vendas_raw['tipo'].isin(['SAE', 'Socioemocional'])]['matricula'].unique()
        )]
        if len(_sem_serie_sae) > 0:
            _alertas_tech.append(f"{len(_sem_serie_sae)} aluno(s) Elo Tech sem serie identificada (sem SAE/Socio e fora do PDF)")

    _total_alertas = len(_alertas_sae) + len(_alertas_socio) + len(_alertas_tech)

    if _total_alertas == 0:
        _cor_u = "#059669"
        _badge = "OK"
    elif _total_alertas <= 3:
        _cor_u = "#d97706"
        _badge = f"{_total_alertas} pendencia(s)"
    else:
        _cor_u = "#dc2626"
        _badge = f"{_total_alertas} pendencia(s)"

    with st.expander(f"{_u_nome} — {_badge}  |  SAE: {_totais_u['vendido_sae']} vendidos, Socio: {_totais_u['vendido_socio']}, Tech: {_totais_u['vendido_tech']}", expanded=(_total_alertas > 0)):
        # Resumo numerico
        _col_r1, _col_r2, _col_r3, _col_r4 = st.columns(4)
        with _col_r1:
            st.metric("Vendido SAE", _totais_u['vendido_sae'])
        with _col_r2:
            st.metric("Enviado SAE", _totais_u['enviado_sae'])
        with _col_r3:
            st.metric("Estoque SAE", _totais_u['estoque_sae'],
                       delta=f"{_totais_u['estoque_sae']:+d}" if _totais_u['estoque_sae'] != 0 else None)
        with _col_r4:
            st.metric("Socio + Tech", f"{_totais_u['vendido_socio']}+{_totais_u['vendido_tech']}")

        if _alertas_sae:
            st.markdown("**SAE:**")
            for _a in _alertas_sae:
                st.markdown(f"- {_a}")

        if _alertas_socio:
            st.markdown("**Socioemocional:**")
            for _a in _alertas_socio:
                st.markdown(f"- {_a}")

        if _alertas_tech:
            st.markdown("**Elo Tech:**")
            for _a in _alertas_tech:
                st.markdown(f"- {_a}")

        if _total_alertas > 0:
            # Gerar resumo para Karol
            _itens = []
            _neg_sae = [a for a in _alertas_sae if 'NEGATIVO' in a]
            _crit_sae = [a for a in _alertas_sae if 'critico' in a]
            _aj_sae = [a for a in _alertas_sae if 'Ajuste' in a]
            if _neg_sae:
                _itens.append(f"cobrar entrega de {len(_neg_sae)} serie(s) com estoque negativo")
            if _crit_sae:
                _itens.append(f"verificar necessidade de reposicao em {len(_crit_sae)} serie(s) com estoque critico")
            if _aj_sae:
                _itens.append(f"conferir {len(_aj_sae)} serie(s) onde ajuste 2025 anula as vendas — ha contratos assinados em 2026?")
            if _alertas_socio:
                _n_socio_prob = len([a for a in _alertas_socio if 'zero' not in a])
                _n_socio_zero = len([a for a in _alertas_socio if 'zero' in a])
                if _n_socio_prob:
                    _itens.append(f"Socioemocional: {_n_socio_prob} serie(s) com ajuste >= vendido")
                if _n_socio_zero:
                    _itens.append(f"Socioemocional: {_n_socio_zero} serie(s) sem vendas — alunos receberam livro sem contrato?")
            if _alertas_tech:
                _itens.append(f"Elo Tech: identificar serie dos alunos listados")

            st.markdown(f"---\n**Karol, falar com {_u_nome}:**")
            for _i, _item in enumerate(_itens, 1):
                st.markdown(f"{_i}. {_item}")
        else:
            st.success(f"{_u_nome}: nenhuma pendencia. Numeros de vendido e estoque consistentes.")

st.divider()


# ===========================================
# CONFERENCIA DE DADOS
# ===========================================

# Indicador de frescor dos dados
_json_check = OUTPUT_DIR / "recebimento_final.json"
if _json_check.exists():
    _mtime = datetime.fromtimestamp(_json_check.stat().st_mtime)
    _horas_atraso = (datetime.now() - _mtime).total_seconds() / 3600

    if _horas_atraso < 6:
        _cor_fresh = "#059669"
        _icon_fresh = "&#9989;"
        _msg_fresh = f"Dados atualizados ha {_horas_atraso:.0f}h"
    elif _horas_atraso < 48:
        _cor_fresh = "#d97706"
        _icon_fresh = "&#9888;&#65039;"
        _msg_fresh = f"Dados com {_horas_atraso:.0f}h de atraso"
    else:
        _dias_atraso = _horas_atraso / 24
        _cor_fresh = "#dc2626"
        _icon_fresh = "&#128680;"
        _msg_fresh = f"DADOS DESATUALIZADOS! {_dias_atraso:.0f} dias sem atualizar — clique em Atualizar Dados"

    st.markdown(f"""
    <div style='background: {_cor_fresh}15; border-left: 4px solid {_cor_fresh}; padding: 10px 16px; border-radius: 4px; margin: 12px 0;'>
        <span style='font-size: 1.1rem;'>{_icon_fresh}</span>
        <span style='color: {_cor_fresh}; font-weight: 700; margin-left: 8px;'>{_msg_fresh}</span>
        <span style='color: #64748b; margin-left: 16px;'>Ultima extracao: {ultima_atualizacao}</span>
    </div>
    """, unsafe_allow_html=True)

with st.expander("Conferencia de Dados — Auditoria de Vendas", expanded=False):
    st.markdown("**Valide os numeros antes de cobrancas ou relatorios.**")

    # 1. Resumo por tipo
    st.markdown("##### Registros Brutos vs Alunos Unicos")
    _resumo = []
    for _tipo in ['SAE', 'Socioemocional', 'Elo Tech']:
        _df_t = df_vendas_raw[df_vendas_raw['tipo'] == _tipo]
        _regs = len(_df_t)
        _uni = _df_t['matricula'].nunique()
        _resumo.append({
            'Tipo': _tipo,
            'Registros SIGA': _regs,
            'Alunos Unicos': _uni,
            'Parc./Aluno': round(_regs / _uni, 1) if _uni else 0,
            'BV': _df_t[_df_t['unidade'] == 'BV']['matricula'].nunique(),
            'CD': _df_t[_df_t['unidade'] == 'CD']['matricula'].nunique(),
            'JG': _df_t[_df_t['unidade'] == 'JG']['matricula'].nunique(),
            'CDR': _df_t[_df_t['unidade'] == 'CDR']['matricula'].nunique(),
        })
    _resumo.append({
        'Tipo': 'TOTAL',
        'Registros SIGA': sum(r['Registros SIGA'] for r in _resumo),
        'Alunos Unicos': sum(r['Alunos Unicos'] for r in _resumo),
        'Parc./Aluno': '',
        'BV': sum(r['BV'] for r in _resumo),
        'CD': sum(r['CD'] for r in _resumo),
        'JG': sum(r['JG'] for r in _resumo),
        'CDR': sum(r['CDR'] for r in _resumo),
    })
    st.dataframe(pd.DataFrame(_resumo), use_container_width=True, hide_index=True)
    st.caption("**Registros SIGA** = linhas extraidas (parcelas). **Alunos Unicos** = matriculas distintas (numero usado como 'vendido').")

    # 2. Detalhamento SAE por serie x unidade
    st.markdown("##### SAE: Contagem por Serie x Unidade")
    _det = []
    for cod, (serie, segmento, _, _) in PEDIDO_SAE.items():
        _row = {'Cod': cod, 'Serie': f"{segmento} - {serie}"}
        _df_cod = df_vendas_raw[
            (df_vendas_raw['tipo'] == 'SAE') &
            (df_vendas_raw['segmento'] == segmento) &
            (df_vendas_raw['serie'] == serie)
        ]
        for u in ['BV', 'CD', 'JG', 'CDR']:
            _sub = _df_cod[_df_cod['unidade'] == u]
            _row[u] = _sub['matricula'].nunique()
        _row['Total'] = sum(_row[u] for u in ['BV', 'CD', 'JG', 'CDR'])
        _det.append(_row)
    _df_det_audit = pd.DataFrame(_det)
    _totais_det = {'Cod': '', 'Serie': 'TOTAL'}
    for c in ['BV', 'CD', 'JG', 'CDR', 'Total']:
        _totais_det[c] = int(_df_det_audit[c].sum())
    _df_det_audit = pd.concat([_df_det_audit, pd.DataFrame([_totais_det])], ignore_index=True)
    st.dataframe(_df_det_audit, use_container_width=True, hide_index=True)

    # 3. Confronto Vendido vs Ajuste — divergencias
    st.markdown("##### Divergencias: Vendido vs Ajuste 2025 vs Estoque")
    _problemas = []
    for cod, (serie, segmento, _, _) in PEDIDO_SAE.items():
        for u_cod in ['BV', 'CD', 'JG', 'CDR']:
            _aj = AJUSTE_ANO_PASSADO.get(cod, {}).get(u_cod, 0)
            _f = (
                (df_vendas_sae['segmento'] == segmento) &
                (df_vendas_sae['serie'] == serie) &
                (df_vendas_sae['unidade'] == u_cod)
            )
            _vend = int(df_vendas_sae.loc[_f, 'vendido'].values[0]) if _f.any() else 0
            _liq = _vend - _aj
            _env = ESTOQUE_ENVIADO.get(cod, {}).get(u_cod, 0)
            _est = _env - _vend + _aj
            if _liq < 0 or _est < 0 or (_aj > 0 and _aj >= _vend):
                _alerta = []
                if _aj >= _vend and _aj > 0:
                    _alerta.append('Ajuste >= Vendido')
                if _est < 0:
                    _alerta.append('Estoque negativo')
                if _liq < 0:
                    _alerta.append('Venda liq. negativa')
                _problemas.append({
                    'Serie': f"{segmento} - {serie}",
                    'Unidade': u_cod,
                    'Enviado': _env,
                    'Vendido': _vend,
                    'Ajuste 2025': _aj,
                    'Venda Liq.': _liq,
                    'Estoque': _est,
                    'Alerta': ' | '.join(_alerta),
                })
    if _problemas:
        st.warning(f"{len(_problemas)} divergencia(s) encontrada(s):")
        st.dataframe(pd.DataFrame(_problemas), use_container_width=True, hide_index=True)
    else:
        st.success("Nenhuma divergencia encontrada nos dados SAE.")

    # 4. Qualidade dos dados
    st.markdown("##### Qualidade dos Dados")
    _sem_mat = len(df_vendas_raw[df_vendas_raw['matricula'].str.strip() == ''])
    _sem_dt = len(df_vendas_raw[df_vendas_raw['dt_baixa'].str.strip() == ''])
    _injetados = len(df_vendas_raw[df_vendas_raw['parcela'] == 'TAXA'])

    col_q1, col_q2, col_q3 = st.columns(3)
    with col_q1:
        if _sem_mat > 0:
            st.error(f"{_sem_mat} registros sem matricula")
        else:
            st.success("0 sem matricula")
    with col_q2:
        if _sem_dt > 0:
            st.warning(f"{_sem_dt} sem data de baixa")
        else:
            st.success("0 sem data")
    with col_q3:
        st.info(f"{_injetados} injetados (Em Aberto Elo Tech)")

    # 5. Cruzamento entre tipos
    st.markdown("##### Cruzamento entre Tipos")
    _m_sae = set(df_vendas_raw[df_vendas_raw['tipo'] == 'SAE']['matricula'].unique())
    _m_socio = set(df_vendas_raw[df_vendas_raw['tipo'] == 'Socioemocional']['matricula'].unique())
    _m_tech = set(df_vendas_raw[df_vendas_raw['tipo'] == 'Elo Tech']['matricula'].unique())
    _tech_sem_ref = _m_tech - _m_sae - _m_socio

    col_x1, col_x2, col_x3 = st.columns(3)
    with col_x1:
        st.metric("SAE + Socio", len(_m_sae & _m_socio), help="Alunos com ambos (esperado)")
    with col_x2:
        st.metric("SAE + Elo Tech", len(_m_sae & _m_tech), help="Elo Tech com SAE (serie identificada)")
    with col_x3:
        _n_sem = len(_tech_sem_ref)
        st.metric("Elo Tech SEM SAE", _n_sem, help="Serie sera do PDF manual ou 'Sem serie'")
        if _n_sem > 0:
            st.caption(f"{_n_sem} alunos Elo Tech sem SAE/Socio — serie vem do mapa PDF")

st.divider()


# ===========================================
# COMPUTACAO DE DADOS POR SECAO (antes do render)
# ===========================================

# --- Tabela Socioemocional ---
if filtro_unidade and filtro_unidade != "Todas":
    unidades_socio = [filtro_unidade]
else:
    unidades_socio = list(UNIDADES.values())

registros_socio = []
for segmento, serie in ORDEM_SERIES_SOCIO:
    registro = {'Descrição': f"{segmento} - {serie}"}
    for unidade_nome in unidades_socio:
        unidade_cod = UNIDADES_INV.get(unidade_nome, "")
        filtro = (
            (df_vendas_socio['segmento'] == segmento) &
            (df_vendas_socio['serie'] == serie) &
            (df_vendas_socio['unidade'] == unidade_cod)
        )
        registro[unidade_nome] = int(df_vendas_socio.loc[filtro, 'vendido'].sum()) if filtro.any() else 0
    registros_socio.append(registro)

df_socio = pd.DataFrame(registros_socio)
colunas_socio = [c for c in df_socio.columns if c != 'Descrição']
df_socio['TOTAL'] = df_socio[colunas_socio].sum(axis=1)

totais_socio = {'Descrição': 'TOTAL'}
for col in colunas_socio:
    totais_socio[col] = df_socio[col].sum()
totais_socio['TOTAL'] = df_socio['TOTAL'].sum()
df_socio = pd.concat([df_socio, pd.DataFrame([totais_socio])], ignore_index=True)

total_socio_vendido = int(df_socio.loc[df_socio['Descrição'] == 'TOTAL', 'TOTAL'].values[0])

# --- Tabela Elo Tech ---
if filtro_unidade and filtro_unidade != "Todas":
    unidades_robo = [filtro_unidade]
else:
    unidades_robo = list(UNIDADES.values())

registros_robo = []
for segmento, serie in ORDEM_SERIES_ELOTECH:
    registro = {'Descrição': f"{segmento} - {serie}"}
    for unidade_nome in unidades_robo:
        unidade_cod = UNIDADES_INV.get(unidade_nome, "")
        filtro = (
            (df_vendas_elotech['segmento'] == segmento) &
            (df_vendas_elotech['serie'] == serie) &
            (df_vendas_elotech['unidade'] == unidade_cod)
        )
        registro[unidade_nome] = int(df_vendas_elotech.loc[filtro, 'vendido'].sum()) if filtro.any() else 0
    registros_robo.append(registro)

# Linha "Outros" (Sem serie + series fora do 2o-5o Ano)
series_elotech = {s for _, s in ORDEM_SERIES_ELOTECH}
reg_outros = {'Descrição': 'Outros'}
for unidade_nome in unidades_robo:
    unidade_cod = UNIDADES_INV.get(unidade_nome, "")
    filtro = (~df_vendas_elotech['serie'].isin(series_elotech)) & (df_vendas_elotech['unidade'] == unidade_cod)
    reg_outros[unidade_nome] = int(df_vendas_elotech.loc[filtro, 'vendido'].sum()) if filtro.any() else 0
if any(v for k, v in reg_outros.items() if k != 'Descrição'):
    registros_robo.append(reg_outros)

df_robo = pd.DataFrame(registros_robo)
colunas_robo = [c for c in df_robo.columns if c != 'Descrição']
df_robo['TOTAL'] = df_robo[colunas_robo].sum(axis=1)

totais_robo = {'Descrição': 'TOTAL'}
for col in colunas_robo:
    totais_robo[col] = df_robo[col].sum()
totais_robo['TOTAL'] = df_robo['TOTAL'].sum()
df_robo = pd.concat([df_robo, pd.DataFrame([totais_robo])], ignore_index=True)

total_robo_vendido = int(totais_robo['TOTAL'])

# --- Balanco Fisico ---
registros_balanco = []
for cod, (serie, segmento, _, _) in PEDIDO_SAE.items():
    balanco = BALANCO_FISICO.get(cod, {})
    estoque_env = ESTOQUE_ENVIADO.get(cod, {})

    for unidade_cod, unidade_nome in UNIDADES.items():
        fisico_datas = balanco.get(unidade_cod, {})
        fisico = list(fisico_datas.values())[-1] if fisico_datas else None

        enviado = estoque_env.get(unidade_cod, 0)

        # Vendido SAE
        vendido = 0
        filtro = (
            (df_vendas_sae['segmento'] == segmento) &
            (df_vendas_sae['serie'] == serie) &
            (df_vendas_sae['unidade'] == unidade_cod)
        )
        if filtro.any():
            vendido = int(df_vendas_sae.loc[filtro, 'vendido'].values[0])

        teorico = enviado - vendido
        diferenca = (fisico - teorico) if fisico is not None else None

        registros_balanco.append({
            'Descrição': f"{segmento} - {serie}",
            'Unidade': unidade_nome,
            'Enviado': enviado,
            'Vendido': vendido,
            'Teórico': teorico,
            'Físico': fisico if fisico is not None else "-",
            'Diferença': diferenca if diferenca is not None else "-",
        })

df_balanco = pd.DataFrame(registros_balanco)
if filtro_unidade != "Todas":
    df_balanco = df_balanco[df_balanco['Unidade'] == filtro_unidade]
df_balanco_valido = df_balanco[df_balanco['Físico'] != "-"].copy()

# --- Alertas ---
df_negativo = df_estoque[df_estoque['estoque'] < 0][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']]
df_negativo = df_negativo.rename(columns={
    'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade',
    'enviado': 'Enviado', 'vendido': 'Vendido', 'estoque': 'Falta'
})
df_baixo = df_estoque[(df_estoque['estoque'] > 0) & (df_estoque['estoque'] <= 5)][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']]
df_baixo = df_baixo.rename(columns={
    'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade',
    'enviado': 'Enviado', 'vendido': 'Vendido', 'estoque': 'Restante'
})

# --- Elo Tech detail (Outros e Em Aberto) ---
_series_ok = {s for _, s in ORDEM_SERIES_ELOTECH}
df_outros_detail = df_vendas_elotech_detail[~df_vendas_elotech_detail['serie_real'].isin(_series_ok)]\
    .drop_duplicates(subset=['matricula', 'unidade'])\
    .sort_values(['unidade', 'serie_real', 'nome'])

# --- Dados de ata (preparados antes do render) ---
df_ata_sae = df_vendas_raw[df_vendas_raw['tipo'] == 'SAE']\
    .drop_duplicates(subset=['matricula', 'unidade'])\
    .sort_values(['unidade', 'serie', 'nome'])
df_ata_sae = _filtrar_unidade_ata(df_ata_sae, filtro_unidade)

df_ata_socio = df_vendas_raw[df_vendas_raw['tipo'] == 'Socioemocional']\
    .drop_duplicates(subset=['matricula', 'unidade'])\
    .sort_values(['unidade', 'serie', 'nome'])
df_ata_socio = _filtrar_unidade_ata(df_ata_socio, filtro_unidade)

df_ata_tech = df_vendas_elotech_detail.drop_duplicates(subset=['matricula', 'unidade'])\
    .sort_values(['unidade', 'serie_real', 'nome'])
df_ata_tech = _filtrar_unidade_ata(df_ata_tech, filtro_unidade)

# --- Excel por secao ---
xlsx_estoque = _to_excel_bytes(df_estoque_export, "Estoque SAE")
xlsx_socio = _to_excel_bytes(df_socio, "Socioemocional")
xlsx_robo = _to_excel_bytes(df_robo, "Elo Tech")
xlsx_tudo = _to_excel_multi_sheets({
    "Estoque SAE": df_estoque_export,
    "Vendas Socioemocional": df_socio,
    "Vendas Elo Tech": df_robo,
    "Estoque Detalhado": df_estoque_completo,
})

# --- Grafico dados ---
df_seg = df_estoque.groupby('segmento').agg({
    'enviado': 'sum', 'vendido': 'sum', 'estoque': 'sum'
}).reset_index()
df_seg['segmento_ordem'] = df_seg['segmento'].map({s: i for i, s in enumerate(ORDEM_SEGMENTOS)})
df_seg = df_seg.sort_values('segmento_ordem')

# Data string para impressao
data_str_impressao = datetime.now().strftime("%d/%m/%Y %H:%M")


# ###########################################################
# SECAO SAE
# ###########################################################

st.markdown("---")
st.markdown("### 📦 Secao SAE - Estoque por Serie e Unidade")

col_config1, col_config2 = st.columns([1, 3])
with col_config1:
    altura_tabela = st.selectbox(
        "Linhas visiveis",
        options=[10, 17, "Todas (17 + Total)"],
        index=1,
        help="Quantas linhas mostrar por vez"
    )

if altura_tabela == "Todas (17 + Total)":
    altura_px = None
else:
    altura_px = 35 * (int(altura_tabela) + 1) + 40

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Visao Completa", "Estoque Restante", "Enviado", "Vendido", "Visao Geral"])

with tab1:
    if filtro_unidade == "Todas":
        st.markdown("**Pedido / Enviado / Vendido / Ajuste 2025 / Estoque por Unidade**")
    else:
        st.markdown(f"**Pedido / Enviado / Vendido / Ajuste 2025 / Estoque — {filtro_unidade}**")
    registros_detail = []
    for segmento, serie in ORDEM_SERIES_COMPLETA:
        filtro = (df_estoque['segmento'] == segmento) & (df_estoque['serie'] == serie)
        if filtro.any():
            pedido = df_estoque.loc[filtro, 'pedido_total'].iloc[0]
            enviado = df_estoque.loc[filtro, 'enviado'].sum()
            vendido = df_estoque.loc[filtro, 'vendido'].sum()
            ajuste = df_estoque.loc[filtro, 'ajuste'].sum()
        else:
            filtro_c = (df_estoque_completo['segmento'] == segmento) & (df_estoque_completo['serie'] == serie)
            pedido = df_estoque_completo.loc[filtro_c, 'pedido_total'].iloc[0] if filtro_c.any() else 0
            enviado = vendido = ajuste = 0

        reg = {
            'Descrição': f"{segmento} - {serie}",
            'Pedido': int(pedido), 'Enviado': int(enviado),
            'Vendido': int(vendido), 'Ajuste 2025': int(ajuste),
        }
        if filtro_unidade == "Todas":
            for u_nome in UNIDADES.values():
                f_u = filtro & (df_estoque['unidade'] == u_nome)
                reg[u_nome] = int(df_estoque.loc[f_u, 'estoque'].sum()) if f_u.any() else 0
            reg['Total'] = sum(reg[u] for u in UNIDADES.values())
        else:
            f_est = filtro
            reg['Estoque'] = int(df_estoque.loc[f_est, 'estoque'].sum()) if f_est.any() else 0
        registros_detail.append(reg)

    df_detail = pd.DataFrame(registros_detail)
    totais = {'Descrição': 'TOTAL'}
    for c in df_detail.columns:
        if c != 'Descrição':
            totais[c] = df_detail[c].sum()
    df_detail = pd.concat([df_detail, pd.DataFrame([totais])], ignore_index=True)
    _det_cols_est = list(UNIDADES.values()) + ['Total'] if filtro_unidade == "Todas" else ['Estoque']
    styled_detail = df_detail.style.map(colorir_estoque, subset=_det_cols_est)
    st.dataframe(styled_detail, use_container_width=True, hide_index=True, height=altura_px)

with tab2:
    st.markdown("**Estoque Restante = Enviado - Vendido**")
    df_display = criar_tabela_completa(df_estoque, 'estoque', filtro_unidade)
    styled = df_display.style.map(
        colorir_estoque,
        subset=[c for c in df_display.columns if c != 'Descrição']
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=altura_px)
    st.markdown("""
    <div style='font-size: 0.85rem; color: #64748b; margin-top: 1rem;'>
        <strong style='color:#dc2626'>Vermelho</strong>: Estoque negativo (falta) &nbsp;&nbsp;
        <strong style='color:#d97706'>Amarelo</strong>: Estoque baixo (5) &nbsp;&nbsp;
        <strong style='color:#059669'>Verde</strong>: Estoque OK (>5)
    </div>
    """, unsafe_allow_html=True)

with tab3:
    st.markdown("**Quantidade Enviada para cada Unidade**")
    df_display = criar_tabela_completa(df_estoque, 'enviado', filtro_unidade)
    st.dataframe(df_display, use_container_width=True, hide_index=True, height=altura_px)

with tab4:
    st.markdown("**Quantidade Vendida (Alunos Unicos SAE)**")
    df_display = criar_tabela_completa(df_estoque, 'vendido', filtro_unidade)
    st.dataframe(df_display, use_container_width=True, hide_index=True, height=altura_px)

with tab5:
    st.markdown("**SAE: Pedido x Venda Real x Saldo — por Unidade**")
    registros_vg = []
    for segmento, serie in ORDEM_SERIES_COMPLETA:
        filtro_c = (df_estoque_completo['segmento'] == segmento) & (df_estoque_completo['serie'] == serie)
        if not filtro_c.any():
            continue
        pedido = int(df_estoque_completo.loc[filtro_c, 'pedido_total'].iloc[0])
        vend_uni = {}
        aj_uni = {}
        for u_cod in ["BV", "CD", "JG", "CDR"]:
            u_nome = UNIDADES[u_cod]
            f_u = filtro_c & (df_estoque_completo['unidade'] == u_nome)
            vend_uni[u_cod] = int(df_estoque_completo.loc[f_u, 'vendido'].sum()) if f_u.any() else 0
            aj_uni[u_cod] = int(df_estoque_completo.loc[f_u, 'ajuste'].sum()) if f_u.any() else 0
        vendido_total = sum(vend_uni.values())
        ajuste_total = sum(aj_uni.values())
        vend_real = vendido_total - ajuste_total
        saldo_estoque = pedido - vend_real
        registros_vg.append({
            'Descricao': f"{segmento} - {serie}",
            'Pedido': pedido,
            'Vendido': vendido_total,
            'Ajuste': ajuste_total,
            'Vend. Real': vend_real,
            'Saldo Estoque': saldo_estoque,
            'BV': vend_uni['BV'] - aj_uni['BV'],
            'CD': vend_uni['CD'] - aj_uni['CD'],
            'JG': vend_uni['JG'] - aj_uni['JG'],
            'CDR': vend_uni['CDR'] - aj_uni['CDR'],
        })
    df_vg = pd.DataFrame(registros_vg)
    totais_vg = {'Descricao': 'TOTAL'}
    for c in df_vg.columns:
        if c != 'Descricao':
            totais_vg[c] = int(df_vg[c].sum())
    df_vg = pd.concat([df_vg, pd.DataFrame([totais_vg])], ignore_index=True)
    styled_vg = df_vg.style.map(colorir_estoque, subset=['Saldo Estoque'])
    st.dataframe(styled_vg, use_container_width=True, hide_index=True, height=altura_px)
    st.markdown("""
    <div style='font-size: 0.85rem; color: #64748b; margin-top: 0.5rem;'>
        <b>Vend. Real</b> = Vendido - Ajuste 2025 &nbsp;|&nbsp;
        <b>Saldo Estoque</b> = Pedido - Vend. Real &nbsp;|&nbsp;
        <b>BV/CD/JG/CDR</b> = Venda liquida por unidade (Vendido - Ajuste)
    </div>
    """, unsafe_allow_html=True)


# --- Conferencia de Contratos por Unidade ---
st.markdown("#### Conferencia de Contratos por Unidade")
st.caption("Imprima esta secao para conferir contratos assinados e assinaturas no caderno por unidade.")

for u_cod_conf, u_nome_conf in [("BV", "Boa Viagem"), ("CD", "Candeias"), ("JG", "Janga"), ("CDR", "Cordeiro")]:
    with st.expander(f"📋 {u_nome_conf}", expanded=False):
        regs_conf = []
        for segmento, serie in ORDEM_SERIES_COMPLETA:
            filtro_c = (df_estoque_completo['segmento'] == segmento) & (df_estoque_completo['serie'] == serie)
            if not filtro_c.any():
                continue
            pedido = int(df_estoque_completo.loc[filtro_c, 'pedido_total'].iloc[0])
            f_u = filtro_c & (df_estoque_completo['unidade'] == u_nome_conf)
            vendido = int(df_estoque_completo.loc[f_u, 'vendido'].sum()) if f_u.any() else 0
            ajuste = int(df_estoque_completo.loc[f_u, 'ajuste'].sum()) if f_u.any() else 0
            venda_liq = vendido - ajuste
            regs_conf.append({
                'Descricao': f"{segmento} - {serie}",
                'Pedido Geral': pedido,
                'Vendido': vendido,
                'Ajuste': ajuste,
                'Venda Liquida': venda_liq,
                'Contratos': venda_liq,
                'Assinaturas Caderno': '',
            })
        df_conf = pd.DataFrame(regs_conf)
        totais_conf = {
            'Descricao': 'TOTAL',
            'Pedido Geral': int(df_conf['Pedido Geral'].sum()),
            'Vendido': int(df_conf['Vendido'].sum()),
            'Ajuste': int(df_conf['Ajuste'].sum()),
            'Venda Liquida': int(df_conf['Venda Liquida'].sum()),
            'Contratos': int(df_conf['Contratos'].sum()),
            'Assinaturas Caderno': '',
        }
        df_conf = pd.concat([df_conf, pd.DataFrame([totais_conf])], ignore_index=True)
        st.dataframe(df_conf, use_container_width=True, hide_index=True)
        st.markdown("""
        <div style='font-size: 0.85rem; color: #64748b;'>
            <b>Contratos</b> = quantidade esperada de contratos assinados (= Venda Liquida) &nbsp;|&nbsp;
            <b>Assinaturas Caderno</b> = preencher manualmente apos contagem
        </div>
        """, unsafe_allow_html=True)

# --- Alertas (dentro da secao SAE) ---
st.markdown("#### Alertas de Estoque")

col_alert1, col_alert2 = st.columns(2)

with col_alert1:
    st.markdown("**Estoque Negativo (Falta)**")
    if not df_negativo.empty:
        st.dataframe(df_negativo, use_container_width=True, hide_index=True)
    else:
        st.success("Nenhum estoque negativo!")

with col_alert2:
    st.markdown("**Estoque Baixo (5 unidades)**")
    if not df_baixo.empty:
        st.dataframe(df_baixo, use_container_width=True, hide_index=True)
    else:
        st.success("Nenhum estoque baixo!")

# --- Barra de acoes SAE ---
st.markdown("---")
col_dl_sae, col_ata_sae, col_print_sae = st.columns(3)

with col_dl_sae:
    st.download_button(
        label="Baixar Estoque SAE (Excel)",
        data=xlsx_estoque,
        file_name=f"estoque_sae_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_estoque_xlsx",
        use_container_width=True,
    )

with col_ata_sae:
    if st.button(f"Ata SAE ({len(df_ata_sae)})", use_container_width=True, key="btn_ata_sae"):
        st.session_state['ata_tipo'] = 'sae'

with col_print_sae:
    if st.button("Imprimir SAE", use_container_width=True, key="btn_print_sae"):
        st.session_state['print_secao'] = 'sae'

# Renderizar ata SAE se selecionada
if st.session_state.get('ata_tipo') == 'sae':
    components.html(gerar_ata_html(df_ata_sae, "Livros SAE", 'serie'), height=800, scrolling=True)

# Renderizar impressao SAE se selecionada
if st.session_state.get('print_secao') == 'sae':
    # Preparar HTML de impressao SAE
    df_est_print = criar_tabela_completa(df_estoque_completo, 'estoque', "Todas")
    df_est_print.columns = [c.replace('ã', 'a').replace('í', 'i').replace('ó', 'o').replace('ç', 'c') for c in df_est_print.columns]
    tbl_sae = df_to_html_colorido(df_est_print, "Estoque Restante SAE (Enviado - Vendido)")
    kpis_sae = [
        ("Pedido Total", f"{pedido_total:,}"),
        ("Enviado Total", f"{enviado_total:,}"),
        ("Vendido SAE", f"{vendido_total:,}"),
        ("Estoque Restante", f"{estoque_total:,}"),
        ("% Vendido", f"{percentual_venda:.1f}%"),
    ]
    html_sae = gerar_html_impressao_secao("Estoque SAE", kpis_sae, [tbl_sae], data_str_impressao, filtro_unidade)
    _abrir_impressao(html_sae)
    st.session_state['print_secao'] = None


# ###########################################################
# SECAO SOCIOEMOCIONAL
# ###########################################################

st.markdown("---")
st.markdown("### 💜 Secao Socioemocional")

PEDIDO_SOCIO = {
    'Infantil IV': 100, 'Infantil V': 125,
    '1º Ano': 147, '2º Ano': 140,
    '3º Ano': 135, '4º Ano': 130,
    '5º Ano': 155, '6º Ano': 160,
}

regs_socio_vg = []
for segmento, serie in ORDEM_SERIES_SOCIO:
    pedido = PEDIDO_SOCIO.get(serie, 0)
    aj_uni = AJUSTE_SOCIO_2025.get(serie, {})
    vend_uni = {}
    for u_cod in ["BV", "CD", "JG", "CDR"]:
        filtro = (
            (df_vendas_socio['segmento'] == segmento) &
            (df_vendas_socio['serie'] == serie) &
            (df_vendas_socio['unidade'] == u_cod)
        )
        vend_uni[u_cod] = int(df_vendas_socio.loc[filtro, 'vendido'].sum()) if filtro.any() else 0
    vendido_total = sum(vend_uni.values())
    ajuste_total = sum(aj_uni.values())
    vend_real = vendido_total - ajuste_total
    saldo = pedido - vend_real
    liq_uni = {}
    for u_cod in ["BV", "CD", "JG", "CDR"]:
        liq_uni[u_cod] = vend_uni[u_cod] - aj_uni.get(u_cod, 0)
    regs_socio_vg.append({
        'Descricao': f"{segmento} - {serie}",
        'Pedido': pedido,
        'Vendido': vendido_total,
        'Ajuste': ajuste_total,
        'Vend. Real': vend_real,
        'Saldo Estoque': saldo,
        'BV': liq_uni['BV'],
        'CD': liq_uni['CD'],
        'JG': liq_uni['JG'],
        'CDR': liq_uni['CDR'],
        '_vend_uni': vend_uni,
        '_aj_uni': aj_uni,
    })

tot_s_pedido = sum(r['Pedido'] for r in regs_socio_vg)
tot_s_vendido = sum(r['Vendido'] for r in regs_socio_vg)
tot_s_ajuste = sum(r['Ajuste'] for r in regs_socio_vg)
tot_s_vend_real = tot_s_vendido - tot_s_ajuste
tot_s_saldo = tot_s_pedido - tot_s_vend_real

rows_socio_vg = [{k: v for k, v in r.items() if not k.startswith('_')} for r in regs_socio_vg]
rows_socio_vg.append({
    'Descricao': 'TOTAL',
    'Pedido': tot_s_pedido,
    'Vendido': tot_s_vendido,
    'Ajuste': tot_s_ajuste,
    'Vend. Real': tot_s_vend_real,
    'Saldo Estoque': tot_s_saldo,
    'BV': sum(r['BV'] for r in rows_socio_vg),
    'CD': sum(r['CD'] for r in rows_socio_vg),
    'JG': sum(r['JG'] for r in rows_socio_vg),
    'CDR': sum(r['CDR'] for r in rows_socio_vg),
})
df_socio_vg = pd.DataFrame(rows_socio_vg)

col_s1, col_s2, col_s3, col_s4 = st.columns(4)
with col_s1:
    st.metric("Pedido", f"{tot_s_pedido:,}")
with col_s2:
    st.metric("Vendido", f"{tot_s_vendido:,}")
with col_s3:
    st.metric("Ajuste 2025", f"{tot_s_ajuste:,}")
with col_s4:
    st.metric("Saldo Estoque", f"{tot_s_saldo:,}", delta=f"{tot_s_saldo:+,}")

styled_socio_vg = df_socio_vg.style.map(colorir_estoque, subset=['Saldo Estoque'])
st.dataframe(styled_socio_vg, use_container_width=True, hide_index=True)
st.markdown("""
<div style='font-size: 0.85rem; color: #64748b; margin-top: 0.5rem;'>
    <b>Vend. Real</b> = Vendido - Ajuste 2025 &nbsp;|&nbsp;
    <b>Saldo Estoque</b> = Pedido - Vend. Real &nbsp;|&nbsp;
    <b>BV/CD/JG/CDR</b> = Venda liquida por unidade (Vendido - Ajuste)
</div>
""", unsafe_allow_html=True)

# Conferencia de Contratos Socioemocional por Unidade
st.markdown("#### Conferencia de Contratos Socioemocional por Unidade")
for u_cod_sc, u_nome_sc in [("BV", "Boa Viagem"), ("CD", "Candeias"), ("JG", "Janga"), ("CDR", "Cordeiro")]:
    with st.expander(f"📋 {u_nome_sc}", expanded=False):
        regs_conf_sc = []
        for segmento, serie in ORDEM_SERIES_SOCIO:
            pedido = PEDIDO_SOCIO.get(serie, 0)
            aj = AJUSTE_SOCIO_2025.get(serie, {}).get(u_cod_sc, 0)
            filtro = (
                (df_vendas_socio['segmento'] == segmento) &
                (df_vendas_socio['serie'] == serie) &
                (df_vendas_socio['unidade'] == u_cod_sc)
            )
            vendido = int(df_vendas_socio.loc[filtro, 'vendido'].sum()) if filtro.any() else 0
            venda_liq = vendido - aj
            regs_conf_sc.append({
                'Descricao': f"{segmento} - {serie}",
                'Pedido Geral': pedido,
                'Vendido': vendido,
                'Ajuste': aj,
                'Venda Liquida': venda_liq,
                'Contratos': venda_liq,
                'Assinaturas Caderno': '',
            })
        df_conf_sc = pd.DataFrame(regs_conf_sc)
        totais_conf_sc = {
            'Descricao': 'TOTAL',
            'Pedido Geral': int(df_conf_sc['Pedido Geral'].sum()),
            'Vendido': int(df_conf_sc['Vendido'].sum()),
            'Ajuste': int(df_conf_sc['Ajuste'].sum()),
            'Venda Liquida': int(df_conf_sc['Venda Liquida'].sum()),
            'Contratos': int(df_conf_sc['Contratos'].sum()),
            'Assinaturas Caderno': '',
        }
        df_conf_sc = pd.concat([df_conf_sc, pd.DataFrame([totais_conf_sc])], ignore_index=True)
        st.dataframe(df_conf_sc, use_container_width=True, hide_index=True)

st.markdown("---")

# Vendido por unidade (tabela original)
st.markdown("**Vendido por Unidade (detalhe)**")
st.dataframe(df_socio, use_container_width=True, hide_index=True)

# --- Barra de acoes Socioemocional ---
st.markdown("---")
col_dl_socio, col_ata_socio, col_print_socio = st.columns(3)

with col_dl_socio:
    st.download_button(
        label="Baixar Vendas Socioemocional (Excel)",
        data=xlsx_socio,
        file_name=f"vendas_socioemocional_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_socio_xlsx",
        use_container_width=True,
    )

with col_ata_socio:
    if st.button(f"Ata Socioemocional ({len(df_ata_socio)})", use_container_width=True, key="btn_ata_socio"):
        st.session_state['ata_tipo'] = 'socio'

with col_print_socio:
    if st.button("Imprimir Socioemocional", use_container_width=True, key="btn_print_socio"):
        st.session_state['print_secao'] = 'socio'

# Renderizar ata Socio se selecionada
if st.session_state.get('ata_tipo') == 'socio':
    components.html(gerar_ata_html(df_ata_socio, "Livros Socioemocional", 'serie'), height=800, scrolling=True)

# Renderizar impressao Socio se selecionada
if st.session_state.get('print_secao') == 'socio':
    df_socio_print = df_socio.copy()
    df_socio_print.columns = [c.replace('ã', 'a').replace('í', 'i').replace('ó', 'o').replace('ç', 'c') for c in df_socio_print.columns]
    tbl_socio = df_to_html_colorido(df_socio_print, "Vendas Socioemocional (Alunos Unicos)")
    kpis_socio = [
        ("Total Vendido Socio.", f"{total_socio_vendido}"),
    ]
    html_socio = gerar_html_impressao_secao("Vendas Socioemocional", kpis_socio, [tbl_socio], data_str_impressao, filtro_unidade)
    _abrir_impressao(html_socio)
    st.session_state['print_secao'] = None


# ###########################################################
# SECAO ELO TECH
# ###########################################################

st.markdown("---")
st.markdown("### 🤖 Secao Elo Tech")
st.caption("Nota: Cordeiro (CDR) nao possui Elo Tech. Serie identificada via cruzamento com matricula SAE.")

PEDIDO_ELOTECH = {"2º Ano": 150, "3º Ano": 155, "4º Ano": 135, "5º Ano": 150}

# Visao Geral Elo Tech
regs_tech_vg = []
for segmento, serie in ORDEM_SERIES_ELOTECH:
    pedido = PEDIDO_ELOTECH.get(serie, 0)
    vend_uni = {}
    for u_cod in ["BV", "CD", "JG", "CDR"]:
        filtro = (
            (df_vendas_elotech['segmento'] == segmento) &
            (df_vendas_elotech['serie'] == serie) &
            (df_vendas_elotech['unidade'] == u_cod)
        )
        vend_uni[u_cod] = int(df_vendas_elotech.loc[filtro, 'vendido'].sum()) if filtro.any() else 0
    vendido_total = sum(vend_uni.values())
    saldo = pedido - vendido_total
    regs_tech_vg.append({
        'Descricao': f"Fund1 - {serie}",
        'Pedido': pedido,
        'Vendido': vendido_total,
        'Saldo Estoque': saldo,
        'BV': vend_uni['BV'],
        'CD': vend_uni['CD'],
        'JG': vend_uni['JG'],
        'CDR': vend_uni['CDR'],
    })

tot_pedido = sum(r['Pedido'] for r in regs_tech_vg)
tot_vendido = sum(r['Vendido'] for r in regs_tech_vg)
tot_saldo = tot_pedido - tot_vendido

rows_tech = list(regs_tech_vg)
rows_tech.append({
    'Descricao': 'TOTAL',
    'Pedido': tot_pedido,
    'Vendido': tot_vendido,
    'Saldo Estoque': tot_saldo,
    'BV': sum(r['BV'] for r in regs_tech_vg),
    'CD': sum(r['CD'] for r in regs_tech_vg),
    'JG': sum(r['JG'] for r in regs_tech_vg),
    'CDR': sum(r['CDR'] for r in regs_tech_vg),
})

df_tech_vg = pd.DataFrame(rows_tech)

col_r1, col_r2, col_r3 = st.columns(3)
with col_r1:
    st.metric("Pedido", f"{tot_pedido:,}")
with col_r2:
    st.metric("Vendido", f"{tot_vendido:,}")
with col_r3:
    st.metric("Saldo Estoque", f"{tot_saldo:,}", delta=f"{tot_saldo:+,}")

styled_tech_vg = df_tech_vg.style.map(colorir_estoque, subset=['Saldo Estoque'])
st.dataframe(styled_tech_vg, use_container_width=True, hide_index=True)
st.markdown("""
<div style='font-size: 0.85rem; color: #64748b; margin-top: 0.5rem;'>
    <b>Saldo Estoque</b> = Pedido - Vendido &nbsp;|&nbsp;
    <b>BV/CD/JG/CDR</b> = Vendido por unidade
</div>
""", unsafe_allow_html=True)

# Conferencia de Contratos Elo Tech por Unidade
st.markdown("#### Conferencia de Contratos Elo Tech por Unidade")
for u_cod_et, u_nome_et in [("BV", "Boa Viagem"), ("CD", "Candeias"), ("JG", "Janga")]:
    with st.expander(f"📋 {u_nome_et}", expanded=False):
        regs_conf_et = []
        for segmento, serie in ORDEM_SERIES_ELOTECH:
            pedido = PEDIDO_ELOTECH.get(serie, 0)
            filtro = (
                (df_vendas_elotech['segmento'] == segmento) &
                (df_vendas_elotech['serie'] == serie) &
                (df_vendas_elotech['unidade'] == u_cod_et)
            )
            vendido = int(df_vendas_elotech.loc[filtro, 'vendido'].sum()) if filtro.any() else 0
            regs_conf_et.append({
                'Descricao': f"Fund1 - {serie}",
                'Pedido Geral': pedido,
                'Vendido': vendido,
                'Contratos': vendido,
                'Assinaturas Caderno': '',
            })
        df_conf_et = pd.DataFrame(regs_conf_et)
        totais_conf_et = {
            'Descricao': 'TOTAL',
            'Pedido Geral': int(df_conf_et['Pedido Geral'].sum()),
            'Vendido': int(df_conf_et['Vendido'].sum()),
            'Contratos': int(df_conf_et['Contratos'].sum()),
            'Assinaturas Caderno': '',
        }
        df_conf_et = pd.concat([df_conf_et, pd.DataFrame([totais_conf_et])], ignore_index=True)
        st.dataframe(df_conf_et, use_container_width=True, hide_index=True)

st.markdown("---")

# Tabela vendido por unidade (original)
st.markdown("**Vendido por Unidade (detalhe)**")
st.dataframe(df_robo, use_container_width=True, hide_index=True)

# Expander: alunos Elo Tech "Outros"
if not df_outros_detail.empty:
    with st.expander(f"Alunos classificados como Outros ({len(df_outros_detail)})"):
        st.dataframe(
            df_outros_detail[['matricula', 'nome', 'unidade', 'turma', 'serie_real']]\
                .rename(columns={'serie_real': 'Serie Identificada'}),
            use_container_width=True, hide_index=True,
        )

# Expander: Alunos Elo Tech "Em aberto" (pagamento pendente)
if ELOTECH_EM_ABERTO:
    with st.expander(f"Alunos Elo Tech com pagamento Em Aberto ({len(ELOTECH_EM_ABERTO)})"):
        st.caption("Alunos faturados no SIGA mas com titulo nao liquidado. Incluidos na contagem conforme PDF financeiro.")
        _aberto_data = [{'Matricula': m, 'Nome': n, 'Unidade': UNIDADES.get(u, u), 'Serie': s}
                        for m, n, u, s in ELOTECH_EM_ABERTO]
        st.dataframe(pd.DataFrame(_aberto_data), use_container_width=True, hide_index=True)

# Expander: Ajuste manual de vendas 2025
if AJUSTE_ANO_PASSADO:
    with st.expander("Ajuste Manual - Vendas 2025 (alunos do ano passado)"):
        st.caption("Alunos que compraram livro em 2025 e nao contam como venda nova 2026. Somados ao estoque.")
        _aj_rows = []
        for cod, (serie, segmento, _, _) in PEDIDO_SAE.items():
            ajuste = AJUSTE_ANO_PASSADO.get(cod, {})
            if any(v > 0 for v in ajuste.values()):
                for u_cod, u_nome in UNIDADES.items():
                    val = ajuste.get(u_cod, 0)
                    if val > 0:
                        _aj_rows.append({'Codigo': cod, 'Segmento': segmento, 'Serie': serie,
                                         'Unidade': u_nome, 'Ajuste': val})
        # Elo Tech (995)
        ajuste_tech = AJUSTE_ANO_PASSADO.get(995, {})
        for u_cod, u_nome in UNIDADES.items():
            val = ajuste_tech.get(u_cod, 0)
            if val > 0:
                _aj_rows.append({'Codigo': 995, 'Segmento': 'Geral', 'Serie': 'Elo Tech',
                                 'Unidade': u_nome, 'Ajuste': val})
        if _aj_rows:
            df_aj = pd.DataFrame(_aj_rows)
            st.dataframe(df_aj, use_container_width=True, hide_index=True)
            st.metric("Total Ajustes 2025", int(df_aj['Ajuste'].sum()))

# --- Barra de acoes Elo Tech ---
st.markdown("---")
col_dl_tech, col_ata_tech, col_print_tech = st.columns(3)

with col_dl_tech:
    st.download_button(
        label="Baixar Vendas Elo Tech (Excel)",
        data=xlsx_robo,
        file_name=f"vendas_elotech_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_robo_xlsx",
        use_container_width=True,
    )

with col_ata_tech:
    if st.button(f"Ata Elo Tech ({len(df_ata_tech)})", use_container_width=True, key="btn_ata_tech"):
        st.session_state['ata_tipo'] = 'tech'

with col_print_tech:
    if st.button("Imprimir Elo Tech", use_container_width=True, key="btn_print_tech"):
        st.session_state['print_secao'] = 'tech'

# Renderizar ata Elo Tech se selecionada
if st.session_state.get('ata_tipo') == 'tech':
    components.html(gerar_ata_elotech_html(df_ata_tech), height=800, scrolling=True)

# Renderizar impressao Elo Tech se selecionada
if st.session_state.get('print_secao') == 'tech':
    df_robo_print = df_robo.copy()
    df_robo_print.columns = [c.replace('ã', 'a').replace('í', 'i').replace('ó', 'o').replace('ç', 'c') for c in df_robo_print.columns]
    tbl_tech = df_to_html_colorido(df_robo_print, "Vendas Elo Tech por Serie (Alunos Unicos) - CDR excluido")
    kpis_tech = [
        ("Total Vendido Elo Tech", f"{total_robo_vendido}"),
    ]
    html_tech = gerar_html_impressao_secao("Vendas Elo Tech", kpis_tech, [tbl_tech], data_str_impressao, filtro_unidade)
    _abrir_impressao(html_tech)
    st.session_state['print_secao'] = None


# ###########################################################
# VISAO GERAL
# ###########################################################

st.markdown("---")
st.markdown("### 📊 Visao Geral - Comparativo por Segmento")

fig = go.Figure()
fig.add_trace(go.Bar(name='Enviado', x=df_seg['segmento'], y=df_seg['enviado'],
                      marker_color='#3b82f6', text=df_seg['enviado'], textposition='outside'))
fig.add_trace(go.Bar(name='Vendido', x=df_seg['segmento'], y=df_seg['vendido'],
                      marker_color='#22c55e', text=df_seg['vendido'], textposition='outside'))
fig.add_trace(go.Bar(name='Estoque', x=df_seg['segmento'], y=df_seg['estoque'],
                      marker_color='#f59e0b', text=df_seg['estoque'], textposition='outside'))

fig.update_layout(
    barmode='group',
    paper_bgcolor='#ffffff', plot_bgcolor='#f8fafc',
    font=dict(color='#1e293b'),
    legend=dict(orientation='h', y=1.1),
    xaxis=dict(gridcolor='#e2e8f0'),
    yaxis=dict(gridcolor='#e2e8f0', title='Quantidade'),
    margin=dict(t=60, b=20)
)
st.plotly_chart(fig, use_container_width=True)

# --- Barra de acoes Visao Geral ---
st.markdown("---")
col_dl_tudo, col_print_completo = st.columns(2)

with col_dl_tudo:
    st.download_button(
        label="Baixar Tudo (Excel)",
        data=xlsx_tudo,
        file_name=f"controle_estoque_completo_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_tudo_xlsx",
        use_container_width=True,
    )

with col_print_completo:
    if st.button("Imprimir Completo", use_container_width=True, key="btn_print_completo"):
        st.session_state['print_secao'] = 'completo'

# Renderizar impressao completa se selecionada
if st.session_state.get('print_secao') == 'completo':
    # Tabela estoque restante (todas unidades, sem filtro)
    df_est_print = criar_tabela_completa(df_estoque_completo, 'estoque', "Todas")
    df_est_print.columns = [c.replace('ã', 'a').replace('í', 'i').replace('ó', 'o').replace('ç', 'c') for c in df_est_print.columns]
    tbl_estoque_full = df_to_html_colorido(df_est_print, "Estoque Restante SAE (Enviado - Vendido)")

    df_socio_print = df_socio.copy()
    df_socio_print.columns = [c.replace('ã', 'a').replace('í', 'i').replace('ó', 'o').replace('ç', 'c') for c in df_socio_print.columns]
    tbl_socio_full = df_to_html_colorido(df_socio_print, "Vendas Socioemocional (Alunos Unicos)")

    df_robo_print = df_robo.copy()
    df_robo_print.columns = [c.replace('ã', 'a').replace('í', 'i').replace('ó', 'o').replace('ç', 'c') for c in df_robo_print.columns]
    tbl_robo_full = df_to_html_colorido(df_robo_print, "Vendas Elo Tech por Serie (Alunos Unicos) - CDR excluido")

    kpis_all = [
        ("Pedido Total", f"{pedido_total:,}"),
        ("Enviado Total", f"{enviado_total:,}"),
        ("Vendido SAE", f"{vendido_total:,}"),
        ("Estoque Restante", f"{estoque_total:,}"),
        ("% Vendido", f"{percentual_venda:.1f}%"),
        ("Socioemocional", f"{total_socio_vendido}"),
        ("Elo Tech", f"{total_robo_vendido}"),
    ]
    html_completo = gerar_html_impressao_completo(kpis_all, [tbl_estoque_full, tbl_socio_full, tbl_robo_full], data_str_impressao, filtro_unidade)
    _abrir_impressao(html_completo)
    st.session_state['print_secao'] = None
