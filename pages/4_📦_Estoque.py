"""Página de Estoque SAE + Socioemocional - Controle de Livros."""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import json
import re
import io
from pathlib import Path
from datetime import datetime
from utils.theme import aplicar_tema

st.set_page_config(
    page_title="Estoque - Colégio Elo",
    page_icon="📦",
    layout="wide",
)

aplicar_tema()

OUTPUT_DIR = Path(__file__).parent.parent / "output"

# Importa dados_estoque
from utils.dados_estoque import (
    PEDIDO_SAE, ESTOQUE_ENVIADO, UNIDADES, ORDEM_SEGMENTOS,
    BALANCO_FISICO, AJUSTE_ANO_PASSADO, ELOTECH_EM_ABERTO,
    ELOTECH_SERIE_PDF,
)

CODIGOS_SERVICOS = {
    "901", "902", "903", "904",
    "913", "914", "915", "920", "921",
    "916", "917", "918", "919",
    "912", "991", "992",
    "933", "934", "941", "942", "943", "944", "945", "946",
    "995",
}

# Cordeiro NÃO tem Elo Tech
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

UNIDADES_INV = {v: k for k, v in UNIDADES.items()}


# =====================================================
# FUNCOES
# =====================================================

def parse_tsv(tsv_content, unidade_codigo):
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
                    "unidade": unidade_codigo, "servico_codigo": servico_codigo,
                    "turma": turma_atual, "matricula": matricula,
                    "nome": campos[1].strip(), "titulo": campos[2].strip(),
                    "parcela": campos[3].strip(), "dt_baixa": campos[4].strip(),
                    "valor": campos[5].strip(),
                    "recebido": campos[-1].strip() if len(campos) > 6 else "",
                    "segmento": info[0], "serie": info[1], "tipo": info[2],
                })
    return registros


def atualizar_dados():
    todos_registros = []
    agora = datetime.now()
    for arquivo in sorted(OUTPUT_DIR.glob("dados_*.tsv")):
        codigo = arquivo.stem.replace("dados_", "").upper()
        with open(arquivo, 'r', encoding='utf-8') as f:
            registros = parse_tsv(f.read(), codigo)
            todos_registros.extend(registros)
    json_path = OUTPUT_DIR / "recebimento_final.json"
    if json_path.exists():
        with open(json_path, 'r', encoding='utf-8') as f:
            dados_atuais = json.load(f)
        registros_atuais = len(dados_atuais.get('registros', []))
        if len(todos_registros) < registros_atuais * 0.8:
            return -1, agora
    dados_finais = {
        "data_extracao": agora.isoformat(),
        "ultima_atualizacao": agora.strftime("%d/%m/%Y %H:%M:%S"),
        "periodo": {"data_inicial": "01/08/2025", "data_final": agora.strftime("%d/%m/%Y")},
        "total_registros": len(todos_registros),
        "registros": todos_registros,
    }
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(dados_finais, f, ensure_ascii=False, indent=2)
    return len(todos_registros), agora


@st.cache_data(ttl=300)
def carregar_vendas():
    json_path = OUTPUT_DIR / "recebimento_final.json"
    if not json_path.exists():
        return None, None
    with open(json_path, 'r', encoding='utf-8') as f:
        dados = json.load(f)
    return pd.DataFrame(dados.get('registros', [])), dados.get('ultima_atualizacao', 'N/A')


def calcular_vendas(df, tipo="SAE"):
    df_f = df[df['tipo'] == tipo].copy()
    vendas = df_f.groupby(['segmento', 'serie', 'unidade'])['matricula'].nunique().reset_index()
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
    registros = []
    for cod, (serie, segmento, ped_ini, ped_comp) in PEDIDO_SAE.items():
        ped_total = ped_ini + ped_comp
        env = ESTOQUE_ENVIADO.get(cod, {})
        ajuste_ap = AJUSTE_ANO_PASSADO.get(cod, {})
        for u_cod, u_nome in UNIDADES.items():
            enviado = env.get(u_cod, 0)
            ajuste = ajuste_ap.get(u_cod, 0)
            vendido = 0
            if df_vendas is not None:
                f = (df_vendas['segmento'] == segmento) & (df_vendas['serie'] == serie) & (df_vendas['unidade'] == u_cod)
                if f.any():
                    vendido = df_vendas.loc[f, 'vendido'].values[0]
            registros.append({
                'codigo': cod, 'segmento': segmento, 'serie': serie,
                'unidade_cod': u_cod, 'unidade': u_nome,
                'pedido_inicial': ped_ini, 'pedido_compl': ped_comp,
                'pedido_total': ped_total, 'enviado': enviado,
                'ajuste': ajuste,
                'vendido': vendido, 'estoque': enviado - vendido + ajuste,
            })
    df = pd.DataFrame(registros)
    df['_ord'] = df['segmento'].map({s: i for i, s in enumerate(ORDEM_SEGMENTOS)})
    df = df.sort_values(['_ord', 'serie', 'unidade_cod']).drop(columns=['_ord'])
    return df


def criar_tabela_completa(df_dados, coluna_valor, unidade_sel=None):
    unidades_m = [unidade_sel] if unidade_sel and unidade_sel != "Todas" else list(UNIDADES.values())
    registros = []
    for seg, serie in ORDEM_SERIES_COMPLETA:
        reg = {'segmento': seg, 'serie': serie, 'Descrição': f"{seg} - {serie}"}
        for u in unidades_m:
            f = (df_dados['segmento'] == seg) & (df_dados['serie'] == serie) & (df_dados['unidade'] == u)
            reg[u] = int(df_dados.loc[f, coluna_valor].sum()) if f.any() else 0
        registros.append(reg)
    df_r = pd.DataFrame(registros)
    cols_v = [c for c in df_r.columns if c not in ['segmento', 'serie', 'Descrição']]
    df_r['TOTAL'] = df_r[cols_v].sum(axis=1)
    totais = {'segmento': '', 'serie': '', 'Descrição': 'TOTAL'}
    for c in cols_v:
        totais[c] = df_r[c].sum()
    totais['TOTAL'] = df_r['TOTAL'].sum()
    df_r = pd.concat([df_r, pd.DataFrame([totais])], ignore_index=True)
    df_r = df_r.drop(columns=['segmento', 'serie'])
    return df_r[['Descrição'] + cols_v + ['TOTAL']]


def colorir_estoque(val):
    if isinstance(val, str):
        return ''
    if val < 0:
        return 'background-color: rgba(220,38,38,0.1); color: #dc2626; font-weight: bold'
    elif val <= 5:
        return 'background-color: rgba(202,138,4,0.1); color: #b45309; font-weight: bold'
    return 'background-color: rgba(22,163,74,0.1); color: #16a34a'


def colorir_dif(val):
    if isinstance(val, str):
        return ''
    if val < 0:
        return 'color: #dc2626; font-weight: bold'
    elif val > 0:
        return 'color: #b45309; font-weight: bold'
    return 'color: #16a34a'


def gerar_excel(dfs_dict):
    """Gera um arquivo Excel com múltiplas abas a partir de um dicionário {nome_aba: DataFrame}."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for nome_aba, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=nome_aba[:31], index=False)
    buffer.seek(0)
    return buffer


def gerar_csv(df):
    """Gera um CSV a partir de um DataFrame."""
    return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')


def cor_celula_print(val):
    """Retorna o estilo inline para células na versão de impressão."""
    if isinstance(val, str):
        return ''
    try:
        v = float(val)
    except (ValueError, TypeError):
        return ''
    if v < 0:
        return 'background-color: #fecaca; color: #991b1b; font-weight: bold;'
    elif v <= 5:
        return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
    return 'background-color: #dcfce7; color: #166534;'


def df_to_html_colorido(df, titulo, colunas_colorir=None):
    """Converte DataFrame para HTML com células coloridas para impressão."""
    html = f'<h3 style="color:#1e293b; margin-top:25px; border-bottom:2px solid #667eea; padding-bottom:6px;">{titulo}</h3>\n'
    html += '<table style="width:100%; border-collapse:collapse; font-size:12px; margin-bottom:15px;">\n'
    html += '<thead><tr style="background:#f0f2f6;">'
    for col in df.columns:
        html += f'<th style="padding:8px 10px; text-align:left; border:1px solid #d1d5db; font-weight:600; color:#374151;">{col}</th>'
    html += '</tr></thead>\n<tbody>\n'
    for idx, row in df.iterrows():
        is_total = str(row.iloc[0]).strip().upper() == 'TOTAL'
        row_style = 'background:#e0e7ff; font-weight:700;' if is_total else ''
        html += f'<tr style="{row_style}">'
        for col in df.columns:
            val = row[col]
            cell_style = 'padding:6px 10px; border:1px solid #e5e7eb;'
            if colunas_colorir and col in colunas_colorir:
                cell_style += ' ' + cor_celula_print(val)
            html += f'<td style="{cell_style}">{val}</td>'
        html += '</tr>\n'
    html += '</tbody></table>\n'
    return html


def gerar_html_impressao_secao(titulo, kpis, tabelas_html, data_str, filtro_unidade="Todas"):
    """Gera HTML de impressao para uma secao individual, respeitando filtro de unidade."""
    unidade_label = f" - {filtro_unidade}" if filtro_unidade != "Todas" else ""
    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>{titulo} - Colégio Elo</title>
<style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #1e293b;
        background: #ffffff;
        padding: 20px 30px;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
    }}
    @media print {{
        body {{ padding: 10px 15px; }}
        * {{
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
            color-adjust: exact !important;
        }}
        .no-print {{ display: none !important; }}
        h2 {{ page-break-before: auto; }}
        table {{ page-break-inside: auto; }}
        tr {{ page-break-inside: avoid; }}
    }}
    .header {{
        text-align: center;
        border-bottom: 3px solid #667eea;
        padding-bottom: 15px;
        margin-bottom: 20px;
    }}
    .header h1 {{ font-size: 22px; color: #1e293b; margin-bottom: 4px; }}
    .header p {{ font-size: 13px; color: #64748b; }}
    .kpis {{
        display: flex;
        gap: 12px;
        justify-content: space-around;
        margin-bottom: 20px;
        flex-wrap: wrap;
    }}
    .kpi-box {{
        flex: 1;
        min-width: 120px;
        text-align: center;
        padding: 12px 8px;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        background: #f8fafc;
    }}
    .kpi-box .label {{ font-size: 10px; text-transform: uppercase; color: #64748b; letter-spacing: 1px; }}
    .kpi-box .value {{ font-size: 22px; font-weight: 700; color: #1e293b; }}
    .section-break {{ page-break-before: auto; margin-top: 25px; }}
    .btn-print {{
        display: inline-block;
        padding: 10px 24px;
        background: #667eea;
        color: #ffffff;
        border: none;
        border-radius: 8px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 600;
        margin: 15px 0;
    }}
    .btn-print:hover {{ background: #5a6fd6; }}
</style>
</head>
<body>
<div class="header">
    <h1>Colegio Elo - {titulo}{unidade_label}</h1>
    <p>Relatorio gerado em {data_str}</p>
</div>
<button class="btn-print no-print" onclick="window.print()">Imprimir</button>
<div class="kpis">
"""
    for label, valor in kpis:
        html += f'<div class="kpi-box"><div class="label">{label}</div><div class="value">{valor}</div></div>\n'
    html += '</div>\n'
    for secao_html in tabelas_html:
        html += f'<div class="section-break">{secao_html}</div>\n'
    html += """
<script>
    window.onload = function() { window.print(); };
</script>
</body>
</html>"""
    return html


def gerar_ata_html(df_ata_src, titulo, col_serie='serie'):
    """Gera HTML de ata de entrega agrupada por unidade e turma."""
    data_atual = datetime.now().strftime("%d/%m/%Y")
    unidades_ata = sorted(df_ata_src['unidade'].unique())
    total_alunos = len(df_ata_src)
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
    """Filtra DataFrame de ata por unidade."""
    if filtro_un != "Todas":
        u_cod = UNIDADES_INV.get(filtro_un, "")
        return df_src[df_src['unidade'] == u_cod]
    return df_src


# =====================================================
# CARREGAR DADOS
# =====================================================
df_raw, ultima_att = carregar_vendas()

# Injetar alunos Elo Tech com pagamento "Em aberto" (faturados, não pagos)
if df_raw is not None and ELOTECH_EM_ABERTO:
    _mats_existentes = set(df_raw[df_raw['tipo'] == 'Elo Tech']['matricula'].unique())
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
        df_raw = pd.concat([df_raw, pd.DataFrame(_novos)], ignore_index=True)

# Header + Atualizar
col_t, col_b = st.columns([4, 1])
with col_t:
    st.markdown("# 📦 Controle de Estoque SAE")
with col_b:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Atualizar Dados", type="primary", use_container_width=True):
        with st.spinner("Atualizando..."):
            total_r, dt = atualizar_dados()
        if total_r == -1:
            st.warning("TSVs corrompidos (menos de 80% dos registros atuais). JSON nao foi sobrescrito. Re-execute a extracao completa.")
        else:
            st.cache_data.clear()
            st.success(f"{total_r} registros atualizados")
            st.rerun()

if ultima_att:
    st.caption(f"Dados: {ultima_att}")

if df_raw is None:
    st.error("Dados não encontrados. Clique em 'Atualizar Dados'.")
    st.stop()

# =====================================================
# COMPUTACAO BASE (antes dos filtros)
# =====================================================
df_vendas_sae = calcular_vendas(df_raw, "SAE")
df_vendas_socio = calcular_vendas(df_raw, "Socioemocional")
df_estoque_completo = criar_tabela_estoque(df_vendas_sae)
df_vendas_elotech = calcular_vendas_elotech_por_serie(df_raw)

# Elo Tech detail (para atas e expanders)
df_tech_detail = df_raw[df_raw['tipo'] == 'Elo Tech'].copy()
df_ref_detail = df_raw[df_raw['tipo'].isin(['SAE', 'Socioemocional'])].copy()
_mat_serie = df_ref_detail.drop_duplicates(subset=['matricula', 'unidade'])\
    .set_index(['matricula', 'unidade'])['serie'].to_dict()
df_tech_detail['serie_real'] = df_tech_detail.apply(
    lambda r: ELOTECH_SERIE_PDF.get((r['matricula'], r['unidade']),
              _mat_serie.get((r['matricula'], r['unidade']),
                             r['serie'] if r['serie'] != 'Todas' else 'Sem serie')), axis=1
)
_mat_turma = df_ref_detail[df_ref_detail['turma'].str.strip() != '']\
    .drop_duplicates(subset=['matricula', 'unidade'])\
    .set_index(['matricula', 'unidade'])['turma'].to_dict()
df_tech_detail['turma_real'] = df_tech_detail.apply(
    lambda r: _mat_turma.get((r['matricula'], r['unidade']),
                             r['turma'].strip() if r['turma'].strip() else ''), axis=1
)
_TURNO_MAP = {'A': 'Manha', 'B': 'Tarde', 'C': 'Integral'}
df_tech_detail['turno'] = df_tech_detail['turma_real'].map(
    lambda t: _TURNO_MAP.get(t.strip().upper(), '') if t.strip() else ''
)

st.divider()

# =====================================================
# FILTROS
# =====================================================
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    filtro_unidade = st.selectbox("Unidade", ["Todas"] + list(UNIDADES.values()))
with col_f2:
    filtro_segmento = st.selectbox("Segmento", ["Todos"] + ORDEM_SEGMENTOS)
with col_f3:
    if filtro_segmento != "Todos":
        _series_disp = sorted(df_estoque_completo[df_estoque_completo['segmento'] == filtro_segmento]['serie'].unique().tolist())
    else:
        _series_disp = sorted(df_estoque_completo['serie'].unique().tolist())
    filtro_serie = st.selectbox("Série", ["Todas"] + _series_disp)

# Aplicar filtros
df_estoque = df_estoque_completo.copy()
if filtro_unidade != "Todas":
    df_estoque = df_estoque[df_estoque['unidade'] == filtro_unidade]
if filtro_segmento != "Todos":
    df_estoque = df_estoque[df_estoque['segmento'] == filtro_segmento]
if filtro_serie != "Todas":
    df_estoque = df_estoque[df_estoque['serie'] == filtro_serie]

# =====================================================
# KPIs GLOBAIS
# =====================================================
pedido_total = df_estoque.drop_duplicates(subset=['codigo'])['pedido_total'].sum()
enviado_total = df_estoque['enviado'].sum()
vendido_total = df_estoque['vendido'].sum()
estoque_total = df_estoque['estoque'].sum()
ajuste_total = df_estoque['ajuste'].sum()
pct_venda = (vendido_total / pedido_total * 100) if pedido_total > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Pedido", f"{pedido_total:,}".replace(",", "."))
with col2:
    st.metric("Enviado", f"{enviado_total:,}".replace(",", "."))
with col3:
    st.metric("Vendido SAE", f"{vendido_total:,}".replace(",", "."))
with col4:
    st.metric("Estoque", f"{estoque_total:,}".replace(",", "."))
with col5:
    st.metric("% Vendido", f"{pct_venda:.1f}%")

# =====================================================
# COMPUTACAO FILTRADA (tudo antes de renderizar)
# =====================================================

# Socioemocional - tabela display
unidades_socio = [filtro_unidade] if filtro_unidade != "Todas" else list(UNIDADES.values())
registros_socio = []
for seg, serie in ORDEM_SERIES_SOCIO:
    reg = {'Descrição': f"{seg} - {serie}"}
    for u_nome in unidades_socio:
        u_cod = UNIDADES_INV.get(u_nome, "")
        f = (df_vendas_socio['segmento'] == seg) & (df_vendas_socio['serie'] == serie) & (df_vendas_socio['unidade'] == u_cod)
        reg[u_nome] = int(df_vendas_socio.loc[f, 'vendido'].sum()) if f.any() else 0
    registros_socio.append(reg)
df_socio = pd.DataFrame(registros_socio)
cols_s = [c for c in df_socio.columns if c != 'Descrição']
df_socio['TOTAL'] = df_socio[cols_s].sum(axis=1)
totais_s = {'Descrição': 'TOTAL'}
for c in cols_s:
    totais_s[c] = df_socio[c].sum()
totais_s['TOTAL'] = df_socio['TOTAL'].sum()
df_socio = pd.concat([df_socio, pd.DataFrame([totais_s])], ignore_index=True)
total_socio = int(df_socio.loc[df_socio['Descrição'] == 'TOTAL', 'TOTAL'].values[0])

# Elo Tech - tabela display
if not df_vendas_elotech.empty:
    unidades_rob = [filtro_unidade] if filtro_unidade != "Todas" else list(UNIDADES.values())
    registros_rob = []
    for seg, serie in ORDEM_SERIES_ELOTECH:
        reg = {'Descrição': f"{seg} - {serie}"}
        for u_nome in unidades_rob:
            u_cod = UNIDADES_INV.get(u_nome, "")
            f = (df_vendas_elotech['segmento'] == seg) & (df_vendas_elotech['serie'] == serie) & (df_vendas_elotech['unidade'] == u_cod)
            reg[u_nome] = int(df_vendas_elotech.loc[f, 'vendido'].sum()) if f.any() else 0
        registros_rob.append(reg)
    series_elotech = {s for _, s in ORDEM_SERIES_ELOTECH}
    reg_outros = {'Descrição': 'Outros'}
    for u_nome in unidades_rob:
        u_cod = UNIDADES_INV.get(u_nome, "")
        f = (~df_vendas_elotech['serie'].isin(series_elotech)) & (df_vendas_elotech['unidade'] == u_cod)
        reg_outros[u_nome] = int(df_vendas_elotech.loc[f, 'vendido'].sum()) if f.any() else 0
    if any(v for k, v in reg_outros.items() if k != 'Descrição'):
        registros_rob.append(reg_outros)
    df_rob = pd.DataFrame(registros_rob)
    cols_r = [c for c in df_rob.columns if c != 'Descrição']
    df_rob['TOTAL'] = df_rob[cols_r].sum(axis=1)
    totais_r = {'Descrição': 'TOTAL'}
    for c in cols_r:
        totais_r[c] = df_rob[c].sum()
    totais_r['TOTAL'] = df_rob['TOTAL'].sum()
    df_rob = pd.concat([df_rob, pd.DataFrame([totais_r])], ignore_index=True)
    total_rob = int(totais_r['TOTAL'])
else:
    df_rob = pd.DataFrame()
    total_rob = 0

# Download DFs
df_download_estoque = df_estoque[['segmento', 'serie', 'unidade', 'pedido_total', 'enviado', 'vendido', 'ajuste', 'estoque']].rename(
    columns={'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade',
             'pedido_total': 'Pedido', 'enviado': 'Enviado', 'vendido': 'Vendido',
             'ajuste': 'Ajuste 2025', 'estoque': 'Estoque'}
)
df_socio_dl = df_socio.copy()
df_rob_dl = df_rob.copy() if not df_rob.empty else pd.DataFrame()

# Ata data
df_ata_sae = _filtrar_unidade_ata(
    df_raw[df_raw['tipo'] == 'SAE'].drop_duplicates(subset=['matricula', 'unidade']).sort_values(['unidade', 'serie', 'nome']),
    filtro_unidade)
df_ata_socio = _filtrar_unidade_ata(
    df_raw[df_raw['tipo'] == 'Socioemocional'].drop_duplicates(subset=['matricula', 'unidade']).sort_values(['unidade', 'serie', 'nome']),
    filtro_unidade)
df_ata_tech = _filtrar_unidade_ata(
    df_tech_detail.drop_duplicates(subset=['matricula', 'unidade']).sort_values(['unidade', 'serie_real', 'nome']),
    filtro_unidade)

# Balanço Físico data
registros_bal = []
for cod, (serie, segmento, _, _) in PEDIDO_SAE.items():
    bal = BALANCO_FISICO.get(cod, {})
    env = ESTOQUE_ENVIADO.get(cod, {})
    for u_cod, u_nome in UNIDADES.items():
        fisico_d = bal.get(u_cod, {})
        fisico = list(fisico_d.values())[-1] if fisico_d else None
        enviado = env.get(u_cod, 0)
        vendido = 0
        f = (df_vendas_sae['segmento'] == segmento) & (df_vendas_sae['serie'] == serie) & (df_vendas_sae['unidade'] == u_cod)
        if f.any():
            vendido = int(df_vendas_sae.loc[f, 'vendido'].values[0])
        teorico = enviado - vendido
        dif = (fisico - teorico) if fisico is not None else None
        registros_bal.append({
            'Descrição': f"{segmento} - {serie}", 'Unidade': u_nome,
            'Enviado': enviado, 'Vendido': vendido, 'Teórico': teorico,
            'Físico': fisico if fisico is not None else "-",
            'Diferença': dif if dif is not None else "-",
        })
df_bal = pd.DataFrame(registros_bal)
if filtro_unidade != "Todas":
    df_bal = df_bal[df_bal['Unidade'] == filtro_unidade]

# SAE Detail (Visao Completa) - pre-compute para uso em tab e impressao
registros_det = []
for seg, serie in ORDEM_SERIES_COMPLETA:
    f = (df_estoque['segmento'] == seg) & (df_estoque['serie'] == serie)
    if f.any():
        reg = {
            'Descrição': f"{seg} - {serie}",
            'Pedido': int(df_estoque.loc[f, 'pedido_total'].iloc[0]),
            'Enviado': int(df_estoque.loc[f, 'enviado'].sum()),
            'Vendido': int(df_estoque.loc[f, 'vendido'].sum()),
            'Ajuste 2025': int(df_estoque.loc[f, 'ajuste'].sum()),
        }
        if filtro_unidade == "Todas":
            for u_nome in UNIDADES.values():
                f_u = f & (df_estoque['unidade'] == u_nome)
                reg[u_nome] = int(df_estoque.loc[f_u, 'estoque'].sum()) if f_u.any() else 0
            reg['Total'] = sum(reg[u] for u in UNIDADES.values())
        else:
            reg['Estoque'] = int(df_estoque.loc[f, 'estoque'].sum())
        registros_det.append(reg)
df_det = pd.DataFrame(registros_det)
if not df_det.empty:
    totais = {'Descrição': 'TOTAL'}
    for c in df_det.columns:
        if c != 'Descrição':
            totais[c] = df_det[c].sum()
    df_det = pd.concat([df_det, pd.DataFrame([totais])], ignore_index=True)
# Colunas de estoque para colorir
_det_cols_estoque = list(UNIDADES.values()) + ['Total'] if filtro_unidade == "Todas" else ['Estoque']


# #####################################################
# SECAO SAE
# #####################################################
st.divider()
st.header("Livros SAE")

tab1, tab2, tab3, tab4 = st.tabs(["Visão Completa", "Estoque Restante", "Enviado", "Vendido"])

with tab1:
    if filtro_unidade == "Todas":
        st.markdown("**Pedido / Enviado / Vendido / Ajuste 2025 / Estoque por Unidade**")
    else:
        st.markdown(f"**Pedido / Enviado / Vendido / Ajuste 2025 / Estoque — {filtro_unidade}**")
    if not df_det.empty:
        styled_det = df_det.style.map(colorir_estoque, subset=_det_cols_estoque)
        st.dataframe(styled_det, width="stretch", hide_index=True)

with tab2:
    df_d = criar_tabela_completa(df_estoque, 'estoque', filtro_unidade)
    styled = df_d.style.map(colorir_estoque, subset=[c for c in df_d.columns if c != 'Descrição'])
    st.dataframe(styled, width="stretch", hide_index=True)

with tab3:
    df_d = criar_tabela_completa(df_estoque, 'enviado', filtro_unidade)
    st.dataframe(df_d, width="stretch", hide_index=True)

with tab4:
    df_d = criar_tabela_completa(df_estoque, 'vendido', filtro_unidade)
    st.dataframe(df_d, width="stretch", hide_index=True)

# Balanço Físico (dentro da secao SAE)
st.markdown("#### Balanço Físico vs Estoque Teórico")
st.caption("Contagem física (29/01) vs estoque calculado (enviado - vendido)")
df_bal_v = df_bal[df_bal['Físico'] != "-"].copy()
if not df_bal_v.empty:
    styled_bal = df_bal_v.style.map(colorir_dif, subset=['Diferença'])
    st.dataframe(styled_bal, width="stretch", hide_index=True)
else:
    st.info("Nenhum balanço físico registrado.")

# Alertas (dentro da secao SAE)
st.markdown("#### Alertas de Estoque")
col_a1, col_a2 = st.columns(2)
with col_a1:
    st.markdown("**Estoque Negativo**")
    df_neg = df_estoque[df_estoque['estoque'] < 0][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']]
    if not df_neg.empty:
        st.dataframe(df_neg.rename(columns={'segmento': 'Seg.', 'serie': 'Série', 'unidade': 'Unidade', 'enviado': 'Env.', 'vendido': 'Vend.', 'estoque': 'Falta'}), width="stretch", hide_index=True)
    else:
        st.success("Nenhum estoque negativo!")
with col_a2:
    st.markdown("**Estoque Baixo (<=5)**")
    df_bx = df_estoque[(df_estoque['estoque'] > 0) & (df_estoque['estoque'] <= 5)][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']]
    if not df_bx.empty:
        st.dataframe(df_bx.rename(columns={'segmento': 'Seg.', 'serie': 'Série', 'unidade': 'Unidade', 'enviado': 'Env.', 'vendido': 'Vend.', 'estoque': 'Rest.'}), width="stretch", hide_index=True)
    else:
        st.success("Nenhum estoque baixo!")

# Acoes SAE
st.markdown("---")
col_dl, col_ata, col_print = st.columns(3)
with col_dl:
    excel_sae = gerar_excel({"Estoque SAE": df_download_estoque})
    st.download_button(
        label="Download SAE (Excel)",
        data=excel_sae,
        file_name=f"estoque_sae_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_ata:
    if st.button(f"Ata SAE ({len(df_ata_sae)})", use_container_width=True):
        st.session_state['ata_tipo'] = 'sae'
        st.session_state['print_secao'] = None
with col_print:
    if st.button("Imprimir SAE", use_container_width=True):
        st.session_state['print_secao'] = 'sae'
        st.session_state['ata_tipo'] = None

if st.session_state.get('ata_tipo') == 'sae':
    components.html(gerar_ata_html(df_ata_sae, "Livros SAE", 'serie'), height=800, scrolling=True)

if st.session_state.get('print_secao') == 'sae':
    _data_imp = datetime.now().strftime("%d/%m/%Y %H:%M")
    _kpis_sae = [
        ("Pedido", f"{pedido_total:,}".replace(",", ".")),
        ("Enviado", f"{enviado_total:,}".replace(",", ".")),
        ("Vendido SAE", f"{vendido_total:,}".replace(",", ".")),
        ("Estoque", f"{estoque_total:,}".replace(",", ".")),
        ("% Vendido", f"{pct_venda:.1f}%"),
    ]
    _tabelas_sae = []
    _df_est_p = criar_tabela_completa(df_estoque, 'estoque', filtro_unidade)
    _cols_num = [c for c in _df_est_p.columns if c != 'Descrição']
    _tabelas_sae.append(df_to_html_colorido(_df_est_p, "Estoque SAE por Serie e Unidade", _cols_num))
    if not df_det.empty:
        _tabelas_sae.append(df_to_html_colorido(df_det, "Detalhado (Pedido / Enviado / Vendido / Estoque)", _det_cols_estoque))
    _df_neg_p = df_estoque[df_estoque['estoque'] < 0][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']].rename(
        columns={'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade', 'enviado': 'Enviado', 'vendido': 'Vendido', 'estoque': 'Falta'})
    if not _df_neg_p.empty:
        _tabelas_sae.append(df_to_html_colorido(_df_neg_p, "Alerta: Estoque Negativo", ['Falta']))
    _df_bx_p = df_estoque[(df_estoque['estoque'] > 0) & (df_estoque['estoque'] <= 5)][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']].rename(
        columns={'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade', 'enviado': 'Enviado', 'vendido': 'Vendido', 'estoque': 'Restante'})
    if not _df_bx_p.empty:
        _tabelas_sae.append(df_to_html_colorido(_df_bx_p, "Alerta: Estoque Baixo (<=5)", ['Restante']))
    components.html(gerar_html_impressao_secao("Estoque SAE", _kpis_sae, _tabelas_sae, _data_imp, filtro_unidade), height=800, scrolling=True)


# #####################################################
# SECAO SOCIOEMOCIONAL
# #####################################################
st.divider()
st.header("Livros Socioemocional")

st.metric("Total Socioemocional Vendido", f"{total_socio}")
st.dataframe(df_socio, width="stretch", hide_index=True)

# Acoes Socioemocional
st.markdown("---")
col_dl2, col_ata2, col_print2 = st.columns(3)
with col_dl2:
    excel_socio = gerar_excel({"Socioemocional": df_socio_dl})
    st.download_button(
        label="Download Socioemocional (Excel)",
        data=excel_socio,
        file_name=f"vendas_socioemocional_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_ata2:
    if st.button(f"Ata Socioemocional ({len(df_ata_socio)})", use_container_width=True):
        st.session_state['ata_tipo'] = 'socio'
        st.session_state['print_secao'] = None
with col_print2:
    if st.button("Imprimir Socioemocional", use_container_width=True):
        st.session_state['print_secao'] = 'socio'
        st.session_state['ata_tipo'] = None

if st.session_state.get('ata_tipo') == 'socio':
    components.html(gerar_ata_html(df_ata_socio, "Livros Socioemocional", 'serie'), height=800, scrolling=True)

if st.session_state.get('print_secao') == 'socio':
    _data_imp = datetime.now().strftime("%d/%m/%Y %H:%M")
    _kpis_socio = [("Total Vendido", f"{total_socio}")]
    _tabelas_socio = [df_to_html_colorido(df_socio, "Vendas Socioemocional (Alunos Unicos)", None)]
    components.html(gerar_html_impressao_secao("Socioemocional", _kpis_socio, _tabelas_socio, _data_imp, filtro_unidade), height=800, scrolling=True)


# #####################################################
# SECAO ELO TECH
# #####################################################
st.divider()
st.header("Livros Elo Tech")
st.caption("Nota: Cordeiro (CDR) nao possui Elo Tech. Serie identificada via cruzamento com matricula SAE.")

if not df_vendas_elotech.empty:
    st.metric("Total Elo Tech Vendido", f"{total_rob}")
    st.dataframe(df_rob, width="stretch", hide_index=True)

    # Expander: alunos Elo Tech "Outros"
    _series_ok = {s for _, s in ORDEM_SERIES_ELOTECH}
    df_outros_detail = df_tech_detail[~df_tech_detail['serie_real'].isin(_series_ok)]\
        .drop_duplicates(subset=['matricula', 'unidade'])\
        .sort_values(['unidade', 'serie_real', 'nome'])
    if not df_outros_detail.empty:
        with st.expander(f"Alunos classificados como Outros ({len(df_outros_detail)})"):
            st.dataframe(
                df_outros_detail[['matricula', 'nome', 'unidade', 'turma', 'serie_real']]\
                    .rename(columns={'serie_real': 'Serie Identificada'}),
                width="stretch", hide_index=True,
            )
else:
    st.info("Nenhuma venda de Elo Tech registrada ainda.")

# Expander: Alunos Elo Tech "Em aberto" (pagamento pendente)
if ELOTECH_EM_ABERTO:
    with st.expander(f"Alunos Elo Tech com pagamento Em Aberto ({len(ELOTECH_EM_ABERTO)})"):
        st.caption("Alunos faturados no SIGA mas com titulo nao liquidado. Incluidos na contagem conforme PDF financeiro.")
        _aberto_data = [{'Matricula': m, 'Nome': n, 'Unidade': UNIDADES.get(u, u), 'Serie': s}
                        for m, n, u, s in ELOTECH_EM_ABERTO]
        st.dataframe(pd.DataFrame(_aberto_data), width="stretch", hide_index=True)

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
        ajuste_tech = AJUSTE_ANO_PASSADO.get(995, {})
        for u_cod, u_nome in UNIDADES.items():
            val = ajuste_tech.get(u_cod, 0)
            if val > 0:
                _aj_rows.append({'Codigo': 995, 'Segmento': 'Geral', 'Serie': 'Elo Tech',
                                 'Unidade': u_nome, 'Ajuste': val})
        if _aj_rows:
            df_aj = pd.DataFrame(_aj_rows)
            st.dataframe(df_aj, width="stretch", hide_index=True)
            st.metric("Total Ajustes 2025", int(df_aj['Ajuste'].sum()))

# Acoes Elo Tech
st.markdown("---")
col_dl3, col_ata3, col_print3 = st.columns(3)
with col_dl3:
    if not df_rob_dl.empty:
        excel_rob = gerar_excel({"Elo Tech": df_rob_dl})
        st.download_button(
            label="Download Elo Tech (Excel)",
            data=excel_rob,
            file_name=f"vendas_elotech_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.button("Download Elo Tech", disabled=True, use_container_width=True, help="Sem dados")
with col_ata3:
    if st.button(f"Ata Elo Tech ({len(df_ata_tech)})", use_container_width=True):
        st.session_state['ata_tipo'] = 'tech'
        st.session_state['print_secao'] = None
with col_print3:
    if st.button("Imprimir Elo Tech", use_container_width=True):
        st.session_state['print_secao'] = 'elotech'
        st.session_state['ata_tipo'] = None

if st.session_state.get('ata_tipo') == 'tech':
    components.html(gerar_ata_elotech_html(df_ata_tech), height=800, scrolling=True)

if st.session_state.get('print_secao') == 'elotech':
    _data_imp = datetime.now().strftime("%d/%m/%Y %H:%M")
    _kpis_tech = [("Total Vendido", f"{total_rob}")]
    _tabelas_tech = []
    if not df_rob.empty:
        _tabelas_tech.append(df_to_html_colorido(df_rob, "Vendas Elo Tech por Serie (Alunos Unicos) - CDR excluido", None))
    components.html(gerar_html_impressao_secao("Elo Tech", _kpis_tech, _tabelas_tech, _data_imp, filtro_unidade), height=800, scrolling=True)


# #####################################################
# VISAO GERAL
# #####################################################
st.divider()
st.header("Visão Geral")

st.markdown("### Comparativo por Segmento")
df_seg = df_estoque.groupby('segmento').agg({'enviado': 'sum', 'vendido': 'sum', 'estoque': 'sum'}).reset_index()
df_seg['_ord'] = df_seg['segmento'].map({s: i for i, s in enumerate(ORDEM_SEGMENTOS)})
df_seg = df_seg.sort_values('_ord')

fig = go.Figure()
fig.add_trace(go.Bar(name='Enviado', x=df_seg['segmento'], y=df_seg['enviado'], marker_color='#3b82f6', text=df_seg['enviado'], textposition='outside'))
fig.add_trace(go.Bar(name='Vendido', x=df_seg['segmento'], y=df_seg['vendido'], marker_color='#22c55e', text=df_seg['vendido'], textposition='outside'))
fig.add_trace(go.Bar(name='Estoque', x=df_seg['segmento'], y=df_seg['estoque'], marker_color='#f59e0b', text=df_seg['estoque'], textposition='outside'))
fig.update_layout(
    barmode='group', template='plotly_white',
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#334155'),
    legend=dict(orientation='h', y=1.1),
    margin=dict(t=60, b=20), height=400,
)
fig.update_xaxes(gridcolor='#e2e8f0')
fig.update_yaxes(gridcolor='#e2e8f0')
st.plotly_chart(fig, width="stretch")

# Acoes Gerais
st.markdown("---")
col_dl_all, col_print_all = st.columns(2)
with col_dl_all:
    sheets_all = {"Estoque SAE": df_download_estoque, "Socioemocional": df_socio_dl}
    if not df_rob_dl.empty:
        sheets_all["Elo Tech"] = df_rob_dl
    excel_all = gerar_excel(sheets_all)
    st.download_button(
        label="Download Tudo (Excel)",
        data=excel_all,
        file_name=f"estoque_completo_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_print_all:
    if st.button("Imprimir Completo", use_container_width=True):
        st.session_state['print_secao'] = 'all'
        st.session_state['ata_tipo'] = None

if st.session_state.get('print_secao') == 'all':
    _data_imp = datetime.now().strftime("%d/%m/%Y %H:%M")
    _kpis_all = [
        ("Pedido", f"{pedido_total:,}".replace(",", ".")),
        ("Enviado", f"{enviado_total:,}".replace(",", ".")),
        ("Vendido SAE", f"{vendido_total:,}".replace(",", ".")),
        ("Estoque", f"{estoque_total:,}".replace(",", ".")),
        ("% Vendido", f"{pct_venda:.1f}%"),
    ]
    _tabelas_all = []
    _df_est_p = criar_tabela_completa(df_estoque, 'estoque', filtro_unidade)
    _cols_num = [c for c in _df_est_p.columns if c != 'Descrição']
    _tabelas_all.append(df_to_html_colorido(_df_est_p, "Estoque SAE por Serie e Unidade", _cols_num))
    if not df_det.empty:
        _tabelas_all.append(df_to_html_colorido(df_det, "Detalhado (Pedido / Enviado / Vendido / Estoque)", _det_cols_estoque))
    _tabelas_all.append(df_to_html_colorido(df_socio, "Vendas Socioemocional (Alunos Unicos)", None))
    if not df_rob.empty:
        _tabelas_all.append(df_to_html_colorido(df_rob, "Vendas Elo Tech por Serie (Alunos Unicos) - CDR excluido", None))
    _df_neg_p = df_estoque[df_estoque['estoque'] < 0][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']].rename(
        columns={'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade', 'enviado': 'Enviado', 'vendido': 'Vendido', 'estoque': 'Falta'})
    if not _df_neg_p.empty:
        _tabelas_all.append(df_to_html_colorido(_df_neg_p, "Alerta: Estoque Negativo", ['Falta']))
    _df_bx_p = df_estoque[(df_estoque['estoque'] > 0) & (df_estoque['estoque'] <= 5)][['segmento', 'serie', 'unidade', 'enviado', 'vendido', 'estoque']].rename(
        columns={'segmento': 'Segmento', 'serie': 'Serie', 'unidade': 'Unidade', 'enviado': 'Enviado', 'vendido': 'Vendido', 'estoque': 'Restante'})
    if not _df_bx_p.empty:
        _tabelas_all.append(df_to_html_colorido(_df_bx_p, "Alerta: Estoque Baixo (<=5)", ['Restante']))
    components.html(gerar_html_impressao_secao("Controle de Estoque", _kpis_all, _tabelas_all, _data_imp, filtro_unidade), height=800, scrolling=True)
