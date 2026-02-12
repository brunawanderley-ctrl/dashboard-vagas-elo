"""Página de Matrículas - Comparativo 2025 vs 2026 (estilo dashboard HTML)."""

import streamlit as st
import pandas as pd
from utils.theme import aplicar_tema
from utils.database import (
    get_matriculas_2025, get_matriculas_2026,
    get_integral_2025, get_integral_2026,
    get_matriculas_por_turma_2026, get_matriculas_por_turma_2025,
    get_ultima_extracao,
)
from utils.constants import (
    UNIDADES_MAP, ORDEM_UNIDADES, SEGMENTOS,
    METAS_2026, META_TOTAL, SERIES_POR_SEGMENTO,
)
from utils.calculations import (
    calc_variacao, format_variacao, format_diferenca,
    extrair_serie_ensalamento,
)

st.set_page_config(
    page_title="Matrículas - Colégio Elo",
    page_icon="📊",
    layout="wide",
)

aplicar_tema()


# --- Helpers HTML ---

def _cor_valor(val):
    """Retorna classe CSS baseada no sinal do valor."""
    if val > 0:
        return "val-pos"
    elif val < 0:
        return "val-neg"
    return ""


def _fmt_dif(val):
    """Formata diferença com sinal."""
    if val > 0:
        return f"+{val}"
    return str(val)


def _fmt_var(val):
    """Formata variação com sinal e %."""
    if val > 0:
        return f"+{val}%"
    return f"{val}%"


def _cor_meta(pct):
    """Cor da barra de progresso de meta."""
    if pct >= 95:
        return "#4ade80"
    if pct >= 85:
        return "#60a5fa"
    if pct >= 70:
        return "#fbbf24"
    return "#f87171"


# --- Sidebar ---
with st.sidebar:
    st.markdown("### 📊 Matrículas")
    st.caption(f"Última extração: {get_ultima_extracao()}")

# --- Dados ---
df_2025 = get_matriculas_2025().copy()
df_2026 = get_matriculas_2026().copy()
if not df_2025.empty:
    df_2025["unidade"] = df_2025["unidade_codigo"].map(UNIDADES_MAP)
if not df_2026.empty:
    df_2026["unidade"] = df_2026["unidade_codigo"].map(UNIDADES_MAP)

df_int_2025 = get_integral_2025()
df_int_2026 = get_integral_2026()

total_2025 = int(df_2025["matriculados"].sum()) if not df_2025.empty else 0
total_2026 = int(df_2026["matriculados"].sum()) if not df_2026.empty else 0
diferenca = total_2026 - total_2025
variacao = calc_variacao(total_2025, total_2026)
total_int_2025 = int(df_int_2025["matriculados"].sum()) if not df_int_2025.empty else 0
total_int_2026 = int(df_int_2026["matriculados"].sum()) if not df_int_2026.empty else 0
dif_int = total_int_2026 - total_int_2025
var_int = calc_variacao(total_int_2025, total_int_2026)

# ============================================================
# HEADER + KPIs
# ============================================================
st.markdown("# Dashboard Matrículas - Colégio Elo")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total 2025", f"{total_2025:,}".replace(",", "."), help="alunos matriculados")
with col2:
    st.metric("Total 2026", f"{total_2026:,}".replace(",", "."), help="alunos matriculados")
with col3:
    st.metric("Diferença", _fmt_dif(diferenca), delta=_fmt_var(variacao))
with col4:
    st.metric("Integral 2026", str(total_int_2026),
              delta=f"{_fmt_dif(dif_int)} vs 2025 ({_fmt_var(var_int)})")

st.divider()

# ============================================================
# PROGRESSO DAS METAS 2026
# ============================================================
st.markdown("### Progresso das Metas 2026")

cards_html = '<div style="display:grid; grid-template-columns: repeat(4,1fr); gap:15px; min-width:0;">'
for unidade in ORDEM_UNIDADES:
    meta = METAS_2026.get(unidade, 0)
    atual = int(df_2026[df_2026["unidade"] == unidade]["matriculados"].sum()) if not df_2026.empty else 0
    pct = round(atual / meta * 100, 1) if meta > 0 else 0
    faltam = max(0, meta - atual)
    cor = _cor_meta(pct)
    pct_bar = min(pct, 100)

    cards_html += f"""
    <div class="meta-card">
        <div class="meta-card-top">
            <span class="meta-card-name">{unidade}</span>
            <span class="meta-card-pct" style="color:{cor};">{pct}%</span>
        </div>
        <div class="meta-bar"><div class="meta-bar-fill" style="width:{pct_bar}%; background:{cor};"></div></div>
        <div class="meta-card-bottom">
            <span>{atual:,} / {meta:,} alunos</span>
            <span>Faltam {faltam} alunos</span>
        </div>
    </div>"""

cards_html += "</div>"

# Total rede
pct_total = round(total_2026 / META_TOTAL * 100, 1) if META_TOTAL > 0 else 0
faltam_total = max(0, META_TOTAL - total_2026)
cor_total = _cor_meta(pct_total)
pct_bar_total = min(pct_total, 100)

cards_html += f"""
<div class="meta-total">
    <div class="meta-total-top">
        <span class="meta-total-name">TOTAL REDE</span>
        <span class="meta-total-pct" style="color:{cor_total};">{pct_total}%</span>
    </div>
    <div class="meta-total-bar"><div class="meta-total-bar-fill" style="width:{pct_bar_total}%; background: linear-gradient(90deg, {cor_total}, {cor_total}dd);"></div></div>
    <div class="meta-total-bottom">
        <span>{total_2026:,} / {META_TOTAL:,} alunos</span>
        <span>Faltam {faltam_total} alunos</span>
    </div>
</div>"""

st.markdown(cards_html.replace(",", "."), unsafe_allow_html=True)

st.divider()

# ============================================================
# MATRÍCULAS POR UNIDADE E SEGMENTO (4 tabelas)
# ============================================================
st.markdown("### Matrículas por Unidade e Segmento")

tables_html = '<div style="display:grid; grid-template-columns: repeat(2,1fr); gap:20px;">'

for unidade in ORDEM_UNIDADES:
    rows = ""
    total_u25 = 0
    total_u26 = 0
    for seg in SEGMENTOS:
        v25 = int(df_2025[(df_2025["unidade"] == unidade) & (df_2025["segmento"] == seg)]["matriculados"].sum()) if not df_2025.empty else 0
        v26 = int(df_2026[(df_2026["unidade"] == unidade) & (df_2026["segmento"] == seg)]["matriculados"].sum()) if not df_2026.empty else 0
        dif = v26 - v25
        var = calc_variacao(v25, v26)
        total_u25 += v25
        total_u26 += v26
        rows += f"""<tr>
            <td>{seg}</td>
            <td class="num">{v25}</td>
            <td class="num">{v26}</td>
            <td class="num {_cor_valor(dif)}">{_fmt_dif(dif)}</td>
            <td class="num {_cor_valor(var)}">{_fmt_var(var)}</td>
        </tr>"""

    dif_u = total_u26 - total_u25
    var_u = calc_variacao(total_u25, total_u26)
    rows += f"""<tr class="total-row">
        <td>TOTAL</td>
        <td class="num">{total_u25:,}</td>
        <td class="num">{total_u26:,}</td>
        <td class="num {_cor_valor(dif_u)}">{_fmt_dif(dif_u)}</td>
        <td class="num {_cor_valor(var_u)}">{_fmt_var(var_u)}</td>
    </tr>"""

    tables_html += f"""
    <div class="dark-table-container">
        <div class="dark-table-header">{unidade}</div>
        <table class="dark-table">
            <thead><tr>
                <th>Segmento</th><th class="num">2025</th><th class="num">2026</th>
                <th class="num">Dif.</th><th class="num">Var.</th>
            </tr></thead>
            <tbody>{rows}</tbody>
        </table>
    </div>"""

tables_html += "</div>"
st.markdown(tables_html.replace(",", "."), unsafe_allow_html=True)

st.divider()

# ============================================================
# RESUMO GERAL (2 tabelas lado a lado)
# ============================================================
st.markdown("### Resumo Geral")

resumo_html = '<div style="display:grid; grid-template-columns: repeat(2,1fr); gap:20px;">'

# Por Unidade (com Metas)
rows_u = ""
for unidade in ORDEM_UNIDADES:
    v25 = int(df_2025[df_2025["unidade"] == unidade]["matriculados"].sum()) if not df_2025.empty else 0
    v26 = int(df_2026[df_2026["unidade"] == unidade]["matriculados"].sum()) if not df_2026.empty else 0
    dif = v26 - v25
    var = calc_variacao(v25, v26)
    meta = METAS_2026.get(unidade, 0)
    pct_meta = round(v26 / meta * 100, 1) if meta > 0 else 0
    cor_m = _cor_meta(pct_meta)
    rows_u += f"""<tr>
        <td>{unidade}</td>
        <td class="num">{v25:,}</td><td class="num">{v26:,}</td>
        <td class="num {_cor_valor(dif)}">{_fmt_dif(dif)}</td>
        <td class="num {_cor_valor(var)}">{_fmt_var(var)}</td>
        <td class="num">{meta:,}</td>
        <td class="num" style="color:{cor_m}">{pct_meta}%</td>
    </tr>"""

var_total = calc_variacao(total_2025, total_2026)
pct_meta_total = round(total_2026 / META_TOTAL * 100, 1) if META_TOTAL > 0 else 0
cor_mt = _cor_meta(pct_meta_total)
rows_u += f"""<tr class="total-row">
    <td>TOTAL REDE</td>
    <td class="num">{total_2025:,}</td><td class="num">{total_2026:,}</td>
    <td class="num {_cor_valor(diferenca)}">{_fmt_dif(diferenca)}</td>
    <td class="num {_cor_valor(var_total)}">{_fmt_var(var_total)}</td>
    <td class="num">{META_TOTAL:,}</td>
    <td class="num" style="color:{cor_mt}">{pct_meta_total}%</td>
</tr>"""

resumo_html += f"""
<div class="dark-table-container">
    <div class="dark-table-header">Por Unidade (com Metas 2026)</div>
    <table class="dark-table">
        <thead><tr>
            <th>Unidade</th><th class="num">2025</th><th class="num">2026</th>
            <th class="num">Dif.</th><th class="num">Var.</th>
            <th class="num">Meta</th><th class="num">% Alcançado</th>
        </tr></thead>
        <tbody>{rows_u}</tbody>
    </table>
</div>"""

# Por Segmento
rows_s = ""
for seg in SEGMENTOS:
    v25 = int(df_2025[df_2025["segmento"] == seg]["matriculados"].sum()) if not df_2025.empty else 0
    v26 = int(df_2026[df_2026["segmento"] == seg]["matriculados"].sum()) if not df_2026.empty else 0
    dif = v26 - v25
    var = calc_variacao(v25, v26)
    rows_s += f"""<tr>
        <td>{seg}</td>
        <td class="num">{v25:,}</td><td class="num">{v26:,}</td>
        <td class="num {_cor_valor(dif)}">{_fmt_dif(dif)}</td>
        <td class="num {_cor_valor(var)}">{_fmt_var(var)}</td>
    </tr>"""

rows_s += f"""<tr class="total-row">
    <td>TOTAL</td>
    <td class="num">{total_2025:,}</td><td class="num">{total_2026:,}</td>
    <td class="num {_cor_valor(diferenca)}">{_fmt_dif(diferenca)}</td>
    <td class="num {_cor_valor(var_total)}">{_fmt_var(var_total)}</td>
</tr>"""

resumo_html += f"""
<div class="dark-table-container">
    <div class="dark-table-header">Por Segmento (Toda a Rede)</div>
    <table class="dark-table">
        <thead><tr>
            <th>Segmento</th><th class="num">2025</th><th class="num">2026</th>
            <th class="num">Dif.</th><th class="num">Var.</th>
        </tr></thead>
        <tbody>{rows_s}</tbody>
    </table>
</div>"""

resumo_html += "</div>"
st.markdown(resumo_html.replace(",", "."), unsafe_allow_html=True)

st.divider()

# ============================================================
# MATRICULADOS POR SÉRIE - 2026
# ============================================================
st.markdown("### Matriculados por Série - 2026")

df_turmas_26 = get_matriculas_por_turma_2026()
if not df_turmas_26.empty:
    df_turmas_26["unidade"] = df_turmas_26["unidade_codigo"].map(UNIDADES_MAP)
    df_turmas_26["serie"] = df_turmas_26["turma"].apply(extrair_serie_ensalamento)
    df_turmas_26 = df_turmas_26[df_turmas_26["matriculados"] > 0]

    # Agregar por série e unidade
    pivot = df_turmas_26.groupby(["serie", "unidade"])["matriculados"].sum().reset_index()

    serie_rows = ""
    grand_totals = {u: 0 for u in ORDEM_UNIDADES}
    grand_total = 0

    for seg, series in SERIES_POR_SEGMENTO.items():
        # Header do segmento
        serie_rows += f'<tr class="segment-header"><td colspan="{len(ORDEM_UNIDADES) + 2}">{seg.upper()}</td></tr>'

        seg_totals = {u: 0 for u in ORDEM_UNIDADES}
        seg_total = 0

        for serie in series:
            row_vals = []
            row_total = 0
            for unidade in ORDEM_UNIDADES:
                val = int(pivot[(pivot["serie"] == serie) & (pivot["unidade"] == unidade)]["matriculados"].sum())
                row_vals.append(val)
                row_total += val
                seg_totals[unidade] += val
                grand_totals[unidade] += val

            seg_total += row_total
            grand_total += row_total

            cells = "".join(f'<td class="num">{v}</td>' for v in row_vals)
            serie_rows += f'<tr><td>{serie}</td>{cells}<td class="num" style="font-weight:600;">{row_total}</td></tr>'

        # Subtotal segmento
        cells_sub = "".join(f'<td class="num">{seg_totals[u]}</td>' for u in ORDEM_UNIDADES)
        serie_rows += f'<tr class="subtotal-row"><td>Subtotal {seg.upper()}</td>{cells_sub}<td class="num">{seg_total}</td></tr>'

    # Total geral
    cells_grand = "".join(f'<td class="num">{grand_totals[u]}</td>' for u in ORDEM_UNIDADES)
    serie_rows += f'<tr class="total-row"><td>TOTAL GERAL</td>{cells_grand}<td class="num">{grand_total}</td></tr>'

    headers = "".join(f"<th class='num'>{u.split()[0] if ' ' in u else u}</th>" for u in ORDEM_UNIDADES)
    serie_html = f"""
    <div class="dark-table-container">
        <table class="dark-table">
            <thead><tr><th>Série</th>{headers}<th class="num">Total</th></tr></thead>
            <tbody>{serie_rows}</tbody>
        </table>
    </div>"""
    st.markdown(serie_html.replace(",", "."), unsafe_allow_html=True)

st.divider()

# ============================================================
# COMPARATIVO NOVATOS/VETERANOS 2025 vs 2026
# ============================================================
st.markdown("### Comparativo Novatos/Veteranos 2025 vs 2026")

tipo_exibir = st.radio("Exibir:", ["Novatos", "Veteranos", "Ambos"], horizontal=True)

df_turmas_25 = get_matriculas_por_turma_2025()
if not df_turmas_25.empty and not df_turmas_26.empty:
    df_t25 = df_turmas_25.copy()
    df_t25["unidade"] = df_t25["unidade_codigo"].map(UNIDADES_MAP)

    # Agregar por unidade e segmento
    campo = "novatos" if tipo_exibir == "Novatos" else "veteranos"

    if tipo_exibir == "Ambos":
        agg_25 = df_t25.groupby(["unidade", "segmento"]).agg({"novatos": "sum", "veteranos": "sum"}).reset_index()
        agg_26 = df_turmas_26.groupby(["unidade", "segmento"]).agg({"novatos": "sum", "veteranos": "sum"}).reset_index()
    else:
        agg_25 = df_t25.groupby(["unidade", "segmento"])[campo].sum().reset_index()
        agg_26 = df_turmas_26.groupby(["unidade", "segmento"])[campo].sum().reset_index()

    # Construir tabela
    if tipo_exibir == "Ambos":
        # Headers com sub-colunas: Segmento → 2025 Nov | 2025 Vet | 2026 Nov | 2026 Vet
        headers_comp = "<th>Unidade</th>"
        for seg in SEGMENTOS:
            headers_comp += f'<th class="num" colspan="2">{seg}</th>'
        headers_comp += '<th class="num" colspan="2">Total</th>'

        sub_headers = "<th></th>"
        for _ in SEGMENTOS:
            sub_headers += '<th class="num">Nov.</th><th class="num">Vet.</th>'
        sub_headers += '<th class="num">Nov.</th><th class="num">Vet.</th>'

        rows_comp = ""
        total_nov25 = 0
        total_vet25 = 0
        total_nov26 = 0
        total_vet26 = 0
        grand_nov = {seg: [0, 0] for seg in SEGMENTOS}

        for unidade in ORDEM_UNIDADES:
            cells = f"<td>{unidade}</td>"
            u_nov_total = 0
            u_vet_total = 0
            for seg in SEGMENTOS:
                n26 = int(agg_26[(agg_26["unidade"] == unidade) & (agg_26["segmento"] == seg)]["novatos"].sum())
                v26 = int(agg_26[(agg_26["unidade"] == unidade) & (agg_26["segmento"] == seg)]["veteranos"].sum())
                u_nov_total += n26
                u_vet_total += v26
                grand_nov[seg][0] += n26
                grand_nov[seg][1] += v26
                cells += f'<td class="num">{n26}</td><td class="num">{v26}</td>'
            total_nov26 += u_nov_total
            total_vet26 += u_vet_total
            cells += f'<td class="num" style="font-weight:600;">{u_nov_total}</td><td class="num" style="font-weight:600;">{u_vet_total}</td>'
            rows_comp += f"<tr>{cells}</tr>"

        # Total
        cells_t = "<td>TOTAL</td>"
        for seg in SEGMENTOS:
            cells_t += f'<td class="num">{grand_nov[seg][0]}</td><td class="num">{grand_nov[seg][1]}</td>'
        cells_t += f'<td class="num">{total_nov26}</td><td class="num">{total_vet26}</td>'
        rows_comp += f'<tr class="total-row">{cells_t}</tr>'

        comp_html = f"""
        <div class="dark-table-container">
            <table class="dark-table">
                <thead>
                    <tr>{headers_comp}</tr>
                    <tr>{sub_headers}</tr>
                </thead>
                <tbody>{rows_comp}</tbody>
            </table>
        </div>"""
    else:
        # Tabela simples: Unidade | Seg1 2025/2026/Dif | Seg2 ... | Total
        headers_comp = "<th>Unidade</th>"
        for seg in SEGMENTOS:
            headers_comp += f'<th class="num" colspan="3">{seg}</th>'
        headers_comp += '<th class="num" colspan="3">Total</th>'

        sub_headers = "<th></th>"
        for _ in SEGMENTOS:
            sub_headers += '<th class="num">2025</th><th class="num">2026</th><th class="num">Dif</th>'
        sub_headers += '<th class="num">2025</th><th class="num">2026</th><th class="num">Dif</th>'

        rows_comp = ""
        grand = {seg: [0, 0] for seg in SEGMENTOS}

        for unidade in ORDEM_UNIDADES:
            cells = f"<td>{unidade}</td>"
            u_total_25 = 0
            u_total_26 = 0
            for seg in SEGMENTOS:
                v25 = int(agg_25[(agg_25["unidade"] == unidade) & (agg_25["segmento"] == seg)][campo].sum())
                v26 = int(agg_26[(agg_26["unidade"] == unidade) & (agg_26["segmento"] == seg)][campo].sum())
                dif = v26 - v25
                u_total_25 += v25
                u_total_26 += v26
                grand[seg][0] += v25
                grand[seg][1] += v26
                cells += f'<td class="num">{v25}</td><td class="num">{v26}</td><td class="num {_cor_valor(dif)}">{_fmt_dif(dif)}</td>'

            dif_u = u_total_26 - u_total_25
            cells += f'<td class="num">{u_total_25}</td><td class="num">{u_total_26}</td><td class="num {_cor_valor(dif_u)}">{_fmt_dif(dif_u)}</td>'
            rows_comp += f"<tr>{cells}</tr>"

        # Total
        cells_t = "<td>TOTAL</td>"
        gt25 = 0
        gt26 = 0
        for seg in SEGMENTOS:
            d = grand[seg][1] - grand[seg][0]
            gt25 += grand[seg][0]
            gt26 += grand[seg][1]
            cells_t += f'<td class="num">{grand[seg][0]}</td><td class="num">{grand[seg][1]}</td><td class="num {_cor_valor(d)}">{_fmt_dif(d)}</td>'
        dt = gt26 - gt25
        cells_t += f'<td class="num">{gt25}</td><td class="num">{gt26}</td><td class="num {_cor_valor(dt)}">{_fmt_dif(dt)}</td>'
        rows_comp += f'<tr class="total-row">{cells_t}</tr>'

        comp_html = f"""
        <div class="dark-table-container">
            <table class="dark-table">
                <thead>
                    <tr>{headers_comp}</tr>
                    <tr>{sub_headers}</tr>
                </thead>
                <tbody>{rows_comp}</tbody>
            </table>
        </div>"""

    st.markdown(comp_html, unsafe_allow_html=True)

st.divider()

# ============================================================
# INTEGRAL / COMPLEMENTAR
# ============================================================
st.markdown("### Integral/Complementar - Comparativo 2025 vs 2026")

int_cards = '<div style="display:grid; grid-template-columns: repeat(4,1fr); gap:15px; margin-bottom:15px;">'
for unidade in ORDEM_UNIDADES:
    codigos = [k for k, v in UNIDADES_MAP.items() if v == unidade]
    v25 = int(df_int_2025[df_int_2025["unidade_codigo"].isin(codigos)]["matriculados"].sum()) if not df_int_2025.empty else 0
    v26 = int(df_int_2026[df_int_2026["unidade_codigo"].isin(codigos)]["matriculados"].sum()) if not df_int_2026.empty else 0
    var = calc_variacao(v25, v26)
    cor_v = "val-pos" if var >= 0 else "val-neg"

    int_cards += f"""
    <div class="meta-card" style="text-align:center;">
        <div style="font-size:0.85rem; color:#64748b; margin-bottom:8px;">{unidade}</div>
        <div style="display:flex; justify-content:center; gap:20px; font-size:1.1rem; color:#1e293b;">
            <div><span style="font-size:0.7rem; color:#94a3b8;">2025</span><br>{v25}</div>
            <div><span style="font-size:0.7rem; color:#94a3b8;">2026</span><br>{v26}</div>
            <div><span style="font-size:0.7rem; color:#94a3b8;">Var.</span><br><span class="{cor_v}">{_fmt_var(var)}</span></div>
        </div>
    </div>"""

int_cards += "</div>"

# Total integral
var_int_total = calc_variacao(total_int_2025, total_int_2026)
cor_int = "val-pos" if dif_int >= 0 else "val-neg"
int_cards += f"""
<div style="text-align:center; padding:15px; background:rgba(220,38,38,0.06); border-radius:10px; border:1px solid #e2e8f0; color:#1e293b;">
    <strong>TOTAL INTEGRAL:</strong> {total_int_2025} (2025) &rarr; {total_int_2026} (2026) =
    <span class="{cor_int}">{_fmt_dif(dif_int)} alunos ({_fmt_var(var_int_total)})</span>
</div>"""

st.markdown(int_cards, unsafe_allow_html=True)

st.divider()
st.caption("Dashboard Matrículas - Colégio Elo | Dados extraídos do SIGA Activesoft")
