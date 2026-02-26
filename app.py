"""Dashboard Vagas - Colégio Elo - Página Inicial."""

import streamlit as st
import subprocess
import os
import sys
import pandas as pd
import plotly.graph_objects as go
from utils.theme import aplicar_tema, kpi_card
from utils.database import (
    get_matriculas_2025, get_matriculas_2026,
    get_integral_2025, get_integral_2026,
    get_evasao_2025, get_evasao_2026,
    get_turmas_detalhadas_2026,
    get_ultima_extracao,
)
from utils.constants import UNIDADES_MAP, ORDEM_UNIDADES, SEGMENTOS, CORES_UNIDADES, CORES_SEGMENTOS
from utils.calculations import (
    calc_variacao, format_variacao, format_diferenca,
    calcular_evasao, calcular_status_ocupacao,
)

st.set_page_config(
    page_title="Dashboard Vagas - Colégio Elo",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_tema()

# Sidebar
with st.sidebar:
    st.markdown("### 🎓 Colégio Elo")
    st.markdown("**Dashboard de Vagas**")
    st.divider()
    ultima = get_ultima_extracao()
    st.caption(f"Última extração: {ultima}")
    st.divider()
    st.markdown("#### Navegação")
    st.page_link("app.py", label="🏠 Home", icon="🏠")
    st.page_link("pages/1_📊_Matrículas.py", label="📊 Matrículas")
    st.page_link("pages/2_📉_Evasão.py", label="📉 Evasão")
    st.page_link("pages/3_🏫_Ensalamento.py", label="🏫 Ensalamento")
    st.page_link("pages/4_📦_Estoque.py", label="📦 Estoque SAE")
    st.divider()

    # Botão de atualização de dados
    if st.button("🔄 Atualizar Dados do SIGA", use_container_width=True):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        venv_python = sys.executable

        with st.status("Extraindo dados do SIGA...", expanded=True) as status:
            erros = []

            st.write("Extraindo vagas 2026...")
            r1 = subprocess.run(
                [venv_python, "extrair_vagas_otimizado.py", "--periodo", "2026"],
                capture_output=True, text=True, timeout=600, cwd=script_dir,
            )
            if r1.returncode != 0:
                erros.append(f"Vagas 2026: {r1.stderr[:200]}")
            else:
                st.write("Vagas 2026 OK")

            st.write("Extraindo integral 2026...")
            r2 = subprocess.run(
                [venv_python, "extrair_integral.py", "--periodo", "2026"],
                capture_output=True, text=True, timeout=600, cwd=script_dir,
            )
            if r2.returncode != 0:
                erros.append(f"Integral: {r2.stderr[:200]}")
            else:
                st.write("Integral 2026 OK")

            if erros:
                status.update(label="Extração com erros", state="error")
                for e in erros:
                    st.error(e)
            else:
                status.update(label="Dados atualizados!", state="complete")
                st.cache_data.clear()
                st.rerun()

# Header
st.markdown("# 🎓 Dashboard Vagas - Colégio Elo")
st.markdown("Acompanhamento de matrículas, evasão e ensalamento em tempo real.")
st.divider()

# Carregar dados
df_2025 = get_matriculas_2025()
df_2026 = get_matriculas_2026()
df_int_2025 = get_integral_2025()
df_int_2026 = get_integral_2026()

# Aplicar nomes de unidade
for df in [df_2025, df_2026]:
    if not df.empty:
        df["unidade"] = df["unidade_codigo"].map(UNIDADES_MAP)

# Calcular totais
total_2025 = int(df_2025["matriculados"].sum()) if not df_2025.empty else 0
total_2026 = int(df_2026["matriculados"].sum()) if not df_2026.empty else 0
total_int_2025 = int(df_int_2025["matriculados"].sum()) if not df_int_2025.empty else 0
total_int_2026 = int(df_int_2026["matriculados"].sum()) if not df_int_2026.empty else 0

diferenca = total_2026 - total_2025
variacao = calc_variacao(total_2025, total_2026)

# KPIs
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total 2025", f"{total_2025:,}".replace(",", "."), help="Alunos matriculados em 2025")
with col2:
    st.metric("Total 2026", f"{total_2026:,}".replace(",", "."), help="Alunos matriculados em 2026")
with col3:
    delta_color = "normal" if diferenca >= 0 else "inverse"
    st.metric("Diferença", format_diferenca(diferenca), delta=format_variacao(variacao), delta_color=delta_color)
with col4:
    dif_int = total_int_2026 - total_int_2025
    st.metric("Integral 2026", f"{total_int_2026:,}".replace(",", "."), delta=format_diferenca(dif_int))

st.divider()

# --- GRÁFICOS RESUMO ---

col_g1, col_g2 = st.columns(2)

# Gráfico: Matrículas por Unidade (2025 vs 2026)
with col_g1:
    st.markdown("### Matrículas por Unidade")
    dados_unid = []
    for unidade in ORDEM_UNIDADES:
        v2025 = int(df_2025[df_2025["unidade"] == unidade]["matriculados"].sum()) if not df_2025.empty else 0
        v2026 = int(df_2026[df_2026["unidade"] == unidade]["matriculados"].sum()) if not df_2026.empty else 0
        dados_unid.append({"unidade": unidade, "2025": v2025, "2026": v2026})

    df_g = pd.DataFrame(dados_unid)
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        name="2025", x=df_g["unidade"], y=df_g["2025"],
        marker_color="rgba(102, 126, 234, 0.4)",
        text=df_g["2025"], textposition="auto",
    ))
    fig1.add_trace(go.Bar(
        name="2026", x=df_g["unidade"], y=df_g["2026"],
        marker_color="#667eea",
        text=df_g["2026"], textposition="auto",
    ))
    fig1.update_layout(
        barmode="group",
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=10, b=20),
        height=350,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(color="#334155"),
    )
    fig1.update_xaxes(gridcolor="#e2e8f0")
    fig1.update_yaxes(gridcolor="#e2e8f0")
    st.plotly_chart(fig1, width="stretch")

# Gráfico: Matrículas por Segmento (2025 vs 2026)
with col_g2:
    st.markdown("### Matrículas por Segmento")
    dados_seg = []
    for seg in SEGMENTOS:
        v2025 = int(df_2025[df_2025["segmento"] == seg]["matriculados"].sum()) if not df_2025.empty else 0
        v2026 = int(df_2026[df_2026["segmento"] == seg]["matriculados"].sum()) if not df_2026.empty else 0
        dados_seg.append({"segmento": seg, "2025": v2025, "2026": v2026})

    df_s = pd.DataFrame(dados_seg)
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        name="2025", x=df_s["segmento"], y=df_s["2025"],
        marker_color="rgba(118, 75, 162, 0.4)",
        text=df_s["2025"], textposition="auto",
    ))
    fig2.add_trace(go.Bar(
        name="2026", x=df_s["segmento"], y=df_s["2026"],
        marker_color="#764ba2",
        text=df_s["2026"], textposition="auto",
    ))
    fig2.update_layout(
        barmode="group",
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=10, b=20),
        height=350,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(color="#334155"),
    )
    fig2.update_xaxes(gridcolor="#e2e8f0")
    fig2.update_yaxes(gridcolor="#e2e8f0")
    st.plotly_chart(fig2, width="stretch")

st.divider()

# Segunda linha de gráficos
col_g3, col_g4 = st.columns(2)

# Gráfico: Distribuição por Segmento 2026 (donut)
with col_g3:
    st.markdown("### Distribuição 2026 por Segmento")
    seg_vals = []
    seg_labels = []
    seg_cores = []
    for seg in SEGMENTOS:
        v = int(df_2026[df_2026["segmento"] == seg]["matriculados"].sum()) if not df_2026.empty else 0
        if v > 0:
            seg_vals.append(v)
            seg_labels.append(seg)
            seg_cores.append(CORES_SEGMENTOS.get(seg, "#667eea"))

    fig3 = go.Figure(data=[go.Pie(
        labels=seg_labels,
        values=seg_vals,
        hole=0.5,
        marker_colors=seg_cores,
        textinfo="label+percent",
        textfont=dict(size=12, color="#334155"),
        hovertemplate="%{label}: %{value:,} alunos (%{percent})<extra></extra>",
    )])
    fig3.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=10, b=20),
        height=350,
        font=dict(color="#334155"),
        showlegend=False,
        annotations=[dict(text=f"<b>{total_2026:,}</b>".replace(",", "."), x=0.5, y=0.5,
                          font_size=22, font_color="#1e293b", showarrow=False)],
    )
    st.plotly_chart(fig3, width="stretch")

# Gráfico: Evasão resumo por unidade
with col_g4:
    st.markdown("### Evasão por Unidade")
    df_ev_2025 = get_evasao_2025()
    df_ev_2026 = get_evasao_2026()
    if not df_ev_2025.empty and not df_ev_2026.empty:
        df_evasao = calcular_evasao(df_ev_2025, df_ev_2026)
        if not df_evasao.empty:
            ev_unid = df_evasao.groupby("unidade").agg({
                "alunos_2025": "sum", "evasao": "sum",
            }).reset_index()
            ev_unid["pct_evasao"] = (ev_unid["evasao"] / ev_unid["alunos_2025"] * 100).round(1)
            ev_unid["pct_retencao"] = (100 - ev_unid["pct_evasao"]).round(1)

            # Ordenar como ORDEM_UNIDADES
            ev_unid["_ord"] = ev_unid["unidade"].apply(lambda x: ORDEM_UNIDADES.index(x) if x in ORDEM_UNIDADES else 99)
            ev_unid = ev_unid.sort_values("_ord")

            fig4 = go.Figure()
            fig4.add_trace(go.Bar(
                name="Retenção",
                x=ev_unid["unidade"], y=ev_unid["pct_retencao"],
                marker_color="#4ade80",
                text=ev_unid["pct_retencao"].apply(lambda x: f"{x}%"),
                textposition="auto",
            ))
            fig4.add_trace(go.Bar(
                name="Evasão",
                x=ev_unid["unidade"], y=ev_unid["pct_evasao"],
                marker_color="#f87171",
                text=ev_unid["pct_evasao"].apply(lambda x: f"{x}%"),
                textposition="auto",
            ))
            fig4.update_layout(
                barmode="stack",
                template="plotly_white",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                margin=dict(l=20, r=20, t=10, b=20),
                height=350,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                font=dict(color="#334155"),
                yaxis=dict(range=[0, 105]),
            )
            fig4.update_xaxes(gridcolor="#e2e8f0")
            fig4.update_yaxes(gridcolor="#e2e8f0")
            st.plotly_chart(fig4, width="stretch")
        else:
            st.info("Dados de evasão indisponíveis.")
    else:
        st.info("Dados de evasão indisponíveis.")

st.divider()

# Resumo por unidade (cards)
st.markdown("### Resumo por Unidade")

cols = st.columns(4)
for i, unidade in enumerate(ORDEM_UNIDADES):
    with cols[i]:
        v2025 = int(df_2025[df_2025["unidade"] == unidade]["matriculados"].sum()) if not df_2025.empty else 0
        v2026 = int(df_2026[df_2026["unidade"] == unidade]["matriculados"].sum()) if not df_2026.empty else 0
        dif = v2026 - v2025
        var = calc_variacao(v2025, v2026)
        delta_color = "normal" if dif >= 0 else "inverse"
        st.metric(unidade, f"{v2026:,}".replace(",", "."), delta=f"{format_diferenca(dif)} ({format_variacao(var)})", delta_color=delta_color)

st.divider()

# Ocupação resumo (gauge)
st.markdown("### Ocupação Geral (Ensalamento)")
df_turmas = get_turmas_detalhadas_2026()
if not df_turmas.empty:
    df_turmas_valid = df_turmas[df_turmas["matriculados"] > 0]
    total_m = int(df_turmas_valid["matriculados"].sum())
    total_v = int(df_turmas_valid["vagas"].sum())
    ocup_geral = round(total_m / total_v * 100, 1) if total_v > 0 else 0

    col_oc1, col_oc2, col_oc3 = st.columns([2, 1, 1])
    with col_oc1:
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=ocup_geral,
            number={"suffix": "%", "font": {"size": 40, "color": "#1e293b"}},
            gauge={
                "axis": {"range": [0, 120], "tickcolor": "#94a3b8"},
                "bar": {"color": "#667eea"},
                "bgcolor": "#f1f5f9",
                "steps": [
                    {"range": [0, 40], "color": "rgba(220,38,38,0.1)"},
                    {"range": [40, 70], "color": "rgba(202,138,4,0.1)"},
                    {"range": [70, 90], "color": "rgba(22,163,74,0.1)"},
                    {"range": [90, 120], "color": "rgba(37,99,235,0.1)"},
                ],
                "threshold": {
                    "line": {"color": "#1e293b", "width": 2},
                    "thickness": 0.75,
                    "value": ocup_geral,
                },
            },
        ))
        fig_gauge.update_layout(
            template="plotly_white",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=30, r=30, t=30, b=10),
            height=250,
            font=dict(color="#334155"),
        )
        st.plotly_chart(fig_gauge, width="stretch")

    with col_oc2:
        total_disp = int(df_turmas_valid["disponiveis"].apply(lambda x: max(0, x)).sum())
        total_pre = int(df_turmas_valid["pre_matriculados"].sum())
        st.metric("Vagas Disponíveis", f"{total_disp:,}".replace(",", "."))
        st.metric("Pré-Matriculados", f"{total_pre:,}".replace(",", "."))

    with col_oc3:
        # Contar turmas por status
        df_turmas_valid = df_turmas_valid.copy()
        df_turmas_valid["status"] = df_turmas_valid.apply(
            lambda r: calcular_status_ocupacao(r["matriculados"], r["vagas"])[0], axis=1
        )
        criticas = int((df_turmas_valid["status"].isin(["critica", "atencao"])).sum())
        lotadas = int((df_turmas_valid["status"].isin(["lotada", "super"])).sum())
        st.metric("Turmas Atenção", criticas, help="Críticas + Atenção (<70%)")
        st.metric("Turmas Lotadas", lotadas, help="Lotadas + Superlotadas (>90%)")

st.divider()

# Links de navegação
st.markdown("### Páginas do Dashboard")
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown("""
    #### 📊 Matrículas
    Comparativo 2025 vs 2026 por unidade e segmento.
    """)
    st.page_link("pages/1_📊_Matrículas.py", label="Abrir Matrículas →")

with col2:
    st.markdown("""
    #### 📉 Evasão
    Evasão e retenção com progressão de séries.
    """)
    st.page_link("pages/2_📉_Evasão.py", label="Abrir Evasão →")

with col3:
    st.markdown("""
    #### 🏫 Ensalamento
    Ocupação das turmas e vagas disponíveis.
    """)
    st.page_link("pages/3_🏫_Ensalamento.py", label="Abrir Ensalamento →")

with col4:
    st.markdown("""
    #### 📦 Estoque SAE
    Estoque de livros, vendas e balanço físico.
    """)
    st.page_link("pages/4_📦_Estoque.py", label="Abrir Estoque →")

st.divider()
st.caption("Dashboard Vagas - Colégio Elo | Dados extraídos do SIGA Activesoft")
