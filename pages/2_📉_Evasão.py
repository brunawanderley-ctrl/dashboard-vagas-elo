"""Página de Evasão - Análise de retenção e evasão de alunos."""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from utils.theme import aplicar_tema
from utils.database import get_evasao_2025, get_evasao_2026, get_ultima_extracao
from utils.constants import ORDEM_UNIDADES, SEGMENTOS, CORES_UNIDADES, CORES_SEGMENTOS
from utils.calculations import calcular_evasao

st.set_page_config(
    page_title="Evasão - Colégio Elo",
    page_icon="📉",
    layout="wide",
)

aplicar_tema()

# Carregar e calcular dados
df_2025_raw = get_evasao_2025()
df_2026_raw = get_evasao_2026()

if df_2025_raw.empty or df_2026_raw.empty:
    st.error("Dados não disponíveis. Verifique os bancos SQLite.")
    st.stop()

df_evasao = calcular_evasao(df_2025_raw, df_2026_raw)

if df_evasao.empty:
    st.error("Não foi possível calcular a evasão.")
    st.stop()

# Sidebar
with st.sidebar:
    st.markdown("### 📉 Evasão")
    st.caption(f"Última extração: {get_ultima_extracao()}")
    st.divider()

    filtro_unidade = st.selectbox("Unidade", ["Todas"] + ORDEM_UNIDADES)
    filtro_segmento = st.selectbox("Segmento", ["Todos"] + SEGMENTOS)

    series_disponiveis = sorted(df_evasao["serie_2025"].unique().tolist())
    filtro_serie = st.selectbox("Série (2025)", ["Todas"] + series_disponiveis)

# Aplicar filtros
df = df_evasao.copy()
if filtro_unidade != "Todas":
    df = df[df["unidade"] == filtro_unidade]
if filtro_segmento != "Todos":
    df = df[df["segmento"] == filtro_segmento]
if filtro_serie != "Todas":
    df = df[df["serie_2025"] == filtro_serie]

# Header
st.markdown("# 📉 Análise de Evasão - 2025 → 2026")
st.divider()

# Separar concluintes
total_concluintes = int(df["concluintes"].sum()) if "concluintes" in df.columns else 0
df_sem_concluintes = df[df["serie_2026"] != "Formado"] if not df.empty else df

# KPIs (sem concluintes no cálculo de evasão)
total_alunos = int(df["alunos_2025"].sum())
total_alunos_rematricula = total_alunos - total_concluintes
total_veteranos = int(df_sem_concluintes["veteranos_2026"].sum())
total_evasao = int(df_sem_concluintes["evasao"].sum())
taxa_evasao = round(total_evasao / total_alunos_rematricula * 100, 1) if total_alunos_rematricula > 0 else 0
taxa_retencao = round(100 - taxa_evasao, 1)

col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    st.metric("Alunos 2025", f"{total_alunos:,}".replace(",", "."))
with col2:
    st.metric("Concluintes 3ª Série", f"{total_concluintes:,}".replace(",", "."),
              help="Alunos da 3ª Série EM que se formaram")
with col3:
    st.metric("Esperado Rematrícula", f"{total_alunos_rematricula:,}".replace(",", "."),
              help="Total 2025 menos concluintes")
with col4:
    st.metric("Veteranos 2026", f"{total_veteranos:,}".replace(",", "."))
with col5:
    st.metric("Evasão", f"{total_evasao:,}".replace(",", "."),
              delta=f"{taxa_evasao}%", delta_color="inverse")
with col6:
    st.metric("Retenção", f"{taxa_retencao}%")

st.divider()

# Gráficos
col_g1, col_g2 = st.columns(2)

# Evasão por Série (sem concluintes)
with col_g1:
    st.markdown("### Evasão por Série")

    df_serie = df_sem_concluintes.groupby("serie_2025").agg({
        "alunos_2025": "sum",
        "veteranos_2026": "sum",
        "evasao": "sum",
    }).reset_index()
    df_serie["pct_evasao"] = (df_serie["evasao"] / df_serie["alunos_2025"] * 100).round(1)

    # Ordenar
    ordem_series = [
        "Infantil II", "Infantil III", "Infantil IV", "Infantil V",
        "1º Ano", "2º Ano", "3º Ano", "4º Ano", "5º Ano",
        "6º Ano", "7º Ano", "8º Ano", "9º Ano",
        "1ª Série", "2ª Série",
    ]
    df_serie["ordem"] = df_serie["serie_2025"].apply(lambda x: ordem_series.index(x) if x in ordem_series else 99)
    df_serie = df_serie.sort_values("ordem")

    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        x=df_serie["serie_2025"],
        y=df_serie["evasao"],
        marker_color="#f87171",
        text=df_serie["pct_evasao"].apply(lambda x: f"{x}%"),
        textposition="auto",
        name="Evasão",
    ))
    fig1.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=30, b=60),
        height=400,
        font=dict(color="#334155"),
        xaxis_tickangle=-45,
    )
    fig1.update_xaxes(gridcolor="#e2e8f0")
    fig1.update_yaxes(gridcolor="#e2e8f0")
    st.plotly_chart(fig1, width="stretch")

# Evasão por Unidade (sem concluintes)
with col_g2:
    st.markdown("### Evasão por Unidade")

    df_unid = df_sem_concluintes.groupby("unidade").agg({
        "alunos_2025": "sum",
        "veteranos_2026": "sum",
        "evasao": "sum",
    }).reset_index()
    df_unid["pct_evasao"] = (df_unid["evasao"] / df_unid["alunos_2025"] * 100).round(1)
    df_unid = df_unid.sort_values("pct_evasao", ascending=False)

    cores_barras = [CORES_UNIDADES.get(u, "#667eea") for u in df_unid["unidade"]]

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=df_unid["unidade"],
        y=df_unid["pct_evasao"],
        marker_color=cores_barras,
        text=df_unid["pct_evasao"].apply(lambda x: f"{x}%"),
        textposition="auto",
        name="% Evasão",
    ))
    fig2.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=30, b=20),
        height=400,
        font=dict(color="#334155"),
        yaxis_title="% Evasao",
    )
    fig2.update_xaxes(gridcolor="#e2e8f0")
    fig2.update_yaxes(gridcolor="#e2e8f0")
    st.plotly_chart(fig2, width="stretch")

st.divider()

# Heatmap Série x Unidade
st.markdown("### Heatmap: Evasão por Série x Unidade")

df_heat = df_sem_concluintes.groupby(["unidade", "serie_2025"]).agg({"pct_evasao": "mean"}).reset_index()
if not df_heat.empty:
    pivot = df_heat.pivot_table(index="serie_2025", columns="unidade", values="pct_evasao", fill_value=0)

    # Ordenar séries
    ordem_idx = [s for s in ordem_series if s in pivot.index]
    outros = [s for s in pivot.index if s not in ordem_series]
    pivot = pivot.reindex(ordem_idx + outros)

    fig3 = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale=[[0, "#f0fdf4"], [0.5, "#fef3c7"], [1, "#fecaca"]],
        text=pivot.values.round(1),
        texttemplate="%{text}%",
        textfont=dict(size=11),
        hovertemplate="Série: %{y}<br>Unidade: %{x}<br>Evasão: %{z:.1f}%<extra></extra>",
    ))
    fig3.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=20, b=20),
        height=500,
        font=dict(color="#334155"),
    )
    st.plotly_chart(fig3, width="stretch")

st.divider()

# Resumo Concluintes por Unidade
if total_concluintes > 0:
    st.markdown("### Concluintes (3ª Série EM - 2025)")
    df_concl = df[df["serie_2026"] == "Formado"].groupby("unidade")["concluintes"].sum().reset_index()
    df_concl.columns = ["Unidade", "Concluintes"]
    cols_concl = st.columns(len(df_concl) + 1)
    for i, row in df_concl.iterrows():
        with cols_concl[i]:
            st.metric(row["Unidade"], row["Concluintes"])
    with cols_concl[len(df_concl)]:
        st.metric("TOTAL", total_concluintes)
    st.divider()

# Tabela detalhada (sem concluintes)
st.markdown("### Detalhamento Completo")

df_display = df_sem_concluintes[[
    "unidade", "serie_2025", "serie_2026", "segmento",
    "alunos_2025", "veteranos_2026", "novatos_2026", "total_2026",
    "evasao", "pct_evasao", "pct_retencao",
]].copy()

df_display.columns = [
    "Unidade", "Série 2025", "Série 2026", "Segmento",
    "Alunos 2025", "Veteranos 2026", "Novatos 2026", "Total 2026",
    "Evasão", "% Evasão", "% Retenção",
]

df_display = df_display.sort_values(["Unidade", "Segmento", "Série 2025"])

st.dataframe(
    df_display,
    width="stretch",
    hide_index=True,
    column_config={
        "% Evasão": st.column_config.NumberColumn(format="%.1f%%"),
        "% Retenção": st.column_config.NumberColumn(format="%.1f%%"),
    },
    height=500,
)

# Download CSV
csv = df_display.to_csv(index=False, sep=";", encoding="utf-8-sig")
st.download_button(
    label="Download CSV",
    data=csv,
    file_name="evasao_analise.csv",
    mime="text/csv",
)
