"""Página de Ensalamento - Ocupação das turmas."""

import streamlit as st
import pandas as pd
from utils.theme import aplicar_tema
from utils.database import get_turmas_detalhadas_2026, get_ultima_extracao
from utils.constants import (
    UNIDADES_MAP, ORDEM_UNIDADES, SEGMENTOS,
    ORDEM_SERIES, STATUS_OCUPACAO,
)
from utils.calculations import (
    extrair_serie_ensalamento, extrair_turno, extrair_letra_turma,
    calcular_status_ocupacao,
)

st.set_page_config(
    page_title="Ensalamento - Colégio Elo",
    page_icon="🏫",
    layout="wide",
)

aplicar_tema()

# Carregar dados
df_raw = get_turmas_detalhadas_2026()

if df_raw.empty:
    st.error("Dados não disponíveis. Verifique o banco SQLite.")
    st.stop()

# Processar turmas
df = df_raw.copy()
df["unidade"] = df["unidade_codigo"].map(UNIDADES_MAP)
df["serie"] = df["turma"].apply(extrair_serie_ensalamento)
df["turno"] = df["turma"].apply(extrair_turno)
df["letra_turma"] = df["turma"].apply(extrair_letra_turma)

# Calcular ocupação e status
def _calc_status(row):
    status, ocupacao = calcular_status_ocupacao(row["matriculados"], row["vagas"])
    return pd.Series({"status": status, "ocupacao": round(ocupacao, 1)})

df[["status", "ocupacao"]] = df.apply(_calc_status, axis=1)

# Filtrar turmas com matriculados > 0
df = df[df["matriculados"] > 0].copy()

# Sidebar
with st.sidebar:
    st.markdown("### 🏫 Ensalamento")
    st.caption(f"Última extração: {get_ultima_extracao()}")
    st.divider()

    unidades_disponiveis = ["Todas"] + [u for u in ORDEM_UNIDADES if u in df["unidade"].values]
    filtro_unidade = st.selectbox("Unidade", unidades_disponiveis)

    segmentos_disponiveis = ["Todos"] + [s for s in SEGMENTOS if s in df["segmento"].values]
    filtro_segmento = st.selectbox("Segmento", segmentos_disponiveis)

    series_disponiveis = ["Todas"] + sorted(
        [s for s in df["serie"].unique() if s != "Outros"],
        key=lambda x: ORDEM_SERIES.index(x) if x in ORDEM_SERIES else 99
    )
    filtro_serie = st.selectbox("Série", series_disponiveis)

    status_opcoes = {
        "Todos": "",
        "Crítica (<40%)": "critica",
        "Atenção (40-70%)": "atencao",
        "Normal (70-90%)": "normal",
        "Lotada (≥90%)": "lotada",
        "Superlotada (>100%)": "super",
    }
    filtro_status = st.selectbox("Status", list(status_opcoes.keys()))

    filtro_busca = st.text_input("Buscar turma", placeholder="Digite para buscar...")

# Aplicar filtros
df_filtrado = df.copy()

if filtro_unidade != "Todas":
    df_filtrado = df_filtrado[df_filtrado["unidade"] == filtro_unidade]
if filtro_segmento != "Todos":
    df_filtrado = df_filtrado[df_filtrado["segmento"] == filtro_segmento]
if filtro_serie != "Todas":
    df_filtrado = df_filtrado[df_filtrado["serie"] == filtro_serie]
if status_opcoes[filtro_status]:
    df_filtrado = df_filtrado[df_filtrado["status"] == status_opcoes[filtro_status]]
if filtro_busca:
    mask = df_filtrado["turma"].str.lower().str.contains(filtro_busca.lower(), na=False)
    df_filtrado = df_filtrado[mask]

# Header
st.markdown("# 🏫 Dashboard Ensalamento 2026")
st.divider()

# KPIs
total_matric = int(df_filtrado["matriculados"].sum())
total_vagas = int(df_filtrado["vagas"].sum())
total_disp = int(df_filtrado["disponiveis"].apply(lambda x: max(0, x)).sum())
total_pre = int(df_filtrado["pre_matriculados"].sum())
total_novatos = int(df_filtrado["novatos"].sum())
total_veteranos = int(df_filtrado["veteranos"].sum())
taxa_ocup = round(total_matric / total_vagas * 100, 1) if total_vagas > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Matriculados", f"{total_matric:,}".replace(",", "."), help=f"{len(df_filtrado)} turmas")
with col2:
    st.metric("Ocupação", f"{taxa_ocup}%", help=f"{total_vagas:,} vagas totais".replace(",", "."))
with col3:
    st.metric("Vagas Disponíveis", f"{total_disp:,}".replace(",", "."))
with col4:
    st.metric("Pré-Matriculados", f"{total_pre:,}".replace(",", "."))
with col5:
    st.metric("Novatos", f"{total_novatos:,}".replace(",", "."), help=f"{total_veteranos:,} veteranos".replace(",", "."))

st.divider()

# Info de filtro
st.caption(f"Mostrando **{len(df_filtrado)}** de **{len(df)}** turmas")

# Tabela principal
df_display = df_filtrado[[
    "unidade", "segmento", "serie", "letra_turma", "turno",
    "vagas", "novatos", "veteranos", "matriculados",
    "pre_matriculados", "disponiveis", "ocupacao", "status",
]].copy()

df_display.columns = [
    "Unidade", "Segmento", "Série", "Turma", "Turno",
    "Vagas", "Novatos", "Veteranos", "Matric.",
    "Pré-Mat.", "Disp.", "Ocupação %", "Status",
]

# Mapear status para labels
status_labels = {
    "critica": "🔴 Crítica",
    "atencao": "🟡 Atenção",
    "normal": "🟢 Normal",
    "lotada": "🔵 Lotada",
    "super": "🟣 Super",
}
df_display["Status"] = df_display["Status"].map(status_labels)

# Ordenar
ordem_unid = {u: i for i, u in enumerate(ORDEM_UNIDADES)}
df_display["_ord_unid"] = df_display["Unidade"].map(ordem_unid).fillna(99)
ordem_ser = {s: i for i, s in enumerate(ORDEM_SERIES)}
df_display["_ord_serie"] = df_display["Série"].map(ordem_ser).fillna(99)
df_display = df_display.sort_values(["_ord_unid", "Segmento", "_ord_serie", "Turma"])
df_display = df_display.drop(columns=["_ord_unid", "_ord_serie"])

st.dataframe(
    df_display,
    width="stretch",
    hide_index=True,
    height=600,
    column_config={
        "Ocupação %": st.column_config.ProgressColumn(
            "Ocupação %",
            format="%.0f%%",
            min_value=0,
            max_value=120,
        ),
        "Matric.": st.column_config.NumberColumn("Matric.", format="%d"),
        "Vagas": st.column_config.NumberColumn("Vagas", format="%d"),
        "Disp.": st.column_config.NumberColumn("Disp.", format="%d"),
    },
)

# Download CSV
csv = df_display.to_csv(index=False, sep=";", encoding="utf-8-sig")
st.download_button(
    label="📥 Download CSV",
    data=csv,
    file_name="ensalamento_2026.csv",
    mime="text/csv",
)

# Resumo por status
st.divider()
st.markdown("### Resumo por Status")

status_count = df_filtrado["status"].value_counts()
cols_status = st.columns(5)

status_info = [
    ("critica", "🔴 Crítica (<40%)", "#f87171"),
    ("atencao", "🟡 Atenção (40-70%)", "#fbbf24"),
    ("normal", "🟢 Normal (70-90%)", "#4ade80"),
    ("lotada", "🔵 Lotada (≥90%)", "#60a5fa"),
    ("super", "🟣 Superlotada (>100%)", "#a78bfa"),
]

for i, (key, label, cor) in enumerate(status_info):
    with cols_status[i]:
        count = int(status_count.get(key, 0))
        st.metric(label, count)
