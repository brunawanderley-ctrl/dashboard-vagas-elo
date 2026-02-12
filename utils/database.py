"""Queries SQLite com cache para o dashboard."""

import sqlite3
import pandas as pd
import streamlit as st
from pathlib import Path
from datetime import datetime

OUTPUT_DIR = Path(__file__).parent.parent / "output"


def _get_connection(db_name):
    """Cria conexão SQLite read-only."""
    db_path = OUTPUT_DIR / db_name
    if not db_path.exists():
        return None
    return sqlite3.connect(db_path, check_same_thread=False)


@st.cache_data(ttl=300)
def get_matriculas_2026():
    """Retorna dados de matrículas 2026 agrupados por unidade e segmento."""
    conn = _get_connection("vagas.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, segmento, SUM(matriculados) as matriculados
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
        GROUP BY unidade_codigo, segmento
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_matriculas_2025():
    """Retorna dados de matrículas 2025 agrupados por unidade e segmento."""
    conn = _get_connection("vagas_2025.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, segmento, SUM(matriculados) as matriculados
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
        GROUP BY unidade_codigo, segmento
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_integral_2026():
    """Retorna dados do integral 2026 agrupados por unidade."""
    conn = _get_connection("integral.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, SUM(matriculados) as matriculados
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
        GROUP BY unidade_codigo
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_integral_2025():
    """Retorna dados do integral 2025 agrupados por unidade."""
    conn = _get_connection("integral_2025.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, SUM(matriculados) as matriculados
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
        GROUP BY unidade_codigo
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_turmas_detalhadas_2026():
    """Retorna dados detalhados por turma 2026 (para ensalamento)."""
    conn = _get_connection("vagas.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, unidade_nome, segmento, turma,
               vagas, novatos, veteranos, matriculados,
               pre_matriculados, disponiveis
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_evasao_2025():
    """Retorna dados por turma 2025 para cálculo de evasão."""
    conn = _get_connection("vagas_2025.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, turma,
               SUM(matriculados) as total_2025,
               SUM(veteranos) as veteranos_2025,
               SUM(novatos) as novatos_2025
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
        GROUP BY unidade_codigo, turma
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_evasao_2026():
    """Retorna dados por turma 2026 para cálculo de evasão."""
    conn = _get_connection("vagas.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, turma,
               SUM(matriculados) as total_2026,
               SUM(veteranos) as veteranos_2026,
               SUM(novatos) as novatos_2026
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
        GROUP BY unidade_codigo, turma
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_matriculas_por_turma_2026():
    """Retorna dados por turma 2026 (para tabela por série)."""
    conn = _get_connection("vagas.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, segmento, turma,
               novatos, veteranos, matriculados
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_matriculas_por_turma_2025():
    """Retorna dados por turma 2025 (para comparativo novatos/veteranos)."""
    conn = _get_connection("vagas_2025.db")
    if conn is None:
        return pd.DataFrame()
    df = pd.read_sql_query("""
        SELECT unidade_codigo, segmento, turma,
               novatos, veteranos, matriculados
        FROM vagas
        WHERE extracao_id = (SELECT MAX(id) FROM extrações)
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=300)
def get_ultima_extracao():
    """Retorna a data da última extração."""
    conn = _get_connection("vagas.db")
    if conn is None:
        return "N/A"
    cursor = conn.cursor()
    cursor.execute("SELECT data_extracao FROM extrações ORDER BY id DESC LIMIT 1")
    row = cursor.fetchone()
    conn.close()
    if row and row[0]:
        try:
            dt = datetime.fromisoformat(row[0])
            return dt.strftime("%d/%m/%Y %H:%M")
        except (ValueError, TypeError):
            return row[0]
    return "N/A"
