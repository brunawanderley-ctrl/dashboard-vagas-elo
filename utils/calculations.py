"""Lógica de cálculo compartilhada do dashboard."""

import re
import pandas as pd
from .constants import UNIDADES_MAP, PROGRESSAO, SERIE_SEGMENTO


def nome_unidade(codigo):
    """Converte código de unidade para nome."""
    return UNIDADES_MAP.get(codigo, codigo)


def calc_variacao(v2025, v2026):
    """Calcula variação percentual."""
    if v2025 == 0:
        return 0.0
    return round((v2026 - v2025) / v2025 * 100, 1)


def format_variacao(var):
    """Formata variação com sinal."""
    if var > 0:
        return f"+{var}%"
    return f"{var}%"


def format_diferenca(dif):
    """Formata diferença com sinal."""
    if dif > 0:
        return f"+{dif}"
    return str(dif)


def extrair_serie(nome_turma):
    """Extrai e padroniza a série do nome da turma."""
    if not nome_turma:
        return None
    nome = nome_turma.lower()

    # Infantil (ordem importa!)
    if "infantil v" in nome and "infantil vi" not in nome:
        return "Infantil V"
    if "infantil iv" in nome:
        return "Infantil IV"
    if "infantil iii" in nome:
        return "Infantil III"
    if "infantil ii" in nome:
        return "Infantil II"

    # Ensino Médio (série)
    match = re.search(r"(\d)\s*[ºª°]?\s*s[eé]rie", nome)
    if match:
        return f"{match.group(1)}ª Série"

    # Ensino Médio (ano + médio) → normaliza para "Xª Série"
    if "médio" in nome or "medio" in nome:
        match = re.search(r"(\d)\s*[ºª°]?\s*ano", nome)
        if match:
            return f"{match.group(1)}ª Série"

    # Fundamental (ano)
    match = re.search(r"(\d)\s*[ºª°]?\s*ano", nome)
    if match:
        return f"{match.group(1)}º Ano"

    return None


def extrair_serie_ensalamento(nome_turma):
    """Extrai série para ensalamento (usa 'ª Série' para EM)."""
    if not nome_turma:
        return "Outros"
    nome = nome_turma.lower()

    if "infantil v" in nome and "infantil vi" not in nome:
        return "Infantil V"
    if "infantil iv" in nome:
        return "Infantil IV"
    if "infantil iii" in nome:
        return "Infantil III"
    if "infantil ii" in nome:
        return "Infantil II"

    match = re.search(r"(\d)\s*[ºª°]?\s*s[eé]rie", nome)
    if match:
        return f"{match.group(1)}ª Série"

    match = re.search(r"(\d)\s*[ºª°]?\s*ano", nome)
    if match:
        return f"{match.group(1)}º Ano"

    return "Outros"


def extrair_turno(nome_turma):
    """Extrai o turno do nome da turma."""
    if not nome_turma:
        return "Manhã"
    if "tarde" in nome_turma.lower():
        return "Tarde"
    return "Manhã"


def extrair_letra_turma(nome_turma):
    """Extrai a letra da turma (A, B, C)."""
    if not nome_turma:
        return "-"
    match = re.search(r"Turma\s*([A-Z])|([A-Z])\s*(Manhã|Tarde|M\s|T\s)", nome_turma, re.IGNORECASE)
    if match:
        return match.group(1) or match.group(2) or "-"
    return "-"


def calcular_status_ocupacao(matriculados, vagas):
    """Calcula o status de ocupação de uma turma."""
    if vagas <= 0:
        return "critica", 0.0
    ocupacao = matriculados / vagas * 100
    if matriculados > vagas:
        return "super", ocupacao
    if ocupacao >= 90:
        return "lotada", ocupacao
    if ocupacao >= 70:
        return "normal", ocupacao
    if ocupacao >= 40:
        return "atencao", ocupacao
    return "critica", ocupacao


def calcular_evasao(df_2025, df_2026):
    """Calcula evasão por série e unidade.

    Recebe DataFrames com colunas: unidade_codigo, turma, total_20XX, veteranos_20XX, novatos_20XX
    Retorna DataFrame com análise de evasão.
    """
    # Extrair série
    df_2025 = df_2025.copy()
    df_2026 = df_2026.copy()
    df_2025["serie"] = df_2025["turma"].apply(extrair_serie)
    df_2026["serie"] = df_2026["turma"].apply(extrair_serie)

    df_2025 = df_2025[df_2025["serie"].notna()]
    df_2026 = df_2026[df_2026["serie"].notna()]

    # Agregar por unidade e série
    dados_2025 = df_2025.groupby(["unidade_codigo", "serie"]).agg({
        "total_2025": "sum",
        "veteranos_2025": "sum",
        "novatos_2025": "sum",
    }).reset_index()

    dados_2026 = df_2026.groupby(["unidade_codigo", "serie"]).agg({
        "total_2026": "sum",
        "veteranos_2026": "sum",
        "novatos_2026": "sum",
    }).reset_index()

    resultados = []
    for _, row_2025 in dados_2025.iterrows():
        unidade = row_2025["unidade_codigo"]
        serie_2025 = row_2025["serie"]
        serie_esperada = PROGRESSAO.get(serie_2025)

        if not serie_esperada:
            continue

        alunos_2025 = row_2025["total_2025"]
        is_concluinte = (serie_esperada == "Formado")

        if is_concluinte:
            # Concluintes (3ª Série): não são evasão
            resultados.append({
                "unidade_codigo": unidade,
                "unidade": nome_unidade(unidade),
                "serie_2025": serie_2025,
                "serie_2026": "Formado",
                "segmento": SERIE_SEGMENTO.get(serie_2025, ""),
                "alunos_2025": int(alunos_2025),
                "veteranos_2026": 0,
                "novatos_2026": 0,
                "total_2026": 0,
                "evasao": 0,
                "pct_evasao": 0.0,
                "pct_retencao": 0.0,
                "concluintes": int(alunos_2025),
            })
            continue

        filtro = (dados_2026["unidade_codigo"] == unidade) & (dados_2026["serie"] == serie_esperada)
        dados_serie_2026 = dados_2026[filtro]

        if len(dados_serie_2026) > 0:
            veteranos_2026 = dados_serie_2026["veteranos_2026"].values[0]
            total_2026 = dados_serie_2026["total_2026"].values[0]
            novatos_2026 = dados_serie_2026["novatos_2026"].values[0]
        else:
            veteranos_2026 = 0
            total_2026 = 0
            novatos_2026 = 0

        evasao = max(0, alunos_2025 - veteranos_2026)
        pct_evasao = (evasao / alunos_2025 * 100) if alunos_2025 > 0 else 0
        pct_retencao = 100 - pct_evasao

        resultados.append({
            "unidade_codigo": unidade,
            "unidade": nome_unidade(unidade),
            "serie_2025": serie_2025,
            "serie_2026": serie_esperada,
            "segmento": SERIE_SEGMENTO.get(serie_2025, ""),
            "alunos_2025": int(alunos_2025),
            "veteranos_2026": int(veteranos_2026),
            "novatos_2026": int(novatos_2026),
            "total_2026": int(total_2026),
            "evasao": int(evasao),
            "pct_evasao": round(pct_evasao, 1),
            "pct_retencao": round(pct_retencao, 1),
            "concluintes": 0,
        })

    return pd.DataFrame(resultados)
