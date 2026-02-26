#!/usr/bin/env python3
"""
SIGA API - Modulo de login e extracao via requests.
Extraido de extrair_recebimento_api.py para uso no dashboard Streamlit Cloud.

Funcoes publicas:
    atualizar_via_api(instituicao, login, senha, progress_cb=None) -> (list, str|None)
"""

import re
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

# =============================================================================
# CONFIGURACOES
# =============================================================================

SIGA_URL = "https://siga02.activesoft.com.br"

UNIDADES_API = [
    {"pk": 2, "codigo": "BV", "nome": "1 - BV (Boa Viagem)"},
    {"pk": 3, "codigo": "CD", "nome": "2 - CD (Jaboatao)"},
    {"pk": 4, "codigo": "JG", "nome": "3 - JG (Paulista)"},
    {"pk": 5, "codigo": "CDR", "nome": "4 - CDR (Cordeiro)"},
]

DATA_INICIAL = "01/08/2025"

CODIGOS_SERVICOS = {
    "901", "902", "903", "904",
    "913", "914", "915", "920", "921",
    "916", "917", "918", "919",
    "912", "991", "992",
    "933", "934", "941", "942", "943", "944", "945", "946",
    "995",
}

CODIGOS_EXCLUIR_POR_UNIDADE = {"CDR": {"995"}}

SERVICOS = {
    "901": {"nome": "SISTEM. ELO - INF. 2", "segmento": "Infantil", "serie": "Infantil II", "tipo": "SAE"},
    "902": {"nome": "SISTEM. ELO - INF. 3", "segmento": "Infantil", "serie": "Infantil III", "tipo": "SAE"},
    "903": {"nome": "SISTEM. ELO - INF. 4", "segmento": "Infantil", "serie": "Infantil IV", "tipo": "SAE"},
    "904": {"nome": "SISTEM. ELO - INF. 5", "segmento": "Infantil", "serie": "Infantil V", "tipo": "SAE"},
    "913": {"nome": "SISTEM. ELO - 3 ANO", "segmento": "Fund1", "serie": "3\u00ba Ano", "tipo": "SAE"},
    "914": {"nome": "SISTEM. ELO - 4 ANO", "segmento": "Fund1", "serie": "4\u00ba Ano", "tipo": "SAE"},
    "915": {"nome": "SISTEM. ELO - 5 ANO", "segmento": "Fund1", "serie": "5\u00ba Ano", "tipo": "SAE"},
    "920": {"nome": "SISTEM. ELO - 2 ANO", "segmento": "Fund1", "serie": "2\u00ba Ano", "tipo": "SAE"},
    "921": {"nome": "SISTEM. ELO - 1 ANO", "segmento": "Fund1", "serie": "1\u00ba Ano", "tipo": "SAE"},
    "916": {"nome": "SISTEM. ELO - 6 ANO", "segmento": "Fund2", "serie": "6\u00ba Ano", "tipo": "SAE"},
    "917": {"nome": "SISTEM. ELO - 7 ANO", "segmento": "Fund2", "serie": "7\u00ba Ano", "tipo": "SAE"},
    "918": {"nome": "SISTEM. ELO - 8 ANO", "segmento": "Fund2", "serie": "8\u00ba Ano", "tipo": "SAE"},
    "919": {"nome": "SISTEM. ELO - 9 ANO", "segmento": "Fund2", "serie": "9\u00ba Ano", "tipo": "SAE"},
    "912": {"nome": "SISTEM. ELO - 1 MEDIO", "segmento": "M\u00e9dio", "serie": "1\u00ba Ano", "tipo": "SAE"},
    "991": {"nome": "SISTEM. ELO - 2 MEDIO", "segmento": "M\u00e9dio", "serie": "2\u00ba Ano", "tipo": "SAE"},
    "992": {"nome": "SISTEM. ELO - 3 MEDIO", "segmento": "M\u00e9dio", "serie": "3\u00ba Ano", "tipo": "SAE"},
    "933": {"nome": "PROJETO SOCIO EMOCIONAL 4", "segmento": "Infantil", "serie": "Infantil IV", "tipo": "Socioemocional"},
    "934": {"nome": "PROJETO SOCIO EMOCIONAL 5", "segmento": "Infantil", "serie": "Infantil V", "tipo": "Socioemocional"},
    "941": {"nome": "PROJETO SOCIO EMOCIONAL 1 ANO", "segmento": "Fund1", "serie": "1\u00ba Ano", "tipo": "Socioemocional"},
    "942": {"nome": "PROJETO SOCIO EMOCIONAL 2 ANO", "segmento": "Fund1", "serie": "2\u00ba Ano", "tipo": "Socioemocional"},
    "943": {"nome": "PROJETO SOCIO EMOCIONAL 3 ANO", "segmento": "Fund1", "serie": "3\u00ba Ano", "tipo": "Socioemocional"},
    "944": {"nome": "PROJETO SOCIO EMOCIONAL 4 ANO", "segmento": "Fund1", "serie": "4\u00ba Ano", "tipo": "Socioemocional"},
    "945": {"nome": "PROJETO SOCIO EMOCIONAL 5 ANO", "segmento": "Fund1", "serie": "5\u00ba Ano", "tipo": "Socioemocional"},
    "946": {"nome": "PROJETO SOCIOEMOCIONAL 6 ANO", "segmento": "Fund2", "serie": "6\u00ba Ano", "tipo": "Socioemocional"},
    "995": {"nome": "ELO TECH", "segmento": "Geral", "serie": "Todas", "tipo": "Elo Tech"},
}

_RE_CODIGO_SERVICO = re.compile(r'^(\d{3})\s*-')


# =============================================================================
# FUNCOES AUXILIARES
# =============================================================================

def extract_csrf(html_text):
    """Extrai csrfmiddlewaretoken do HTML do formulario Django."""
    m = re.search(r'name=["\']csrfmiddlewaretoken["\'].*?value=["\']([^"\']+)["\']', html_text)
    if m:
        return m.group(1)
    m = re.search(r'value=["\']([^"\']+)["\'].*?name=["\']csrfmiddlewaretoken["\']', html_text)
    return m.group(1) if m else None


def _formatar_data(dt_str):
    """Converte ISO datetime (2026-01-26T00:00:00) para DD/MM/YYYY."""
    if not dt_str:
        return ""
    dt_str = str(dt_str)
    if "T" in dt_str:
        dt_str = dt_str.split("T")[0]
    try:
        dt_obj = datetime.strptime(dt_str, "%Y-%m-%d")
        return dt_obj.strftime("%d/%m/%Y")
    except ValueError:
        return dt_str


def _formatar_valor(v):
    """Formata valor numerico para string com virgula (ex: 1765.00 -> '1.765,00')."""
    if v is None:
        return "0,00"
    try:
        n = float(v)
        inteira = int(n)
        dec = abs(n - inteira)
        dec_str = f"{dec:.2f}"[2:]
        if inteira >= 1000:
            parts = []
            s = str(abs(inteira))
            while len(s) > 3:
                parts.insert(0, s[-3:])
                s = s[:-3]
            parts.insert(0, s)
            inteira_str = ".".join(parts)
        else:
            inteira_str = str(inteira)
        return f"{inteira_str},{dec_str}"
    except (ValueError, TypeError):
        return str(v)


def _extrair_turma(turmas_vinculadas):
    """Extrai letra da turma de turmas_vinculadas."""
    if not turmas_vinculadas:
        return ""
    texto = str(turmas_vinculadas[0])
    m = re.search(r'Turma\s+(\w)', texto)
    return m.group(1) if m else ""


# =============================================================================
# LOGIN E SESSAO
# =============================================================================

def criar_sessao_logada(instituicao, login, senha, unidade):
    """Cria uma sessao requests autenticada para uma unidade especifica.

    Args:
        instituicao: Codigo da instituicao (ex: "COLEGIOELO")
        login: Usuario SIGA
        senha: Senha SIGA
        unidade: Dict com pk, codigo, nome
    """
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Accept": "application/json, text/html, */*",
    })

    # 1. GET login page -> CSRF token
    r = s.get(f"{SIGA_URL}/login/", timeout=15)
    csrf = s.cookies.get("csrftoken", "")
    if not csrf:
        csrf = extract_csrf(r.text) or ""
    if not csrf:
        return None

    # 2. POST login credentials
    r = s.post(
        f"{SIGA_URL}/login/",
        data={
            "csrfmiddlewaretoken": csrf,
            "codigo": instituicao,
            "login": login,
            "senha": senha,
        },
        headers={"Referer": f"{SIGA_URL}/login/"},
        allow_redirects=True,
        timeout=15,
    )

    if "/login/unidade" not in r.url:
        return None

    # 3. Selecionar unidade via POST
    csrf2 = s.cookies.get("csrftoken", csrf)
    r = s.post(
        f"{SIGA_URL}/login/unidade/",
        data={
            "csrfmiddlewaretoken": csrf2,
            "unidade": str(unidade["pk"]),
        },
        headers={"Referer": f"{SIGA_URL}/login/unidade/"},
        allow_redirects=True,
        timeout=15,
    )

    if "/info" not in r.url and r.status_code != 200:
        return None

    # 4. Verifica sessao
    try:
        r = s.get(f"{SIGA_URL}/api/v1/servico/?limit=1", timeout=10)
        if r.status_code == 200:
            return s
    except Exception:
        pass

    try:
        r = s.get(f"{SIGA_URL}/info/", allow_redirects=False, timeout=10)
        if r.status_code == 200:
            return s
    except Exception:
        pass

    return s


# =============================================================================
# EXTRACAO POR UNIDADE
# =============================================================================

def _processar_unidade(instituicao, login, senha, unidade):
    """Processa uma unidade: login + paginacao streaming com filtro inline."""
    codigo = unidade["codigo"]
    data_final = datetime.now().strftime("%d/%m/%Y")
    LIMIT = 500

    session = criar_sessao_logada(instituicao, login, senha, unidade)
    if not session:
        return {"codigo": codigo, "registros": [], "erro": f"Login falhou para {codigo}"}

    excluir = CODIGOS_EXCLUIR_POR_UNIDADE.get(codigo, set())
    registros = []
    offset = 0
    pagina = 0
    max_retries = 3

    while True:
        params = {
            "limit": LIMIT,
            "offset": offset,
            "ordenacao": "nome_aluno",
            "situacao": "liq",
            "data_baixa_inicial": DATA_INICIAL,
            "data_baixa_final": data_final,
        }

        success = False
        data = None
        for tentativa in range(max_retries):
            try:
                timeout = 60 + (tentativa * 30)
                r = session.get(
                    f"{SIGA_URL}/api/v1/titulos/",
                    params=params,
                    timeout=timeout,
                )
                if r.status_code != 200:
                    break

                data = r.json()
                success = True
                break
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
                if tentativa < max_retries - 1:
                    time.sleep(5 * (tentativa + 1))
                else:
                    return {"codigo": codigo, "registros": registros, "erro": f"Timeout em {codigo}"}
            except Exception as e:
                return {"codigo": codigo, "registros": registros, "erro": str(e)}

        if not success or data is None:
            break

        results = data.get("results", [])

        for titulo in results:
            for srv in titulo.get("servicos", []):
                nome = srv.get("nome_servico", "")
                m = _RE_CODIGO_SERVICO.match(nome)
                if not m:
                    continue
                cod = m.group(1)
                if cod not in CODIGOS_SERVICOS or cod in excluir:
                    continue

                info = SERVICOS.get(cod, {})
                dt_baixa = _formatar_data(
                    titulo.get("data_baixa") or titulo.get("data_pagamento")
                )
                if not dt_baixa:
                    dt_baixa = _formatar_data(titulo.get("data_vencimento"))

                registros.append({
                    "unidade": codigo,
                    "servico_codigo": cod,
                    "turma": _extrair_turma(titulo.get("turmas_vinculadas", [])),
                    "matricula": titulo.get("matricula", ""),
                    "nome": titulo.get("nome_aluno", ""),
                    "titulo": str(titulo.get("id", "")),
                    "parcela": titulo.get("parcela", ""),
                    "dt_baixa": dt_baixa,
                    "valor": _formatar_valor(srv.get("valor_servico")),
                    "recebido": _formatar_valor(titulo.get("valor_recebido")),
                    "segmento": info.get("segmento", ""),
                    "serie": info.get("serie", ""),
                    "tipo": info.get("tipo", ""),
                })

        if not data.get("next"):
            break
        offset += LIMIT
        pagina += 1

    return {"codigo": codigo, "registros": registros, "erro": None}


# =============================================================================
# FUNCAO PUBLICA
# =============================================================================

def atualizar_via_api(instituicao, login, senha, progress_cb=None):
    """Extrai dados de recebimento de todas as unidades via API SIGA.

    Args:
        instituicao: Codigo da instituicao (ex: "COLEGIOELO")
        login: Usuario SIGA
        senha: Senha SIGA
        progress_cb: Callback opcional fn(msg) para progresso

    Returns:
        (registros, erro): Lista de registros e mensagem de erro (None se ok)
    """
    todos_registros = []
    erros = []

    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {
            pool.submit(_processar_unidade, instituicao, login, senha, u): u
            for u in UNIDADES_API
        }
        for future in as_completed(futures):
            u = futures[future]
            try:
                resultado = future.result()
                todos_registros.extend(resultado["registros"])
                if resultado["erro"]:
                    erros.append(resultado["erro"])
                if progress_cb:
                    progress_cb(f"{u['codigo']}: {len(resultado['registros'])} registros")
            except Exception as e:
                erros.append(f"{u['codigo']}: {e}")

    if not todos_registros and erros:
        return [], "; ".join(erros)

    return todos_registros, None
