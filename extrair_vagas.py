#!/usr/bin/env python3
"""
SIGA Activesoft - Extrator de Resumo de Vagas por Turma
Extrai dados das 4 unidades do Colégio Elo e salva em SQLite/JSON
"""

import json
import sqlite3
import re
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# Configurações
CONFIG = {
    "url": "https://siga.activesoft.com.br/login/",
    "instituicao": "COLEGIOELO",
    "login": "bruna",
    "senha": "Sucesso@25",
    "periodo": "2026",
    "unidades": [
        {"id": "17", "nome": "1 - BV (Boa Viagem)", "codigo": "01-BV"},
        {"id": "18", "nome": "2 - CD (Jaboatão)", "codigo": "02-CD"},
        {"id": "19", "nome": "3 - JG (Paulista)", "codigo": "03-JG"},
        {"id": "20", "nome": "4 - CDR (Cordeiro)", "codigo": "04-CDR"},
    ],
    # Cursos a IGNORAR (Esportes, Integral, Lanche, Cursos Livres, Transporte)
    "cursos_ignorar": [
        "esporte", "ballet", "futsal", "judô", "judo", "voleibol", "basquete",
        "ginástica", "ginastica", "karatê", "karate",
        "integral", "complementar",
        "lanche saudável", "lanche saudavel",
        "curso livre", "cursos livres",
        "transporte"
    ]
}

# Diretório de saída
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def deve_ignorar_curso(nome_turma: str) -> bool:
    """Verifica se a turma deve ser ignorada baseado no nome"""
    nome_lower = nome_turma.lower()
    for termo in CONFIG["cursos_ignorar"]:
        if termo in nome_lower:
            return True
    return False


def identificar_segmento(nome_curso: str) -> str:
    """Identifica o segmento educacional pelo nome do curso"""
    nome_lower = nome_curso.lower()

    if "infantil" in nome_lower:
        return "Ed. Infantil"
    # Fund. 2 ANTES de Fund. 1 (senão "Fundamental II" casa com "Fundamental I")
    elif "fundamental ii" in nome_lower or "fundamental 2" in nome_lower:
        return "Fund. 2"
    elif "fundamental i" in nome_lower or "fundamental 1" in nome_lower:
        return "Fund. 1"
    elif "médio" in nome_lower or "medio" in nome_lower:
        return "Ens. Médio"
    else:
        return "Outro"


def parse_numero(texto: str) -> int:
    """Converte texto para número, tratando valores negativos e vazios"""
    try:
        return int(texto.replace(".", "").replace(",", "").strip())
    except (ValueError, AttributeError):
        return 0


def extrair_dados_relatorio(page) -> list:
    """Extrai os dados do relatório renderizado no iframe"""

    # Aguarda o iframe carregar
    page.wait_for_selector("iframe", timeout=10000)

    # Pega o conteúdo do iframe
    iframe = page.frame_locator("iframe").first

    # Aguarda o conteúdo carregar
    iframe.locator("text=Total geral").wait_for(timeout=15000)

    # Extrai o HTML do relatório
    html_content = iframe.locator("body").inner_html()

    # Parse dos dados usando regex (mais robusto que DOM parsing)
    turmas = []

    # Encontra todas as linhas de turma (não são "Total da série" nem "Total geral")
    # Padrão: Nome da turma seguido de 7 números
    linhas = html_content.split("<tr")

    curso_atual = ""

    for linha in linhas:
        # Detecta header de curso/série
        if "/ 2026" in linha:
            match = re.search(r'>([^<]+/ 2026)<', linha)
            if match:
                curso_atual = match.group(1)

        # Detecta linha de turma (tem exatamente 7 células de dados após o nome)
        # Ignora linhas de total
        if "Total da série" in linha or "Total geral" in linha:
            continue

        # Extrai dados da turma
        cells = re.findall(r'<td[^>]*>([^<]*)</td>', linha)
        if len(cells) >= 7:
            nome_turma = cells[0].strip() if cells else ""

            if nome_turma and not deve_ignorar_curso(nome_turma) and not deve_ignorar_curso(curso_atual):
                try:
                    turma_data = {
                        "turma": nome_turma,
                        "curso": curso_atual,
                        "segmento": identificar_segmento(curso_atual),
                        "vagas": parse_numero(cells[1]) if len(cells) > 1 else 0,
                        "novatos": parse_numero(cells[2]) if len(cells) > 2 else 0,
                        "veteranos": parse_numero(cells[3]) if len(cells) > 3 else 0,
                        "matriculados": parse_numero(cells[4]) if len(cells) > 4 else 0,
                        "vagas_restantes": parse_numero(cells[5]) if len(cells) > 5 else 0,
                        "pre_matriculados": parse_numero(cells[6]) if len(cells) > 6 else 0,
                        "disponiveis": parse_numero(cells[7]) if len(cells) > 7 else 0,
                    }

                    # Só adiciona se for segmento válido
                    if turma_data["segmento"] != "Outro":
                        turmas.append(turma_data)
                except Exception as e:
                    print(f"  Erro ao processar turma: {e}")

    return turmas


def extrair_dados_simples(page) -> list:
    """Método alternativo: usa o botão de copiar dados"""

    # Clica no botão de copiar
    page.click("button:has-text('COPIAR DADOS')")
    page.wait_for_timeout(1000)

    # Pega do clipboard (requer permissão)
    # Alternativa: extrai direto do snapshot

    turmas = []

    # Pega snapshot do iframe
    iframe = page.frame_locator("iframe").first
    content = iframe.locator("body").inner_text()

    linhas = content.split("\n")
    curso_atual = ""

    i = 0
    while i < len(linhas):
        linha = linhas[i].strip()

        # Detecta header de curso
        if "/ 2026" in linha:
            curso_atual = linha
            i += 1
            continue

        # Pula totais
        if "Total da série" in linha or "Total geral" in linha:
            i += 1
            continue

        # Detecta turma (linha com nome seguida de números)
        if linha and not linha.isdigit() and curso_atual:
            # Próximas 7 linhas são os números
            if i + 7 < len(linhas):
                nome_turma = linha

                if not deve_ignorar_curso(nome_turma) and not deve_ignorar_curso(curso_atual):
                    try:
                        turma_data = {
                            "turma": nome_turma,
                            "curso": curso_atual,
                            "segmento": identificar_segmento(curso_atual),
                            "vagas": parse_numero(linhas[i+1]),
                            "novatos": parse_numero(linhas[i+2]),
                            "veteranos": parse_numero(linhas[i+3]),
                            "matriculados": parse_numero(linhas[i+4]),
                            "vagas_restantes": parse_numero(linhas[i+5]),
                            "pre_matriculados": parse_numero(linhas[i+6]),
                            "disponiveis": parse_numero(linhas[i+7]),
                        }

                        if turma_data["segmento"] != "Outro" and turma_data["vagas"] > 0:
                            turmas.append(turma_data)

                        i += 8
                        continue
                    except:
                        pass

        i += 1

    return turmas


def extrair_via_snapshot(page) -> list:
    """Extrai dados usando o texto do snapshot (mais confiável)"""

    turmas = []

    # Tenta extrair do iframe primeiro, se existir
    try:
        frames = page.frames
        # Procura por frame que contenha os dados do relatório
        frame_content = None
        for frame in frames:
            try:
                content = frame.locator("body").inner_text(timeout=5000)
                if "Total da série" in content:
                    frame_content = content
                    break
            except:
                continue

        if frame_content:
            texto_completo = frame_content
        else:
            # Se não encontrou em frames, pega da página principal
            texto_completo = page.locator("body").inner_text()
    except:
        texto_completo = page.locator("body").inner_text()

    linhas = texto_completo.split("\n")

    curso_atual = ""

    for linha in linhas:
        linha_raw = linha
        linha = linha.strip()
        if not linha:
            continue

        # Detecta header de curso/série (ex: "1- BV - Educação Infantil - ... / 2026")
        if "/ 2026" in linha and not linha.startswith("Total"):
            curso_atual = linha
            continue

        # Pula linhas de total
        if linha.startswith("Total da série") or linha.startswith("Total geral"):
            continue

        # Verifica se é uma linha de dados (contém tabs com números)
        partes = linha_raw.split("\t")
        if len(partes) >= 2 and curso_atual:
            nome_turma = partes[0].strip()

            # Pula se for cabeçalho
            if nome_turma in ["Turma - Turno", "(A)", "Vagas abertas", "Novatos"]:
                continue

            # Tenta extrair números das partes após o nome
            numeros = []
            for p in partes[1:]:
                p = p.strip()
                # Extrai números, incluindo negativos
                if p and (p.lstrip("-").isdigit()):
                    numeros.append(parse_numero(p))

            if len(numeros) >= 7:
                if not deve_ignorar_curso(nome_turma) and not deve_ignorar_curso(curso_atual):
                    segmento = identificar_segmento(curso_atual)

                    if segmento != "Outro":
                        turma_data = {
                            "turma": nome_turma,
                            "curso": curso_atual,
                            "segmento": segmento,
                            "vagas": numeros[0],
                            "novatos": numeros[1],
                            "veteranos": numeros[2],
                            "matriculados": numeros[3],
                            "vagas_restantes": numeros[4],
                            "pre_matriculados": numeros[5],
                            "disponiveis": numeros[6],
                        }
                        turmas.append(turma_data)

    return turmas


def salvar_sqlite(dados: dict, db_path: Path):
    """Salva os dados em SQLite"""

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Cria tabelas se não existirem
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS extrações (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_extracao TEXT,
            periodo TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vagas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            extracao_id INTEGER,
            unidade_codigo TEXT,
            unidade_nome TEXT,
            segmento TEXT,
            curso TEXT,
            turma TEXT,
            vagas INTEGER,
            novatos INTEGER,
            veteranos INTEGER,
            matriculados INTEGER,
            vagas_restantes INTEGER,
            pre_matriculados INTEGER,
            disponiveis INTEGER,
            FOREIGN KEY (extracao_id) REFERENCES extrações(id)
        )
    """)

    # Insere extração
    cursor.execute(
        "INSERT INTO extrações (data_extracao, periodo) VALUES (?, ?)",
        (dados["data_extracao"], dados["periodo"])
    )
    extracao_id = cursor.lastrowid

    # Insere turmas
    for unidade in dados["unidades"]:
        for turma in unidade["turmas"]:
            cursor.execute("""
                INSERT INTO vagas (
                    extracao_id, unidade_codigo, unidade_nome, segmento, curso, turma,
                    vagas, novatos, veteranos, matriculados, vagas_restantes,
                    pre_matriculados, disponiveis
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                extracao_id,
                unidade["codigo"],
                unidade["nome"],
                turma["segmento"],
                turma["curso"],
                turma["turma"],
                turma["vagas"],
                turma["novatos"],
                turma["veteranos"],
                turma["matriculados"],
                turma["vagas_restantes"],
                turma["pre_matriculados"],
                turma["disponiveis"],
            ))

    conn.commit()
    conn.close()

    print(f"  Dados salvos em SQLite: {db_path}")


def salvar_json(dados: dict, json_path: Path):
    """Salva os dados em JSON"""

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    print(f"  Dados salvos em JSON: {json_path}")


def gerar_resumo(dados: dict) -> dict:
    """Gera resumo consolidado por unidade e segmento"""

    resumo = {
        "data_extracao": dados["data_extracao"],
        "periodo": dados["periodo"],
        "unidades": []
    }

    for unidade in dados["unidades"]:
        resumo_unidade = {
            "codigo": unidade["codigo"],
            "nome": unidade["nome"],
            "segmentos": {},
            "total": {
                "vagas": 0,
                "novatos": 0,
                "veteranos": 0,
                "matriculados": 0,
                "disponiveis": 0
            }
        }

        for turma in unidade["turmas"]:
            seg = turma["segmento"]

            if seg not in resumo_unidade["segmentos"]:
                resumo_unidade["segmentos"][seg] = {
                    "vagas": 0,
                    "novatos": 0,
                    "veteranos": 0,
                    "matriculados": 0,
                    "disponiveis": 0
                }

            resumo_unidade["segmentos"][seg]["vagas"] += turma["vagas"]
            resumo_unidade["segmentos"][seg]["novatos"] += turma["novatos"]
            resumo_unidade["segmentos"][seg]["veteranos"] += turma["veteranos"]
            resumo_unidade["segmentos"][seg]["matriculados"] += turma["matriculados"]
            resumo_unidade["segmentos"][seg]["disponiveis"] += turma["disponiveis"]

            resumo_unidade["total"]["vagas"] += turma["vagas"]
            resumo_unidade["total"]["novatos"] += turma["novatos"]
            resumo_unidade["total"]["veteranos"] += turma["veteranos"]
            resumo_unidade["total"]["matriculados"] += turma["matriculados"]
            resumo_unidade["total"]["disponiveis"] += turma["disponiveis"]

        resumo["unidades"].append(resumo_unidade)

    # Total geral
    resumo["total_geral"] = {
        "vagas": sum(u["total"]["vagas"] for u in resumo["unidades"]),
        "novatos": sum(u["total"]["novatos"] for u in resumo["unidades"]),
        "veteranos": sum(u["total"]["veteranos"] for u in resumo["unidades"]),
        "matriculados": sum(u["total"]["matriculados"] for u in resumo["unidades"]),
        "disponiveis": sum(u["total"]["disponiveis"] for u in resumo["unidades"]),
    }

    return resumo


def main():
    """Função principal"""

    print("=" * 60)
    print("SIGA - Extrator de Resumo de Vagas por Turma")
    print(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 60)

    dados = {
        "data_extracao": datetime.now().isoformat(),
        "periodo": CONFIG["periodo"],
        "unidades": []
    }

    with sync_playwright() as p:
        # Inicia browser (headless=False para debug, True para produção)
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            # 1. Login
            print("\n[1/5] Fazendo login...")
            page.goto(CONFIG["url"], wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            # Preenche formulário de login
            page.fill('#codigoInstituicao', CONFIG["instituicao"])
            page.fill('#id_login', CONFIG["login"])
            page.fill('#id_senha', CONFIG["senha"])
            page.click('button:has-text("ENTRAR")')

            # Aguarda redirect para página de seleção de unidade
            page.wait_for_url("**/login/unidade/**", timeout=15000)
            page.wait_for_timeout(2000)
            print("  Login realizado com sucesso!")

            # 2. Para cada unidade
            for idx, unidade in enumerate(CONFIG["unidades"]):
                print(f"\n[{idx+2}/5] Processando {unidade['nome']}...")

                # Se é a primeira unidade, seleciona diretamente na página de login/unidade
                if idx == 0:
                    # Clica na unidade na página de seleção inicial
                    page.click(f'text={unidade["nome"]}')
                    page.wait_for_timeout(5000)
                else:
                    # Para unidades seguintes, usa o dropdown de troca
                    # Clica no botão que contém "Unidade" e "keyboard_arrow_down"
                    page.locator('button:has-text("Unidade"):has-text("keyboard_arrow_down")').click()
                    page.wait_for_timeout(1000)

                    # Clica na nova unidade no dropdown (botão com swap_vert)
                    page.locator(f'button:has-text("{unidade["nome"]}")').filter(has_text="swap_vert").click()
                    page.wait_for_timeout(5000)

                # Navega diretamente para o relatório de vagas
                # Nota: URL base pode mudar dependendo da unidade
                current_url = page.url
                base_url = current_url.split('/')[0] + '//' + current_url.split('/')[2]
                report_url = f"{base_url}/busca_central_relatorios/?relatorio=aluno_turma/resumo_vagas_por_turma"

                print(f"    Navegando para: {report_url}")
                page.goto(report_url, wait_until="domcontentloaded")
                page.wait_for_timeout(3000)

                # Espera o botão CONSULTAR aparecer e clica
                print("    Aguardando botão CONSULTAR...")
                page.wait_for_selector('button:has-text("CONSULTAR")', timeout=30000)
                print("    Clicando CONSULTAR...")
                page.click('button:has-text("CONSULTAR")')

                # Aguarda relatório carregar
                try:
                    # Aguarda o conteúdo do relatório aparecer
                    print("    Aguardando relatório carregar...")

                    # Espera o título do relatório aparecer (pode estar em iframe)
                    # Usa wait_for_timeout + verificação manual
                    page.wait_for_timeout(5000)

                    # Tenta encontrar iframe com dados
                    frames = page.frames
                    dados_encontrados = False

                    for i in range(12):  # Tenta por até 60 segundos (12 x 5s)
                        for frame in frames:
                            try:
                                content = frame.locator("body").inner_text(timeout=2000)
                                if "Total da série" in content:
                                    dados_encontrados = True
                                    print("    Dados encontrados no frame!")
                                    break
                            except:
                                continue

                        if dados_encontrados:
                            break

                        print(f"    Aguardando dados... ({(i+1)*5}s)")
                        page.wait_for_timeout(5000)
                        frames = page.frames  # Atualiza lista de frames

                    if not dados_encontrados:
                        raise PlaywrightTimeout("Dados não encontrados nos frames")

                    page.wait_for_timeout(2000)
                    print("    Dados carregados!")

                    # Extrai dados
                    turmas = extrair_via_snapshot(page)

                    dados["unidades"].append({
                        "codigo": unidade["codigo"],
                        "nome": unidade["nome"],
                        "turmas": turmas
                    })

                    print(f"  Extraídas {len(turmas)} turmas")

                except PlaywrightTimeout as e:
                    print(f"  ERRO: Timeout ao carregar relatório para {unidade['nome']}")
                    # Salva screenshot para debug
                    screenshot_path = OUTPUT_DIR / f"erro_{unidade['codigo']}_{datetime.now().strftime('%H%M%S')}.png"
                    page.screenshot(path=str(screenshot_path))
                    print(f"    Screenshot salvo: {screenshot_path}")
                    dados["unidades"].append({
                        "codigo": unidade["codigo"],
                        "nome": unidade["nome"],
                        "turmas": [],
                        "erro": str(e)
                    })

        except Exception as e:
            print(f"\nERRO: {e}")
            raise

        finally:
            browser.close()

    # 3. Salva dados
    print("\n[5/5] Salvando dados...")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # JSON completo
    json_path = OUTPUT_DIR / f"vagas_{timestamp}.json"
    salvar_json(dados, json_path)

    # JSON resumo
    resumo = gerar_resumo(dados)
    resumo_path = OUTPUT_DIR / f"resumo_{timestamp}.json"
    salvar_json(resumo, resumo_path)

    # SQLite
    db_path = OUTPUT_DIR / "vagas.db"
    salvar_sqlite(dados, db_path)

    # Link para último arquivo
    ultimo_json = OUTPUT_DIR / "vagas_ultimo.json"
    ultimo_resumo = OUTPUT_DIR / "resumo_ultimo.json"

    import shutil
    shutil.copy(json_path, ultimo_json)
    shutil.copy(resumo_path, ultimo_resumo)

    # 4. Imprime resumo
    print("\n" + "=" * 60)
    print("RESUMO")
    print("=" * 60)

    for unidade in resumo["unidades"]:
        print(f"\n{unidade['nome']}:")
        for seg, valores in unidade["segmentos"].items():
            print(f"  {seg}: {valores['matriculados']} matriculados / {valores['vagas']} vagas")
        print(f"  TOTAL: {unidade['total']['matriculados']} matriculados")

    print(f"\nTOTAL GERAL: {resumo['total_geral']['matriculados']} matriculados / {resumo['total_geral']['vagas']} vagas")
    print("=" * 60)

    return dados


if __name__ == "__main__":
    main()
