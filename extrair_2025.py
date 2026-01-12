#!/usr/bin/env python3
"""
SIGA Activesoft - Extrator de Dados 2025
Extrai dados do ano letivo 2025 para comparação com 2026
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
    "periodo": "2025",  # ANO 2025
    "unidades": [
        {"id": "17", "nome": "1 - BV (Boa Viagem)", "codigo": "01-BV"},
        {"id": "18", "nome": "2 - CD (Jaboatão)", "codigo": "02-CD"},
        {"id": "19", "nome": "3 - JG (Paulista)", "codigo": "03-JG"},
        {"id": "20", "nome": "4 - CDR (Cordeiro)", "codigo": "04-CDR"},
    ],
    # Cursos a IGNORAR
    "cursos_ignorar": [
        "esporte", "ballet", "futsal", "judô", "judo", "voleibol", "basquete",
        "ginástica", "ginastica", "karatê", "karate",
        "integral", "complementar",
        "lanche saudável", "lanche saudavel",
        "curso livre", "cursos livres",
        "transporte"
    ]
}

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def deve_ignorar_curso(nome_turma: str) -> bool:
    nome_lower = nome_turma.lower()
    for termo in CONFIG["cursos_ignorar"]:
        if termo in nome_lower:
            return True
    return False


def identificar_segmento(nome_curso: str) -> str:
    nome_lower = nome_curso.lower()
    if "infantil" in nome_lower:
        return "Ed. Infantil"
    elif "fundamental ii" in nome_lower or "fundamental 2" in nome_lower:
        return "Fund. II"
    elif "fundamental i" in nome_lower or "fundamental 1" in nome_lower:
        return "Fund. I"
    elif "médio" in nome_lower or "medio" in nome_lower:
        return "Ens. Médio"
    else:
        return "Outro"


def parse_numero(texto: str) -> int:
    try:
        return int(texto.replace(".", "").replace(",", "").strip())
    except (ValueError, AttributeError):
        return 0


def extrair_via_snapshot(page, ano: str = "2025") -> list:
    """Extrai dados usando o texto do snapshot"""
    turmas = []

    try:
        frames = page.frames
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

        # Detecta header de curso/série (ex: "1- BV - Educação Infantil - ... / 2025")
        if f"/ {ano}" in linha and not linha.startswith("Total"):
            curso_atual = linha
            continue

        if linha.startswith("Total da série") or linha.startswith("Total geral"):
            continue

        partes = linha_raw.split("\t")
        if len(partes) >= 2 and curso_atual:
            nome_turma = partes[0].strip()

            if nome_turma in ["Turma - Turno", "(A)", "Vagas abertas", "Novatos"]:
                continue

            numeros = []
            for p in partes[1:]:
                p = p.strip()
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


def salvar_json(dados: dict, json_path: Path):
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    print(f"  Dados salvos: {json_path}")


def gerar_resumo(dados: dict) -> dict:
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
            "total": {"vagas": 0, "novatos": 0, "veteranos": 0, "matriculados": 0, "disponiveis": 0}
        }

        for turma in unidade.get("turmas", []):
            seg = turma["segmento"]
            if seg not in resumo_unidade["segmentos"]:
                resumo_unidade["segmentos"][seg] = {"vagas": 0, "novatos": 0, "veteranos": 0, "matriculados": 0, "disponiveis": 0}

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

    resumo["total_geral"] = {
        "vagas": sum(u["total"]["vagas"] for u in resumo["unidades"]),
        "novatos": sum(u["total"]["novatos"] for u in resumo["unidades"]),
        "veteranos": sum(u["total"]["veteranos"] for u in resumo["unidades"]),
        "matriculados": sum(u["total"]["matriculados"] for u in resumo["unidades"]),
        "disponiveis": sum(u["total"]["disponiveis"] for u in resumo["unidades"]),
    }

    return resumo


def main():
    print("=" * 60)
    print("SIGA - Extrator de Dados 2025 (Comparativo)")
    print(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 60)

    dados = {
        "data_extracao": datetime.now().isoformat(),
        "periodo": CONFIG["periodo"],
        "unidades": []
    }

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            # Login
            print("\n[1/5] Fazendo login...")
            page.goto(CONFIG["url"], wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            page.fill('#codigoInstituicao', CONFIG["instituicao"])
            page.fill('#id_login', CONFIG["login"])
            page.fill('#id_senha', CONFIG["senha"])
            page.click('button:has-text("ENTRAR")')

            page.wait_for_url("**/login/unidade/**", timeout=15000)
            page.wait_for_timeout(2000)
            print("  Login realizado!")

            # Para cada unidade
            for idx, unidade in enumerate(CONFIG["unidades"]):
                print(f"\n[{idx+2}/5] Processando {unidade['nome']}...")

                if idx == 0:
                    page.click(f'text={unidade["nome"]}')
                    page.wait_for_timeout(5000)
                else:
                    page.locator('button:has-text("Unidade"):has-text("keyboard_arrow_down")').click()
                    page.wait_for_timeout(1000)
                    page.locator(f'button:has-text("{unidade["nome"]}")').filter(has_text="swap_vert").click()
                    page.wait_for_timeout(5000)

                # Navega para relatório
                current_url = page.url
                base_url = current_url.split('/')[0] + '//' + current_url.split('/')[2]

                # Navega para relatório
                report_url = f"{base_url}/busca_central_relatorios/?relatorio=aluno_turma/resumo_vagas_por_turma"
                page.goto(report_url, wait_until="domcontentloaded")
                page.wait_for_timeout(3000)

                # Captura screenshot para debug
                screenshot_path = OUTPUT_DIR / f"debug_relatorio_{unidade['codigo']}.png"
                page.screenshot(path=str(screenshot_path))
                print(f"    Screenshot salvo: {screenshot_path}")

                # Tenta selecionar período 2025 no filtro do relatório
                # O sistema usa dropdown customizado (não HTML select nativo)
                periodo_selecionado = False

                try:
                    # Procura o campo Período que mostra "2026"
                    # Clica no dropdown de período para abrir as opções
                    periodo_dropdown = page.locator('div:has-text("Período")').locator('..').locator('input, div[role="combobox"], div[role="listbox"]').first

                    # Tenta clicar no campo que contém "2026"
                    campo_2026 = page.locator('input[value="2026"], div:text-is("2026")').first
                    if campo_2026.is_visible():
                        campo_2026.click()
                        page.wait_for_timeout(1000)
                        print("    Clicou no campo de período")

                        # Procura opção 2025 no dropdown aberto
                        opcao_2025 = page.locator('div[role="option"]:has-text("2025"), li:has-text("2025"), span:text-is("2025"), div:text-is("2025")').first
                        if opcao_2025.is_visible():
                            opcao_2025.click()
                            periodo_selecionado = True
                            print("    Período 2025 selecionado!")
                            page.wait_for_timeout(2000)
                except Exception as e:
                    print(f"    Método 1 falhou: {e}")

                # Método 2: Tenta localizar pelo label "Período" e clicar no campo adjacente
                if not periodo_selecionado:
                    try:
                        # Material Design dropdown - clica no container do período
                        page.click('text=Período', timeout=2000)
                        page.wait_for_timeout(500)
                        # Agora clica no valor 2026 para abrir dropdown
                        page.click('text=2026', timeout=2000)
                        page.wait_for_timeout(1000)
                        # Seleciona 2025
                        page.click('text=2025', timeout=2000)
                        periodo_selecionado = True
                        print("    Período 2025 selecionado via cliques!")
                        page.wait_for_timeout(2000)
                    except Exception as e:
                        print(f"    Método 2 falhou: {e}")

                # Método 3: Tenta clicar diretamente no dropdown e digitar
                if not periodo_selecionado:
                    try:
                        # Procura input dentro de um container com label Período
                        inputs = page.locator('input').all()
                        for inp in inputs:
                            try:
                                val = inp.input_value()
                                if val == '2026':
                                    inp.click()
                                    page.wait_for_timeout(500)
                                    inp.fill('2025')
                                    inp.press('Enter')
                                    periodo_selecionado = True
                                    print("    Período 2025 digitado no input!")
                                    page.wait_for_timeout(2000)
                                    break
                            except:
                                continue
                    except Exception as e:
                        print(f"    Método 3 falhou: {e}")

                if not periodo_selecionado:
                    print("    AVISO: Não foi possível selecionar período 2025")
                    print("    O relatório usará o período padrão (2026)")

                # Captura screenshot após tentativa de seleção
                screenshot_path = OUTPUT_DIR / f"debug_apos_periodo_{unidade['codigo']}.png"
                page.screenshot(path=str(screenshot_path))

                page.wait_for_timeout(1000)

                page.wait_for_selector('button:has-text("CONSULTAR")', timeout=30000)
                page.click('button:has-text("CONSULTAR")')

                try:
                    page.wait_for_timeout(5000)
                    frames = page.frames
                    dados_encontrados = False

                    for i in range(12):
                        for frame in frames:
                            try:
                                content = frame.locator("body").inner_text(timeout=2000)
                                if "Total da série" in content:
                                    dados_encontrados = True
                                    break
                            except:
                                continue

                        if dados_encontrados:
                            break
                        print(f"    Aguardando... ({(i+1)*5}s)")
                        page.wait_for_timeout(5000)
                        frames = page.frames

                    if not dados_encontrados:
                        raise PlaywrightTimeout("Dados não encontrados")

                    page.wait_for_timeout(2000)
                    turmas = extrair_via_snapshot(page, "2025")

                    dados["unidades"].append({
                        "codigo": unidade["codigo"],
                        "nome": unidade["nome"],
                        "turmas": turmas
                    })

                    print(f"  Extraídas {len(turmas)} turmas")

                except PlaywrightTimeout as e:
                    print(f"  ERRO: Timeout - {unidade['nome']}")
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

    # Salva dados
    print("\n[5/5] Salvando dados 2025...")

    # JSON
    json_path = OUTPUT_DIR / "dados_2025.json"
    salvar_json(dados, json_path)

    # Resumo
    resumo = gerar_resumo(dados)
    resumo_path = OUTPUT_DIR / "resumo_2025.json"
    salvar_json(resumo, resumo_path)

    # Imprime resumo
    print("\n" + "=" * 60)
    print("RESUMO 2025")
    print("=" * 60)

    for unidade in resumo["unidades"]:
        print(f"\n{unidade['nome']}:")
        for seg, valores in unidade.get("segmentos", {}).items():
            print(f"  {seg}: {valores['matriculados']} matr. | {valores['novatos']} nov. | {valores['veteranos']} vet.")
        print(f"  TOTAL: {unidade['total']['matriculados']} matriculados")

    print(f"\nTOTAL GERAL: {resumo['total_geral']['matriculados']} matriculados")
    print("=" * 60)

    return dados


if __name__ == "__main__":
    main()
