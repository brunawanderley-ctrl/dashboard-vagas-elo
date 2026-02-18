"""Agente de auditoria automatica para dados de matriculas.

Roda apos cada atualizacao para validar integridade dos dados.
Pode ser chamado standalone ou importado por outros scripts.

Uso standalone:
    python -m utils.auditor          # audita tudo
    python -m utils.auditor --fix    # audita e corrige o que puder
"""

import json
import sqlite3
from datetime import datetime
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent.parent / "output"

# Localidades esperadas por unidade
LOCALIDADE_ESPERADA = {
    "01-BV": ["Boa Viagem"],
    "02-CD": ["Candeias", "Jaboatão"],
    "03-JG": ["Janga", "Paulista"],
    "04-CDR": ["Cordeiro"],
}

# Ranges razoaveis de matriculados por unidade (para detectar anomalias)
RANGES_MATRICULADOS = {
    "01-BV": (800, 1500),
    "02-CD": (800, 1500),
    "03-JG": (500, 1100),
    "04-CDR": (400, 1000),
}


class AuditoriaResultado:
    """Resultado de uma auditoria."""

    def __init__(self):
        self.erros_criticos = []
        self.avisos = []
        self.info = []
        self.dados_ok = True

    def critico(self, msg):
        self.erros_criticos.append(msg)
        self.dados_ok = False

    def aviso(self, msg):
        self.avisos.append(msg)

    def ok(self, msg):
        self.info.append(msg)

    def resumo(self):
        linhas = []
        linhas.append("=" * 60)
        linhas.append("AUDITORIA DE DADOS - MATRICULAS")
        linhas.append(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        linhas.append("=" * 60)

        if self.erros_criticos:
            linhas.append(f"\nERROS CRITICOS ({len(self.erros_criticos)}):")
            for e in self.erros_criticos:
                linhas.append(f"  [ERRO] {e}")

        if self.avisos:
            linhas.append(f"\nAVISOS ({len(self.avisos)}):")
            for a in self.avisos:
                linhas.append(f"  [AVISO] {a}")

        if self.info:
            linhas.append(f"\nOK ({len(self.info)}):")
            for i in self.info:
                linhas.append(f"  [OK] {i}")

        status = "APROVADO" if self.dados_ok else "REPROVADO"
        linhas.append(f"\nRESULTADO: {status}")
        linhas.append("=" * 60)
        return "\n".join(linhas)


def auditar_json(json_path: Path) -> AuditoriaResultado:
    """Audita um arquivo JSON de extracao."""
    resultado = AuditoriaResultado()

    if not json_path.exists():
        resultado.critico(f"Arquivo nao encontrado: {json_path}")
        return resultado

    with open(json_path, "r", encoding="utf-8") as f:
        dados = json.load(f)

    unidades = dados.get("unidades", [])

    # 1. Verificar numero de unidades
    if len(unidades) != 4:
        resultado.critico(f"Esperava 4 unidades, encontrou {len(unidades)}")
    else:
        resultado.ok(f"4 unidades presentes")

    # 2. Verificar duplicacao entre unidades
    turmas_por_unidade = {}
    totais_por_unidade = {}
    for unidade in unidades:
        cod = unidade.get("codigo", "?")
        turmas = unidade.get("turmas", [])
        turmas_hash = frozenset(t.get("turma", "") for t in turmas)
        turmas_por_unidade[cod] = turmas_hash
        totais_por_unidade[cod] = sum(t.get("matriculados", 0) for t in turmas)

    codigos = list(turmas_por_unidade.keys())
    for i, cod1 in enumerate(codigos):
        for cod2 in codigos[i + 1:]:
            if turmas_por_unidade[cod1] == turmas_por_unidade[cod2] and len(turmas_por_unidade[cod1]) > 0:
                resultado.critico(
                    f"DUPLICACAO: {cod1} e {cod2} tem turmas identicas "
                    f"({totais_por_unidade[cod1]} matriculados cada)"
                )
            elif totais_por_unidade.get(cod1) == totais_por_unidade.get(cod2) and totais_por_unidade.get(cod1, 0) > 0:
                resultado.aviso(f"{cod1} e {cod2} tem o mesmo total de matriculados ({totais_por_unidade[cod1]})")

    if not any("DUPLICACAO" in e for e in resultado.erros_criticos):
        resultado.ok("Nenhuma duplicacao entre unidades")

    # 3. Verificar localidade nos nomes das turmas
    erros_loc = 0
    for unidade in unidades:
        cod = unidade.get("codigo", "")
        locs = LOCALIDADE_ESPERADA.get(cod, [])
        if not locs:
            continue
        for turma in unidade.get("turmas", []):
            nome = turma.get("turma", "")
            if not nome:
                continue
            # Verifica se contem localidade de OUTRA unidade
            for outro_cod, outras_locs in LOCALIDADE_ESPERADA.items():
                if outro_cod != cod and any(ol.lower() in nome.lower() for ol in outras_locs):
                    resultado.critico(
                        f"LOCALIDADE ERRADA: {cod} contem turma '{nome}' que pertence a {outro_cod}"
                    )
                    erros_loc += 1
                    if erros_loc >= 5:
                        break
            if erros_loc >= 5:
                resultado.critico(f"... e mais turmas com localidade errada (mostrando apenas 5)")
                break

    if erros_loc == 0:
        resultado.ok("Localidades das turmas consistentes")

    # 4. Verificar ranges de matriculados
    for cod, total in totais_por_unidade.items():
        range_esperado = RANGES_MATRICULADOS.get(cod)
        if range_esperado:
            min_val, max_val = range_esperado
            if total < min_val:
                resultado.aviso(f"{cod}: {total} matriculados (abaixo do esperado: {min_val}-{max_val})")
            elif total > max_val:
                resultado.aviso(f"{cod}: {total} matriculados (acima do esperado: {min_val}-{max_val})")
            else:
                resultado.ok(f"{cod}: {total} matriculados (dentro do range)")

    # 5. Verificar unidades vazias
    for unidade in unidades:
        cod = unidade.get("codigo", "?")
        turmas = unidade.get("turmas", [])
        if len(turmas) == 0:
            if unidade.get("erro"):
                resultado.critico(f"{cod}: 0 turmas (erro: {unidade['erro']})")
            else:
                resultado.critico(f"{cod}: 0 turmas extraidas")

    # 6. Verificar turmas fantasma
    fantasmas = 0
    for unidade in unidades:
        for turma in unidade.get("turmas", []):
            if turma.get("vagas", 0) == 0 and turma.get("matriculados", 0) == 0:
                fantasmas += 1

    if fantasmas > 0:
        resultado.aviso(f"{fantasmas} turmas fantasma (vagas=0, matriculados=0)")

    return resultado


def auditar_sqlite(db_path: Path) -> AuditoriaResultado:
    """Audita o banco SQLite."""
    resultado = AuditoriaResultado()

    if not db_path.exists():
        resultado.critico(f"Banco nao encontrado: {db_path}")
        return resultado

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 1. Ultima extracao
    cursor.execute("SELECT MAX(id), data_extracao FROM extrações")
    row = cursor.fetchone()
    if not row or not row[0]:
        resultado.critico("Nenhuma extracao no banco")
        conn.close()
        return resultado

    ext_id, data_ext = row
    resultado.ok(f"Ultima extracao: id={ext_id}, data={data_ext}")

    # 2. Totais por unidade na ultima extracao
    cursor.execute("""
        SELECT unidade_codigo, COUNT(*) as turmas, SUM(matriculados) as total
        FROM vagas WHERE extracao_id = ?
        GROUP BY unidade_codigo
    """, (ext_id,))

    totais = {}
    for cod, turmas, total in cursor.fetchall():
        totais[cod] = {"turmas": turmas, "matriculados": total}

    if len(totais) != 4:
        resultado.critico(f"Ultima extracao tem {len(totais)} unidades (esperava 4)")

    # 3. Verificar duplicacao no banco
    valores = list(totais.values())
    for i, (cod1, v1) in enumerate(totais.items()):
        for cod2, v2 in list(totais.items())[i + 1:]:
            if v1["turmas"] == v2["turmas"] and v1["matriculados"] == v2["matriculados"]:
                resultado.critico(f"DUPLICACAO no banco: {cod1} e {cod2} identicos")

    # 4. Verificar ranges
    for cod, vals in totais.items():
        range_esperado = RANGES_MATRICULADOS.get(cod)
        if range_esperado:
            min_val, max_val = range_esperado
            if vals["matriculados"] < min_val or vals["matriculados"] > max_val:
                resultado.aviso(
                    f"{cod}: {vals['matriculados']} matriculados fora do range ({min_val}-{max_val})"
                )
            else:
                resultado.ok(f"{cod}: {vals['matriculados']} matriculados OK")

    conn.close()
    return resultado


def auditar_consistencia_json_db() -> AuditoriaResultado:
    """Compara JSON ultimo com SQLite para verificar consistencia."""
    resultado = AuditoriaResultado()

    json_path = OUTPUT_DIR / "vagas_ultimo.json"
    db_path = OUTPUT_DIR / "vagas.db"

    if not json_path.exists() or not db_path.exists():
        resultado.aviso("Nao foi possivel comparar JSON vs SQLite (arquivos ausentes)")
        return resultado

    # Totais do JSON
    with open(json_path, "r", encoding="utf-8") as f:
        dados_json = json.load(f)

    totais_json = {}
    for unidade in dados_json.get("unidades", []):
        cod = unidade.get("codigo", "")
        totais_json[cod] = sum(t.get("matriculados", 0) for t in unidade.get("turmas", []))

    # Totais do SQLite
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT unidade_codigo, SUM(matriculados)
        FROM vagas WHERE extracao_id = (SELECT MAX(id) FROM extrações)
        GROUP BY unidade_codigo
    """)
    totais_db = {row[0]: row[1] for row in cursor.fetchall()}
    conn.close()

    # Compara
    for cod in set(list(totais_json.keys()) + list(totais_db.keys())):
        vj = totais_json.get(cod, "N/A")
        vd = totais_db.get(cod, "N/A")
        if vj != vd:
            resultado.aviso(f"{cod}: JSON={vj} vs DB={vd} (divergencia - datas diferentes?)")
        else:
            resultado.ok(f"{cod}: JSON e DB concordam ({vj} matriculados)")

    return resultado


def auditar_tudo(verbose=True) -> bool:
    """Roda auditoria completa. Retorna True se tudo OK."""
    resultados = []

    # 1. JSON ultimo
    json_path = OUTPUT_DIR / "vagas_ultimo.json"
    if json_path.exists():
        r = auditar_json(json_path)
        if verbose:
            print(r.resumo())
        resultados.append(r)

    # 2. SQLite
    db_path = OUTPUT_DIR / "vagas.db"
    if db_path.exists():
        r = auditar_sqlite(db_path)
        if verbose:
            print(r.resumo())
        resultados.append(r)

    # 3. Consistencia
    r = auditar_consistencia_json_db()
    if verbose:
        print(r.resumo())
    resultados.append(r)

    # Salvar log
    log_path = OUTPUT_DIR / "auditoria_log.json"
    log = {
        "timestamp": datetime.now().isoformat(),
        "aprovado": all(r.dados_ok for r in resultados),
        "erros_criticos": sum(len(r.erros_criticos) for r in resultados),
        "avisos": sum(len(r.avisos) for r in resultados),
        "detalhes": {
            "erros": [e for r in resultados for e in r.erros_criticos],
            "avisos": [a for r in resultados for a in r.avisos],
        }
    }
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

    return all(r.dados_ok for r in resultados)


if __name__ == "__main__":
    import sys
    ok = auditar_tudo(verbose=True)
    sys.exit(0 if ok else 1)
