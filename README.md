# SIGA - Extrator de Vagas por Turma

Extrai automaticamente o relatório "Resumo de Vagas por Turma" do SIGA Activesoft para as 4 unidades do Colégio Elo.

## Unidades Processadas
- 01-BV: Boa Viagem
- 02-CD: Jaboatão
- 03-JG: Paulista
- 04-CDR: Cordeiro

## Cursos Extraídos
Apenas cursos regulares:
- Educação Infantil (Infantil II ao V)
- Ensino Fundamental I (1º ao 5º ano)
- Ensino Fundamental II (6º ao 9º ano)
- Ensino Médio (1ª a 3ª série)

**Excluídos:** Esportes, Integral, Lanche Saudável, Cursos Livres, Transporte

## Setup

```bash
cd /Users/anibal37/Documents/siga_vagas_extractor
chmod +x setup.sh
./setup.sh
```

## Uso

```bash
python3 extrair_vagas.py
```

## Output

Arquivos gerados em `./output/`:

| Arquivo | Descrição |
|---------|-----------|
| `vagas_YYYYMMDD_HHMMSS.json` | Dados completos com todas as turmas |
| `resumo_YYYYMMDD_HHMMSS.json` | Resumo por segmento |
| `vagas.db` | SQLite com histórico |
| `vagas_ultimo.json` | Link para última extração |
| `resumo_ultimo.json` | Link para último resumo |

## Cron (execução diária)

```bash
# Editar crontab
crontab -e

# Adicionar linha (executa todo dia às 6h)
0 6 * * * cd /Users/anibal37/Documents/siga_vagas_extractor && /usr/bin/python3 extrair_vagas.py >> output/log.txt 2>&1
```

## Estrutura do JSON de Resumo

```json
{
  "data_extracao": "2026-01-11T05:00:00",
  "periodo": "2026",
  "unidades": [
    {
      "codigo": "01-BV",
      "nome": "1 - BV (Boa Viagem)",
      "segmentos": {
        "Ed. Infantil": {"vagas": 100, "novatos": 20, "veteranos": 30, "matriculados": 50, "disponiveis": 50},
        "Fund. I": {...},
        "Fund. II": {...},
        "Ens. Médio": {...}
      },
      "total": {"vagas": 500, "novatos": 80, "veteranos": 200, "matriculados": 280, "disponiveis": 220}
    }
  ],
  "total_geral": {"vagas": 4000, "novatos": 500, "veteranos": 1400, "matriculados": 1900, "disponiveis": 1200}
}
```

## Para Dashboard

O arquivo `resumo_ultimo.json` é atualizado a cada execução e pode ser consumido diretamente pelo dashboard.

Exemplo com Streamlit:
```python
import json
import streamlit as st

with open("output/resumo_ultimo.json") as f:
    dados = json.load(f)

st.title("Vagas Colégio Elo 2026")
st.json(dados["total_geral"])
```
