"""Constantes compartilhadas do dashboard."""

# Mapeamento de códigos de unidade para nomes
UNIDADES_MAP = {
    "01-BV": "Boa Viagem",
    "02-CD": "Jaboatão",
    "03-JG": "Paulista",
    "04-CDR": "Cordeiro",
}

# Ordem de exibição das unidades
ORDEM_UNIDADES = ["Boa Viagem", "Jaboatão", "Paulista", "Cordeiro"]

# Segmentos
SEGMENTOS = ["Ed. Infantil", "Fund. I", "Fund. II", "Ens. Médio"]

# Ordem de séries
ORDEM_SERIES = [
    "Infantil II", "Infantil III", "Infantil IV", "Infantil V",
    "1º Ano", "2º Ano", "3º Ano", "4º Ano", "5º Ano",
    "6º Ano", "7º Ano", "8º Ano", "9º Ano",
    "1ª Série", "2ª Série", "3ª Série",
]

# Mapeamento legado: normaliza nomes antigos "Xº Ano Médio" para "Xª Série"
NORMALIZAR_SERIE_EM = {
    "1º Ano Médio": "1ª Série",
    "2º Ano Médio": "2ª Série",
    "3º Ano Médio": "3ª Série",
}

# Progressão de séries (2025 → 2026)
PROGRESSAO = {
    "Infantil II": "Infantil III",
    "Infantil III": "Infantil IV",
    "Infantil IV": "Infantil V",
    "Infantil V": "1º Ano",
    "1º Ano": "2º Ano",
    "2º Ano": "3º Ano",
    "3º Ano": "4º Ano",
    "4º Ano": "5º Ano",
    "5º Ano": "6º Ano",
    "6º Ano": "7º Ano",
    "7º Ano": "8º Ano",
    "8º Ano": "9º Ano",
    "9º Ano": "1ª Série",
    "1ª Série": "2ª Série",
    "2ª Série": "3ª Série",
    "3ª Série": "Formado",
}

# Mapeamento de série para segmento
SERIE_SEGMENTO = {
    "Infantil II": "Ed. Infantil",
    "Infantil III": "Ed. Infantil",
    "Infantil IV": "Ed. Infantil",
    "Infantil V": "Ed. Infantil",
    "1º Ano": "Fund. I",
    "2º Ano": "Fund. I",
    "3º Ano": "Fund. I",
    "4º Ano": "Fund. I",
    "5º Ano": "Fund. I",
    "6º Ano": "Fund. II",
    "7º Ano": "Fund. II",
    "8º Ano": "Fund. II",
    "9º Ano": "Fund. II",
    "1ª Série": "Ens. Médio",
    "2ª Série": "Ens. Médio",
    "3ª Série": "Ens. Médio",
}

# Cores do tema (Power BI dark style)
CORES = {
    "primaria": "#667eea",
    "secundaria": "#764ba2",
    "verde": "#4ade80",
    "vermelho": "#f87171",
    "amarelo": "#fbbf24",
    "azul": "#60a5fa",
    "roxo": "#a78bfa",
    "texto": "#1e293b",
    "texto_secundario": "#64748b",
    "fundo": "#f8f9fa",
    "fundo_card": "#ffffff",
}

# Paleta de cores para gráficos
CORES_UNIDADES = {
    "Boa Viagem": "#667eea",
    "Jaboatão": "#764ba2",
    "Paulista": "#4ade80",
    "Cordeiro": "#fbbf24",
}

CORES_SEGMENTOS = {
    "Ed. Infantil": "#a78bfa",
    "Fund. I": "#4ade80",
    "Fund. II": "#60a5fa",
    "Ens. Médio": "#fbbf24",
}

# Metas de matrícula 2026 por unidade
METAS_2026 = {
    "Boa Viagem": 1230,
    "Jaboatão": 1200,
    "Paulista": 850,
    "Cordeiro": 800,
}
META_TOTAL = 4080

# Séries por segmento (para tabela cruzada)
SERIES_POR_SEGMENTO = {
    "Ed. Infantil": ["Infantil II", "Infantil III", "Infantil IV", "Infantil V"],
    "Fund. I": ["1º Ano", "2º Ano", "3º Ano", "4º Ano", "5º Ano"],
    "Fund. II": ["6º Ano", "7º Ano", "8º Ano", "9º Ano"],
    "Ens. Médio": ["1ª Série", "2ª Série", "3ª Série"],
}

# Status de ocupação para ensalamento
STATUS_OCUPACAO = {
    "critica": {"label": "Crítica", "cor": "#f87171", "min": 0, "max": 40},
    "atencao": {"label": "Atenção", "cor": "#fbbf24", "min": 40, "max": 70},
    "normal": {"label": "Normal", "cor": "#4ade80", "min": 70, "max": 90},
    "lotada": {"label": "Lotada", "cor": "#60a5fa", "min": 90, "max": 100},
    "super": {"label": "Superlotada", "cor": "#a78bfa", "min": 100, "max": 999},
}
