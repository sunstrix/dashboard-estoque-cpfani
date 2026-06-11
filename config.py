"""
Configuração centralizada das planilhas Google Sheets
Todos os IDs e URLs devem ser definidos aqui
"""

# IDs das planilhas Google Sheets
PLANILHAS = {
    "draft_pdvs": {
        "id": "1EDDyKie9UiugMLMowcPzHfViqzziFcSgxVPvZ2Rx3L0",
        "nome": "DRAFT_PDVS",
        "descricao": "Dados principais de PDVs, estoque e produtos",
        "abas_esperadas": ["BOTICARIO", "EUDORA", "QUEM_DISSE_BERENICE"]
    },
    "estoque_seguranca": {
        "id": "1uHonFnFM4p7bz4s7YpewhKHNs6fSEfw9rDMTKC7jtHE",
        "nome": "CONSULTA_DE_ESTOQUE",
        "descricao": "Estoque de segurança e mínimos",
        "abas_esperadas": ["BOT", "EUD", "QDB"]
    },
    "retaguarda": {
        "id": "11Z21gFvJ9pm2xSlF3IweC7xcYZwAZWrjcWDnRe5LexY",
        "nome": "Planilha Retaguarda",
        "descricao": "Custos dos produtos",
        "coluna_custo": "CUSTO"  # Nome da coluna que contém o valor de custo
    }
}

# Mapeamento de PDVs
DE_PARA_LOJAS = {
    4842: "4842 - Metrópole", 5152: "5152 - Coração", 6105: "6105 - Assai Anchieta",
    6106: "6106 - Direita", 6110: "6110 - Arouche", 8001: "8001 - Dom José",
    11576: "11576 - Davó", 12055: "12055 - São Bento", 12056: "12056 - Marechal",
    12605: "12605 - Coop", 12645: "12645 - Light", 14120: "14120 - VD SBC",
    14353: "14353 - VD SP", 20371: "20371 - Luz", 21502: "21502 - Bem Barato",
    23000: "23000 - Outlet", 23379: "23379 - Assai Piraporinha"
}

# Mapeamento reverso
DE_PARA_LOJAS_REVERSO = {v: k for k, v in DE_PARA_LOJAS.items()}

# Mapeamento de nomes de loja da planilha de retaguarda para PDVs
MAPEAMENTO_PDV_DRAFT_RAW = {
    'Loja: 4842 - N. S. F. COSMETICOS E PRESENTES LTDA': 4842,
    'Loja: 5152 - N. S. F. COSMETICOS E PRESENTES LTDA': 5152,
    'Loja: 6105 - N. S. F. COSMETICOS E PRESENTES LTDA': 6105,
    'Loja: 6106 - N. S. F. COSMETICOS E PRESENTES LTDA': 6106,
    'Loja: 6110 - N. S. F. COSMETICOS E PRESENTES LTDA': 6110,
    'Loja: 8001 - N. S. F. COSMETICOS E PRESENTES LTDA': 8001,
    'Loja: 11576 - N. S. F. COSMETICOS E PRESENTES LTDA': 11576,
    'Loja: 12055 - N. S. F. COSMETICOS E PRESENTES LTDA': 12055,
    'Loja: 12056 - S. P. ARON COSMETICOS EPP': 12056,
    'Loja: 12605 - N.S.F. COSMETICOS E PRESENTES LTDA.': 12605,
    'Loja: 12645 - N. S. F. COSMETICOS E PRESENTES LTDA': 12645,
    'Loja: 14120 - ARPEL DISTRIBUIDORA DE COSMETICOS LTDA - EPP': 14120,
    'Loja: 14353 - ARPEL DISTRIBUIDORA DE COSMETICOS LTDA - EPP': 14353,
    'Loja: 20371 - N. S. F. COSMÉTICOS E PRESENTES LTDA.': 20371,
    'Loja: 21502 - N. S. F. COSMETICOS E PRESENTES LTD': 21502,
    'Loja: 23000 - N. S. F. COSMETICOS E PRESENTES LTD': 23000,
    'Loja: 23379 - N. S. F. COSMETICOS E PRESENTES LTD': 23379
}

# Nomes das marcas
NOMES_MARCAS = {
    'BOTICARIO': 'O Boticário',
    'EUDORA': 'Eudora',
    'QUEM_DISSE_BERENICE': 'Quem Disse, Berenice?'
}

# Abas de estoque de segurança
ABAS_SEGURANCA = {'BOT': 'O Boticário', 'EUD': 'Eudora', 'QDB': 'Quem Disse, Berenice?'}

# Logos das marcas
LOGOS_MARCAS = {
    'O Boticário': 'logo_boticario.png',
    'Eudora': 'logo_eudora.png',
    'Quem Disse, Berenice?': 'logo_qdb.png'
}

# Cores das marcas
CORES_MARCAS = {
    'O Boticário': '#007A33',
    'Eudora': '#a855f7',
    'Quem Disse, Berenice?': '#ff4b4b'
}

# Regras de estoque mínimo por classe
REGRAS_ESTOQUE_MINIMO = {'A': 15, 'B': 10, 'C': 5, 'E': 2}

# Colunas obrigatórias em cada planilha
COLUNAS_OBRIGATORIAS = {
    "draft_pdvs": ['PDV', 'SKU', 'Estoque Atual', 'Preço tabela'],
    "estoque_seguranca": ['PDV', 'SKU'],
    "retaguarda": ['SKU']  # Pelo menos SKU é obrigatório
}

# Timeout para downloads (segundos)
TIMEOUT_DOWNLOAD = 120

# TTL do cache (segundos) - 1 hora
CACHE_TTL = 3600