"""
Configuração centralizada das planilhas Google Sheets
Todos os IDs, URLs, mapeamentos e autenticação devem ser definidos aqui.

================================================================================
MODO DE ACESSO ÀS PLANILHAS
================================================================================

Este projeto suporta DOIS modos de acesso às planilhas:

1. MODO PÚBLICO (atual - padrão):
   - Usa URLs de exportação pública do Google Sheets
   - Não requer autenticação
   - Ideal para planilhas compartilhadas como "Qualquer pessoa com o link"
   - Mais rápido e simples

2. MODO PRIVADO (preparado para o futuro):
   - Usa a API oficial do Google Sheets via gspread
   - Requer arquivo de credenciais (service account JSON)
   - Necessário quando as planilhas forem tornadas privadas
   - Mais seguro e robusto

Para alternar entre os modos, altere a constante MODO_ACESSO abaixo:
    MODO_ACESSO = "publico"   # Usa URLs públicas (atual)
    MODO_ACESSO = "privado"   # Usa gspread com autenticação (futuro)

Para usar o modo privado:
1. Crie uma Conta de Serviço no Google Cloud Console
2. Baixe o arquivo JSON de credenciais
3. Coloque o arquivo na raiz do projeto (ou caminho configurado em CREDENTIALS_PATH)
4. Compartilhe as planilhas com o e-mail da Conta de Serviço (como Leitor ou Editor)
5. Altere MODO_ACESSO para "privado"
================================================================================
"""

import os
import logging

# Imports opcionais para gspread (modo privado)
# Se não instalados, o modo público continua funcionando normalmente
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False
    gspread = None
    Credentials = None

logger = logging.getLogger(__name__)

# ==============================================================================
# MODO DE ACESSO ÀS PLANILHAS
# ==============================================================================
# Valores possíveis: "publico" ou "privado"
# - "publico": Usa URLs de exportação (atual, recomendado para planilhas públicas)
# - "privado": Usa gspread com autenticação (para planilhas privadas)
MODO_ACESSO = "publico"

# ==============================================================================
# CONFIGURAÇÕES DE AUTENTICAÇÃO (gspread - modo privado)
# ==============================================================================
# Escopos de permissão necessários para leitura de planilhas
GSPREAD_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive.readonly"
]

# Caminho para o arquivo de credenciais (service account JSON)
# Pode ser sobrescrito pela variável de ambiente GOOGLE_CREDENTIALS_PATH
CREDENTIALS_PATH = os.getenv(
    "GOOGLE_CREDENTIALS_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
)

# ==============================================================================
# IDs DAS PLANILHAS GOOGLE SHEETS
# ==============================================================================
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
    },
    "ignorados": {
        "id": "13QBNlk9M435Jos0Q-U77tuF-BINXSycmJmV8rZOk92o",
        "nome": "SKUs Ignorados",
        "descricao": "SKUs que devem ser excluídos de todos os cálculos",
        "coluna_sku": "SKU"  # Coluna que contém os SKUs ignorados
    }
}

# ==============================================================================
# URLs CALCULADAS (evita reconstrução no app.py)
# ==============================================================================
# URL base para exportação em formato Excel (modo público)
URL_EXPORT_BASE = "https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

# URL base para edição no navegador (referência)
URL_EDIT_BASE = "https://docs.google.com/spreadsheets/d/{sheet_id}/edit"

# URLs de exportação pré-calculadas para cada planilha
URLS_EXPORTACAO = {
    chave: URL_EXPORT_BASE.format(sheet_id=config["id"])
    for chave, config in PLANILHAS.items()
}

# URLs de edição pré-calculadas (para referência/debug)
URLS_EDICAO = {
    chave: URL_EDIT_BASE.format(sheet_id=config["id"])
    for chave, config in PLANILHAS.items()
}

# ==============================================================================
# MAPEAMENTOS E CONSTANTES DO DOMÍNIO
# ==============================================================================

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
    "retaguarda": ['SKU'],  # Pelo menos SKU é obrigatório
    "ignorados": ['SKU']  # Coluna com SKUs a serem ignorados
}

# ==============================================================================
# CONSTANTES DE CÁLCULO DE DDV (Demanda Diária de Venda)
# ==============================================================================
# O DDV é calculado somando as colunas de histórico (I até Z) e dividindo por 365 dias
# Colunas I até Z correspondem aos índices 8 até 25 (0-indexed)
COLUNAS_HISTORICO_INICIO = 8   # Índice da coluna I (0-indexed)
COLUNAS_HISTORICO_FIM = 26     # Índice após coluna Z (exclusive, 0-indexed)
DIAS_ANO = 365                 # Dias no ano para cálculo do DDV

# Fórmula do DDV:
# DDV = Soma(colunas_I_ate_Z) / DIAS_ANO
#
# Fórmula da Cobertura de Estoque:
# Cobertura = Estoque_Atual / DDV
#
# SKUs com DDV = 0 devem ser tratados como caso especial (evitar divisão por zero)

# Timeout para downloads (segundos)
TIMEOUT_DOWNLOAD = 120

# TTL do cache (segundos) - 1 hora
CACHE_TTL = 3600

# ==============================================================================
# VERSÃO DO PROJETO
# ==============================================================================
VERSAO = "2.2.0"
DATA_VERSAO = "2026-06-18"

# ==============================================================================
# FUNÇÕES AUXILIARES DE AUTENTICAÇÃO (modo privado)
# ==============================================================================

def obter_url_exportacao(chave_planilha: str) -> str:
    """
    Retorna a URL de exportação em formato Excel para uma planilha.
    
    Args:
        chave_planilha: Chave da planilha no dicionário PLANILHAS 
                       (ex: "draft_pdvs", "estoque_seguranca", "retaguarda", "ignorados")
    
    Returns:
        URL completa de exportação em formato Excel
    
    Raises:
        KeyError: Se a chave não existir em PLANILHAS
    """
    if chave_planilha not in PLANILHAS:
        raise KeyError(f"Planilha '{chave_planilha}' não encontrada. "
                      f"Disponíveis: {list(PLANILHAS.keys())}")
    return URLS_EXPORTACAO[chave_planilha]


def obter_url_edicao(chave_planilha: str) -> str:
    """
    Retorna a URL de edição no navegador para uma planilha.
    
    Args:
        chave_planilha: Chave da planilha no dicionário PLANILHAS
    
    Returns:
        URL completa de edição
    """
    if chave_planilha not in PLANILHAS:
        raise KeyError(f"Planilha '{chave_planilha}' não encontrada.")
    return URLS_EDICAO[chave_planilha]


def obter_id_planilha(chave_planilha: str) -> str:
    """
    Retorna o ID da planilha no Google Sheets.
    
    Args:
        chave_planilha: Chave da planilha no dicionário PLANILHAS
    
    Returns:
        ID da planilha (string alfanumérica)
    """
    if chave_planilha not in PLANILHAS:
        raise KeyError(f"Planilha '{chave_planilha}' não encontrada.")
    return PLANILHAS[chave_planilha]["id"]


def esta_no_modo_privado() -> bool:
    """
    Verifica se o sistema está configurado para usar o modo privado (gspread).
    
    Returns:
        True se MODO_ACESSO for "privado", False caso contrário
    """
    return MODO_ACESSO.lower().strip() == "privado"


def verificar_disponibilidade_gspread() -> bool:
    """
    Verifica se as bibliotecas necessárias para o modo privado estão instaladas.
    
    Returns:
        True se gspread e google-auth estiverem disponíveis
    """
    return GSPREAD_AVAILABLE


def obter_cliente_gspread():
    """
    Cria e retorna um cliente gspread autenticado via service account.
    
    Esta função é usada APENAS quando MODO_ACESSO = "privado".
    Para o modo público, as URLs de exportação são usadas diretamente.
    
    Returns:
        gspread.Client: Cliente autenticado do gspread
    
    Raises:
        RuntimeError: Se gspread não estiver instalado
        FileNotFoundError: Se o arquivo de credenciais não existir
        Exception: Se houver erro na autenticação
    """
    if not GSPREAD_AVAILABLE:
        raise RuntimeError(
            "gspread e/ou google-auth não estão instalados. "
            "Instale com: pip install gspread google-auth"
        )
    
    if not os.path.exists(CREDENTIALS_PATH):
        raise FileNotFoundError(
            f"Arquivo de credenciais não encontrado em: {CREDENTIALS_PATH}\n"
            f"Para usar o modo privado, crie uma Conta de Serviço no Google Cloud Console "
            f"e coloque o arquivo JSON em: {CREDENTIALS_PATH}\n"
            f"Ou defina a variável de ambiente GOOGLE_CREDENTIALS_PATH com o caminho correto."
        )
    
    try:
        creds = Credentials.from_service_account_file(
            CREDENTIALS_PATH, 
            scopes=GSPREAD_SCOPES
        )
        client = gspread.authorize(creds)
        logger.info(f"Cliente gspread autenticado com sucesso via {CREDENTIALS_PATH}")
        return client
    except Exception as e:
        logger.error(f"Erro ao autenticar com gspread: {str(e)}")
        raise


def obter_workbook(sheet_id: str):
    """
    Abre e retorna um workbook do Google Sheets via gspread.
    
    Esta função é usada APENAS quando MODO_ACESSO = "privado".
    
    Args:
        sheet_id: ID da planilha no Google Sheets
    
    Returns:
        gspread.Spreadsheet: Objeto da planilha aberta
    
    Raises:
        gspread.SpreadsheetNotFound: Se a planilha não for encontrada
        Exception: Se houver erro na abertura
    """
    client = obter_cliente_gspread()
    try:
        workbook = client.open_by_key(sheet_id)
        logger.info(f"Planilha aberta com sucesso: {sheet_id} ({workbook.title})")
        return workbook
    except Exception as e:
        logger.error(f"Erro ao abrir planilha {sheet_id}: {str(e)}")
        raise


def obter_workbook_por_nome(nome_planilha: str):
    """
    Abre e retorna um workbook do Google Sheets pelo nome.
    
    Args:
        nome_planilha: Nome da planilha no Google Drive
    
    Returns:
        gspread.Spreadsheet: Objeto da planilha aberta
    """
    client = obter_cliente_gspread()
    try:
        workbook = client.open(nome_planilha)
        logger.info(f"Planilha aberta com sucesso: {nome_planilha}")
        return workbook
    except Exception as e:
        logger.error(f"Erro ao abrir planilha '{nome_planilha}': {str(e)}")
        raise


def diagnosticar_configuracao() -> dict:
    """
    Retorna um dicionário com o diagnóstico da configuração atual.
    Útil para debug e para exibir informações no dashboard.
    
    Returns:
        dict: Dicionário com informações de diagnóstico
    """
    return {
        "modo_acesso": MODO_ACESSO,
        "usando_gspread": esta_no_modo_privado(),
        "gspread_disponivel": GSPREAD_AVAILABLE,
        "credentials_path": CREDENTIALS_PATH,
        "credentials_existe": os.path.exists(CREDENTIALS_PATH) if CREDENTIALS_PATH else False,
        "versao": VERSAO,
        "data_versao": DATA_VERSAO,
        "planilhas_configuradas": list(PLANILHAS.keys()),
        "total_pdvs": len(DE_PARA_LOJAS),
        "total_marcas": len(NOMES_MARCAS),
    }