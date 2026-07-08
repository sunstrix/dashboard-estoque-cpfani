"""
Utilitários para download de arquivos do SharePoint Online.

Este módulo implementa conversão de URLs de compartilhamento do SharePoint
em URLs de download direto, adicionando o parâmetro download=1.

Solução baseada em: adicionar download=1 à URL original preserva todos os
parâmetros (especialmente o token e=) e força o download do arquivo binário
em vez de abrir o visualizador web HTML.
"""

import logging
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
from io import BytesIO
import traceback

logger = logging.getLogger(__name__)

# ==============================================================================
# CONFIGURAÇÕES GLOBAIS
# ==============================================================================

# User-Agent de navegador real (evita bloqueios do SharePoint)
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# Headers padrão para todas as requisições
HEADERS_PADRAO = {
    "User-Agent": USER_AGENT,
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*;q=0.9",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}

# Timeout padrão (segundos)
TIMEOUT_REQUEST = 120

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def criar_sessao_sharepoint():
    """
    Cria uma sessão requests otimizada para SharePoint com retry automático.
    
    Returns:
        requests.Session: Sessão configurada com retry e headers
    """
    session = requests.Session()
    
    # Configura retry automático
    retry = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    
    # Define headers padrão
    session.headers.update(HEADERS_PADRAO)
    
    return session


def converter_sharepoint_para_download(sharepoint_url: str) -> str:
    """
    Converte link de compartilhamento do SharePoint em URL de download direto
    adicionando parâmetro download=1, preservando todos os parâmetros originais.
    
    Args:
        sharepoint_url: URL de compartilhamento do SharePoint
                       (ex: https://tenant.sharepoint.com/:x:/s/Site/TOKEN?e=abc123)
    
    Returns:
        str: URL de download direto com parâmetro download=1
             ou None se falhar na conversão
    
    Example:
        >>> url = "https://didiernsf.sharepoint.com/:x:/s/NSFcosmticosepresentesLTDA/IQCujrbIbWZLT50lUu7tb2V7Aew2WFZQK1Uo2c4T583mDnU?e=5RIBrD"
        >>> converter_sharepoint_para_download(url)
        'https://didiernsf.sharepoint.com/:x:/s/NSFcosmticosepresentesLTDA/IQCujrbIbWZLT50lUu7tb2V7Aew2WFZQK1Uo2c4T583mDnU?e=5RIBrD&download=1'
    """
    try:
        if not sharepoint_url:
            logger.warning("URL vazia ou None fornecida")
            return None
        
        parsed = urlparse(sharepoint_url)
        
        # Se já é URL de download, retornar como está
        if 'download.aspx' in sharepoint_url or 'download=1' in sharepoint_url:
            logger.info(f"URL já é de download: {sharepoint_url[:80]}...")
            return sharepoint_url
        
        # Extrair query string existente
        query_params = parse_qs(parsed.query, keep_blank_values=True)
        
        # Adicionar parâmetro download=1
        query_params['download'] = ['1']
        
        # Reconstruir URL
        new_query = urlencode(query_params, doseq=True)
        new_url = urlunparse((
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            new_query,
            parsed.fragment
        ))
        
        logger.info(f"✅ URL convertida com sucesso")
        logger.info(f"   Original: {sharepoint_url[:80]}...")
        logger.info(f"   Convertida: {new_url[:80]}...")
        
        return new_url
        
    except Exception as e:
        logger.error(f"❌ Erro ao converter URL: {str(e)}")
        logger.error(traceback.format_exc())
        return None


def baixar_arquivo_sharepoint(url_compartilhamento: str, nome_arquivo: str = "arquivo.xlsx") -> BytesIO:
    """
    Baixa arquivo do SharePoint convertendo automaticamente a URL para download direto.
    
    Args:
        url_compartilhamento: URL de compartilhamento do SharePoint
        nome_arquivo: Nome do arquivo para logging
    
    Returns:
        BytesIO: Conteúdo do arquivo em buffer ou None se falhar
    
    Example:
        >>> url = "https://didiernsf.sharepoint.com/:x:/s/NSFcosmticosepresentesLTDA/..."
        >>> buffer = baixar_arquivo_sharepoint(url, "DRAFT_PDVS.xlsx")
        >>> if buffer:
        ...     df = pd.read_excel(buffer)
    """
    logger.info(f"📥 Iniciando download: {nome_arquivo}")
    logger.info(f"URL original: {url_compartilhamento}")
    
    # Converte URL para download direto
    url_download = converter_sharepoint_para_download(url_compartilhamento)
    
    if not url_download:
        logger.error(f"❌ Não foi possível converter URL para download: {nome_arquivo}")
        return None
    
    # Baixa o arquivo
    session = criar_sessao_sharepoint()
    
    try:
        logger.info(f"Fazendo requisição para URL de download...")
        response = session.get(url_download, timeout=TIMEOUT_REQUEST, stream=True)
        
        # Verifica se o download foi bem-sucedido
        if response.status_code != 200:
            logger.error(f"❌ HTTP {response.status_code}: {response.reason}")
            logger.error(f"URL tentada: {url_download}")
            return None
        
        # Verifica se o conteúdo é realmente um arquivo (não HTML)
        content_type = response.headers.get('Content-Type', '')
        logger.info(f"Content-Type: {content_type}")
        
        if 'text/html' in content_type:
            logger.error(f"❌ SharePoint retornou HTML em vez de arquivo binário")
            logger.error(f"   Content-Type: {content_type}")
            logger.error(f"   Possível causa: arquivo não está público ou URL inválida")
            return None
        
        content = response.content
        tamanho_kb = len(content) / 1024
        
        # Verifica se o arquivo não está vazio
        if len(content) < 100:
            logger.error(f"❌ Arquivo muito pequeno ({len(content)} bytes) - possivelmente vazio ou inválido")
            return None
        
        logger.info(f"✅ Download concluído: {nome_arquivo} ({tamanho_kb:.1f} KB)")
        
        return BytesIO(content)
    
    except requests.exceptions.Timeout:
        logger.error(f"❌ Timeout ao baixar {nome_arquivo} (>{TIMEOUT_REQUEST}s)")
        return None
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ Erro de requisição ao baixar {nome_arquivo}: {str(e)}")
        logger.error(traceback.format_exc())
        return None
    except Exception as e:
        logger.error(f"❌ Erro inesperado ao baixar {nome_arquivo}: {str(e)}")
        logger.error(traceback.format_exc())
        return None


def testar_url_sharepoint(url_compartilhamento: str) -> dict:
    """
    Testa uma URL do SharePoint e retorna informações de diagnóstico.
    
    Args:
        url_compartilhamento: URL de compartilhamento do SharePoint
    
    Returns:
        dict: Dicionário com informações de diagnóstico
    """
    resultado = {
        "url_original": url_compartilhamento,
        "url_convertida": None,
        "sucesso_conversao": False,
        "status_code": None,
        "content_type": None,
        "tamanho_bytes": 0,
        "erro": None
    }
    
    try:
        # Converte URL
        url_download = converter_sharepoint_para_download(url_compartilhamento)
        resultado["url_convertida"] = url_download
        resultado["sucesso_conversao"] = url_download is not None
        
        if not url_download:
            resultado["erro"] = "Falha na conversão da URL"
            return resultado
        
        # Testa download
        session = criar_sessao_sharepoint()
        response = session.get(url_download, timeout=TIMEOUT_REQUEST, stream=True)
        
        resultado["status_code"] = response.status_code
        resultado["content_type"] = response.headers.get('Content-Type', '')
        resultado["tamanho_bytes"] = len(response.content)
        
        if response.status_code == 200 and 'text/html' not in resultado["content_type"]:
            logger.info(f"✅ URL testada com sucesso: {url_compartilhamento[:80]}...")
        else:
            logger.warning(f"⚠️ URL retornou status {response.status_code} ou HTML")
        
    except Exception as e:
        resultado["erro"] = str(e)
        logger.error(f"❌ Erro ao testar URL: {str(e)}")
    
    return resultado