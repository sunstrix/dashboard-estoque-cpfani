"""
Utilitários para download de arquivos do SharePoint Online.

SOLUÇÃO: Adicionar parâmetro download=1 à URL de compartilhamento,
preservando todos os parâmetros originais (especialmente o token e=).

Exemplo:
    Entrada: https://tenant.sharepoint.com/:x:/s/Site/TOKEN?e=abc123
    Saída:   https://tenant.sharepoint.com/:x:/s/Site/TOKEN?e=abc123&download=1

Esta abordagem funciona para arquivos compartilhados publicamente
("Qualquer pessoa com o link pode visualizar").
"""

import re
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

HEADERS_PADRAO = {
    "User-Agent": USER_AGENT,
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
              "application/vnd.ms-excel,application/octet-stream,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
}

TIMEOUT_REQUEST = 120

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def criar_sessao_sharepoint():
    """
    Cria uma sessão requests otimizada para SharePoint com retry automático.
    """
    session = requests.Session()
    
    retry = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    
    session.headers.update(HEADERS_PADRAO)
    
    return session


# ==============================================================================
# FUNÇÃO PRINCIPAL: CONVERSÃO DE URL (SOLUÇÃO download=1)
# ==============================================================================

def converter_sharepoint_para_download(sharepoint_url: str) -> str:
    """
    Converte link de compartilhamento do SharePoint em URL de download direto
    adicionando parâmetro download=1, preservando todos os parâmetros originais.
    
    Args:
        sharepoint_url: URL completa de compartilhamento do SharePoint
    
    Returns:
        str: URL de download direta ou None se falhar
    
    Exemplo:
        >>> url = "https://tenant.sharepoint.com/:x:/s/Site/TOKEN?e=abc123"
        >>> converter_sharepoint_para_download(url)
        'https://tenant.sharepoint.com/:x:/s/Site/TOKEN?e=abc123&download=1'
    """
    try:
        if not sharepoint_url:
            logger.warning("URL vazia fornecida")
            return None
        
        parsed = urlparse(sharepoint_url)
        
        # Validação básica: deve ser HTTPS e ter path
        if not parsed.scheme or not parsed.netloc or not parsed.path:
            logger.error(f"URL inválida: {sharepoint_url}")
            return None
        
        # Se já é URL de download, retornar como está
        if 'download.aspx' in sharepoint_url or 'download=1' in sharepoint_url:
            logger.info(f"URL já é de download: {sharepoint_url}")
            return sharepoint_url
        
        # Extrair query string existente preservando todos os parâmetros
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
        logger.info(f"   Original: {sharepoint_url}")
        logger.info(f"   Download: {new_url}")
        
        return new_url
        
    except Exception as e:
        logger.error(f"❌ Erro ao converter URL: {str(e)}")
        logger.error(traceback.format_exc())
        return None


# Mantida para retrocompatibilidade com código existente
def resolver_url_download_sharepoint(url_compartilhamento: str) -> str:
    """
    Alias para converter_sharepoint_para_download.
    Mantida para compatibilidade com código que já usa este nome.
    """
    return converter_sharepoint_para_download(url_compartilhamento)


# ==============================================================================
# FUNÇÃO DE DOWNLOAD
# ==============================================================================

def baixar_arquivo_sharepoint(url_compartilhamento: str, nome_arquivo: str = "arquivo.xlsx") -> BytesIO:
    """
    Baixa arquivo do SharePoint convertendo automaticamente a URL de download.
    
    Args:
        url_compartilhamento: URL completa de compartilhamento do SharePoint
        nome_arquivo: Nome do arquivo para logging
    
    Returns:
        BytesIO: Conteúdo do arquivo em buffer ou None se falhar
    """
    logger.info(f"📥 Iniciando download SharePoint: {nome_arquivo}")
    logger.info(f"   URL original: {url_compartilhamento}")
    
    # Converter URL para formato de download
    url_download = converter_sharepoint_para_download(url_compartilhamento)
    
    if not url_download:
        logger.error(f"❌ Não foi possível converter URL para download: {nome_arquivo}")
        return None
    
    # Baixar o arquivo
    session = criar_sessao_sharepoint()
    
    try:
        logger.info(f"📡 Fazendo requisição GET para URL de download...")
        response = session.get(url_download, timeout=TIMEOUT_REQUEST, stream=True)
        
        # Verificar status
        if response.status_code != 200:
            logger.error(f"❌ Status code {response.status_code} para {nome_arquivo}")
            logger.error(f"   Content-Type: {response.headers.get('Content-Type', 'N/A')}")
            return None
        
        # Verificar Content-Type (deve ser binário, não HTML)
        content_type = response.headers.get('Content-Type', '').lower()
        logger.info(f"   Content-Type: {content_type}")
        
        # Se retornou HTML, falhou
        if 'text/html' in content_type:
            logger.error(f"❌ SharePoint retornou HTML ao invés do arquivo binário")
            logger.error(f"   Possíveis causas:")
            logger.error(f"   1. Arquivo não está compartilhado como 'Qualquer pessoa com o link'")
            logger.error(f"   2. Link expirou ou foi revogado")
            logger.error(f"   3. SharePoint requer autenticação SSO")
            return None
        
        # Ler conteúdo
        content = response.content
        tamanho_kb = len(content) / 1024
        
        # Validação de tamanho mínimo (arquivos Excel válidos têm pelo menos alguns KB)
        if tamanho_kb < 1:
            logger.warning(f"⚠️ Arquivo muito pequeno ({tamanho_kb:.2f} KB), pode estar corrompido")
        
        logger.info(f"✅ Download SharePoint concluído: {nome_arquivo} ({tamanho_kb:.1f} KB)")
        
        return BytesIO(content)
    
    except requests.exceptions.Timeout:
        logger.error(f"❌ Timeout ao baixar {nome_arquivo} (>{TIMEOUT_REQUEST}s)")
        return None
    except requests.exceptions.ConnectionError as e:
        logger.error(f"❌ Erro de conexão ao baixar {nome_arquivo}: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"❌ Erro inesperado ao baixar {nome_arquivo}: {str(e)}")
        logger.error(traceback.format_exc())
        return None


# ==============================================================================
# FUNÇÃO DE DIAGNÓSTICO
# ==============================================================================

def diagnosticar_url_sharepoint(url_compartilhamento: str) -> dict:
    """
    Realiza diagnóstico completo de uma URL do SharePoint.
    
    Testa a conversão e faz uma requisição HEAD para validar.
    
    Returns:
        dict: Relatório de diagnóstico
    """
    logger.info(f"🔬 Iniciando diagnóstico: {url_compartilhamento}")
    
    resultado = {
        "url_original": url_compartilhamento,
        "url_convertida": None,
        "conversao_ok": False,
        "head_request_ok": False,
        "content_type": None,
        "status_code": None,
        "erro": None
    }
    
    # Etapa 1: Converter URL
    url_convertida = converter_sharepoint_para_download(url_compartilhamento)
    resultado["url_convertida"] = url_convertida
    resultado["conversao_ok"] = url_convertida is not None
    
    if not resultado["conversao_ok"]:
        resultado["erro"] = "Falha na conversão da URL"
        return resultado
    
    # Etapa 2: Fazer HEAD request para validar
    try:
        session = criar_sessao_sharepoint()
        response = session.head(
            url_convertida,
            timeout=30,
            allow_redirects=True
        )
        
        resultado["status_code"] = response.status_code
        resultado["content_type"] = response.headers.get('Content-Type')
        resultado["head_request_ok"] = (
            response.status_code == 200 and 
            'text/html' not in resultado["content_type"].lower()
        )
        
        if not resultado["head_request_ok"]:
            resultado["erro"] = f"HEAD request falhou (status={response.status_code}, type={resultado['content_type']})"
    
    except Exception as e:
        resultado["erro"] = f"Erro no HEAD request: {str(e)}"
    
    logger.info(f"✅ Diagnóstico concluído: {'OK' if resultado['head_request_ok'] else 'FALHA'}")
    
    return resultado