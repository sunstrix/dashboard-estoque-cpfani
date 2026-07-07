"""
Utilitários para resolução de URLs de download do SharePoint.

Este módulo implementa 4 estratégias em cascata para converter URLs de compartilhamento
do SharePoint (formato /:x:/s/<site>/<token>) em URLs de download direto.

Estratégias (em ordem de tentativa):
1. Conversão direta para /_layouts/15/download.aspx?share=<token>
2. Extração de URL via análise do HTML da página de compartilhamento
3. Captura de header Location em redirects 301/302
4. Fallback com ?download=1 adicionado à URL original

Todas as estratégias usam User-Agent de navegador real e retry automático.
"""

import re
import logging
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
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
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

# Timeout padrão (segundos)
TIMEOUT_REQUEST = 60

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


def parse_url_sharepoint(url_compartilhamento: str) -> dict:
    """
    Extrai componentes de uma URL de compartilhamento do SharePoint.
    
    Formato esperado:
    https://<host>/:x:/s/<site>/<token>?e=<param>
    
    Args:
        url_compartilhamento: URL completa de compartilhamento
    
    Returns:
        dict: Dicionário com 'host', 'site', 'token', 'param' ou None se falhar
    
    Exemplo:
        >>> parse_url_sharepoint("https://didiernsf.sharepoint.com/:x:/s/NSFcosmticosepresentesLTDA/IQCujrbIbWZLT50lUu7tb2V7Aew2WFZQK1Uo2c4T583mDnU?e=5RIBrD")
        {
            'host': 'didiernsf.sharepoint.com',
            'site': 'NSFcosmticosepresentesLTDA',
            'token': 'IQCujrbIbWZLT50lUu7tb2V7Aew2WFZQK1Uo2c4T583mDnU',
            'param': '5RIBrD'
        }
    """
    # Regex para extrair componentes
    pattern = r'https?://([^/]+)/:x:/s/([^/]+)/([^?]+)(?:\?e=([^&]+))?'
    match = re.match(pattern, url_compartilhamento)
    
    if not match:
        logger.error(f"URL não corresponde ao formato esperado do SharePoint: {url_compartilhamento}")
        return None
    
    return {
        'host': match.group(1),
        'site': match.group(2),
        'token': match.group(3),
        'param': match.group(4) if match.group(4) else None
    }


# ==============================================================================
# ESTRATÉGIA 1: CONVERSÃO DIRETA
# ==============================================================================

def estrategia_1_conversao_direta(url_compartilhamento: str) -> str:
    """
    Estratégia 1: Converte URL de compartilhamento para URL de download direto.
    
    Transforma:
    https://<host>/:x:/s/<site>/<token>?e=<param>
    
    Em:
    https://<host>/sites/<site>/_layouts/15/download.aspx?share=<token>&e=<param>
    
    Args:
        url_compartilhamento: URL completa de compartilhamento
    
    Returns:
        str: URL de download direto ou None se falhar
    """
    logger.info("🔍 Estratégia 1: Conversão direta para download.aspx")
    
    componentes = parse_url_sharepoint(url_compartilhamento)
    if not componentes:
        return None
    
    # Monta URL de download direto
    url_download = f"https://{componentes['host']}/sites/{componentes['site']}/_layouts/15/download.aspx?share={componentes['token']}"
    
    if componentes['param']:
        url_download += f"&e={componentes['param']}"
    
    logger.info(f"URL gerada: {url_download}")
    
    # Testa se a URL é válida com HEAD request
    session = criar_sessao_sharepoint()
    try:
        response = session.head(url_download, timeout=TIMEOUT_REQUEST, allow_redirects=True)
        
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '')
            if 'application' in content_type or 'octet-stream' in content_type:
                logger.info("✅ Estratégia 1: URL de download válida")
                return url_download
            else:
                logger.warning(f"⚠️ Estratégia 1: Content-Type inesperado: {content_type}")
        else:
            logger.warning(f"⚠️ Estratégia 1: Status code {response.status_code}")
    except Exception as e:
        logger.error(f"❌ Estratégia 1 falhou: {str(e)}")
    
    return None


# ==============================================================================
# ESTRATÉGIA 2: EXTRAÇÃO VIA HTML
# ==============================================================================

def estrategia_2_extracao_html(url_compartilhamento: str) -> str:
    """
    Estratégia 2: Baixa o HTML da página de compartilhamento e extrai a URL real.
    
    Procura por padrões no HTML:
    - "@downloadUrl": "..."
    - "@content.downloadUrl": "..."
    - /_layouts/15/download.aspx?share=...
    - URLs terminando em .xlsx
    
    Args:
        url_compartilhamento: URL completa de compartilhamento
    
    Returns:
        str: URL de download extraída ou None se falhar
    """
    logger.info("🔍 Estratégia 2: Extração via análise do HTML")
    
    session = criar_sessao_sharepoint()
    
    try:
        # Baixa o HTML da página
        response = session.get(url_compartilhamento, timeout=TIMEOUT_REQUEST)
        response.raise_for_status()
        
        html_content = response.text
        
        # Lista de padrões regex para procurar no HTML
        padroes = [
            # Padrão 1: "@downloadUrl": "URL"
            r'"@downloadUrl"\s*:\s*"([^"]+)"',
            
            # Padrão 2: "@content.downloadUrl": "URL"
            r'"@content\.downloadUrl"\s*:\s*"([^"]+)"',
            
            # Padrão 3: download.aspx?share=TOKEN
            r'(https?://[^"\s]+/_layouts/15/download\.aspx\?share=[^"\s&]+)',
            
            # Padrão 4: URLs terminando em .xlsx
            r'(https?://[^"\s]+\.xlsx[^"\s]*)',
        ]
        
        for i, padrao in enumerate(padroes, 1):
            matches = re.findall(padrao, html_content)
            
            if matches:
                # Pega o primeiro match
                url_extraida = matches[0]
                
                # Remove escapes de URL se houver
                url_extraida = url_extraida.replace('\\u0026', '&')
                url_extraida = url_extraida.replace('&amp;', '&')
                
                logger.info(f"✅ Estratégia 2: URL extraída via padrão {i}: {url_extraida[:100]}...")
                return url_extraida
        
        logger.warning("⚠️ Estratégia 2: Nenhum padrão encontrado no HTML")
        
    except Exception as e:
        logger.error(f"❌ Estratégia 2 falhou: {str(e)}")
    
    return None


# ==============================================================================
# ESTRATÉGIA 3: CAPTURA DE REDIRECT
# ==============================================================================

def estrategia_3_captura_redirect(url_compartilhamento: str) -> str:
    """
    Estratégia 3: Captura o header Location de redirects 301/302.
    
    Faz requisição SEM seguir redirect e captura o header Location.
    
    Args:
        url_compartilhamento: URL completa de compartilhamento
    
    Returns:
        str: URL de redirect capturada ou None se falhar
    """
    logger.info("🔍 Estratégia 3: Captura de redirect 301/302")
    
    session = criar_sessao_sharepoint()
    
    try:
        # Requisição SEM seguir redirect
        response = session.get(
            url_compartilhamento,
            timeout=TIMEOUT_REQUEST,
            allow_redirects=False
        )
        
        # Verifica se é redirect
        if response.status_code in [301, 302, 303, 307, 308]:
            location = response.headers.get('Location')
            
            if location:
                logger.info(f"✅ Estratégia 3: Redirect capturado: {location[:100]}...")
                return location
            else:
                logger.warning("⚠️ Estratégia 3: Redirect sem header Location")
        else:
            logger.warning(f"⚠️ Estratégia 3: Status code {response.status_code} (não é redirect)")
    
    except Exception as e:
        logger.error(f"❌ Estratégia 3 falhou: {str(e)}")
    
    return None


# ==============================================================================
# ESTRATÉGIA 4: FALLBACK COM ?download=1
# ==============================================================================

def estrategia_4_fallback_download(url_compartilhamento: str) -> str:
    """
    Estratégia 4: Adiciona ?download=1 à URL original e tenta download direto.
    
    Args:
        url_compartilhamento: URL completa de compartilhamento
    
    Returns:
        str: URL com ?download=1 ou None se falhar
    """
    logger.info("🔍 Estratégia 4: Fallback com ?download=1")
    
    # Adiciona ?download=1
    if '?' in url_compartilhamento:
        url_download = url_compartilhamento + '&download=1'
    else:
        url_download = url_compartilhamento + '?download=1'
    
    logger.info(f"URL gerada: {url_download}")
    
    # Testa com HEAD request
    session = criar_sessao_sharepoint()
    try:
        response = session.head(url_download, timeout=TIMEOUT_REQUEST, allow_redirects=True)
        
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '')
            if 'application' in content_type or 'octet-stream' in content_type:
                logger.info("✅ Estratégia 4: URL com ?download=1 válida")
                return url_download
            else:
                logger.warning(f"⚠️ Estratégia 4: Content-Type inesperado: {content_type}")
        else:
            logger.warning(f"⚠️ Estratégia 4: Status code {response.status_code}")
    except Exception as e:
        logger.error(f"❌ Estratégia 4 falhou: {str(e)}")
    
    return None


# ==============================================================================
# FUNÇÃO PRINCIPAL: ORQUESTRAÇÃO DAS ESTRATÉGIAS
# ==============================================================================

def resolver_url_download_sharepoint(url_compartilhamento: str) -> str:
    """
    Resolve URL de download do SharePoint usando 4 estratégias em cascata.
    
    Tenta cada estratégia em ordem. A primeira que retornar sucesso é usada.
    Se todas falharem, retorna None.
    
    Args:
        url_compartilhamento: URL completa de compartilhamento do SharePoint
    
    Returns:
        str: URL de download válida ou None se todas as estratégias falharem
    
    Exemplo:
        >>> url = "https://didiernsf.sharepoint.com/:x:/s/NSFcosmticosepresentesLTDA/IQCujrbIbWZLT50lUu7tb2V7Aew2WFZQK1Uo2c4T583mDnU?e=5RIBrD"
        >>> url_download = resolver_url_download_sharepoint(url)
        >>> print(url_download)
        'https://didiernsf.sharepoint.com/sites/NSFcosmticosepresentesLTDA/_layouts/15/download.aspx?share=IQCujrbIbWZLT50lUu7tb2V7Aew2WFZQK1Uo2c4T583mDnU&e=5RIBrD'
    """
    logger.info("=" * 80)
    logger.info(f"🚀 Iniciando resolução de URL SharePoint")
    logger.info(f"URL original: {url_compartilhamento}")
    logger.info("=" * 80)
    
    # Lista de estratégias em ordem
    estrategias = [
        ("Conversão Direta", estrategia_1_conversao_direta),
        ("Extração via HTML", estrategia_2_extracao_html),
        ("Captura de Redirect", estrategia_3_captura_redirect),
        ("Fallback ?download=1", estrategia_4_fallback_download),
    ]
    
    # Tenta cada estratégia
    for i, (nome, funcao) in enumerate(estrategias, 1):
        logger.info(f"\n--- Tentando Estratégia {i}/{len(estrategias)}: {nome} ---")
        
        try:
            url_resolvida = funcao(url_compartilhamento)
            
            if url_resolvida:
                logger.info("=" * 80)
                logger.info(f"✅ SUCESSO: Estratégia {i} ({nome}) funcionou!")
                logger.info(f"URL resolvida: {url_resolvida}")
                logger.info("=" * 80)
                return url_resolvida
            else:
                logger.warning(f"⚠️ Estratégia {i} ({nome}) não retornou URL válida")
        
        except Exception as e:
            logger.error(f"❌ Erro inesperado na Estratégia {i} ({nome}): {str(e)}")
            logger.error(traceback.format_exc())
    
    # Todas as estratégias falharam
    logger.error("=" * 80)
    logger.error("❌ FALHA: Todas as 4 estratégias falharam")
    logger.error(f"URL original: {url_compartilhamento}")
    logger.error("=" * 80)
    
    return None


# ==============================================================================
# FUNÇÃO DE DOWNLOAD COM RESOLUÇÃO AUTOMÁTICA
# ==============================================================================

def baixar_arquivo_sharepoint(url_compartilhamento: str, nome_arquivo: str = "arquivo.xlsx") -> BytesIO:
    """
    Baixa arquivo do SharePoint resolvendo automaticamente a URL de download.
    
    Args:
        url_compartilhamento: URL completa de compartilhamento do SharePoint
        nome_arquivo: Nome do arquivo para logging
    
    Returns:
        BytesIO: Conteúdo do arquivo em buffer ou None se falhar
    
    Exemplo:
        >>> url = "https://didiernsf.sharepoint.com/:x:/s/NSFcosmticosepresentesLTDA/..."
        >>> buffer = baixar_arquivo_sharepoint(url, "DRAFT_PDVS.xlsx")
        >>> if buffer:
        ...     df = pd.read_excel(buffer)
    """
    logger.info(f"📥 Iniciando download: {nome_arquivo}")
    
    # Resolve URL de download
    url_download = resolver_url_download_sharepoint(url_compartilhamento)
    
    if not url_download:
        logger.error(f"❌ Não foi possível resolver URL de download para: {nome_arquivo}")
        return None
    
    # Baixa o arquivo
    session = criar_sessao_sharepoint()
    
    try:
        response = session.get(url_download, timeout=TIMEOUT_REQUEST, stream=True)
        response.raise_for_status()
        
        content = response.content
        tamanho_kb = len(content) / 1024
        
        logger.info(f"✅ Download concluído: {nome_arquivo} ({tamanho_kb:.1f} KB)")
        
        return BytesIO(content)
    
    except Exception as e:
        logger.error(f"❌ Erro ao baixar {nome_arquivo}: {str(e)}")
        logger.error(traceback.format_exc())
        return None


# ==============================================================================
# FUNÇÃO DE DIAGNÓSTICO
# ==============================================================================

def diagnosticar_url_sharepoint(url_compartilhamento: str) -> dict:
    """
    Realiza diagnóstico completo de uma URL do SharePoint.
    
    Testa todas as 4 estratégias e retorna relatório detalhado.
    
    Args:
        url_compartilhamento: URL completa de compartilhamento
    
    Returns:
        dict: Relatório de diagnóstico com resultados de cada estratégia
    """
    logger.info(f"🔬 Iniciando diagnóstico: {url_compartilhamento}")
    
    resultado = {
        "url_original": url_compartilhamento,
        "componentes": parse_url_sharepoint(url_compartilhamento),
        "estrategias": {}
    }
    
    estrategias = [
        ("1_conversao_direta", estrategia_1_conversao_direta),
        ("2_extracao_html", estrategia_2_extracao_html),
        ("3_captura_redirect", estrategia_3_captura_redirect),
        ("4_fallback_download", estrategia_4_fallback_download),
    ]
    
    for nome, funcao in estrategias:
        try:
            url_resolvida = funcao(url_compartilhamento)
            resultado["estrategias"][nome] = {
                "sucesso": url_resolvida is not None,
                "url_resolvida": url_resolvida
            }
        except Exception as e:
            resultado["estrategias"][nome] = {
                "sucesso": False,
                "erro": str(e)
            }
    
    # Verifica se alguma estratégia funcionou
    resultado["sucesso_geral"] = any(
        estrat.get("sucesso", False) 
        for estrat in resultado["estrategias"].values()
    )
    
    logger.info(f"✅ Diagnóstico concluído: {'SUCESSO' if resultado['sucesso_geral'] else 'FALHA'}")
    
    return resultado