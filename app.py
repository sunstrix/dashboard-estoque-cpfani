import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timezone, timedelta
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from io import BytesIO
import logging
import traceback
import re
import os
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

from config import (
    PLANILHAS, DE_PARA_LOJAS, DE_PARA_LOJAS_REVERSO, 
    MAPEAMENTO_PDV_DRAFT_RAW, NOMES_MARCAS, ABAS_SEGURANCA,
    LOGOS_MARCAS, CORES_MARCAS, REGRAS_ESTOQUE_MINIMO,
    COLUNAS_OBRIGATORIAS, TIMEOUT_DOWNLOAD, CACHE_TTL,
    obter_url_exportacao, VERSAO, DATA_VERSAO, diagnosticar_configuracao,
    esta_no_modo_privado, verificar_disponibilidade_gspread,
    DIAS_ANO,
    MODO_ACESSO, PLANILHAS_SHAREPOINT, obter_url_sharepoint,
    esta_no_modo_sharepoint, esta_no_modo_publico,
    detectar_colunas_historico, COLUNA_CLASSE_SEGMENTADA
)

try:
    from sharepoint_utils import baixar_arquivo_sharepoint, resolver_url_download_sharepoint
    SHAREPOINT_UTILS_AVAILABLE = True
except ImportError:
    SHAREPOINT_UTILS_AVAILABLE = False
    logger = logging.getLogger(__name__)
    logger.warning("⚠️ sharepoint_utils.py não encontrado. Modo SharePoint indisponível.")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="Painel de Performance de Estoque NSF - CP Fani",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.main { background-color: #0e1117; }
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
div[data-testid="stMetricValue"] {
    font-size: 28px; font-weight: 700; color: #D4AF37;
    text-shadow: 0 0 12px rgba(212,175,55,0.35);
}
div[data-testid="stMetricLabel"] { font-size: 13px; color: #8da9be; }
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #111820 0%, #0d1f14 100%);
    border: 1px solid #1a3d25; border-left: 4px solid #007A33;
    border-radius: 10px; padding: 16px 20px;
    box-shadow: 0 4px 16px rgba(0,122,51,0.15);
}
.stTabs [data-baseweb="tab-list"] {
    background: #0a0d12; border-bottom: 2px solid #007A33;
    border-radius: 6px 6px 0 0; gap: 4px; padding: 4px 8px 0;
}
.stTabs [data-baseweb="tab"] {
    color: #8da9be; font-size: 15px; padding: 8px 18px;
    border-radius: 6px 6px 0 0; border: none; background: transparent;
    transition: color 0.2s, background 0.2s;
}
.stTabs [data-baseweb="tab"]:hover { color: #D4AF37; background: #0d1f14; }
.stTabs [data-baseweb="tab"][aria-selected="true"] {
    color: #D4AF37; font-weight: 700;
    background: linear-gradient(180deg, #0d1f14 0%, #0a0d12 100%);
    border-top: 2px solid #007A33;
}
div.stButton > button {
    background: linear-gradient(135deg, #007A33, #005a26);
    color: #fff; border: none; border-radius: 8px;
    padding: 0.55rem 1.2rem; font-weight: 600; letter-spacing: 0.3px;
    box-shadow: 0 2px 8px rgba(0,122,51,0.4); transition: all 0.2s;
}
div.stButton > button:hover {
    background: linear-gradient(135deg, #009940, #007A33);
    color: #D4AF37; box-shadow: 0 4px 14px rgba(0,122,51,0.55);
    transform: translateY(-1px);
}
div.stDownloadButton > button {
    background: linear-gradient(135deg, #1a3d25, #0d1f14);
    color: #D4AF37; border: 1px solid #007A33; border-radius: 8px;
    padding: 0.45rem 1rem; font-weight: 600; transition: all 0.2s;
}
div.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #007A33, #005a26);
    color: #fff; border-color: #D4AF37;
}
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #080b0f 0%, #0a1510 100%);
    border-right: 2px solid #007A33;
}
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stSubheader { color: #D4AF37 !important; }
div[data-baseweb="select"] > div {
    background-color: #111820 !important;
    border: 1px solid #1a3d25 !important; border-radius: 8px !important;
}
div[data-baseweb="select"] > div:focus-within {
    border-color: #007A33 !important;
    box-shadow: 0 0 0 2px rgba(0,122,51,0.3) !important;
}
div[data-testid="stDataFrame"] {
    border: 1px solid #1a3d25; border-radius: 8px; overflow: hidden;
}
h1 { color: #D4AF37 !important; font-size: 2rem !important; letter-spacing: -0.5px; }
h2, h3 { color: #D4AF37 !important; }
hr { border-color: #1a3d25 !important; margin: 1.5rem 0; }
div[data-testid="stSpinner"] { color: #007A33; }
details summary {
    color: #D4AF37 !important; background: #0d1f14;
    border: 1px solid #1a3d25; border-radius: 6px; padding: 6px 12px;
}
.stProgress > div > div > div > div { background-color: #007A33; }
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0e1117; }
::-webkit-scrollbar-thumb { background: #007A33; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #D4AF37; }

.footer-rodape {
    position: fixed;
    bottom: 0;
    left: 0;
    right: 0;
    background: linear-gradient(135deg, #0d1f14 0%, #080b0f 100%);
    border-top: 2px solid #007A33;
    padding: 8px 20px;
    text-align: center;
    font-size: 11px;
    color: #6b7e8a;
    z-index: 999;
}
.footer-rodape span {
    color: #D4AF37;
    font-weight: 600;
}

.logo-placeholder {
    width: 40px;
    height: 40px;
    background: linear-gradient(135deg, #1a3d25, #0d1f14);
    border-radius: 6px;
    display: inline-block;
    animation: pulse 1.5s ease-in-out infinite;
}
@keyframes pulse {
    0%, 100% { opacity: 0.6; }
    50% { opacity: 1; }
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# FUNÇÕES AUXILIARES DE NORMALIZAÇÃO
# ==============================================================================

def normalizar_texto(texto):
    """
    Normaliza texto removendo acentos e convertendo para uppercase.
    Usado para matching robusto de nomes de colunas.
    
    Exemplo:
        >>> normalizar_texto("Preço Unitário")
        'PRECO UNITARIO'
    """
    if pd.isna(texto) or texto is None:
        return ""
    texto = str(texto).strip()
    # Remove acentos usando Unicode normalization
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto.upper()


def normalizar_sku(valor):
    """Normaliza SKUs para matching entre planilhas."""
    if pd.isna(valor) or valor is None:
        return ""
    s = str(valor).strip()
    if s.endswith('.0'):
        s = s[:-2]
    try:
        num = float(s)
        if num == int(num):
            s = str(int(num))
    except (ValueError, TypeError):
        pass
    s = re.sub(r'\s+', '', s).upper()
    return s


def normalizar_pdv(valor):
    """
    BUG 1 FIX: Normaliza PDV para aceitar formatos variados.
    
    Aceita: 4842, 4842.0, "4842", " 4842 ", 4842,0, "4842.0"
    Retorna: string limpa do PDV ou None se inválido
    """
    if pd.isna(valor) or valor is None:
        return None
    
    # Converte para string e limpa
    s = str(valor).strip()
    
    # Remove .0 final (ex: "4842.0" → "4842")
    if s.endswith('.0'):
        s = s[:-2]
    
    # Remove vírgula e substitui por ponto (ex: "4842,0" → "4842.0")
    s = s.replace(',', '.')
    
    # Tenta converter para número
    try:
        num = float(s)
        if num == int(num):
            return str(int(num))
        else:
            return str(num)
    except (ValueError, TypeError):
        pass
    
    # Se não conseguiu converter, retorna a string limpa se não estiver vazia
    if s and s != 'nan' and s != 'None':
        return s
    
    return None


def normalizar_nome_loja(nome):
    """Normaliza nomes de loja para matching."""
    if pd.isna(nome) or nome is None:
        return ""
    s = str(nome).strip().lower()
    s = re.sub(r'[^a-z0-9\s]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


MAPEAMENTO_PDV_DRAFT_NORMALIZADO = {
    normalizar_nome_loja(k): v for k, v in MAPEAMENTO_PDV_DRAFT_RAW.items()
}


# ==============================================================================
# FUNÇÕES DE DETECÇÃO DE COLUNAS (BUG 1 FIX)
# ==============================================================================

def detectar_coluna_custo(colunas):
    """
    BUG 1 FIX: Detecta coluna de custo por NOME (não por posição).
    
    Busca por variações normalizadas:
    - CUSTO
    - PRECO CUSTO / PREÇO CUSTO
    - CUSTO UNITARIO / CUSTO UNITÁRIO
    - VALOR CUSTO
    - CUSTO (R$)
    - CUSTO R$
    
    Args:
        colunas: Lista de nomes das colunas do DataFrame
    
    Returns:
        str: Nome da coluna de custo encontrada ou None se não encontrar
    """
    # Variações possíveis do nome da coluna de custo (normalizadas)
    variacoes_custo = [
        'CUSTO',
        'PRECO CUSTO',
        'PRECO DE CUSTO',
        'CUSTO UNITARIO',
        'CUSTO UNIT',
        'VALOR CUSTO',
        'CUSTO R$',
        'CUSTO (R$)',
        'VL CUSTO',
        'VALOR DE CUSTO',
        'PRECO MEDIO CUSTO',
        'CUSTO MEDIO',
    ]
    
    # Normaliza todas as colunas para comparação
    colunas_normalizadas = {col: normalizar_texto(col) for col in colunas}
    
    # Busca match exato
    for col_original, col_norm in colunas_normalizadas.items():
        if col_norm in variacoes_custo:
            logger.info(f"✅ Coluna de custo encontrada por nome: '{col_original}' (normalizado: '{col_norm}')")
            return col_original
    
    # Se não encontrou match exato, busca por substring
    for col_original, col_norm in colunas_normalizadas.items():
        if 'CUSTO' in col_norm and 'TOTAL' not in col_norm:
            logger.info(f"✅ Coluna de custo encontrada por substring: '{col_original}' (normalizado: '{col_norm}')")
            return col_original
    
    logger.error(f"❌ Coluna de custo NÃO encontrada. Colunas disponíveis: {list(colunas)}")
    return None


def detectar_coluna_marca(colunas):
    """
    BUG 2 FIX: Detecta coluna que identifica a marca (Boticário/Eudora/QDB).
    
    Busca por: MARCA, BRAND, LINHA, PRODUTO_LINHA, etc.
    
    Returns:
        str: Nome da coluna de marca ou None
    """
    variacoes_marca = [
        'MARCA',
        'BRAND',
        'LINHA',
        'LINHA PRODUTO',
        'PRODUTO LINHA',
        'CATEGORIA MARCA',
        'TIPO MARCA',
    ]
    
    colunas_normalizadas = {col: normalizar_texto(col) for col in colunas}
    
    for col_original, col_norm in colunas_normalizadas.items():
        if col_norm in variacoes_marca:
            logger.info(f"✅ Coluna de marca encontrada: '{col_original}'")
            return col_original
    
    return None


# ==============================================================================
# SESSÃO E DOWNLOAD
# ==============================================================================

def criar_sessao_com_retry():
    session = requests.Session()
    retry = Retry(
        total=3, 
        backoff_factor=1, 
        status_forcelist=[429, 500, 502, 503, 504], 
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def download_planilha_excel(url_planilha, nome_planilha):
    logger.info(f"Iniciando download: {nome_planilha}")
    session = criar_sessao_com_retry()
    
    try:
        response = session.get(url_planilha, timeout=TIMEOUT_DOWNLOAD, stream=True)
        response.raise_for_status()
        
        content = response.content
        tamanho_kb = len(content) / 1024
        
        logger.info(f"✅ Download concluído: {nome_planilha} ({tamanho_kb:.1f} KB)")
        return BytesIO(content)
        
    except Exception as e:
        logger.error(f"❌ Erro ao baixar {nome_planilha}: {str(e)}")
        logger.error(traceback.format_exc())
        st.error(f"Falha ao carregar {nome_planilha}: {str(e)[:200]}")
        return None


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def download_planilha_sharepoint(url_compartilhamento, nome_planilha):
    if not SHAREPOINT_UTILS_AVAILABLE:
        logger.error("❌ sharepoint_utils.py não disponível")
        st.error("Módulo SharePoint não encontrado. Verifique sharepoint_utils.py")
        return None
    
    logger.info(f"Iniciando download SharePoint: {nome_planilha}")
    
    try:
        buffer = baixar_arquivo_sharepoint(url_compartilhamento, nome_planilha)
        
        if buffer:
            logger.info(f"✅ Download SharePoint concluído: {nome_planilha}")
            return buffer
        else:
            logger.error(f"❌ Falha no download SharePoint: {nome_planilha}")
            st.error(f"Falha ao carregar {nome_planilha} do SharePoint")
            return None
    
    except Exception as e:
        logger.error(f"❌ Erro ao baixar {nome_planilha} do SharePoint: {str(e)}")
        logger.error(traceback.format_exc())
        st.error(f"Falha ao carregar {nome_planilha}: {str(e)[:200]}")
        return None


# ==============================================================================
# CARREGAMENTO DE PLANILHAS
# ==============================================================================

@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_skus_ignorados(ignorados_buffer):
    logger.info("Carregando planilha de SKUs Ignorados...")
    
    try:
        if ignorados_buffer is None:
            logger.warning("Planilha de Ignorados não fornecida")
            return set()
        
        excel_file = pd.ExcelFile(ignorados_buffer)
        
        if not excel_file.sheet_names:
            logger.error("Ignorados: nenhuma aba encontrada")
            return set()
        
        df = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
        
        if df.empty:
            logger.warning("Ignorados: planilha vazia")
            return set()
        
        df.columns = [str(col).strip().upper() for col in df.columns]
        
        colunas_sku = ['SKU', 'CÓDIGO', 'CODIGO', 'CÓDIGO SKU', 'CODIGO SKU', 
                       'EAN', 'COD PRODUTO', 'CÓDIGO PRODUTO', 'IGNORAR']
        coluna_sku = next((c for c in colunas_sku if c in df.columns), None)
        
        if coluna_sku is None:
            logger.error(f"Ignorados: coluna SKU não encontrada. Colunas: {list(df.columns)}")
            return set()
        
        skus_ignorados = set(df[coluna_sku].apply(normalizar_sku))
        skus_ignorados.discard("")
        
        logger.info(f"✅ SKUs Ignorados carregados: {len(skus_ignorados)} SKUs")
        return skus_ignorados
        
    except Exception as e:
        logger.error(f"Erro ao carregar SKUs Ignorados: {str(e)}")
        logger.error(traceback.format_exc())
        return set()


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_planilha_retaguarda(custos_buffer):
    """
    BUG 1 FIX: Carrega planilha de Retaguarda com busca de coluna de custo por NOME.
    
    Melhorias:
    1. Busca coluna de custo por nome (não por posição fixa)
    2. Normaliza PDV para aceitar formatos variados
    3. Diagnóstico detalhado de linhas rejeitadas
    4. Alerta visual quando SKUs matcheados = 0
    """
    logger.info("Carregando planilha Retaguarda (custos)...")
    
    try:
        if custos_buffer is None:
            logger.warning("Planilha Retaguarda não fornecida")
            return pd.DataFrame()
        
        excel_file = pd.ExcelFile(custos_buffer)
        logger.info(f"Abas encontradas na Retaguarda: {excel_file.sheet_names}")
        
        if not excel_file.sheet_names:
            logger.error("Retaguarda: nenhuma aba encontrada")
            return pd.DataFrame()
        
        df = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
        
        if df.empty:
            logger.warning("Retaguarda: planilha vazia")
            return pd.DataFrame()
        
        df.columns = [str(col).strip().upper() for col in df.columns]
        logger.info(f"Colunas encontradas: {list(df.columns)}")
        
        # ==========================================================================
        # BUG 1 FIX: Busca coluna de SKU por nome
        # ==========================================================================
        colunas_sku = ['SKU', 'CÓDIGO', 'CODIGO', 'CÓDIGO SKU', 'CODIGO SKU', 
                       'EAN', 'COD PRODUTO', 'CÓDIGO PRODUTO', 'CODIGO', 'CÓDIGO']
        coluna_sku = next((c for c in colunas_sku if c in df.columns), None)
        
        # ==========================================================================
        # BUG 1 FIX: Busca coluna de CUSTO por NOME (não por posição!)
        # ==========================================================================
        coluna_custo = detectar_coluna_custo(df.columns)
        
        if coluna_sku is None:
            logger.error(f"Retaguarda: coluna SKU não encontrada. Colunas: {list(df.columns)}")
            return pd.DataFrame()
        
        if coluna_custo is None:
            logger.error(f"Retaguarda: coluna CUSTO não encontrada!")
            logger.error(f"Colunas disponíveis: {list(df.columns)}")
            st.warning("⚠️ **Planilha Retaguarda sem coluna de CUSTO.** Verifique se a planilha correta está configurada.")
            return pd.DataFrame()
        
        logger.info(f"✅ Colunas detectadas - SKU: '{coluna_sku}', Custo: '{coluna_custo}'")
        
        # ==========================================================================
        # BUG 1 FIX: Normaliza SKU e CUSTO
        # ==========================================================================
        df_resultado = pd.DataFrame()
        df_resultado['SKU'] = df[coluna_sku].apply(normalizar_sku)
        df_resultado['CUSTO_RETARGUARDA'] = pd.to_numeric(df[coluna_custo], errors='coerce').fillna(0)
        
        # ==========================================================================
        # BUG 1 FIX: Busca coluna de PDV/LOJA e normaliza
        # ==========================================================================
        colunas_loja = ['LOJA', 'PDV', 'LOJA/PDV', 'LOJA - PDV', 'CÓDIGO LOJA', 'CODIGO LOJA',
                       'FILIAL', 'COD FILIAL', 'LOJA NOME', 'NOME LOJA', 'PDV_CODIGO', 'COD PDV']
        coluna_loja = next((c for c in colunas_loja if c in df.columns), None)
        
        if coluna_loja is None:
            logger.info("Retaguarda: sem coluna de loja/PDV. Aplicando custos para todos os PDVs.")
            todos_pdvs = list(DE_PARA_LOJAS.keys())
            df_expandido = []
            for pdv in todos_pdvs:
                df_temp = df_resultado.copy()
                df_temp['PDV'] = pdv
                df_expandido.append(df_temp)
            df_resultado = pd.concat(df_expandido, ignore_index=True)
        else:
            # BUG 1 FIX: Usa normalizar_pdv() ao invés de conversão direta
            logger.info(f"Retaguarda: coluna de PDV encontrada: '{coluna_loja}'")
            
            # Diagnóstico: mostra tipos e valores únicos da coluna PDV
            logger.info(f"📊 Diagnóstico PDV - Tipo original: {df[coluna_loja].dtype}")
            valores_unicos = df[coluna_loja].dropna().unique()[:10]
            logger.info(f"📊 Diagnóstico PDV - Primeiros 10 valores únicos: {valores_unicos}")
            
            df_resultado['PDV_RAW'] = df[coluna_loja]
            df_resultado['PDV'] = df[coluna_loja].apply(normalizar_pdv)
            
            # Diagnóstico: quantos PDVs foram normalizados com sucesso
            pdvs_validos = df_resultado['PDV'].notna().sum()
            pdvs_invalidos = df_resultado['PDV'].isna().sum()
            logger.info(f"📊 Diagnóstico PDV - Válidos: {pdvs_validos}, Inválidos: {pdvs_invalidos}")
            
            # Mostra exemplos de PDVs inválidos
            if pdvs_invalidos > 0:
                exemplos_invalidos = df_resultado[df_resultado['PDV'].isna()]['PDV_RAW'].head(5).tolist()
                logger.warning(f"⚠️ Exemplos de PDVs inválidos: {exemplos_invalidos}")
            
            # Mapeia PDV normalizado para código numérico
            df_resultado['PDV_MAPEADO'] = df_resultado['PDV'].apply(
                lambda x: int(x) if x and x.isdigit() and int(x) in DE_PARA_LOJAS else None
            )
            
            # Mantém apenas PDVs válidos
            antes = len(df_resultado)
            df_resultado = df_resultado[df_resultado['PDV_MAPEADO'].notna()].copy()
            depois = len(df_resultado)
            
            if antes != depois:
                logger.warning(f"Retaguarda: {antes - depois} linhas rejeitadas (PDV inválido ou não mapeado)")
            
            df_resultado['PDV'] = df_resultado['PDV_MAPEADO']
        
        if df_resultado.empty:
            logger.error("Retaguarda: nenhum dado válido após processamento")
            return pd.DataFrame()
        
        df_resultado['PDV'] = df_resultado['PDV'].astype(int)
        df_resultado = df_resultado.sort_values('CUSTO_RETARGUARDA', ascending=False)
        df_resultado = df_resultado.drop_duplicates(subset=['PDV', 'SKU'], keep='first')
        
        logger.info(f"✅ Retaguarda carregada: {len(df_resultado)} registros")
        return df_resultado[['PDV', 'SKU', 'CUSTO_RETARGUARDA']].copy()
        
    except Exception as e:
        logger.error(f"Erro ao carregar Retaguarda: {str(e)}")
        logger.error(traceback.format_exc())
        return pd.DataFrame()


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_estoque_seguranca(seguranca_buffer):
    """
    BUG 2 FIX: Carrega planilha de Estoque de Segurança com suporte a:
    1. Formato antigo: 3 abas separadas (BOT, EUD, QDB) - fallback retroativo
    2. Formato novo: 1 aba consolidada (saldo-estoque-*) com coluna de marca
    
    Detecta automaticamente qual formato está presente e processa corretamente.
    """
    logger.info("Carregando planilha de estoque de segurança...")
    
    try:
        if seguranca_buffer is None:
            logger.warning("Planilha de segurança não fornecida")
            return pd.DataFrame()
        
        excel_file = pd.ExcelFile(seguranca_buffer)
        abas_disponiveis = excel_file.sheet_names
        logger.info(f"Abas disponíveis na Segurança: {abas_disponiveis}")
        
        # ==========================================================================
        # FORMATO 1: Tenta formato antigo (3 abas separadas BOT/EUD/QDB)
        # ==========================================================================
        abas_esperadas = ['BOT', 'EUD', 'QDB']
        abas_encontradas = [aba for aba in abas_esperadas if aba.upper() in [a.upper() for a in abas_disponiveis]]
        
        if abas_encontradas:
            logger.info(f"✅ Formato antigo detectado: {len(abas_encontradas)} abas separadas (BOT/EUD/QDB)")
            return _carregar_seguranca_formato_antigo(excel_file, abas_encontradas)
        
        # ==========================================================================
        # FORMATO 2: Tenta formato novo (aba consolidada saldo-estoque-*)
        # ==========================================================================
        aba_consolidada = None
        for aba in abas_disponiveis:
            if aba.lower().startswith('saldo-estoque'):
                aba_consolidada = aba
                break
        
        # Se não encontrou por padrão, usa a primeira aba
        if aba_consolidada is None and abas_disponiveis:
            aba_consolidada = abas_disponiveis[0]
            logger.info(f"⚠️ Nenhuma aba 'saldo-estoque-*' encontrada. Usando primeira aba: '{aba_consolidada}'")
        
        if aba_consolidada:
            logger.info(f"✅ Formato novo detectado: aba consolidada '{aba_consolidada}'")
            return _carregar_seguranca_formato_novo(excel_file, aba_consolidada)
        
        logger.error("Segurança: nenhuma aba válida encontrada")
        return pd.DataFrame()
        
    except Exception as e:
        logger.error(f"Erro ao carregar segurança: {str(e)}")
        logger.error(traceback.format_exc())
        return pd.DataFrame()


def _carregar_seguranca_formato_antigo(excel_file, abas_encontradas):
    """
    BUG 2 FIX: Carrega formato antigo (3 abas separadas BOT/EUD/QDB).
    Mantém a lógica original para compatibilidade retroativa.
    """
    dfs_abas = []
    for aba_nome in abas_encontradas:
        try:
            aba_exata = [nome for nome in excel_file.sheet_names if nome.upper() == aba_nome.upper()][0]
            df_abas = pd.read_excel(excel_file, sheet_name=aba_exata)
            
            if df_abas.empty:
                continue
            
            df_abas.columns = [col.strip().upper() for col in df_abas.columns]
            
            if 'PDV' not in df_abas.columns or 'SKU' not in df_abas.columns:
                logger.warning(f"Segurança aba {aba_nome}: colunas PDV/SKU faltantes")
                continue
            
            colunas_est = ['ESTOQUE DE SEGURANCA', 'ESTOQUE_DE_SEGURANCA', 'ESTOQUE_SEGURANCA',
                          'ESTOQUE MINIMO', 'ESTOQUE_MINIMO', 'MINIMO', 'SEGURANCA', 'QTD_MINIMA']
            coluna_est = next((c for c in colunas_est if c in df_abas.columns), None)
            
            if coluna_est is None:
                df_abas['ESTOQUE_DE_SEGURANCA'] = 0
            else:
                df_abas = df_abas.rename(columns={coluna_est: 'ESTOQUE_DE_SEGURANCA'})
                df_abas['ESTOQUE_DE_SEGURANCA'] = pd.to_numeric(df_abas['ESTOQUE_DE_SEGURANCA'], errors='coerce').fillna(0)
            
            df_abas['PDV'] = pd.to_numeric(df_abas['PDV'], errors='coerce')
            df_abas['SKU'] = df_abas['SKU'].apply(normalizar_sku)
            df_abas['MARCA_REFERENCIA'] = ABAS_SEGURANCA.get(aba_nome, aba_nome)
            
            dfs_abas.append(df_abas[['PDV', 'SKU', 'ESTOQUE_DE_SEGURANCA', 'MARCA_REFERENCIA']].copy())
            
        except Exception as e:
            logger.error(f"Erro ao processar aba {aba_nome} da segurança: {str(e)}")
            logger.error(traceback.format_exc())
            continue
    
    if dfs_abas:
        df_final = pd.concat(dfs_abas, ignore_index=True)
        logger.info(f"✅ Estoque de segurança carregado (formato antigo): {len(df_final)} registros")
        return df_final
    
    logger.warning("Segurança: nenhum dado válido encontrado (formato antigo)")
    return pd.DataFrame()


def _carregar_seguranca_formato_novo(excel_file, aba_consolidada):
    """
    BUG 2 FIX: Carrega formato novo (aba consolidada com coluna de marca/PDV).
    
    Detecta automaticamente a coluna que identifica a marca e mapeia para
    as 3 marcas (Boticário/Eudora/Quem Disse, Berenice?).
    """
    try:
        df = pd.read_excel(excel_file, sheet_name=aba_consolidada)
        
        if df.empty:
            logger.warning(f"Segurança: aba '{aba_consolidada}' está vazia")
            return pd.DataFrame()
        
        df.columns = [col.strip().upper() for col in df.columns]
        logger.info(f"Colunas na aba consolidada: {list(df.columns)}")
        
        # ==========================================================================
        # Detecta colunas necessárias
        # ==========================================================================
        # SKU
        colunas_sku = ['SKU', 'CÓDIGO', 'CODIGO', 'CÓDIGO SKU', 'CODIGO SKU', 'EAN']
        coluna_sku = next((c for c in colunas_sku if c in df.columns), None)
        
        # PDV
        colunas_pdv = ['PDV', 'LOJA', 'COD LOJA', 'CÓDIGO LOJA', 'CODIGO LOJA', 'FILIAL']
        coluna_pdv = next((c for c in colunas_pdv if c in df.columns), None)
        
        # Estoque de Segurança
        colunas_est = ['ESTOQUE DE SEGURANCA', 'ESTOQUE_DE_SEGURANCA', 'ESTOQUE_SEGURANCA',
                      'ESTOQUE MINIMO', 'ESTOQUE_MINIMO', 'MINIMO', 'SEGURANCA', 'QTD_MINIMA',
                      'EST. MINIMO', 'EST. MÍNIMO', 'ESTOQUE MIN']
        coluna_est = next((c for c in colunas_est if c in df.columns), None)
        
        # Marca (opcional - pode ser inferida pelo PDV)
        coluna_marca = detectar_coluna_marca(df.columns)
        
        if coluna_sku is None or coluna_pdv is None:
            logger.error(f"Segurança: colunas SKU ou PDV não encontradas. SKU={coluna_sku}, PDV={coluna_pdv}")
            return pd.DataFrame()
        
        if coluna_est is None:
            logger.warning("Segurança: coluna de estoque de segurança não encontrada. Usando 0.")
            df['ESTOQUE_DE_SEGURANCA'] = 0
        else:
            df['ESTOQUE_DE_SEGURANCA'] = pd.to_numeric(df[coluna_est], errors='coerce').fillna(0)
        
        # ==========================================================================
        # Normaliza PDV e SKU
        # ==========================================================================
        df['PDV'] = df[coluna_pdv].apply(normalizar_pdv)
        df['SKU'] = df[coluna_sku].apply(normalizar_sku)
        
        # Converte PDV para int se possível
        df['PDV_NUM'] = pd.to_numeric(df['PDV'], errors='coerce')
        
        # ==========================================================================
        # Determina MARCA_REFERENCIA
        # ==========================================================================
        if coluna_marca:
            # Usa coluna de marca direta
            logger.info(f"✅ Usando coluna de marca: '{coluna_marca}'")
            df['MARCA_REFERENCIA'] = df[coluna_marca].apply(_mapear_marca)
        else:
            # Tenta inferir marca pelo PDV (se cada PDV pertence a uma marca específica)
            # ou assume que todos os SKUs se aplicam a todas as marcas
            logger.info("⚠️ Coluna de marca não encontrada. Tentando inferir pelo PDV ou aplicando para todas as marcas.")
            
            # Se há poucos PDVs únicos, pode ser que cada PDV tenha uma marca
            pdvs_unicos = df['PDV_NUM'].dropna().unique()
            if len(pdvs_unicos) <= 3:
                logger.info(f"Detectados {len(pdvs_unicos)} PDVs únicos. Mapeando para marcas.")
                # Mapeia PDVs para marcas (assumindo ordem: Boticário, Eudora, QDB)
                marcas_possiveis = list(ABAS_SEGURANCA.values())
                mapeamento_pdv_marca = {}
                for i, pdv in enumerate(sorted(pdvs_unicos)):
                    if i < len(marcas_possiveis):
                        mapeamento_pdv_marca[pdv] = marcas_possiveis[i]
                
                df['MARCA_REFERENCIA'] = df['PDV_NUM'].map(mapeamento_pdv_marca)
            else:
                # Muitos PDVs: aplica para todas as marcas
                logger.info(f"Muitos PDVs ({len(pdvs_unicos)}). Aplicando estoque de segurança para todas as marcas.")
                # Duplica para cada marca
                dfs_marcas = []
                for marca in ABAS_SEGURANCA.values():
                    df_temp = df.copy()
                    df_temp['MARCA_REFERENCIA'] = marca
                    dfs_marcas.append(df_temp)
                df = pd.concat(dfs_marcas, ignore_index=True)
        
        # ==========================================================================
        # Filtra e retorna
        # ==========================================================================
        df_valido = df[df['PDV_NUM'].notna() & (df['SKU'] != '')].copy()
        
        if df_valido.empty:
            logger.error("Segurança: nenhum dado válido após processamento")
            return pd.DataFrame()
        
        df_valido['PDV'] = df_valido['PDV_NUM'].astype(int)
        
        logger.info(f"✅ Estoque de segurança carregado (formato novo): {len(df_valido)} registros")
        return df_valido[['PDV', 'SKU', 'ESTOQUE_DE_SEGURANCA', 'MARCA_REFERENCIA']].copy()
        
    except Exception as e:
        logger.error(f"Erro ao processar aba consolidada: {str(e)}")
        logger.error(traceback.format_exc())
        return pd.DataFrame()


def _mapear_marca(valor_marca):
    """
    BUG 2 FIX: Mapeia valor da coluna de marca para MARCA_REFERENCIA padrão.
    
    Aceita variações como:
    - "Boticário", "O Boticário", "BOTICARIO", "BOT" → "O Boticário"
    - "Eudora", "EUDORA", "EUD" → "Eudora"
    - "Quem Disse", "QDB", "QUEM_DISSE" → "Quem Disse, Berenice?"
    """
    if pd.isna(valor_marca) or valor_marca is None:
        return None
    
    marca_norm = normalizar_texto(valor_marca)
    
    # Mapeamento normalizado
    if 'BOTICARIO' in marca_norm or 'BOT' == marca_norm:
        return 'O Boticário'
    elif 'EUDORA' in marca_norm or 'EUD' == marca_norm:
        return 'Eudora'
    elif 'QUEM DISSE' in marca_norm or 'BERENICE' in marca_norm or 'QDB' == marca_norm:
        return 'Quem Disse, Berenice?'
    
    return None


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_dados_principais(draft_buffer):
    """
    Carrega planilha principal DRAFT_PDVS com DETECÇÃO DINÂMICA de colunas de histórico.
    """
    logger.info("Carregando planilha principal DRAFT_PDVS...")
    
    try:
        if draft_buffer is None:
            logger.error("Planilha principal não fornecida")
            return {}, None
        
        excel_file = pd.ExcelFile(draft_buffer)
        abas_disponiveis = excel_file.sheet_names
        logger.info(f"Abas disponíveis na principal: {abas_disponiveis}")
        
        abas_esperadas = list(NOMES_MARCAS.keys())
        abas_faltantes = [aba for aba in abas_esperadas if aba not in abas_disponiveis]
        
        if abas_faltantes:
            logger.error(f"Principal: abas faltantes: {abas_faltantes}")
            st.error(f"❌ Abas esperadas não encontradas: {abas_faltantes}")
            st.info(f"Abas disponíveis: {abas_disponiveis}")
            return {}, None
        
        dicionario_marcas = {}
        data_atualizacao = None
        
        for aba_excel, nome_exibicao in NOMES_MARCAS.items():
            logger.info(f"Processando aba: {aba_excel} → {nome_exibicao}")
            
            try:
                # Tenta diferentes configurações de header
                df = None
                for header_row in [0, 1, 2]:
                    try:
                        df_temp = pd.read_excel(excel_file, sheet_name=aba_excel, header=header_row)
                        if not df_temp.empty and len(df_temp.columns) > 5:
                            df = df_temp
                            logger.info(f"✅ Usando header={header_row} para aba {aba_excel}")
                            break
                    except Exception as e_header:
                        logger.warning(f"Tentativa header={header_row} falhou: {str(e_header)}")
                        continue
                
                if df is None or df.empty:
                    logger.error(f"❌ Aba {aba_excel}: não foi possível carregar dados (DataFrame vazio)")
                    continue
                
                # DIAGNÓSTICO
                logger.info(f"📊 DataFrame carregado - Shape: {df.shape}")
                logger.info(f"📊 Colunas: {list(df.columns)[:10]}...")
                
                colunas_faltantes = [col for col in COLUNAS_OBRIGATORIAS["draft_pdvs"] if col not in df.columns]
                if colunas_faltantes:
                    logger.error(f"Aba {aba_excel}: colunas faltantes: {colunas_faltantes}")
                    st.error(f"❌ Aba {aba_excel} está faltando colunas: {colunas_faltantes}")
                    continue
                
                df['Marca'] = nome_exibicao
                df['PDV'] = pd.to_numeric(df['PDV'], errors='coerce')
                df['Estoque Atual'] = pd.to_numeric(df['Estoque Atual'], errors='coerce').fillna(0)
                df['Preço tabela'] = pd.to_numeric(df['Preço tabela'], errors='coerce').fillna(0)
                df['SKU'] = df['SKU'].apply(normalizar_sku)
                
                # Remove linhas com PDV ou SKU inválidos
                antes_limpeza = len(df)
                df = df[df['PDV'].notna() & (df['SKU'] != '')]
                depois_limpeza = len(df)
                if antes_limpeza != depois_limpeza:
                    logger.info(f"🧹 Removidas {antes_limpeza - depois_limpeza} linhas com PDV/SKU inválidos")
                
                # Preserva coluna "Classe Segmentada"
                if COLUNA_CLASSE_SEGMENTADA in df.columns:
                    logger.info(f"✅ Coluna '{COLUNA_CLASSE_SEGMENTADA}' encontrada e preservada")
                
                # Coluna Estoque em Trânsito
                colunas_transito = ['Estoque em Trânsito', 'Estoque em Transito', 'Estoque Transito', 
                                   'Em Trânsito', 'Em Transito', 'Trânsito', 'Transito']
                coluna_transito = next((c for c in colunas_transito if c in df.columns), None)
                
                if coluna_transito:
                    df['Estoque em Trânsito'] = pd.to_numeric(df[coluna_transito], errors='coerce').fillna(0)
                    logger.info(f"✅ Coluna 'Estoque em Trânsito' encontrada: {coluna_transito}")
                else:
                    df['Estoque em Trânsito'] = 0
                    logger.warning(f"Aba {aba_excel}: coluna 'Estoque em Trânsito' não encontrada. Usando 0.")
                
                # CÁLCULO DE DDV COM DETECÇÃO DINÂMICA
                try:
                    colunas_historico = detectar_colunas_historico(df.columns)
                    
                    if colunas_historico:
                        df_hist = df[colunas_historico].apply(pd.to_numeric, errors='coerce').fillna(0)
                        df['Historico_Total'] = df_hist.sum(axis=1)
                        df['DDV'] = df['Historico_Total'] / DIAS_ANO
                        logger.info(f"✅ DDV calculado: {len(colunas_historico)} colunas de histórico somadas")
                    else:
                        logger.warning(f"Aba {aba_excel}: nenhuma coluna de histórico detectada. Usando DDV=0.")
                        df['Historico_Total'] = 0
                        df['DDV'] = 0
                except Exception as e_ddv:
                    logger.error(f"Erro ao calcular DDV na aba {aba_excel}: {str(e_ddv)}")
                    df['Historico_Total'] = 0
                    df['DDV'] = 0
                
                # Calcula Cobertura de Estoque
                df['Cobertura_Estoque'] = np.where(
                    df['DDV'] > 0,
                    df['Estoque Atual'] / df['DDV'],
                    np.nan
                )
                
                dicionario_marcas[nome_exibicao] = df
                logger.info(f"✅ {aba_excel}: {len(df)} registros carregados")
                
            except Exception as e:
                logger.error(f"Erro ao processar aba {aba_excel}: {str(e)}")
                logger.error(traceback.format_exc())
                continue
        
        return dicionario_marcas, data_atualizacao
        
    except Exception as e:
        logger.error(f"Erro ao carregar planilha principal: {str(e)}")
        logger.error(traceback.format_exc())
        return {}, None


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_todos_os_dados():
    """
    Carrega todas as planilhas com downloads PARALELIZADOS.
    """
    logger.info("=" * 60)
    logger.info("INICIANDO CARREGAMENTO DE DADOS")
    logger.info(f"MODO_ACESSO: {MODO_ACESSO}")
    logger.info("=" * 60)
    
    stats = {
        'total_skus_retaguarda': 0,
        'skus_matcheados': 0,
        'skus_sem_match': 0,
        'skus_custo_maior': 0,
        'skus_ignorados': 0,
        'skus_ignorados_filtrados': 0,
        'modo_acesso': MODO_ACESSO,
        'retaguarda_sem_custo': False  # BUG 1: Flag para alerta visual
    }
    
    logger.info("Fase 1: Download PARALELO das planilhas...")
    st.session_state['progresso_atual'] = "Baixando planilhas em paralelo..."
    
    if esta_no_modo_sharepoint():
        logger.info("📡 Usando fonte de dados: SHAREPOINT")
        if not SHAREPOINT_UTILS_AVAILABLE:
            st.error("❌ Modo SharePoint selecionado mas sharepoint_utils.py não está disponível")
            return {}, None, stats
        
        urls_planilhas = {
            chave: config['url'] 
            for chave, config in PLANILHAS_SHAREPOINT.items()
        }
        
        buffers = {}
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {
                executor.submit(download_planilha_sharepoint, url, nome): nome
                for nome, url in [
                    ('DRAFT_PDVS', urls_planilhas['draft_pdvs']),
                    ('CONSULTA_DE_ESTOQUE', urls_planilhas['estoque_seguranca']),
                    ('Planilha Retaguarda', urls_planilhas['retaguarda']),
                    ('SKUs Ignorados', urls_planilhas['ignorados'])
                ]
            }
            
            for future in as_completed(futures):
                nome = futures[future]
                try:
                    buffer = future.result()
                    if nome == 'DRAFT_PDVS':
                        buffers['draft'] = buffer
                    elif nome == 'CONSULTA_DE_ESTOQUE':
                        buffers['seguranca'] = buffer
                    elif nome == 'Planilha Retaguarda':
                        buffers['retaguarda'] = buffer
                    elif nome == 'SKUs Ignorados':
                        buffers['ignorados'] = buffer
                except Exception as e:
                    logger.error(f"Erro no download de {nome}: {str(e)}")
                    logger.error(traceback.format_exc())
    
    else:
        logger.info("📡 Usando fonte de dados: GOOGLE SHEETS")
        urls_planilhas = {
            'draft_pdvs': obter_url_exportacao('draft_pdvs'),
            'estoque_seguranca': obter_url_exportacao('estoque_seguranca'),
            'retaguarda': obter_url_exportacao('retaguarda'),
            'ignorados': obter_url_exportacao('ignorados')
        }
        
        buffers = {}
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {
                executor.submit(download_planilha_excel, url, nome): nome
                for nome, url in [
                    ('DRAFT_PDVS', urls_planilhas['draft_pdvs']),
                    ('CONSULTA_DE_ESTOQUE', urls_planilhas['estoque_seguranca']),
                    ('Planilha Retaguarda', urls_planilhas['retaguarda']),
                    ('SKUs Ignorados', urls_planilhas['ignorados'])
                ]
            }
            
            for future in as_completed(futures):
                nome = futures[future]
                try:
                    buffer = future.result()
                    if nome == 'DRAFT_PDVS':
                        buffers['draft'] = buffer
                    elif nome == 'CONSULTA_DE_ESTOQUE':
                        buffers['seguranca'] = buffer
                    elif nome == 'Planilha Retaguarda':
                        buffers['retaguarda'] = buffer
                    elif nome == 'SKUs Ignorados':
                        buffers['ignorados'] = buffer
                except Exception as e:
                    logger.error(f"Erro no download de {nome}: {str(e)}")
                    logger.error(traceback.format_exc())
    
    logger.info("✅ Todos os downloads concluídos")
    
    logger.info("Fase 2: Carregando SKUs ignorados...")
    st.session_state['progresso_atual'] = "Carregando SKUs ignorados..."
    skus_ignorados = carregar_skus_ignorados(buffers.get('ignorados'))
    stats['skus_ignorados'] = len(skus_ignorados)
    
    logger.info("Fase 3: Processamento dos dados...")
    st.session_state['progresso_atual'] = "Processando planilha principal..."
    
    dados_marcas, data_atualizacao = carregar_dados_principais(buffers.get('draft'))
    if not dados_marcas:
        logger.error("Nenhum dado foi carregado da planilha principal")
        return {}, None, stats
    
    st.session_state['progresso_atual'] = "Carregando estoque de segurança..."
    df_estoque_seguranca = carregar_estoque_seguranca(buffers.get('seguranca'))
    
    st.session_state['progresso_atual'] = "Carregando planilha de custos..."
    df_retaguarda = carregar_planilha_retaguarda(buffers.get('retaguarda'))
    stats['total_skus_retaguarda'] = len(df_retaguarda) if not df_retaguarda.empty else 0
    
    # BUG 1: Verifica se Retaguarda não tem coluna de custo
    if df_retaguarda.empty:
        stats['retaguarda_sem_custo'] = True
        logger.warning("⚠️ Retaguarda vazia ou sem coluna de custo. Custos não serão aplicados.")
    
    if skus_ignorados:
        logger.info(f"Fase 4: Filtrando {len(skus_ignorados)} SKUs ignorados...")
        
        if not df_retaguarda.empty:
            antes = len(df_retaguarda)
            df_retaguarda = df_retaguarda[~df_retaguarda['SKU'].isin(skus_ignorados)]
            logger.info(f"Retaguarda: {antes - len(df_retaguarda)} registros removidos (ignorados)")
        
        if not df_estoque_seguranca.empty:
            antes = len(df_estoque_seguranca)
            df_estoque_seguranca = df_estoque_seguranca[~df_estoque_seguranca['SKU'].isin(skus_ignorados)]
            logger.info(f"Segurança: {antes - len(df_estoque_seguranca)} registros removidos (ignorados)")
        
        total_filtrados = 0
        for nome_marca in dados_marcas.keys():
            antes = len(dados_marcas[nome_marca])
            dados_marcas[nome_marca] = dados_marcas[nome_marca][~dados_marcas[nome_marca]['SKU'].isin(skus_ignorados)]
            filtrados = antes - len(dados_marcas[nome_marca])
            total_filtrados += filtrados
            if filtrados > 0:
                logger.info(f"{nome_marca}: {filtrados} registros removidos (ignorados)")
        
        stats['skus_ignorados_filtrados'] = total_filtrados
    
    logger.info("Fase 5: Aplicando custos da Retaguarda...")
    st.session_state['progresso_atual'] = "Aplicando custos e calculando métricas..."
    
    df_update_global = None
    if not df_retaguarda.empty:
        df_update_global = df_retaguarda[['PDV', 'SKU', 'CUSTO_RETARGUARDA']].copy()
        df_update_global = df_update_global.set_index(['PDV', 'SKU'])
        logger.info(f"DataFrame de custos preparado: {len(df_update_global)} registros")
    
    for nome_marca, df in dados_marcas.items():
        df['CUSTO_RETARGUARDA'] = 0.0
        
        if df_update_global is not None:
            idx = pd.MultiIndex.from_arrays([df['PDV'], df['SKU']])
            mask = idx.isin(df_update_global.index)
            
            df.loc[mask, 'CUSTO_RETARGUARDA'] = df_update_global.loc[idx[mask], 'CUSTO_RETARGUARDA'].values
            
            stats['skus_matcheados'] += int(mask.sum())
            stats['skus_sem_match'] += len(df) - int(mask.sum())
        else:
            stats['skus_sem_match'] += len(df)
        
        preco_tabela = df['Preço tabela'].fillna(0)
        custo_retaguarda = df['CUSTO_RETARGUARDA'].fillna(0)
        
        cond1 = (df['Preço tabela'].isna()) | (preco_tabela == 0)
        cond2 = custo_retaguarda > 0
        
        df['Preço de Custo'] = np.where(
            cond1,
            custo_retaguarda,
            np.where(cond2, np.maximum(preco_tabela, custo_retaguarda), preco_tabela)
        )
        
        custo_maior = (custo_retaguarda > preco_tabela) & (custo_retaguarda > 0)
        stats['skus_custo_maior'] += int(custo_maior.sum())
        
        df = df.drop(columns=['CUSTO_RETARGUARDA'], errors='ignore')
        
        if not df_estoque_seguranca.empty:
            df_seg = df_estoque_seguranca[df_estoque_seguranca['MARCA_REFERENCIA'] == nome_marca].copy()
            
            if not df_seg.empty:
                df = df.merge(df_seg[['PDV', 'SKU', 'ESTOQUE_DE_SEGURANCA']], on=['PDV', 'SKU'], how='left')
                df['Estoque_Minimo_Qtd'] = df['ESTOQUE_DE_SEGURANCA'].fillna(0)
                df = df.drop(columns=['ESTOQUE_DE_SEGURANCA'])
            else:
                df['Estoque_Minimo_Qtd'] = df['Classe'].map(REGRAS_ESTOQUE_MINIMO).fillna(2) if 'Classe' in df.columns else 2
        else:
            df['Estoque_Minimo_Qtd'] = df['Classe'].map(REGRAS_ESTOQUE_MINIMO).fillna(2) if 'Classe' in df.columns else 2
        
        df['Valor_Estoque_Atual'] = df['Estoque Atual'] * df['Preço tabela']
        df['Valor_Estoque_Minimo'] = df['Estoque_Minimo_Qtd'] * df['Preço tabela']
        df['Qtd_Excesso'] = (df['Estoque Atual'] - df['Estoque_Minimo_Qtd']).clip(lower=0)
        df['Valor_Excesso'] = df['Qtd_Excesso'] * df['Preço tabela']
        df['Qtd_Falta'] = (df['Estoque_Minimo_Qtd'] - df['Estoque Atual']).clip(lower=0)
        df['Valor_Falta'] = df['Qtd_Falta'] * df['Preço tabela']
        df['Valor_Custo_Estoque_Atual'] = df['Estoque Atual'] * df['Preço de Custo']
        df['Valor_Custo_Estoque_Minimo'] = df['Estoque_Minimo_Qtd'] * df['Preço de Custo']
        
        dados_marcas[nome_marca] = df
    
    logger.info("=" * 60)
    logger.info("CARREGAMENTO CONCLUÍDO")
    logger.info(f"Total SKUs Retaguarda: {stats['total_skus_retaguarda']}")
    logger.info(f"SKUs matcheados: {stats['skus_matcheados']}")
    logger.info(f"SKUs sem custo: {stats['skus_sem_match']}")
    logger.info(f"SKUs com custo > tabela: {stats['skus_custo_maior']}")
    logger.info(f"SKUs ignorados configurados: {stats['skus_ignorados']}")
    logger.info(f"SKUs ignorados filtrados: {stats['skus_ignorados_filtrados']}")
    logger.info(f"Modo de acesso: {stats['modo_acesso']}")
    logger.info(f"Retaguarda sem custo: {stats['retaguarda_sem_custo']}")
    logger.info("=" * 60)
    
    st.session_state['progresso_atual'] = "Carregamento concluído!"
    
    return dados_marcas, data_atualizacao, stats


# ==============================================================================
# FUNÇÕES DE UI
# ==============================================================================

def exibir_kpi_card(col, icone, titulo, valor_fmt, delta_texto=None, cor_delta="#ff4b4b", help_text=None):
    delta_html = f'<div style="font-size:12px;color:{cor_delta};margin-top:4px;">{delta_texto}</div>' if delta_texto else ''
    help_icon = f'<span title="{help_text}" style="float:right;cursor:help;opacity:0.6;">ℹ️</span>' if help_text else ''
    
    col.markdown(f"""
    <div style="
        background: linear-gradient(135deg,#111820,#0d1f14);
        border:1px solid #1a3d25; border-left:4px solid #007A33;
        border-radius:10px; padding:18px 20px;
        box-shadow:0 4px 16px rgba(0,122,51,0.15);
        min-height:110px;
        position:relative;
    ">
        {help_icon}
        <div style="font-size:22px;margin-bottom:4px;">{icone}</div>
        <div style="font-size:12px;color:#8da9be;margin-bottom:6px;">{titulo}</div>
        <div style="font-size:26px;font-weight:700;color:#D4AF37;
                    text-shadow:0 0 10px rgba(212,175,55,0.3);">{valor_fmt}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)


def exibir_titulo_marca(nome_marca, tamanho_logo=30):
    cor = CORES_MARCAS.get(nome_marca, '#007A33')
    col_logo, col_nome = st.columns([0.08, 0.92])
    with col_logo:
        logo_path = LOGOS_MARCAS.get(nome_marca, '')
        if logo_path:
            try:
                if os.path.exists(logo_path):
                    st.image(logo_path, width=tamanho_logo)
                else:
                    st.markdown('<div class="logo-placeholder"></div>', unsafe_allow_html=True)
                    logger.warning(f"Logo não encontrada: {logo_path}")
            except Exception as e:
                st.write("🏷️")
                logger.error(f"Erro ao carregar logo {logo_path}: {str(e)}")
        else:
            st.write("🏷️")
    with col_nome:
        st.markdown(f"""
        <div style="
            border-left: 4px solid {cor};
            padding: 6px 14px;
            background: linear-gradient(90deg, {cor}22, transparent);
            border-radius: 0 8px 8px 0;
            margin-bottom: 4px;
        ">
            <span style="font-size:18px; font-weight:700; color:{cor};">{nome_marca}</span>
        </div>
        """, unsafe_allow_html=True)


def obter_horario_brasilia():
    fuso_brasilia = timezone(timedelta(hours=-3))
    return datetime.now(fuso_brasilia).strftime("%d/%m/%Y às %H:%M:%S")


def gerar_pdf_dashboard(dados_filtrados, pdv_selecionado, loja_selecionada_nome,
                        marca_selecionada, horario_brasilia,
                        v_estoque_atual_total, v_estoque_min_total,
                        v_excesso_total_total, v_falta_total_total, qtd_itens_total):
    if not REPORTLAB_AVAILABLE:
        st.error("Biblioteca reportlab não instalada.")
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)

    COR_VERDE = colors.HexColor('#007A33')
    COR_DOURADO = colors.HexColor('#D4AF37')
    COR_TEXTO = colors.HexColor('#333333')
    COR_VERMELHO = colors.HexColor('#dc2626')

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle('titulo', fontSize=16, textColor=COR_VERDE, fontName='Helvetica-Bold', spaceAfter=4)
    estilo_subtitulo = ParagraphStyle('subtitulo', fontSize=11, textColor=COR_VERDE, fontName='Helvetica-Bold', spaceAfter=4)
    estilo_normal = ParagraphStyle('normal', fontSize=9, textColor=COR_TEXTO, fontName='Helvetica', spaceAfter=2)
    estilo_rodape = ParagraphStyle('rodape', fontSize=8, textColor=colors.HexColor('#6b7e8a'), fontName='Helvetica', alignment=1)

    elementos = []

    try:
        if os.path.exists('logo_cp_fani.png'):
            logo = RLImage('logo_cp_fani.png', width=3*cm, height=2*cm)
            elementos.append(logo)
    except Exception as e:
        logger.error(f"Erro ao carregar logo no PDF: {str(e)}")

    elementos.append(Paragraph("Painel de Performance de Estoque NSF · CP Fani", estilo_titulo))
    elementos.append(Paragraph(f"PDV: {loja_selecionada_nome}  |  Marca: {marca_selecionada}", estilo_subtitulo))
    elementos.append(Paragraph(f"Gerado em: {horario_brasilia}", estilo_normal))
    elementos.append(Spacer(1, 0.5*cm))

    dados_kpi = [
        ['KPI', 'Valor'],
        ['Valor em Estoque (Tabela)', f"R$ {v_estoque_atual_total:,.2f}"],
        ['Estoque Mínimo (Tabela)', f"R$ {v_estoque_min_total:,.2f}"],
        ['Capital Preso (Excesso)', f"R$ {v_excesso_total_total:,.2f}"],
        ['Risco de Ruptura (Falta)', f"R$ {v_falta_total_total:,.2f}"],
        ['Total de Unidades', f"{int(qtd_itens_total):,}"],
    ]
    tabela_kpi = Table(dados_kpi, colWidths=[10*cm, 6*cm])
    tabela_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COR_VERDE),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8faf9')),
        ('TEXTCOLOR', (0, 1), (0, -1), COR_TEXTO),
        ('TEXTCOLOR', (1, 1), (1, -1), COR_VERDE),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f8faf9'), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#1a3d25')),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))
    elementos.append(tabela_kpi)
    elementos.append(Spacer(1, 0.5*cm))

    for nome_marca, df_completo in dados_filtrados.items():
        if pdv_selecionado == 'TODAS':
            df_loja = df_completo
        else:
            df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
        if df_loja.empty:
            continue

        df_flt = df_loja[df_loja['Valor_Falta'] > 0].sort_values('Valor_Falta', ascending=False).head(20)
        if not df_flt.empty:
            elementos.append(Paragraph(f"Produtos Críticos em Falta — {nome_marca}", estilo_subtitulo))
            colunas_flt = ['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Falta', 'Valor_Falta']
            colunas_flt_existentes = [c for c in colunas_flt if c in df_flt.columns]
            dados_flt = [colunas_flt_existentes]
            for _, row in df_flt[colunas_flt_existentes].iterrows():
                linha = []
                for col in colunas_flt_existentes:
                    v = row[col]
                    if col == 'Valor_Falta':
                        linha.append(f"R$ {float(v):,.2f}")
                    else:
                        linha.append(str(v))
                dados_flt.append(linha)
            t_flt = Table(dados_flt, repeatRows=1)
            t_flt.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), COR_VERMELHO),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#fef2f2')),
                ('TEXTCOLOR', (0, 1), (-1, -1), COR_TEXTO),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#fef2f2'), colors.white]),
                ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#fca5a5')),
                ('PADDING', (0, 0), (-1, -1), 4),
                ('TEXTCOLOR', (-1, 1), (-1, -1), COR_VERMELHO),
            ]))
            elementos.append(t_flt)
            elementos.append(Spacer(1, 0.4*cm))

    elementos.append(Spacer(1, 1*cm))
    elementos.append(Paragraph(f"Grupo NSF · CP Fani  |  Gerado em: {horario_brasilia}", estilo_rodape))

    try:
        doc.build(elementos)
        buffer.seek(0)
        return buffer
    except Exception as e:
        logger.error(f"Erro ao gerar PDF: {str(e)[:200]}")
        logger.error(traceback.format_exc())
        return None


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    progresso_placeholder = st.empty()
    progresso_placeholder.info("🔄 Iniciando carregamento das planilhas...")
    
    try:
        dados_marcas, data_atualizacao, stats = carregar_todos_os_dados()
    except Exception as e:
        logger.error(f"Erro crítico no carregamento: {str(e)}")
        logger.error(traceback.format_exc())
        st.error(f"❌ Erro crítico: {str(e)}")
        st.stop()
    
    progresso_placeholder.empty()
    
    horario_carregamento = obter_horario_brasilia()
    
    if not dados_marcas:
        st.error("❌ Nenhum dado foi carregado. Verifique os logs acima.")
        st.info("""
        ### 🔧 Passos para resolver:
        
        1. **Verifique o MODO_ACESSO no config.py:**
           - "publico" = Google Sheets público
           - "sharepoint" = SharePoint público
        
        2. **Verifique se as planilhas estão acessíveis:**
           - Google: "Qualquer pessoa com o link" → "Leitor"
           - SharePoint: Link de compartilhamento público
        
        3. **Verifique os IDs/URLs no config.py**
        
        4. **Verifique os nomes das abas:**
           - DRAFT_PDVS: BOTICARIO, EUDORA, QUEM_DISSE_BERENICE
           - CONSULTA_DE_ESTOQUE: BOT, EUD, QDB
        """)
        st.stop()
    
    if data_atualizacao:
        horario_exibicao = data_atualizacao
        info_timestamp = "🕒 Última atualização"
    else:
        horario_exibicao = horario_carregamento
        info_timestamp = "🕒 Carregado em"
    
    col_logo, col_info = st.columns([1, 3])
    
    with col_logo:
        try:
            if os.path.exists("logo_cp_fani.png"):
                st.image("logo_cp_fani.png", width=180)
            else:
                st.markdown('<div class="logo-placeholder" style="width:180px;height:120px;"></div>', unsafe_allow_html=True)
                logger.warning("Logo CP Fani não encontrada")
        except Exception as e:
            logger.error(f"Erro ao carregar logo principal: {str(e)}")
    
    with col_info:
        if esta_no_modo_sharepoint():
            modo_badge = '<span style="background:#0078d4;color:white;padding:2px 8px;border-radius:4px;font-size:10px;">SHAREPOINT</span>'
        elif esta_no_modo_privado():
            modo_badge = '<span style="background:#f59e0b;color:white;padding:2px 8px;border-radius:4px;font-size:10px;">GOOGLE PRIVADO</span>'
        else:
            modo_badge = '<span style="background:#007A33;color:white;padding:2px 8px;border-radius:4px;font-size:10px;">GOOGLE PÚBLICO</span>'
        
        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, #0d1f14 0%, #080b0f 100%);
            border: 1px solid #1a3d25; border-left: 5px solid #007A33;
            border-radius: 12px; padding: 20px 28px; margin-bottom: 8px;
        ">
            <div style="font-size:13px; color:#8da9be; margin-bottom:4px; letter-spacing:1px; text-transform:uppercase;">
                Grupo NSF · CP Fani {modo_badge}
            </div>
            <div style="font-size:26px; font-weight:700; color:#D4AF37; line-height:1.2;">
                📊 Painel de Controle de Estoques e Ruptura
            </div>
            <div style="font-size:12px; color:#6b7e8a; margin-top:6px;">
                {info_timestamp}: <span style="color:#a3b8cc;">{horario_exibicao}</span> · Horário de Brasília
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ==========================================================================
    # BUG 1 FIX: ALERTA VISUAL QUANDO RETAGUARDA NÃO TEM CUSTOS
    # ==========================================================================
    if stats.get('retaguarda_sem_custo', False):
        st.error("""
        ⚠️ **ALERTA CRÍTICO: Custos da Retaguarda não estão sendo aplicados!**
        
        A planilha de Retaguarda foi carregada mas **não contém coluna de CUSTO** ou está vazia.
        Todos os cálculos de "Preço de Custo" estão usando o **Preço de Tabela** como fallback.
        
        **Impacto:**
        - KPI "Capital Preso" pode estar incorreto
        - Análise de custos por curva pode estar distorcida
        
        **Ação necessária:**
        1. Verifique se a planilha correta está configurada no `config.py`
        2. Confirme se a planilha tem uma coluna chamada "CUSTO", "PREÇO DE CUSTO", etc.
        3. Verifique os logs para mais detalhes
        """)
    
    # BUG 1 FIX: Alerta quando SKUs matcheados = 0 mas Retaguarda tem dados
    if stats.get('total_skus_retaguarda', 0) > 0 and stats.get('skus_matcheados', 0) == 0:
        st.warning("""
        ⚠️ **ALERTA: Nenhum SKU da Retaguarda foi casado com a planilha principal!**
        
        A planilha de Retaguarda tem dados, mas **nenhum SKU encontrou correspondência** na planilha principal.
        
        **Possíveis causas:**
        1. Formatos de SKU diferentes entre as planilhas (ex: "123" vs "0123")
        2. Planilhas de períodos diferentes
        3. Problema na normalização de SKUs
        
        **Ação:** Verifique os logs para diagnóstico detalhado dos SKUs.
        """)
    
    with st.expander("🔍 Diagnóstico de Carregamento (clique para ver)", expanded=False):
        col_d1, col_d2, col_d3, col_d4 = st.columns(4)
        col_d1.metric("SKUs na Retaguarda", f"{stats.get('total_skus_retaguarda', 0):,}",
                     help="Total de registros de custo encontrados na planilha Retaguarda")
        col_d2.metric("SKUs com custo", f"{stats.get('skus_matcheados', 0):,}",
                     help="SKUs que encontraram correspondência entre planilha principal e Retaguarda")
        col_d3.metric("SKUs sem custo", f"{stats.get('skus_sem_match', 0):,}",
                     help="SKUs sem custo na Retaguarda (usando Preço de Tabela como fallback)")
        col_d4.metric("Custo > Tabela", f"{stats.get('skus_custo_maior', 0):,}",
                     help="SKUs onde o custo da Retaguarda é maior que o Preço de Tabela")
        
        if stats.get('skus_ignorados', 0) > 0:
            st.markdown("---")
            col_i1, col_i2, col_i3, _ = st.columns(4)
            col_i1.metric("SKUs Ignorados Configurados", f"{stats['skus_ignorados']:,}",
                         help="Total de SKUs na planilha de ignorados")
            col_i2.metric("SKUs Filtrados", f"{stats['skus_ignorados_filtrados']:,}",
                         help="SKUs ignorados efetivamente removidos dos cálculos")
            col_i3.metric("SKUs Ativos", f"{stats.get('skus_matcheados', 0) + stats.get('skus_sem_match', 0):,}",
                         help="SKUs ativos após filtro de ignorados")
        else:
            st.warning("⚠️ **Planilha de SKUs Ignorados não carregada.** Todos os SKUs serão incluídos nos cálculos.")
    
    if stats.get('total_skus_retaguarda', 0) == 0:
        st.warning("⚠️ **Planilha Retaguarda não carregada.** Usando Preço de Tabela como custo para todos os SKUs.")
    
    st.markdown("---")
    
    st.sidebar.title("Filtros de Visualização")
    
    todos_pdvs = sorted(set(
        int(pdv)
        for df in dados_marcas.values()
        for pdv in df['PDV'].dropna()
    ))
    
    opcoes_selectbox = ["Todas as Lojas"] + [DE_PARA_LOJAS.get(pdv, f"PDV {pdv}") for pdv in todos_pdvs]
    
    loja_selecionada_nome = st.sidebar.selectbox(
        "Selecione a Loja / PDV:", 
        opcoes_selectbox,
        help="Escolha um PDV específico ou 'Todas as Lojas' para ver o consolidado geral"
    )
    
    if loja_selecionada_nome == "Todas as Lojas":
        pdv_selecionado = 'TODAS'
    else:
        pdv_selecionado = DE_PARA_LOJAS_REVERSO.get(loja_selecionada_nome)
        if pdv_selecionado is None:
            st.error("PDV não reconhecido.")
            st.stop()
    
    st.sidebar.markdown("---")
    
    st.sidebar.subheader("Filtro de Marca")
    opcoes_marca = ["Todas as Marcas"] + list(dados_marcas.keys())
    marca_selecionada = st.sidebar.selectbox(
        "Selecione a Marca:", 
        opcoes_marca,
        help="Filtre por uma marca específica ou veja todas consolidadas"
    )
    
    st.sidebar.markdown("**Marcas disponíveis:**")
    col_logos_sidebar = st.sidebar.columns(len(dados_marcas))
    for idx, (nome_marca, df_marca) in enumerate(dados_marcas.items()):
        with col_logos_sidebar[idx]:
            logo_path = LOGOS_MARCAS.get(nome_marca, '')
            if logo_path:
                try:
                    if os.path.exists(logo_path):
                        st.image(logo_path, width=40)
                    else:
                        st.markdown('<div class="logo-placeholder"></div>', unsafe_allow_html=True)
                except Exception as e:
                    logger.error(f"Erro ao carregar logo {logo_path}: {str(e)}")
            st.caption(nome_marca)
    
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Forçar Atualização", help="Limpa o cache e recarrega todos os dados das planilhas"):
        st.cache_data.clear()
        st.rerun()
    
    if pdv_selecionado == 'TODAS':
        st.markdown(f"""
        <div style="
            display:inline-block;
            background: linear-gradient(135deg,#0d1f14,#111820);
            border:1px solid #007A33; border-radius:20px;
            padding:6px 20px; margin-bottom:12px;
        ">
            <span style="color:#8da9be;font-size:13px;">🏪 PDV: </span>
            <span style="color:#D4AF37;font-weight:700;font-size:15px;">Todas as Lojas ({len(todos_pdvs)})</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div style="
            display:inline-block;
            background: linear-gradient(135deg,#0d1f14,#111820);
            border:1px solid #007A33; border-radius:20px;
            padding:6px 20px; margin-bottom:12px;
        ">
            <span style="color:#8da9be;font-size:13px;">🏪 PDV: </span>
            <span style="color:#D4AF37;font-weight:700;font-size:15px;">{loja_selecionada_nome}</span>
        </div>
        """, unsafe_allow_html=True)
    
    if marca_selecionada == "Todas as Marcas":
        dados_filtrados = dados_marcas
        titulo_secao = "Consolidado Geral"
    else:
        dados_filtrados = {marca_selecionada: dados_marcas[marca_selecionada]}
        titulo_secao = f"Análise: {marca_selecionada}"
    
    st.markdown(f"### {titulo_secao}")
    st.markdown("---")
    
    v_estoque_atual_total = 0
    v_estoque_min_total = 0
    v_excesso_total_total = 0
    v_falta_total_total = 0
    qtd_itens_total = 0
    
    for nome_marca, df_completo in dados_filtrados.items():
        if pdv_selecionado == 'TODAS':
            df_loja = df_completo
        else:
            df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
        
        if not df_loja.empty:
            v_estoque_atual_total += df_loja['Valor_Estoque_Atual'].sum()
            v_estoque_min_total += df_loja['Valor_Estoque_Minimo'].sum()
            v_excesso_total_total += df_loja['Valor_Excesso'].sum()
            v_falta_total_total += df_loja['Valor_Falta'].sum()
            qtd_itens_total += df_loja['Est
