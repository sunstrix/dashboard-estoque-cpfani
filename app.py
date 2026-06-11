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

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Importa configuração centralizada
from config import (
    PLANILHAS, DE_PARA_LOJAS, DE_PARA_LOJAS_REVERSO, 
    MAPEAMENTO_PDV_DRAFT_RAW, NOMES_MARCAS, ABAS_SEGURANCA,
    LOGOS_MARCAS, CORES_MARCAS, REGRAS_ESTOQUE_MINIMO,
    COLUNAS_OBRIGATORIAS, TIMEOUT_DOWNLOAD, CACHE_TTL
)

# ==========================================
# CONFIGURAÇÃO DE LOGGING
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E CSS
# ==========================================
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
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE NORMALIZAÇÃO
# ==========================================

def normalizar_sku(valor):
    """Normaliza SKUs para garantir correspondência entre planilhas."""
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


def normalizar_nome_loja(nome):
    """Normaliza nomes de loja para comparação tolerante."""
    if pd.isna(nome) or nome is None:
        return ""
    s = str(nome).strip().lower()
    s = re.sub(r'[^a-z0-9\s]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


MAPEAMENTO_PDV_DRAFT_NORMALIZADO = {
    normalizar_nome_loja(k): v for k, v in MAPEAMENTO_PDV_DRAFT_RAW.items()
}


# ==========================================
# FUNÇÕES DE DOWNLOAD OTIMIZADAS
# ==========================================

def criar_sessao_com_retry():
    """Cria sessão requests com retry automático."""
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
    """
    Download otimizado de planilha Excel.
    Cada planilha é baixada APENAS UMA VEZ por ciclo de cache.
    """
    logger.info(f"Iniciando download: {nome_planilha}")
    session = criar_sessao_com_retry()
    
    try:
        # Download direto sem HEAD prévio (performance)
        response = session.get(url_planilha, timeout=TIMEOUT_DOWNLOAD, stream=True)
        response.raise_for_status()
        
        content = response.content
        tamanho_kb = len(content) / 1024
        
        logger.info(f"✅ Download concluído: {nome_planilha} ({tamanho_kb:.1f} KB)")
        return BytesIO(content)
        
    except Exception as e:
        logger.error(f"❌ Erro ao baixar {nome_planilha}: {str(e)}")
        st.error(f"Falha ao carregar {nome_planilha}: {str(e)[:200]}")
        return None


# ==========================================
# FUNÇÕES DE CARREGAMENTO DE DADOS
# ==========================================

@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_planilha_retaguarda(custos_buffer):
    """
    Carrega planilha de custos (Retaguarda).
    """
    logger.info("Carregando planilha Retaguarda (custos)...")
    
    try:
        if custos_buffer is None:
            logger.warning("Planilha Retaguarda não fornecida")
            return pd.DataFrame()
        
        excel_file = pd.ExcelFile(custos_buffer)
        logger.info(f"Abas encontradas na Retaguarda: {excel_file.sheet_names}")
        
        # Usa a primeira aba disponível
        if not excel_file.sheet_names:
            logger.error("Retaguarda: nenhuma aba encontrada")
            return pd.DataFrame()
        
        df = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
        
        if df.empty:
            logger.warning("Retaguarda: planilha vazia")
            return pd.DataFrame()
        
        # Normaliza nomes das colunas
        df.columns = [str(col).strip().upper() for col in df.columns]
        logger.info(f"Colunas encontradas: {list(df.columns)}")
        
        # Detecta coluna de SKU
        colunas_sku = ['SKU', 'CÓDIGO', 'CODIGO', 'CÓDIGO SKU', 'CODIGO SKU', 
                       'EAN', 'COD PRODUTO', 'CÓDIGO PRODUTO']
        coluna_sku = next((c for c in colunas_sku if c in df.columns), None)
        
        # Detecta coluna de Custo
        colunas_custo = ['CUSTO', 'PREÇO DE CUSTO', 'PRECO DE CUSTO', 'CUSTO UNITÁRIO',
                        'CUSTO UNITARIO', 'VALOR CUSTO', 'CUSTO (R$)', 'CUSTO R$']
        coluna_custo = next((c for c in colunas_custo if c in df.columns), None)
        
        # Fallback: tenta coluna J (índice 9)
        if coluna_custo is None and len(df.columns) > 9:
            coluna_custo = df.columns[9]
            logger.info(f"Usando coluna J como custo: {coluna_custo}")
        
        if coluna_sku is None or coluna_custo is None:
            logger.error(f"Retaguarda: colunas faltantes. SKU={coluna_sku}, Custo={coluna_custo}")
            return pd.DataFrame()
        
        # Prepara DataFrame
        df_resultado = pd.DataFrame()
        df_resultado['SKU'] = df[coluna_sku].apply(normalizar_sku)
        df_resultado['CUSTO_RETARGUARDA'] = pd.to_numeric(df[coluna_custo], errors='coerce').fillna(0)
        
        # Detecta coluna de loja/PDV (opcional)
        colunas_loja = ['LOJA', 'PDV', 'LOJA/PDV', 'LOJA - PDV', 'CÓDIGO LOJA', 'CODIGO LOJA',
                       'FILIAL', 'COD FILIAL', 'LOJA NOME', 'NOME LOJA']
        coluna_loja = next((c for c in colunas_loja if c in df.columns), None)
        
        if coluna_loja is None:
            # Sem coluna de loja → aplica custos para TODOS os PDVs
            logger.info("Retaguarda: sem coluna de loja. Aplicando custos para todos os PDVs.")
            todos_pdvs = list(DE_PARA_LOJAS.keys())
            df_expandido = []
            for pdv in todos_pdvs:
                df_temp = df_resultado.copy()
                df_temp['PDV'] = pdv
                df_expandido.append(df_temp)
            df_resultado = pd.concat(df_expandido, ignore_index=True)
        else:
            # Com coluna de loja → mapeia normalmente
            df_resultado['LOJA_NOME'] = df[coluna_loja].astype(str)
            df_resultado['LOJA_NORM'] = df_resultado['LOJA_NOME'].apply(normalizar_nome_loja)
            df_resultado['PDV'] = df_resultado['LOJA_NORM'].map(MAPEAMENTO_PDV_DRAFT_NORMALIZADO)
            
            # Filtra apenas PDVs válidos
            total_antes = len(df_resultado)
            df_resultado = df_resultado[df_resultado['PDV'].notna()].copy()
            total_depois = len(df_resultado)
            
            if total_depois < total_antes:
                logger.warning(f"Retaguarda: {total_antes - total_depois} linhas sem PDV válido")
        
        if df_resultado.empty:
            logger.error("Retaguarda: nenhum dado válido após processamento")
            return pd.DataFrame()
        
        df_resultado['PDV'] = df_resultado['PDV'].astype(int)
        
        # Remove duplicatas (mantém maior custo)
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
    """Carrega planilha de estoque de segurança."""
    logger.info("Carregando planilha de estoque de segurança...")
    
    try:
        if seguranca_buffer is None:
            logger.warning("Planilha de segurança não fornecida")
            return pd.DataFrame()
        
        excel_file = pd.ExcelFile(seguranca_buffer)
        abas_esperadas = ['BOT', 'EUD', 'QDB']
        abas_encontradas = [aba for aba in abas_esperadas if aba.upper() in [a.upper() for a in excel_file.sheet_names]]
        
        if not abas_encontradas:
            logger.error(f"Segurança: abas esperadas não encontradas. Disponíveis: {excel_file.sheet_names}")
            return pd.DataFrame()
        
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
                
                # Detecta coluna de estoque de segurança
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
                continue
        
        if dfs_abas:
            df_final = pd.concat(dfs_abas, ignore_index=True)
            logger.info(f"✅ Estoque de segurança carregado: {len(df_final)} registros")
            return df_final
        
        logger.warning("Segurança: nenhum dado válido encontrado")
        return pd.DataFrame()
        
    except Exception as e:
        logger.error(f"Erro ao carregar segurança: {str(e)}")
        logger.error(traceback.format_exc())
        return pd.DataFrame()


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_dados_principais(draft_buffer):
    """
    Carrega planilha principal (DRAFT_PDVS).
    Valida colunas obrigatórias antes de processar.
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
                df = pd.read_excel(excel_file, sheet_name=aba_excel)
                
                # Valida colunas obrigatórias
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


# ==========================================
# FUNÇÃO PRINCIPAL DE CARREGAMENTO
# ==========================================

@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_todos_os_dados():
    """
    Função principal que orquestra o carregamento de todas as planilhas.
    Performance: cada planilha é baixada apenas UMA vez.
    """
    logger.info("=" * 60)
    logger.info("INICIANDO CARREGAMENTO DE DADOS")
    logger.info("=" * 60)
    
    stats = {
        'total_skus_retaguarda': 0,
        'skus_matcheados': 0,
        'skus_sem_match': 0,
        'skus_custo_maior': 0
    }
    
    # 1. Download das 3 planilhas (cada uma apenas UMA vez)
    logger.info("Fase 1: Download das planilhas...")
    
    buffer_draft = download_planilha_excel(
        f"https://docs.google.com/spreadsheets/d/{PLANILHAS['draft_pdvs']['id']}/export?format=xlsx",
        "DRAFT_PDVS"
    )
    
    buffer_seguranca = download_planilha_excel(
        f"https://docs.google.com/spreadsheets/d/{PLANILHAS['estoque_seguranca']['id']}/export?format=xlsx",
        "CONSULTA_DE_ESTOQUE"
    )
    
    buffer_retaguarda = download_planilha_excel(
        f"https://docs.google.com/spreadsheets/d/{PLANILHAS['retaguarda']['id']}/export?format=xlsx",
        "Planilha Retaguarda"
    )
    
    # 2. Carregamento e processamento
    logger.info("Fase 2: Processamento dos dados...")
    
    dados_marcas, data_atualizacao = carregar_dados_principais(buffer_draft)
    if not dados_marcas:
        logger.error("Nenhum dado foi carregado da planilha principal")
        return {}, None, stats
    
    df_estoque_seguranca = carregar_estoque_seguranca(buffer_seguranca)
    df_retaguarda = carregar_planilha_retaguarda(buffer_retaguarda)
    stats['total_skus_retaguarda'] = len(df_retaguarda) if not df_retaguarda.empty else 0
    
    # 3. Aplicação dos custos da Retaguarda (CORREÇÃO: usando abordagem mais robusta)
    logger.info("Fase 3: Aplicando custos da Retaguarda...")
    
    for nome_marca, df in dados_marcas.items():
        # CORREÇÃO: Inicializa a coluna ANTES de qualquer operação
        df['CUSTO_RETARGUARDA'] = 0.0
        
        if not df_retaguarda.empty:
            # Cria um DataFrame com os valores que queremos atualizar
            df_update = df_retaguarda[['PDV', 'SKU', 'CUSTO_RETARGUARDA']].copy()
            df_update = df_update.set_index(['PDV', 'SKU'])
            
            # Atualiza os valores no DataFrame principal
            idx = pd.MultiIndex.from_arrays([df['PDV'], df['SKU']])
            
            # Filtra os índices que estão no df_update
            mask = idx.isin(df_update.index)
            
            # Atualiza apenas os valores que existem
            df.loc[mask, 'CUSTO_RETARGUARDA'] = df_update.loc[idx[mask], 'CUSTO_RETARGUARDA'].values
            
            # Conta matches
            stats['skus_matcheados'] += int(mask.sum())
            stats['skus_sem_match'] += len(df) - int(mask.sum())
        else:
            stats['skus_sem_match'] += len(df)
        
        # Regra de custo: usa o MAIOR entre preço tabela e custo retaguarda
        preco_tabela = df['Preço tabela'].fillna(0)
        custo_retaguarda = df['CUSTO_RETARGUARDA'].fillna(0)
        
        cond1 = (df['Preço tabela'].isna()) | (preco_tabela == 0)
        cond2 = custo_retaguarda > 0
        
        df['Preço de Custo'] = np.where(
            cond1,
            custo_retaguarda,
            np.where(cond2, np.maximum(preco_tabela, custo_retaguarda), preco_tabela)
        )
        
        # Conta quantos SKUs tiveram custo maior
        custo_maior = (custo_retaguarda > preco_tabela) & (custo_retaguarda > 0)
        stats['skus_custo_maior'] += int(custo_maior.sum())
        
        # Remove coluna temporária
        df = df.drop(columns=['CUSTO_RETARGUARDA'], errors='ignore')
        
        # Merge com estoque de segurança
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
        
        # Cálculos financeiros
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
    logger.info("=" * 60)
    
    return dados_marcas, data_atualizacao, stats


# ==========================================
# FUNÇÕES DE UI
# ==========================================

def exibir_kpi_card(col, icone, titulo, valor_fmt, delta_texto=None, cor_delta="#ff4b4b"):
    """Exibe card de KPI customizado."""
    delta_html = f'<div style="font-size:12px;color:{cor_delta};margin-top:4px;">{delta_texto}</div>' if delta_texto else ''
    col.markdown(f"""
    <div style="
        background: linear-gradient(135deg,#111820,#0d1f14);
        border:1px solid #1a3d25; border-left:4px solid #007A33;
        border-radius:10px; padding:18px 20px;
        box-shadow:0 4px 16px rgba(0,122,51,0.15);
        min-height:110px;
    ">
        <div style="font-size:22px;margin-bottom:4px;">{icone}</div>
        <div style="font-size:12px;color:#8da9be;margin-bottom:6px;">{titulo}</div>
        <div style="font-size:26px;font-weight:700;color:#D4AF37;
                    text-shadow:0 0 10px rgba(212,175,55,0.3);">{valor_fmt}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)


def exibir_titulo_marca(nome_marca, tamanho_logo=30):
    """Exibe título de marca com logo e faixa colorida."""
    cor = CORES_MARCAS.get(nome_marca, '#007A33')
    col_logo, col_nome = st.columns([0.08, 0.92])
    with col_logo:
        logo_path = LOGOS_MARCAS.get(nome_marca, '')
        if logo_path:
            try:
                st.image(logo_path, width=tamanho_logo)
            except Exception:
                st.write("🏷️")
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
    """Retorna horário atual em Brasília (UTC-3)."""
    fuso_brasilia = timezone(timedelta(hours=-3))
    return datetime.now(fuso_brasilia).strftime("%d/%m/%Y às %H:%M:%S")


# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================

def main():
    """Função principal do dashboard."""
    
    # Carrega dados
    with st.spinner("🔄 Carregando dados das planilhas..."):
        dados_marcas, data_atualizacao, stats = carregar_todos_os_dados()
    
    horario_carregamento = obter_horario_brasilia()
    
    # Verifica se houve erro no carregamento
    if not dados_marcas:
        st.error("❌ Nenhum dado foi carregado. Verifique os logs acima.")
        st.info("""
        ### 🔧 Passos para resolver:
        
        1. **Verifique se as planilhas estão públicas:**
           - Abra cada planilha no Google Sheets
           - Clique em "Compartilhar"
           - Selecione "Qualquer pessoa com o link"
           - Permissão: "Leitor"
        
        2. **Verifique os IDs no config.py**
        
        3. **Verifique os nomes das abas:**
           - DRAFT_PDVS: BOTICARIO, EUDORA, QUEM_DISSE_BERENICE
           - CONSULTA_DE_ESTOQUE: BOT, EUD, QDB
        """)
        st.stop()
    
    # Determina timestamp
    if data_atualizacao:
        horario_exibicao = data_atualizacao
        info_timestamp = "🕒 Última atualização"
    else:
        horario_exibicao = horario_carregamento
        info_timestamp = "🕒 Carregado em"
    
    # ==========================================
    # CABEÇALHO
    # ==========================================
    col_logo, col_info = st.columns([1, 3])
    
    with col_logo:
        try:
            st.image("logo_cp_fani.png", width=180)
        except Exception:
            pass
    
    with col_info:
        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, #0d1f14 0%, #080b0f 100%);
            border: 1px solid #1a3d25; border-left: 5px solid #007A33;
            border-radius: 12px; padding: 20px 28px; margin-bottom: 8px;
        ">
            <div style="font-size:13px; color:#8da9be; margin-bottom:4px; letter-spacing:1px; text-transform:uppercase;">
                Grupo NSF · CP Fani
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
    
    # ==========================================
    # DIAGNÓSTICO DE CUSTOS
    # ==========================================
    if stats.get('total_skus_retaguarda', 0) > 0:
        with st.expander("🔍 Diagnóstico de Custos (clique para ver)", expanded=False):
            col_d1, col_d2, col_d3, col_d4 = st.columns(4)
            col_d1.metric("SKUs na Retaguarda", f"{stats['total_skus_retaguarda']:,}")
            col_d2.metric("SKUs com custo", f"{stats['skus_matcheados']:,}")
            col_d3.metric("SKUs sem custo", f"{stats['skus_sem_match']:,}")
            col_d4.metric("Custo > Tabela", f"{stats['skus_custo_maior']:,}")
    else:
        st.warning("⚠️ Planilha Retaguarda não carregada. Usando Preço de Tabela como custo.")
    
    st.markdown("---")
    
    # ==========================================
    # SIDEBAR - FILTROS
    # ==========================================
    st.sidebar.title("Filtros de Visualização")
    
    # Coleta todos os PDVs disponíveis
    todos_pdvs = sorted(set(
        int(pdv)
        for df in dados_marcas.values()
        for pdv in df['PDV'].dropna()
    ))
    
    # Opções com "Todas as Lojas"
    opcoes_selectbox = ["Todas as Lojas"] + [DE_PARA_LOJAS.get(pdv, f"PDV {pdv}") for pdv in todos_pdvs]
    
    loja_selecionada_nome = st.sidebar.selectbox("Selecione a Loja / PDV:", opcoes_selectbox)
    
    # Determina PDV selecionado
    if loja_selecionada_nome == "Todas as Lojas":
        pdv_selecionado = 'TODAS'
    else:
        pdv_selecionado = DE_PARA_LOJAS_REVERSO.get(loja_selecionada_nome)
        if pdv_selecionado is None:
            st.error("PDV não reconhecido.")
            st.stop()
    
    st.sidebar.markdown("---")
    
    # Filtro de marca
    st.sidebar.subheader("Filtro de Marca")
    opcoes_marca = ["Todas as Marcas"] + list(dados_marcas.keys())
    marca_selecionada = st.sidebar.selectbox("Selecione a Marca:", opcoes_marca)
    
    # Logos na sidebar
    st.sidebar.markdown("**Marcas disponíveis:**")
    col_logos_sidebar = st.sidebar.columns(len(dados_marcas))
    for idx, (nome_marca, df_marca) in enumerate(dados_marcas.items()):
        with col_logos_sidebar[idx]:
            logo_path = LOGOS_MARCAS.get(nome_marca, '')
            if logo_path:
                try:
                    st.image(logo_path, width=40)
                except Exception:
                    pass
            st.caption(nome_marca)
    
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Forçar Atualização"):
        st.cache_data.clear()
        st.rerun()
    
    # Badge de PDV selecionado
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
    
    # Filtra dados
    if marca_selecionada == "Todas as Marcas":
        dados_filtrados = dados_marcas
        titulo_secao = "Consolidado Geral"
    else:
        dados_filtrados = {marca_selecionada: dados_marcas[marca_selecionada]}
        titulo_secao = f"Análise: {marca_selecionada}"
    
    st.markdown(f"### {titulo_secao}")
    st.markdown("---")
    
    # ==========================================
    # KPIs
    # ==========================================
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
            qtd_itens_total += df_loja['Estoque Atual'].sum()
    
    col1, col2, col3, col4 = st.columns(4)
    
    exibir_kpi_card(col1, "💰", "Valor em Estoque", f"R$ {v_estoque_atual_total:,.2f}")
    exibir_kpi_card(col2, "📉", "Estoque Mínimo", f"R$ {v_estoque_min_total:,.2f}")
    
    pct_excesso = f"{((v_excesso_total_total/v_estoque_atual_total)*100 if v_estoque_atual_total > 0 else 0):.1f}%"
    exibir_kpi_card(col3, "⚠️", "Capital Preso", f"R$ {v_excesso_total_total:,.2f}", 
                   delta_texto=f"{pct_excesso} do estoque", cor_delta="#f59e0b")
    
    exibir_kpi_card(col4, "🚨", "Risco de Ruptura", f"R$ {v_falta_total_total:,.2f}", 
                   delta_texto="Abaixo do Mínimo", cor_delta="#ef4444")
    
    st.markdown("---")
    
    st.success("✅ Dashboard carregado com sucesso!")
    st.info("📊 Use os filtros laterais para selecionar loja e marca.")


if __name__ == "__main__":
    main()