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

from config import (
    PLANILHAS, DE_PARA_LOJAS, DE_PARA_LOJAS_REVERSO, 
    MAPEAMENTO_PDV_DRAFT_RAW, NOMES_MARCAS, ABAS_SEGURANCA,
    LOGOS_MARCAS, CORES_MARCAS, REGRAS_ESTOQUE_MINIMO,
    COLUNAS_OBRIGATORIAS, TIMEOUT_DOWNLOAD, CACHE_TTL
)

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
</style>
""", unsafe_allow_html=True)

def normalizar_sku(valor):
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
    if pd.isna(nome) or nome is None:
        return ""
    s = str(nome).strip().lower()
    s = re.sub(r'[^a-z0-9\s]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


MAPEAMENTO_PDV_DRAFT_NORMALIZADO = {
    normalizar_nome_loja(k): v for k, v in MAPEAMENTO_PDV_DRAFT_RAW.items()
}


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
        st.error(f"Falha ao carregar {nome_planilha}: {str(e)[:200]}")
        return None


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_planilha_retaguarda(custos_buffer):
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
        
        colunas_sku = ['SKU', 'CÓDIGO', 'CODIGO', 'CÓDIGO SKU', 'CODIGO SKU', 
                       'EAN', 'COD PRODUTO', 'CÓDIGO PRODUTO']
        coluna_sku = next((c for c in colunas_sku if c in df.columns), None)
        
        colunas_custo = ['CUSTO', 'PREÇO DE CUSTO', 'PRECO DE CUSTO', 'CUSTO UNITÁRIO',
                        'CUSTO UNITARIO', 'VALOR CUSTO', 'CUSTO (R$)', 'CUSTO R$']
        coluna_custo = next((c for c in colunas_custo if c in df.columns), None)
        
        if coluna_custo is None and len(df.columns) > 9:
            coluna_custo = df.columns[9]
            logger.info(f"Usando coluna J como custo: {coluna_custo}")
        
        if coluna_sku is None or coluna_custo is None:
            logger.error(f"Retaguarda: colunas faltantes. SKU={coluna_sku}, Custo={coluna_custo}")
            return pd.DataFrame()
        
        df_resultado = pd.DataFrame()
        df_resultado['SKU'] = df[coluna_sku].apply(normalizar_sku)
        df_resultado['CUSTO_RETARGUARDA'] = pd.to_numeric(df[coluna_custo], errors='coerce').fillna(0)
        
        colunas_loja = ['LOJA', 'PDV', 'LOJA/PDV', 'LOJA - PDV', 'CÓDIGO LOJA', 'CODIGO LOJA',
                       'FILIAL', 'COD FILIAL', 'LOJA NOME', 'NOME LOJA']
        coluna_loja = next((c for c in colunas_loja if c in df.columns), None)
        
        if coluna_loja is None:
            logger.info("Retaguarda: sem coluna de loja. Aplicando custos para todos os PDVs.")
            todos_pdvs = list(DE_PARA_LOJAS.keys())
            df_expandido = []
            for pdv in todos_pdvs:
                df_temp = df_resultado.copy()
                df_temp['PDV'] = pdv
                df_expandido.append(df_temp)
            df_resultado = pd.concat(df_expandido, ignore_index=True)
        else:
            df_resultado['LOJA_NOME'] = df[coluna_loja].astype(str)
            df_resultado['LOJA_NORM'] = df_resultado['LOJA_NOME'].apply(normalizar_nome_loja)
            df_resultado['PDV'] = df_resultado['LOJA_NORM'].map(MAPEAMENTO_PDV_DRAFT_NORMALIZADO)
            
            total_antes = len(df_resultado)
            df_resultado = df_resultado[df_resultado['PDV'].notna()].copy()
            total_depois = len(df_resultado)
            
            if total_depois < total_antes:
                logger.warning(f"Retaguarda: {total_antes - total_depois} linhas sem PDV válido")
        
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


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def carregar_todos_os_dados():
    logger.info("=" * 60)
    logger.info("INICIANDO CARREGAMENTO DE DADOS")
    logger.info("=" * 60)
    
    stats = {
        'total_skus_retaguarda': 0,
        'skus_matcheados': 0,
        'skus_sem_match': 0,
        'skus_custo_maior': 0
    }
    
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
    
    logger.info("Fase 2: Processamento dos dados...")
    
    dados_marcas, data_atualizacao = carregar_dados_principais(buffer_draft)
    if not dados_marcas:
        logger.error("Nenhum dado foi carregado da planilha principal")
        return {}, None, stats
    
    df_estoque_seguranca = carregar_estoque_seguranca(buffer_seguranca)
    df_retaguarda = carregar_planilha_retaguarda(buffer_retaguarda)
    stats['total_skus_retaguarda'] = len(df_retaguarda) if not df_retaguarda.empty else 0
    
    logger.info("Fase 3: Aplicando custos da Retaguarda...")
    
    for nome_marca, df in dados_marcas.items():
        df['CUSTO_RETARGUARDA'] = 0.0
        
        if not df_retaguarda.empty:
            df_update = df_retaguarda[['PDV', 'SKU', 'CUSTO_RETARGUARDA']].copy()
            df_update = df_update.set_index(['PDV', 'SKU'])
            
            idx = pd.MultiIndex.from_arrays([df['PDV'], df['SKU']])
            mask = idx.isin(df_update.index)
            
            df.loc[mask, 'CUSTO_RETARGUARDA'] = df_update.loc[idx[mask], 'CUSTO_RETARGUARDA'].values
            
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
    logger.info("=" * 60)
    
    return dados_marcas, data_atualizacao, stats


def exibir_kpi_card(col, icone, titulo, valor_fmt, delta_texto=None, cor_delta="#ff4b4b"):
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
    COR_AMARELO = colors.HexColor('#b45309')
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
    except Exception:
        pass

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

        df_exc = df_loja[(df_loja['Valor_Excesso'] > 0) & (df_loja['Estoque_Minimo_Qtd'] > 0)].sort_values('Valor_Excesso', ascending=False).head(20)
        if not df_exc.empty:
            elementos.append(Paragraph(f"Excessos Críticos — {nome_marca}", estilo_subtitulo))
            colunas_exc = ['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Excesso', 'Valor_Excesso']
            colunas_exc_existentes = [c for c in colunas_exc if c in df_exc.columns]
            dados_exc = [colunas_exc_existentes]
            for _, row in df_exc[colunas_exc_existentes].iterrows():
                linha = []
                for col in colunas_exc_existentes:
                    v = row[col]
                    if col == 'Valor_Excesso':
                        linha.append(f"R$ {float(v):,.2f}")
                    else:
                        linha.append(str(v))
                dados_exc.append(linha)
            t_exc = Table(dados_exc, repeatRows=1)
            t_exc.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), COR_VERDE),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8faf9')),
                ('TEXTCOLOR', (0, 1), (-1, -1), COR_TEXTO),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f8faf9'), colors.white]),
                ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#1a3d25')),
                ('PADDING', (0, 0), (-1, -1), 4),
                ('TEXTCOLOR', (-1, 1), (-1, -1), COR_AMARELO),
            ]))
            elementos.append(t_exc)
            elementos.append(Spacer(1, 0.4*cm))

        df_flt = df_loja[df_loja['Valor_Falta'] > 0].sort_values('Valor_Falta', ascending=False).head(20)
        if not df_flt.empty:
            elementos.append(Paragraph(f"Produtos em Falta — {nome_marca}", estilo_subtitulo))
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
        return None


def main():
    with st.spinner("🔄 Carregando dados das planilhas..."):
        dados_marcas, data_atualizacao, stats = carregar_todos_os_dados()
    
    horario_carregamento = obter_horario_brasilia()
    
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
    
    if data_atualizacao:
        horario_exibicao = data_atualizacao
        info_timestamp = "🕒 Última atualização"
    else:
        horario_exibicao = horario_carregamento
        info_timestamp = "🕒 Carregado em"
    
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
    
    st.sidebar.title("Filtros de Visualização")
    
    todos_pdvs = sorted(set(
        int(pdv)
        for df in dados_marcas.values()
        for pdv in df['PDV'].dropna()
    ))
    
    opcoes_selectbox = ["Todas as Lojas"] + [DE_PARA_LOJAS.get(pdv, f"PDV {pdv}") for pdv in todos_pdvs]
    
    loja_selecionada_nome = st.sidebar.selectbox("Selecione a Loja / PDV:", opcoes_selectbox)
    
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
    marca_selecionada = st.sidebar.selectbox("Selecione a Marca:", opcoes_marca)
    
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
    
    # ==========================================
    # GRÁFICO COMPARATIVO POR MARCA
    # ==========================================
    if marca_selecionada == "Todas as Marcas":
        st.subheader("📊 Comparativo entre Marcas")

        dados_grafico = []
        for nome_marca, df_completo in dados_marcas.items():
            if pdv_selecionado == 'TODAS':
                df_loja = df_completo
            else:
                df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
            
            if not df_loja.empty:
                dados_grafico.append({
                    'Marca': nome_marca,
                    'Qtd Itens': int(df_loja['Estoque Atual'].sum()),
                    'Custo Total': df_loja['Valor_Custo_Estoque_Atual'].sum()
                })

        if dados_grafico:
            df_grafico_marcas = pd.DataFrame(dados_grafico)
            col_graf1, col_graf2 = st.columns(2)

            with col_graf1:
                fig_qtd = go.Figure()
                fig_qtd.add_trace(go.Bar(
                    x=df_grafico_marcas['Marca'], y=df_grafico_marcas['Qtd Itens'],
                    name='Quantidade de Itens',
                    marker_color=[CORES_MARCAS.get(m, '#007A33') for m in df_grafico_marcas['Marca']],
                    text=[f"{v:,.0f}" for v in df_grafico_marcas['Qtd Itens']], textposition='auto'
                ))
                fig_qtd.update_layout(template='plotly_dark', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                                      height=400, title='Quantidade de Itens por Marca', yaxis_title='Qtd de Unidades')
                st.plotly_chart(fig_qtd, use_container_width=True)

            with col_graf2:
                fig_custo = go.Figure()
                fig_custo.add_trace(go.Bar(
                    x=df_grafico_marcas['Marca'], y=df_grafico_marcas['Custo Total'],
                    name='Custo Total',
                    marker_color=[CORES_MARCAS.get(m, '#007A33') for m in df_grafico_marcas['Marca']],
                    text=[f"R$ {v:,.0f}" for v in df_grafico_marcas['Custo Total']], textposition='auto'
                ))
                fig_custo.update_layout(template='plotly_dark', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                                        height=400, title='Custo Total por Marca', yaxis_title='Valor (R$)')
                st.plotly_chart(fig_custo, use_container_width=True)

    # ==========================================
    # CUSTO POR CURVA
    # ==========================================
    st.markdown("---")
    st.subheader("📊 Custo Total por Curva de Produto")

    df_curva_consolidado = pd.DataFrame()
    for nome_marca, df_completo in dados_filtrados.items():
        if pdv_selecionado == 'TODAS':
            df_loja = df_completo
        else:
            df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
        
        if not df_loja.empty and 'Classe' in df_loja.columns:
            df_agrupado = df_loja.groupby('Classe').agg({'Valor_Custo_Estoque_Atual': 'sum', 'SKU': 'count'}).reset_index()
            df_agrupado.columns = ['Curva', 'Custo Total', 'Qtd SKUs']
            df_agrupado['Marca'] = nome_marca
            df_curva_consolidado = pd.concat([df_curva_consolidado, df_agrupado], ignore_index=True)

    if not df_curva_consolidado.empty:
        if marca_selecionada == "Todas as Marcas":
            df_pivot = df_curva_consolidado.pivot_table(values='Custo Total', index='Curva', columns='Marca', aggfunc='sum', fill_value=0).reset_index()
            colunas_marcas = [col for col in df_pivot.columns if col != 'Curva']
            df_pivot['Total Geral'] = df_pivot[colunas_marcas].sum(axis=1)
            linha_total = {'Curva': 'TOTAL'}
            for col in colunas_marcas:
                linha_total[col] = df_pivot[col].sum()
            linha_total['Total Geral'] = df_pivot['Total Geral'].sum()
            df_pivot = pd.concat([df_pivot, pd.DataFrame([linha_total])], ignore_index=True)
            for col in colunas_marcas + ['Total Geral']:
                df_pivot[col] = df_pivot[col].apply(lambda x: f"R$ {x:,.2f}")
            st.dataframe(df_pivot, use_container_width=True, hide_index=True)
        else:
            df_exibicao = df_curva_consolidado[['Curva', 'Custo Total', 'Qtd SKUs']].copy().sort_values('Curva')
            total_custo = df_exibicao['Custo Total'].sum()
            total_skus = df_exibicao['Qtd SKUs'].sum()
            df_total = pd.DataFrame([{'Curva': 'TOTAL', 'Custo Total': total_custo, 'Qtd SKUs': total_skus}])
            df_exibicao = pd.concat([df_exibicao, df_total], ignore_index=True)
            df_exibicao['Custo Total'] = df_exibicao['Custo Total'].apply(lambda x: f"R$ {x:,.2f}")
            st.dataframe(df_exibicao, use_container_width=True, hide_index=True)

        fig_custo = go.Figure()
        if marca_selecionada == "Todas as Marcas":
            for nome_marca in dados_filtrados.keys():
                df_mc = df_curva_consolidado[df_curva_consolidado['Marca'] == nome_marca]
                if not df_mc.empty:
                    fig_custo.add_trace(go.Bar(x=df_mc['Curva'], y=df_mc['Custo Total'], name=nome_marca,
                                               marker_color=CORES_MARCAS.get(nome_marca, '#007A33')))
            fig_custo.update_layout(barmode='stack')
        else:
            fig_custo.add_trace(go.Bar(x=df_curva_consolidado['Curva'], y=df_curva_consolidado['Custo Total'],
                                       marker_color=[CORES_MARCAS.get(marca_selecionada, '#007A33')] * len(df_curva_consolidado),
                                       text=[f"R$ {v:,.2f}" for v in df_curva_consolidado['Custo Total']], textposition='auto'))
        fig_custo.update_layout(template='plotly_dark', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                                height=400, xaxis_title='Curva', yaxis_title='Custo Total (R$)', title='Distribuição de Custo por Curva')
        st.plotly_chart(fig_custo, use_container_width=True)

    # ==========================================
    # ANÁLISE POR CATEGORIA
    # ==========================================
    st.markdown("---")
    st.subheader("📊 Análise por Categoria")

    df_categoria_consolidado = pd.DataFrame()
    for nome_marca, df_completo in dados_filtrados.items():
        if pdv_selecionado == 'TODAS':
            df_loja = df_completo
        else:
            df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
        
        if not df_loja.empty and 'Categoria' in df_loja.columns:
            df_cat = df_loja.groupby('Categoria')[['Valor_Estoque_Atual', 'Valor_Estoque_Minimo']].sum().reset_index()
            df_cat['Marca'] = nome_marca
            df_categoria_consolidado = pd.concat([df_categoria_consolidado, df_cat], ignore_index=True)

    if not df_categoria_consolidado.empty:
        fig_cat = go.Figure()
        if marca_selecionada == "Todas as Marcas":
            for nome_marca in dados_filtrados.keys():
                df_mc = df_categoria_consolidado[df_categoria_consolidado['Marca'] == nome_marca]
                if not df_mc.empty:
                    fig_cat.add_trace(go.Bar(x=df_mc['Categoria'], y=df_mc['Valor_Estoque_Atual'], name=nome_marca,
                                             marker_color=CORES_MARCAS.get(nome_marca, '#007A33')))
            fig_cat.update_layout(barmode='group')
        else:
            fig_cat.add_trace(go.Bar(x=df_categoria_consolidado['Categoria'], y=df_categoria_consolidado['Valor_Estoque_Atual'],
                                     name='Estoque Atual', marker_color=CORES_MARCAS.get(marca_selecionada, '#007A33')))
            fig_cat.add_trace(go.Bar(x=df_categoria_consolidado['Categoria'], y=df_categoria_consolidado['Valor_Estoque_Minimo'],
                                     name='Estoque Mínimo', marker_color='#ff4b4b'))
            fig_cat.update_layout(barmode='group')
        fig_cat.update_layout(template='plotly_dark', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                              height=400, xaxis_title='Categoria', yaxis_title='Valor (R$)', title='Estoque por Categoria')
        st.plotly_chart(fig_cat, use_container_width=True)

    # ==========================================
    # TABELAS DE EXCESSOS E FALTAS (COM PREÇO CONSIDERADO)
    # ==========================================
    st.markdown("---")

    for nome_marca, df_completo in dados_filtrados.items():
        if pdv_selecionado == 'TODAS':
            df_loja = df_completo
        else:
            df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
        
        if df_loja.empty:
            continue

        exibir_titulo_marca(nome_marca, tamanho_logo=35)

        col_tab1, col_tab2 = st.columns(2)

        with col_tab1:
            st.write("### 🛑 Excessos Críticos")
            # Seleciona colunas e renomeia 'Preço de Custo' para 'Preço Considerado'
            df_excesso_tabela = df_loja[
                (df_loja['Valor_Excesso'] > 0) & (df_loja['Estoque_Minimo_Qtd'] > 0)
            ][['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Excesso', 'Preço de Custo', 'Valor_Excesso']
            ].rename(columns={'Preço de Custo': 'Preço Considerado'}).sort_values(by='Valor_Excesso', ascending=False)

            total_excesso_qtd = int(df_excesso_tabela['Qtd_Excesso'].sum()) if not df_excesso_tabela.empty else 0
            total_excesso_val = df_excesso_tabela['Valor_Excesso'].sum() if not df_excesso_tabela.empty else 0
            skus_excesso = len(df_excesso_tabela)
            col_e1, col_e2, col_e3 = st.columns(3)
            col_e1.metric("SKUs com excesso", skus_excesso)
            col_e2.metric("Unidades em excesso", f"{total_excesso_qtd:,}")
            col_e3.metric("Valor total em excesso", f"R$ {total_excesso_val:,.2f}")

            # Formata apenas a coluna 'Preço Considerado' e 'Valor_Excesso'
            st.dataframe(df_excesso_tabela.style.format({'Preço Considerado': 'R$ {:.2f}', 'Valor_Excesso': 'R$ {:.2f}'}),
                         use_container_width=True, height=280)

        with col_tab2:
            st.write("### 🚨 Produtos Críticos em Falta / Ruptura")
            # Seleciona colunas e renomeia 'Preço de Custo' para 'Preço Considerado'
            df_falta_tabela = df_loja[df_loja['Valor_Falta'] > 0][
                ['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Falta', 'Preço de Custo', 'Valor_Falta']
            ].rename(columns={'Preço de Custo': 'Preço Considerado'}).sort_values(by='Valor_Falta', ascending=False)

            total_falta_qtd = int(df_falta_tabela['Qtd_Falta'].sum()) if not df_falta_tabela.empty else 0
            total_falta_val = df_falta_tabela['Valor_Falta'].sum() if not df_falta_tabela.empty else 0
            skus_falta = len(df_falta_tabela)
            col_f1, col_f2, col_f3 = st.columns(3)
            col_f1.metric("SKUs em falta", skus_falta)
            col_f2.metric("Unidades em falta", f"{total_falta_qtd:,}")
            col_f3.metric("Risco financeiro", f"R$ {total_falta_val:,.2f}")

            # Formata apenas a coluna 'Preço Considerado' e 'Valor_Falta'
            st.dataframe(df_falta_tabela.style.format({'Preço Considerado': 'R$ {:.2f}', 'Valor_Falta': 'R$ {:.2f}'}),
                         use_container_width=True, height=280)

        st.markdown("---")

    # ==========================================
    # EXPORTAR PDF
    # ==========================================
    st.markdown("---")
    pdf_buffer = gerar_pdf_dashboard(
        dados_filtrados=dados_filtrados,
        pdv_selecionado=pdv_selecionado,
        loja_selecionada_nome=loja_selecionada_nome,
        marca_selecionada=marca_selecionada,
        horario_brasilia=obter_horario_brasilia(),
        v_estoque_atual_total=v_estoque_atual_total,
        v_estoque_min_total=v_estoque_min_total,
        v_excesso_total_total=v_excesso_total_total,
        v_falta_total_total=v_falta_total_total,
        qtd_itens_total=qtd_itens_total
    )

    if pdf_buffer is not None:
        st.download_button(
            label="📥 Exportar Relatório em PDF",
            data=pdf_buffer,
            file_name=f"relatorio_estoque_{pdv_selecionado if pdv_selecionado != 'TODAS' else 'TODAS_LOJAS'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
            type="primary"
        )


if __name__ == "__main__":
    main()