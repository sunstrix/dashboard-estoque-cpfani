import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timezone, timedelta
import openpyxl
from io import BytesIO
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import time

# 1. Configuração Inicial da Página (Visual Tema O Boticário)
st.set_page_config(
    page_title="Painel de Performance de Estoque NSF - CP Fani",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS para tema O Boticário (Verde Escuro + Dourado)
st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    
    /* Cores dos KPIs - Tema Boticário */
    div[data-testid="stMetricValue"] { font-size: 28px; color: #D4AF37; }
    div[data-testid="stMetricLabel"] { font-size: 14px; color: #a3b8cc; }
    
    /* Abas com destaque dourado */
    .stTabs [data-baseweb="tab"] { color: #a3b8cc; font-size: 16px; }
    .stTabs [data-baseweb="tab"][aria-selected="true"] { color: #D4AF37; font-weight: bold; }
    
    /* Botão de atualização com cor Boticário */
    div.stButton > button {
        background-color: #007A33;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 0.5rem 1rem;
    }
    div.stButton > button:hover {
        background-color: #006838;
        color: #D4AF37;
    }
    
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #0a0d12;
        border-right: 2px solid #007A33;
    }    
    /* Títulos */
    h1, h2, h3 { color: #D4AF37; }
    
    /* Barra de progresso e elementos */
    .stProgress > div > div > div > div {
        background-color: #007A33;
    }
    </style>
""", unsafe_allow_html=True)

# IDs OFICIAIS DAS PLANILHAS
SPREADSHEET_ID_PRINCIPAL = "1EDDyKie9UiugMLMowcPzHfViqzziFcSgxVPvZ2Rx3L0"
SPREADSHEET_ID_SEGURANCA = "1uHonFnFM4p7bz4s7YpewhKHNs6fSEfw9rDMTKC7jtHE"
SPREADSHEET_ID_DRAFT = "11Z21gFvJ9pm2xSlF3IweC7xcYZwAZWrjcWDnRe5LexY"

# URLs de exportação direta em formato Excel
URL_EXCEL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID_PRINCIPAL}/export?format=xlsx"
URL_ESTOQUE_SEGURANCA = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID_SEGURANCA}/export?format=xlsx"
URL_DRAFT = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID_DRAFT}/export?format=xlsx"

# 2. Dicionário com os 17 PDVs reais (Nomes atualizados)
DE_PARA_LOJAS = {
    4842: "4842 - Metrópole",
    5152: "5152 - Coração",
    6105: "6105 - Assai Anchieta",
    6106: "6106 - Direita",
    6110: "6110 - Arouche",
    8001: "8001 - Dom José",
    11576: "11576 - Davó",
    12055: "12055 - São Bento",
    12056: "12056 - Marechal",
    12605: "12605 - Coop",
    12645: "12645 - Light",
    14120: "14120 - VD SBC",
    14353: "14353 - VD SP",
    20371: "20371 - Luz",
    21502: "21502 - Bem Barato",
    23000: "23000 - Outlet",
    23379: "23379 - Assai Piraporinha"
}

# Mapeamento dos nomes completos da planilha DRAFT para os códigos de PDV
MAPEAMENTO_PDV_DRAFT = {
    'Loja: 4842 - N. S. F. COSMETICOS E PRESENTES LTDA': 4842,
    'Loja: 5152 - N. S. F. COSMETICOS E PRESENTES LTDA': 5152,
    'Loja: 6105 - N. S. F. COSMETICOS E PRESENTES LTDA': 6105,
    'Loja: 6106 - N. S. F. COSMETICOS E PRESENTES LTDA': 6106,
    'Loja: 6110 - N. S. F. COSMETICOS E PRESENTES LTDA': 6110,
    'Loja: 8001 - N. S. F. COSMETICOS E PRESENTES LTDA': 8001,    'Loja: 11576 - N. S. F. COSMETICOS E PRESENTES LTDA': 11576,
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

# Mapeamento reverso (código -> nome draft) para facilitar busca
MAPEAMENTO_PDV_DRAFT_REVERSO = {v: k for k, v in MAPEAMENTO_PDV_DRAFT.items()}

# Nomes limpos das marcas (sem emojis) - usados como chaves internas
NOMES_MARCAS = {
    'BOTICARIO': 'O Boticário',
    'EUDORA': 'Eudora',
    'QUEM_DISSE_BERENICE': 'Quem Disse, Berenice?'
}

# Mapeamento das abas da planilha de segurança para as marcas
ABAS_SEGURANCA = {
    'BOT': 'O Boticário',
    'EUD': 'Eudora',
    'QDB': 'Quem Disse, Berenice?'
}

# Logos das marcas (arquivos PNG no repositório)
LOGOS_MARCAS = {
    'O Boticário': 'logo_boticario.png',
    'Eudora': 'logo_eudora.png',
    'Quem Disse, Berenice?': 'logo_qdb.png'
}

# Cores das marcas (ajustadas para harmonizar com tema Boticário)
CORES_MARCAS = {
    'O Boticário': '#007A33',
    'Eudora': '#a855f7',
    'Quem Disse, Berenice?': '#ff4b4b'
}

def criar_sessao_com_retry():
    """
    Cria uma sessão requests com retry automático para lidar com falhas de rede.
    """
    session = requests.Session()
    retry = Retry(        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

def download_arquivo_excel_com_retry(url, descricao="arquivo", timeout=60):
    """
    Faz download de arquivo Excel com retry automático e tratamento de erros.
    Retorna BytesIO com o conteúdo ou None em caso de falha.
    """
    session = criar_sessao_com_retry()
    
    try:
        response = session.get(url, timeout=timeout, stream=True)
        response.raise_for_status()
        
        # Verifica se o download foi completo
        if 'content-length' in response.headers:
            expected_size = int(response.headers['content-length'])
            actual_size = len(response.content)
            
            if actual_size < expected_size:
                time.sleep(2)
                # Tenta novamente
                response = session.get(url, timeout=timeout, stream=True)
        
        if response.status_code == 200 and len(response.content) > 0:
            return BytesIO(response.content)
        else:
            return None
            
    except Exception:
        return None

def exibir_titulo_marca(nome_marca, tamanho_logo=30):
    """
    Exibe o título de uma marca com sua logo usando st.columns + st.image.
    Esta é a forma correta de exibir imagens no Streamlit.
    """
    col_logo, col_nome = st.columns([0.1, 0.9])
    with col_logo:
        logo_path = LOGOS_MARCAS.get(nome_marca, '')
        if logo_path:
            try:
                st.image(logo_path, width=tamanho_logo)            except Exception:
                st.write("🏷️")
        else:
            st.write("🏷️")
    with col_nome:
        st.markdown(f"### {nome_marca}")

def obter_horario_brasilia():
    """
    Retorna o horário atual no fuso horário de Brasília (UTC-3).
    Funciona corretamente mesmo em servidores UTC (como Streamlit Cloud).
    """
    fuso_brasilia = timezone(timedelta(hours=-3))
    agora_brasilia = datetime.now(fuso_brasilia)
    return agora_brasilia.strftime("%d/%m/%Y às %H:%M:%S")

def carregar_planilha_draft(url):
    """
    Carrega a planilha DRAFT e retorna um DataFrame com:
    - PDV (código numérico)
    - SKU
    - CUSTO (coluna J ou similar)
    
    Mapeia os nomes completos de loja para os códigos de PDV.
    """
    excel_buffer = download_arquivo_excel_com_retry(url, "planilha draft de custos", timeout=90)
    
    if excel_buffer is None:
        return pd.DataFrame()
    
    try:
        excel_file = pd.ExcelFile(excel_buffer)
        
        # Usa a primeira aba disponível
        if not excel_file.sheet_names:
            return pd.DataFrame()
        
        df_draft = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
        
        if df_draft.empty:
            return pd.DataFrame()
        
        # Normaliza nomes das colunas
        df_draft.columns = [str(col).strip().upper() for col in df_draft.columns]
        
        # Identifica as colunas necessárias
        # Coluna de Loja/PDV (pode ter vários nomes)
        colunas_loja_possiveis = ['LOJA', 'PDV', 'LOJA/PDV', 'LOJA - PDV', 'CÓDIGO LOJA', 'CODIGO LOJA']
        coluna_loja = None
        for col in colunas_loja_possiveis:            if col in df_draft.columns:
                coluna_loja = col
                break
        
        # Coluna de SKU
        colunas_sku_possiveis = ['SKU', 'CÓDIGO', 'CODIGO', 'CÓDIGO SKU', 'CODIGO SKU', 'CÓD. SKU']
        coluna_sku = None
        for col in colunas_sku_possiveis:
            if col in df_draft.columns:
                coluna_sku = col
                break
        
        # Coluna de Custo (coluna J geralmente)
        colunas_custo_possiveis = ['CUSTO', 'PREÇO DE CUSTO', 'PRECO DE CUSTO', 'CUSTO UNITÁRIO', 
                                   'CUSTO UNITARIO', 'VALOR CUSTO', 'CUSTO (R$)', 'CUSTO R$']
        coluna_custo = None
        for col in colunas_custo_possiveis:
            if col in df_draft.columns:
                coluna_custo = col
                break
        
        # Se não encontrou coluna de custo, tenta pela posição (coluna J = índice 9)
        if coluna_custo is None and len(df_draft.columns) > 9:
            coluna_custo = df_draft.columns[9]  # Coluna J (0-indexed = 9)
        
        # Se ainda não encontrou, usa a 10ª coluna por padrão
        if coluna_custo is None and len(df_draft.columns) >= 10:
            coluna_custo = df_draft.columns[9]
        
        if coluna_loja is None or coluna_sku is None or coluna_custo is None:
            return pd.DataFrame()
        
        # Prepara o DataFrame
        df_resultado = pd.DataFrame()
        df_resultado['LOJA_NOME'] = df_draft[coluna_loja].astype(str).str.strip()
        df_resultado['SKU'] = df_draft[coluna_sku].astype(str).str.strip()
        df_resultado['CUSTO_DRAFT'] = pd.to_numeric(df_draft[coluna_custo], errors='coerce').fillna(0)
        
        # Mapeia nomes de loja para códigos de PDV
        df_resultado['PDV'] = df_resultado['LOJA_NOME'].map(MAPEAMENTO_PDV_DRAFT)
        
        # Remove linhas sem PDV válido
        df_resultado = df_resultado[df_resultado['PDV'].notna()].copy()
        df_resultado['PDV'] = df_resultado['PDV'].astype(int)
        
        # Mantém apenas colunas necessárias
        df_resultado = df_resultado[['PDV', 'SKU', 'CUSTO_DRAFT']].copy()
        
        return df_resultado
            except Exception:
        return pd.DataFrame()

def carregar_estoque_seguranca(url):
    """
    Carrega a planilha de estoque de segurança de TODAS as abas (BOT, EUD, QDB)
    e retorna um DataFrame consolidado com as colunas: PDV, SKU, ESTOQUE_DE_SEGURANCA
    """
    excel_buffer = download_arquivo_excel_com_retry(url, "planilha de estoque de segurança", timeout=90)
    
    if excel_buffer is None:
        return pd.DataFrame()
    
    try:
        excel_file = pd.ExcelFile(excel_buffer)
        
        abas_esperadas = ['BOT', 'EUD', 'QDB']
        abas_disponiveis = [aba.upper() for aba in excel_file.sheet_names]
        abas_encontradas = [aba for aba in abas_esperadas if aba in abas_disponiveis]
        
        if not abas_encontradas:
            return pd.DataFrame()
        
        dfs_abas = []
        
        for aba_nome in abas_encontradas:
            try:
                aba_exata = [nome for nome in excel_file.sheet_names if nome.upper() == aba_nome][0]
                df_abas = pd.read_excel(excel_file, sheet_name=aba_exata)
                
                if df_abas.empty:
                    continue
                
                df_abas.columns = [col.strip().upper() for col in df_abas.columns]
                
                colunas_necessarias = ['PDV', 'SKU']
                colunas_faltantes = [col for col in colunas_necessarias if col not in df_abas.columns]
                
                if colunas_faltantes:
                    continue
                
                colunas_possiveis = ['ESTOQUE DE SEGURANCA', 'ESTOQUE_DE_SEGURANCA', 'ESTOQUE_SEGURANCA', 
                                   'ESTOQUE MINIMO', 'ESTOQUE_MINIMO', 'MINIMO', 'SEGURANCA', 'QTD_MINIMA']
                coluna_seguranca = None
                
                for col in colunas_possiveis:
                    if col in df_abas.columns:
                        coluna_seguranca = col
                        break
                                if coluna_seguranca is None:
                    df_abas['ESTOQUE_DE_SEGURANCA'] = 0
                else:
                    df_abas = df_abas.rename(columns={coluna_seguranca: 'ESTOQUE_DE_SEGURANCA'})
                    df_abas['ESTOQUE_DE_SEGURANCA'] = pd.to_numeric(df_abas['ESTOQUE_DE_SEGURANCA'], errors='coerce').fillna(0)
                
                df_abas['PDV'] = pd.to_numeric(df_abas['PDV'], errors='coerce')
                df_abas['SKU'] = df_abas['SKU'].astype(str).str.strip()
                
                marca_correspondente = ABAS_SEGURANCA.get(aba_nome, aba_nome)
                df_abas['MARCA_REFERENCIA'] = marca_correspondente
                
                df_abas = df_abas[['PDV', 'SKU', 'ESTOQUE_DE_SEGURANCA', 'MARCA_REFERENCIA']].copy()
                dfs_abas.append(df_abas)
                
            except Exception:
                continue
        
        if dfs_abas:
            return pd.concat(dfs_abas, ignore_index=True)
        else:
            return pd.DataFrame()
        
    except Exception:
        return pd.DataFrame()

def obter_data_atualizacao_planilha(url_excel):
    """
    Tenta obter a data/hora da última atualização da planilha do Google Sheets.
    """
    excel_buffer = download_arquivo_excel_com_retry(url_excel, "metadados da planilha", timeout=60)
    
    if excel_buffer is None:
        return None
    
    try:
        workbook = openpyxl.load_workbook(excel_buffer, read_only=True, data_only=True)
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
            for row in sheet.iter_rows(min_row=1, max_row=10, max_col=5):
                for cell in row:
                    if cell.value:
                        valor_str = str(cell.value).strip().upper()
                        if any(p in valor_str for p in ['/', ':', '202', '203']):
                            if isinstance(cell.value, datetime):
                                workbook.close()
                                return cell.value.strftime("%d/%m/%Y às %H:%M:%S")
                            elif isinstance(cell.value, str):                                workbook.close()
                                return cell.value
        
        workbook.close()
        return None
    except Exception:
        return None

# 3. Conexão direta via engine do Excel (Otimizado para planilhas públicas)
@st.cache_data(ttl=3600)  # Limpa o cache automaticamente a cada 1 hora
def carregar_dados_nuvem(url_principal, url_seguranca, url_draft):
    dicionario_marcas = {}
    data_atualizacao = None
    
    try:
        # Carrega as planilhas auxiliares
        df_estoque_seguranca = carregar_estoque_seguranca(url_seguranca)
        df_draft = carregar_planilha_draft(url_draft)
        
        # Tenta obter a data de atualização da planilha
        data_atualizacao = obter_data_atualizacao_planilha(url_principal)
        
        # Faz download da planilha principal com retry
        excel_buffer = download_arquivo_excel_com_retry(url_principal, "planilha principal de estoque", timeout=120)
        
        if excel_buffer is None:
            return {}, data_atualizacao
        
        try:
            excel_file = pd.ExcelFile(excel_buffer)
            
            for aba_excel, nome_exibicao in NOMES_MARCAS.items():
                if aba_excel in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=aba_excel)
                    
                    # Adiciona coluna de marca para identificação
                    df['Marca'] = nome_exibicao
                    
                    # Garante que colunas críticas sejam tratadas como números
                    df['PDV'] = pd.to_numeric(df['PDV'], errors='coerce')
                    df['Estoque Atual'] = pd.to_numeric(df['Estoque Atual'], errors='coerce').fillna(0)
                    df['Preço tabela'] = pd.to_numeric(df['Preço tabela'], errors='coerce').fillna(0)
                    
                    # Converte SKU para string para merge correto
                    df['SKU'] = df['SKU'].astype(str).str.strip()
                    
                    # ==========================================
                    # REGRA DE CUSTO COM PLANILHA DRAFT
                    # ==========================================
                    # 1. Se não tiver preço tabela → usa custo da draft                    # 2. Se tiver ambos → usa o MAIOR valor
                    
                    df['Custo_Draft_Original'] = 0.0
                    
                    if not df_draft.empty:
                        # Filtra draft apenas para esta marca (se houver referência)
                        df_draft_merge = df_draft[['PDV', 'SKU', 'CUSTO_DRAFT']].copy()
                        
                        # Merge com a planilha draft
                        df = df.merge(
                            df_draft_merge,
                            on=['PDV', 'SKU'],
                            how='left'
                        )
                        
                        # Preenche NaN com 0
                        df['CUSTO_DRAFT'] = df['CUSTO_DRAFT'].fillna(0)
                        df['Custo_Draft_Original'] = df['CUSTO_DRAFT']
                        
                        # APLICA A REGRA DE CUSTO:
                        # - Se preço tabela = 0 ou vazio → usa CUSTO_DRAFT
                        # - Se ambos têm valor → usa o MAIOR
                        def calcular_custo_final(row):
                            preco_tabela = row['Preço tabela']
                            custo_draft = row['CUSTO_DRAFT']
                            
                            # Se não tem preço tabela, usa custo draft
                            if preco_tabela == 0 or pd.isna(preco_tabela):
                                return custo_draft
                            
                            # Se tem ambos, usa o maior
                            if custo_draft > 0:
                                return max(preco_tabela, custo_draft)
                            
                            # Se só tem preço tabela, usa ele
                            return preco_tabela
                        
                        df['Preço de Custo'] = df.apply(calcular_custo_final, axis=1)
                        
                        # Remove coluna temporária
                        df = df.drop(columns=['CUSTO_DRAFT'])
                    else:
                        # Se não tem draft, usa preço tabela como custo
                        df['Preço de Custo'] = df['Preço tabela']
                    
                    # MERGE com a planilha de estoque de segurança
                    if not df_estoque_seguranca.empty:
                        df_seguranca_marca = df_estoque_seguranca[
                            df_estoque_seguranca['MARCA_REFERENCIA'] == nome_exibicao
                        ].copy()                        
                        if not df_seguranca_marca.empty:
                            df = df.merge(
                                df_seguranca_marca[['PDV', 'SKU', 'ESTOQUE_DE_SEGURANCA']], 
                                on=['PDV', 'SKU'], 
                                how='left'
                            )
                            df['Estoque_Minimo_Qtd'] = df['ESTOQUE_DE_SEGURANCA'].fillna(0)
                            df = df.drop(columns=['ESTOQUE_DE_SEGURANCA'])
                        else:
                            regras_minimo = {'A': 15, 'B': 10, 'C': 5, 'E': 2}
                            df['Estoque_Minimo_Qtd'] = df['Classe'].map(regras_minimo).fillna(2)
                    else:
                        regras_minimo = {'A': 15, 'B': 10, 'C': 5, 'E': 2}
                        df['Estoque_Minimo_Qtd'] = df['Classe'].map(regras_minimo).fillna(2)
                    
                    # Cálculos Financeiros Dinâmicos - Preço de Venda (Tabela)
                    df['Valor_Estoque_Atual'] = df['Estoque Atual'] * df['Preço tabela']
                    df['Valor_Estoque_Minimo'] = df['Estoque_Minimo_Qtd'] * df['Preço tabela']
                    
                    df['Qtd_Excesso'] = (df['Estoque Atual'] - df['Estoque_Minimo_Qtd']).clip(lower=0)
                    df['Valor_Excesso'] = df['Qtd_Excesso'] * df['Preço tabela']
                    
                    df['Qtd_Falta'] = (df['Estoque_Minimo_Qtd'] - df['Estoque Atual']).clip(lower=0)
                    df['Valor_Falta'] = df['Qtd_Falta'] * df['Preço tabela']
                    
                    # CÁLCULOS DE CUSTO - Baseado no Preço de Custo (que pode vir da draft)
                    df['Valor_Custo_Estoque_Atual'] = df['Estoque Atual'] * df['Preço de Custo']
                    df['Valor_Custo_Estoque_Minimo'] = df['Estoque_Minimo_Qtd'] * df['Preço de Custo']
                    
                    dicionario_marcas[nome_exibicao] = df
                    
        except Exception:
            return {}, data_atualizacao
            
    except Exception:
        pass
        
    return dicionario_marcas, data_atualizacao

# Carregamento dos dados e captura do horário de Brasília
with st.spinner("Carregando dados..."):
    dados_marcas, data_atualizacao_planilha = carregar_dados_nuvem(URL_EXCEL, URL_ESTOQUE_SEGURANCA, URL_DRAFT)
    horario_carregamento = obter_horario_brasilia()

if not dados_marcas:
    st.error("❌ Nenhum dado foi carregado. Verifique as permissões de compartilhamento da planilha e sua conexão com a internet.")
    st.stop()

# Determina qual data/hora exibirif data_atualizacao_planilha:
    horario_exibicao = data_atualizacao_planilha
    info_timestamp = "🕒 Última atualização da planilha"
else:
    horario_exibicao = horario_carregamento
    info_timestamp = "🕒 Horário de carregamento do dashboard"

# ==========================================
# CABEÇALHO COM LOGO CP FANI E TIMESTAMP
# ==========================================
col_logo, col_info = st.columns([1, 3])

with col_logo:
    try:
        st.image("logo_cp_fani.png", width=180)
    except Exception:
        pass

with col_info:
    st.title("📊 Painel de Controle de Estoques e Ruptura")
    st.caption(f"{info_timestamp}: **{horario_exibicao}** (Horário de Brasília) | Fonte: Google Sheets")

st.markdown("---")

# 4. Barra Lateral - Filtros
st.sidebar.title("Filtros de Visualização")
primeira_marca = list(dados_marcas.keys())[0]

# Extração limpa e segura dos códigos de PDV únicos
df_pdvs = dados_marcas[primeira_marca]['PDV'].dropna()
todos_pdvs = sorted(df_pdvs.unique().astype(int))
opcoes_selectbox = [DE_PARA_LOJAS.get(pdv, f"PDV {pdv}") for pdv in todos_pdvs]

loja_selecionada_nome = st.sidebar.selectbox("Selecione a Loja / PDV:", opcoes_selectbox)
pdv_selecionado = int(loja_selecionada_nome.split(" - ")[0])

st.sidebar.markdown("---")

# Filtro de Marca
st.sidebar.subheader("Filtro de Marca")
opcoes_marca = ["Todas as Marcas"] + list(dados_marcas.keys())
marca_selecionada = st.sidebar.selectbox("Selecione a Marca:", opcoes_marca)

# Exibe as logos das marcas na sidebar como referência visual
st.sidebar.markdown("**Marcas disponíveis:**")
col_logos_sidebar = st.sidebar.columns(len(dados_marcas))
for idx, (nome_marca, df_marca) in enumerate(dados_marcas.items()):
    with col_logos_sidebar[idx]:
        logo_path = LOGOS_MARCAS.get(nome_marca, '')
        if logo_path:            try:
                st.image(logo_path, width=40)
            except Exception:
                pass
        st.caption(nome_marca)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Forçar Atualização dos Dados"):
    st.cache_data.clear()
    st.rerun()

# Subtítulo da análise
st.subheader(f"Análise Atualizada: {loja_selecionada_nome}")

# Filtra dados pela marca selecionada
if marca_selecionada == "Todas as Marcas":
    dados_filtrados = dados_marcas
    titulo_secao = "Consolidado Geral (Todas as Marcas)"
else:
    dados_filtrados = {marca_selecionada: dados_marcas[marca_selecionada]}
    titulo_secao = f"Análise: {marca_selecionada}"

st.markdown(f"### {titulo_secao}")
st.markdown("---")

# ==========================================
# KPIs CONSOLIDADOS
# ==========================================
v_estoque_atual_total = 0
v_estoque_min_total = 0
v_excesso_total_total = 0
v_falta_total_total = 0
qtd_itens_total = 0

for nome_marca, df_completo in dados_filtrados.items():
    df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
    if not df_loja.empty:
        v_estoque_atual_total += df_loja['Valor_Estoque_Atual'].sum()
        v_estoque_min_total += df_loja['Valor_Estoque_Minimo'].sum()
        v_excesso_total_total += df_loja['Valor_Excesso'].sum()
        v_falta_total_total += df_loja['Valor_Falta'].sum()
        qtd_itens_total += df_loja['Estoque Atual'].sum()

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Valor Estoque Atual (Tabela de Preço)", f"R$ {v_estoque_atual_total:,.2f}")
col2.metric("📉 Valor Estoque Mínimo (Tabela de Preço)", f"R$ {v_estoque_min_total:,.2f}")
col3.metric("⚠️ Capital Preso (Excesso)", f"R$ {v_excesso_total_total:,.2f}", delta=f"{((v_excesso_total_total/v_estoque_atual_total)*100 if v_estoque_atual_total > 0 else 0):.1f}% do estoque", delta_color="inverse")
col4.metric("🚨 Risco de Ruptura (Falta)", f"R$ {v_falta_total_total:,.2f}", delta="Abaixo do Mínimo", delta_color="off")

st.markdown("---")
# ==========================================
# GRÁFICO COMPARATIVO POR MARCA (Quantidade de Itens + Custo Total)
# ==========================================
if marca_selecionada == "Todas as Marcas":
    st.markdown("---")
    st.subheader("📊 Comparativo entre Marcas")
    
    dados_grafico = []
    for nome_marca, df_completo in dados_marcas.items():
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
            fig_qtd_marcas = go.Figure()
            fig_qtd_marcas.add_trace(go.Bar(
                x=df_grafico_marcas['Marca'], 
                y=df_grafico_marcas['Qtd Itens'], 
                name='Quantidade de Itens',
                marker_color=[CORES_MARCAS.get(m, '#007A33') for m in df_grafico_marcas['Marca']],
                text=[f"{v:,.0f}" for v in df_grafico_marcas['Qtd Itens']],
                textposition='auto'
            ))
            fig_qtd_marcas.update_layout(
                template='plotly_dark',
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                height=400,
                title='Quantidade de Itens por Marca',
                yaxis_title='Qtd de Unidades em Estoque'
            )
            st.plotly_chart(fig_qtd_marcas, use_container_width=True)
        
        with col_graf2:
            fig_custo_marcas = go.Figure()
            fig_custo_marcas.add_trace(go.Bar(
                x=df_grafico_marcas['Marca'], 
                y=df_grafico_marcas['Custo Total'], 
                name='Custo Total',
                marker_color=[CORES_MARCAS.get(m, '#007A33') for m in df_grafico_marcas['Marca']],                text=[f"R$ {v:,.0f}" for v in df_grafico_marcas['Custo Total']],
                textposition='auto'
            ))
            fig_custo_marcas.update_layout(
                template='plotly_dark',
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                height=400,
                title='Custo Total por Marca (Draft + Tabela)',
                yaxis_title='Valor (R$)'
            )
            st.plotly_chart(fig_custo_marcas, use_container_width=True)

# ==========================================
# CUSTO POR CURVA (CONSOLIDADO OU POR MARCA)
# ==========================================
st.markdown("---")
st.subheader("📊 Custo Total por Curva de Produto")

df_curva_consolidado = pd.DataFrame()

for nome_marca, df_completo in dados_filtrados.items():
    df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
    if not df_loja.empty and 'Classe' in df_loja.columns:
        df_agrupado = df_loja.groupby('Classe').agg({
            'Valor_Custo_Estoque_Atual': 'sum',
            'SKU': 'count'
        }).reset_index()
        df_agrupado.columns = ['Curva', 'Custo Total', 'Qtd SKUs']
        df_agrupado['Marca'] = nome_marca
        df_curva_consolidado = pd.concat([df_curva_consolidado, df_agrupado], ignore_index=True)

if not df_curva_consolidado.empty:
    if marca_selecionada == "Todas as Marcas":
        df_pivot = df_curva_consolidado.pivot_table(
            values='Custo Total', 
            index='Curva', 
            columns='Marca', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        colunas_marcas = [col for col in df_pivot.columns if col != 'Curva']
        df_pivot['Total Geral'] = df_pivot[colunas_marcas].sum(axis=1)
        
        linha_total = {'Curva': 'TOTAL'}
        for col in colunas_marcas:
            linha_total[col] = df_pivot[col].sum()
        linha_total['Total Geral'] = df_pivot['Total Geral'].sum()
        df_pivot = pd.concat([df_pivot, pd.DataFrame([linha_total])], ignore_index=True)        
        colunas_valor = [col for col in df_pivot.columns if col != 'Curva']
        for col in colunas_valor:
            df_pivot[col] = df_pivot[col].apply(lambda x: f"R$ {x:,.2f}")
        
        st.dataframe(df_pivot, use_container_width=True, hide_index=True)
    else:
        df_exibicao = df_curva_consolidado[['Curva', 'Custo Total', 'Qtd SKUs']].copy()
        df_exibicao = df_exibicao.sort_values('Curva')
        
        total_custo = df_exibicao['Custo Total'].sum()
        total_skus = df_exibicao['Qtd SKUs'].sum()
        df_total = pd.DataFrame([{'Curva': 'TOTAL', 'Custo Total': total_custo, 'Qtd SKUs': total_skus}])
        df_exibicao = pd.concat([df_exibicao, df_total], ignore_index=True)
        
        df_exibicao['Custo Total'] = df_exibicao['Custo Total'].apply(lambda x: f"R$ {x:,.2f}")
        st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
    
    fig_custo = go.Figure()
    
    if marca_selecionada == "Todas as Marcas":
        for nome_marca in dados_filtrados.keys():
            df_marca_curva = df_curva_consolidado[df_curva_consolidado['Marca'] == nome_marca]
            if not df_marca_curva.empty:
                fig_custo.add_trace(go.Bar(
                    x=df_marca_curva['Curva'], 
                    y=df_marca_curva['Custo Total'], 
                    name=nome_marca,
                    marker_color=CORES_MARCAS.get(nome_marca, '#007A33')
                ))
        fig_custo.update_layout(barmode='stack')
    else:
        fig_custo.add_trace(go.Bar(
            x=df_curva_consolidado['Curva'], 
            y=df_curva_consolidado['Custo Total'], 
            marker_color=[CORES_MARCAS.get(marca_selecionada, '#007A33')] * len(df_curva_consolidado),
            text=[f"R$ {v:,.2f}" for v in df_curva_consolidado['Custo Total']],
            textposition='auto'
        ))
    
    fig_custo.update_layout(
        template='plotly_dark', 
        plot_bgcolor='rgba(0,0,0,0)', 
        paper_bgcolor='rgba(0,0,0,0)', 
        height=400,
        xaxis_title='Curva',
        yaxis_title='Custo Total (R$)',
        title='Distribuição de Custo por Curva'
    )
    st.plotly_chart(fig_custo, use_container_width=True)
# ==========================================
# ANÁLISE POR CATEGORIA
# ==========================================
st.markdown("---")
st.subheader("📊 Análise por Categoria")

df_categoria_consolidado = pd.DataFrame()

for nome_marca, df_completo in dados_filtrados.items():
    df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
    if not df_loja.empty and 'Categoria' in df_loja.columns:
        df_cat = df_loja.groupby('Categoria')[['Valor_Estoque_Atual', 'Valor_Estoque_Minimo']].sum().reset_index()
        df_cat['Marca'] = nome_marca
        df_categoria_consolidado = pd.concat([df_categoria_consolidado, df_cat], ignore_index=True)

if not df_categoria_consolidado.empty:
    if marca_selecionada == "Todas as Marcas":
        fig_categoria = go.Figure()
        for nome_marca in dados_filtrados.keys():
            df_marca_cat = df_categoria_consolidado[df_categoria_consolidado['Marca'] == nome_marca]
            if not df_marca_cat.empty:
                fig_categoria.add_trace(go.Bar(
                    x=df_marca_cat['Categoria'], 
                    y=df_marca_cat['Valor_Estoque_Atual'], 
                    name=nome_marca,
                    marker_color=CORES_MARCAS.get(nome_marca, '#007A33')
                ))
        fig_categoria.update_layout(barmode='group')
    else:
        fig_categoria = go.Figure()
        fig_categoria.add_trace(go.Bar(
            x=df_categoria_consolidado['Categoria'], 
            y=df_categoria_consolidado['Valor_Estoque_Atual'], 
            name='Estoque Atual',
            marker_color=CORES_MARCAS.get(marca_selecionada, '#007A33')
        ))
        fig_categoria.add_trace(go.Bar(
            x=df_categoria_consolidado['Categoria'], 
            y=df_categoria_consolidado['Valor_Estoque_Minimo'], 
            name='Estoque Mínimo',
            marker_color='#ff4b4b'
        ))
        fig_categoria.update_layout(barmode='group')
    
    fig_categoria.update_layout(
        template='plotly_dark', 
        plot_bgcolor='rgba(0,0,0,0)', 
        paper_bgcolor='rgba(0,0,0,0)', 
        height=400,        xaxis_title='Categoria',
        yaxis_title='Valor (R$)',
        title='Estoque por Categoria'
    )
    st.plotly_chart(fig_categoria, use_container_width=True)

# ==========================================
# TABELAS DE EXCESSOS E FALTAS (POR MARCA)
# ==========================================
st.markdown("---")

for nome_marca, df_completo in dados_filtrados.items():
    df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
    
    if df_loja.empty:
        continue
    
    exibir_titulo_marca(nome_marca, tamanho_logo=35)
    
    col_tab1, col_tab2 = st.columns(2)
    with col_tab1:
        st.write("### 🛑 Excessos Críticos")
        
        df_excesso_tabela = df_loja[
            (df_loja['Valor_Excesso'] > 0) & 
            (df_loja['Estoque_Minimo_Qtd'] > 0)
        ][
            ['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Excesso', 'Preço tabela', 'Preço de Custo', 'Valor_Excesso']
        ].sort_values(by='Valor_Excesso', ascending=False)
        
        st.dataframe(df_excesso_tabela.style.format({
            'Preço tabela': 'R$ {:.2f}',
            'Preço de Custo': 'R$ {:.2f}',
            'Valor_Excesso': 'R$ {:.2f}'
        }), use_container_width=True, height=280)
        
    with col_tab2:
        st.write("### 🚨 Produtos Críticos em Falta / Ruptura")
        df_falta_tabela = df_loja[df_loja['Valor_Falta'] > 0][
            ['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Falta', 'Preço tabela', 'Preço de Custo', 'Valor_Falta']
        ].sort_values(by='Valor_Falta', ascending=False)
        st.dataframe(df_falta_tabela.style.format({
            'Preço tabela': 'R$ {:.2f}',
            'Preço de Custo': 'R$ {:.2f}',
            'Valor_Falta': 'R$ {:.2f}'
        }), use_container_width=True, height=280)
    
    st.markdown("---")