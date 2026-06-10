import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timezone, timedelta

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

# ID OFICIAL DA PLANILHA EXTRAÍDO DO LINK FORNECIDO
SPREADSHEET_ID = "1PbNYsNPp6ShErx0U3Ml_dJpN-0MPwoxz"

# URL de exportação direta em formato Excel (Método otimizado para planilhas públicas)
URL_EXCEL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=xlsx"

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

# Nomes limpos das marcas (sem emojis) - usados como chaves internas
NOMES_MARCAS = {
    'BOTICARIO': 'O Boticário',
    'EUDORA': 'Eudora',
    'QUEM_DISSE_BERENICE': 'Quem Disse, Berenice?'
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
                st.image(logo_path, width=tamanho_logo)
            except Exception:
                st.write("🏷️")
        else:
            st.write("️")
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

# 3. Conexão direta via engine do Excel (Otimizado para planilhas públicas)
@st.cache_data(ttl=3600)  # Limpa o cache automaticamente a cada 1 hora
def carregar_dados_nuvem(url):
    dicionario_marcas = {}
    
    try:
        # Baixa o arquivo binário completo do Excel direto da nuvem
        excel_file = pd.ExcelFile(url)
        
        for aba_excel, nome_exibicao in NOMES_MARCAS.items():
            if aba_excel in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=aba_excel)
                
                # Adiciona coluna de marca para identificação
                df['Marca'] = nome_exibicao
                
                # Garante que colunas críticas sejam tratadas como números
                df['PDV'] = pd.to_numeric(df['PDV'], errors='coerce')
                df['Estoque Atual'] = pd.to_numeric(df['Estoque Atual'], errors='coerce').fillna(0)
                df['Preço tabela'] = pd.to_numeric(df['Preço tabela'], errors='coerce').fillna(0)
                
                # UTILIZA PREÇO TABELA COMO BASE PARA CUSTO
                df['Preço de Custo'] = df['Preço tabela']
                
                # Regras de Estoque Mínimo por Curva
                regras_minimo = {'A': 15, 'B': 10, 'C': 5, 'E': 2}
                df['Estoque_Minimo_Qtd'] = df['Classe'].map(regras_minimo).fillna(2)
                
                # Cálculos Financeiros Dinâmicos - Preço de Venda
                df['Valor_Estoque_Atual'] = df['Estoque Atual'] * df['Preço tabela']
                df['Valor_Estoque_Minimo'] = df['Estoque_Minimo_Qtd'] * df['Preço tabela']
                
                df['Qtd_Excesso'] = (df['Estoque Atual'] - df['Estoque_Minimo_Qtd']).clip(lower=0)
                df['Valor_Excesso'] = df['Qtd_Excesso'] * df['Preço tabela']
                
                df['Qtd_Falta'] = (df['Estoque_Minimo_Qtd'] - df['Estoque Atual']).clip(lower=0)
                df['Valor_Falta'] = df['Qtd_Falta'] * df['Preço tabela']
                
                # CÁLCULOS DE CUSTO - Baseado no Preço Tabela
                df['Valor_Custo_Estoque_Atual'] = df['Estoque Atual'] * df['Preço de Custo']
                df['Valor_Custo_Estoque_Minimo'] = df['Estoque_Minimo_Qtd'] * df['Preço de Custo']
                
                dicionario_marcas[nome_exibicao] = df
            else:
                st.error(f"Aba {aba_excel} não encontrada no arquivo do Drive.")
    except Exception as e:
        st.error(f"Erro ao conectar ou ler o arquivo do Google Drive: {e}")
        
    return dicionario_marcas

# Carregamento dos dados e captura do horário de Brasília
with st.spinner("Conectando ao Google Drive e processando bases..."):
    dados_marcas = carregar_dados_nuvem(URL_EXCEL)
    horario_atualizacao = obter_horario_brasilia()

if not dados_marcas:
    st.error("Nenhum dado foi carregado. Verifique as permissões de compartilhamento da planilha.")
    st.stop()

# ==========================================
# CABEÇALHO COM LOGO CP FANI E TIMESTAMP
# ==========================================
col_logo, col_info = st.columns([1, 3])

with col_logo:
    try:
        st.image("logo_cp_fani.png", width=180)
    except Exception:
        st.warning("Logo CP Fani não encontrada. Certifique-se de que o arquivo 'logo_cp_fani.png' está na raiz do projeto.")

with col_info:
    st.title("📊 Painel de Controle de Estoques e Ruptura")
    st.caption(f"🕒 Última atualização da base: **{horario_atualizacao}** (Horário de Brasília) | Fonte: Google Sheets")

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

# Filtro de Marca (apenas nomes, sem HTML - Streamlit não renderiza HTML em selectbox)
st.sidebar.subheader("Filtro de Marca")
opcoes_marca = ["Todas as Marcas"] + list(dados_marcas.keys())
marca_selecionada = st.sidebar.selectbox("Selecione a Marca:", opcoes_marca)

# Exibe as logos das marcas na sidebar como referência visual
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
v_custo_estoque_atual_total = 0
v_custo_estoque_min_total = 0
qtd_itens_total = 0

for nome_marca, df_completo in dados_filtrados.items():
    df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
    if not df_loja.empty:
        v_estoque_atual_total += df_loja['Valor_Estoque_Atual'].sum()
        v_estoque_min_total += df_loja['Valor_Estoque_Minimo'].sum()
        v_excesso_total_total += df_loja['Valor_Excesso'].sum()
        v_falta_total_total += df_loja['Valor_Falta'].sum()
        v_custo_estoque_atual_total += df_loja['Valor_Custo_Estoque_Atual'].sum()
        v_custo_estoque_min_total += df_loja['Valor_Custo_Estoque_Minimo'].sum()
        qtd_itens_total += df_loja['Estoque Atual'].sum()

col1, col2, col3, col4 = st.columns(4)
col1.metric(" Valor Estoque Atual (Venda)", f"R$ {v_estoque_atual_total:,.2f}")
col2.metric("📉 Valor Estoque Mínimo (Venda)", f"R$ {v_estoque_min_total:,.2f}")
col3.metric("⚠️ Capital Preso (Excesso)", f"R$ {v_excesso_total_total:,.2f}", delta=f"{((v_excesso_total_total/v_estoque_atual_total)*100 if v_estoque_atual_total > 0 else 0):.1f}% do estoque", delta_color="inverse")
col4.metric("🚨 Risco de Ruptura (Falta)", f"R$ {v_falta_total_total:,.2f}", delta="Abaixo do Mínimo", delta_color="off")

st.markdown("---")
st.subheader("💵 Análise de Custos (Baseado no Preço Tabela)")
col5, col6 = st.columns(2)
col5.metric("💵 Custo Total do Estoque Atual", f"R$ {v_custo_estoque_atual_total:,.2f}", help="Soma do preço de tabela de todos os produtos em estoque")
col6.metric("💵 Custo Total do Estoque Mínimo", f"R$ {v_custo_estoque_min_total:,.2f}", help="Soma do preço de tabela do estoque mínimo necessário")

# ==========================================
# GRÁFICO COMPARATIVO POR MARCA (Quantidade de Itens + Custo Total)
# ==========================================
if marca_selecionada == "Todas as Marcas":
    st.markdown("---")
    st.subheader("📊 Comparativo entre Marcas")
    
    # Prepara dados para o gráfico
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
                marker_color=[CORES_MARCAS.get(m, '#007A33') for m in df_grafico_marcas['Marca']],
                text=[f"R$ {v:,.0f}" for v in df_grafico_marcas['Custo Total']],
                textposition='auto'
            ))
            fig_custo_marcas.update_layout(
                template='plotly_dark',
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                height=400,
                title='Custo Total por Marca',
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
        # Mostra tabela consolidada com marca
        df_pivot = df_curva_consolidado.pivot_table(
            values='Custo Total', 
            index='Curva', 
            columns='Marca', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        # Adiciona coluna de total geral (soma de todas as marcas por curva)
        colunas_marcas = [col for col in df_pivot.columns if col != 'Curva']
        df_pivot['Total Geral'] = df_pivot[colunas_marcas].sum(axis=1)
        
        # Adiciona linha de total por marca (soma de todas as curvas)
        linha_total = {'Curva': 'TOTAL'}
        for col in colunas_marcas:
            linha_total[col] = df_pivot[col].sum()
        linha_total['Total Geral'] = df_pivot['Total Geral'].sum()
        df_pivot = pd.concat([df_pivot, pd.DataFrame([linha_total])], ignore_index=True)
        
        # Formata valores monetários
        colunas_valor = [col for col in df_pivot.columns if col != 'Curva']
        for col in colunas_valor:
            df_pivot[col] = df_pivot[col].apply(lambda x: f"R$ {x:,.2f}")
        
        st.dataframe(df_pivot, use_container_width=True, hide_index=True)
    else:
        # Mostra apenas a marca selecionada
        df_exibicao = df_curva_consolidado[['Curva', 'Custo Total', 'Qtd SKUs']].copy()
        df_exibicao = df_exibicao.sort_values('Curva')
        
        # Adiciona linha de total
        total_custo = df_exibicao['Custo Total'].sum()
        total_skus = df_exibicao['Qtd SKUs'].sum()
        df_total = pd.DataFrame([{'Curva': 'TOTAL', 'Custo Total': total_custo, 'Qtd SKUs': total_skus}])
        df_exibicao = pd.concat([df_exibicao, df_total], ignore_index=True)
        
        df_exibicao['Custo Total'] = df_exibicao['Custo Total'].apply(lambda x: f"R$ {x:,.2f}")
        st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
    
    # Gráfico de custo por curva
    fig_custo = go.Figure()
    
    if marca_selecionada == "Todas as Marcas":
        # Gráfico empilhado por marca
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
        # Gráfico simples
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
else:
    st.warning("Coluna 'Classe' não encontrada nos dados.")

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
        # Gráfico comparativo por categoria e marca
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
        # Gráfico simples
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
        height=400,
        xaxis_title='Categoria',
        yaxis_title='Valor (R$)',
        title='Estoque por Categoria'
    )
    st.plotly_chart(fig_categoria, use_container_width=True)
else:
    st.warning("Coluna 'Categoria' não encontrada nos dados.")

# ==========================================
# TABELAS DE EXCESSOS E FALTAS (POR MARCA)
# ==========================================
st.markdown("---")

for nome_marca, df_completo in dados_filtrados.items():
    df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
    
    if df_loja.empty:
        st.warning(f"Sem registros de movimentação para este PDV na marca {nome_marca}.")
        continue
    
    # Título da marca com logo (usando st.columns + st.image)
    exibir_titulo_marca(nome_marca, tamanho_logo=35)
    
    col_tab1, col_tab2 = st.columns(2)
    with col_tab1:
        st.write("### 🛑 Excessos Críticos")
        df_excesso_tabela = df_loja[df_loja['Valor_Excesso'] > 0][
            ['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Excesso', 'Preço tabela', 'Valor_Excesso']
        ].sort_values(by='Valor_Excesso', ascending=False)
        st.dataframe(df_excesso_tabela.style.format({'Preço tabela': 'R$ {:.2f}', 'Valor_Excesso': 'R$ {:.2f}'}), use_container_width=True, height=280)
        
    with col_tab2:
        st.write("### 🚨 Produtos Críticos em Falta / Ruptura")
        df_falta_tabela = df_loja[df_loja['Valor_Falta'] > 0][
            ['SKU', 'Descrição', 'Classe', 'Estoque Atual', 'Estoque_Minimo_Qtd', 'Qtd_Falta', 'Preço tabela', 'Valor_Falta']
        ].sort_values(by='Valor_Falta', ascending=False)
        st.dataframe(df_falta_tabela.style.format({'Preço tabela': 'R$ {:.2f}', 'Valor_Falta': 'R$ {:.2f}'}), use_container_width=True, height=280)
    
    st.markdown("---")