import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# 1. Configuração Inicial da Página (Visual Modo Escuro)
st.set_page_config(
    page_title="Painel de Performance de Estoque NSF",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS para forçar o tema escuro premium
st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    div[data-testid="stMetricValue"] { font-size: 28px; color: #00f2fe; }
    div[data-testid="stMetricLabel"] { font-size: 14px; color: #a3b8cc; }
    .stTabs [data-baseweb="tab"] { color: #a3b8cc; font-size: 16px; }
    .stTabs [data-baseweb="tab"][aria-selected="true"] { color: #00f2fe; font-weight: bold; }
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

# 3. Conexão direta via engine do Excel (Otimizado para planilhas públicas)
@st.cache_data(ttl=3600)  # Limpa o cache automaticamente a cada 1 hora
def carregar_dados_nuvem(url):
    dicionario_marcas = {}
    abas = {'BOTICARIO': 'O Boticário 🟢', 'EUDORA': 'Eudora 🟣', 'QUEM_DISSE_BERENICE': 'Quem Disse, Berenice? 💖'}
    
    try:
        # Baixa o arquivo binário completo do Excel direto da nuvem
        excel_file = pd.ExcelFile(url)
        
        for aba_excel, nome_exibicao in abas.items():
            if aba_excel in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=aba_excel)
                
                # Garante que colunas críticas sejam tratadas como números
                df['PDV'] = pd.to_numeric(df['PDV'], errors='coerce')
                df['Estoque Atual'] = pd.to_numeric(df['Estoque Atual'], errors='coerce').fillna(0)
                df['Preço tabela'] = pd.to_numeric(df['Preço tabela'], errors='coerce').fillna(0)
                
                # Tenta ler a coluna de Preço de Custo (nomes possíveis)
                colunas_custo_possiveis = ['Preço de Custo', 'Custo Unitário', 'Custo', 'Preco de Custo', 'Preco_Custo', 'custo']
                df['Preço de Custo'] = 0  # Valor padrão
                
                for col_custo in colunas_custo_possiveis:
                    if col_custo in df.columns:
                        df['Preço de Custo'] = pd.to_numeric(df[col_custo], errors='coerce').fillna(0)
                        break
                
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
                
                # NOVOS CÁLCULOS - Preço de Custo
                df['Valor_Custo_Estoque_Atual'] = df['Estoque Atual'] * df['Preço de Custo']
                df['Valor_Custo_Estoque_Minimo'] = df['Estoque_Minimo_Qtd'] * df['Preço de Custo']
                
                dicionario_marcas[nome_exibicao] = df
            else:
                st.error(f"Aba {aba_excel} não encontrada no arquivo do Drive.")
    except Exception as e:
        st.error(f"Erro ao conectar ou ler o arquivo do Google Drive: {e}")
        
    return dicionario_marcas

# Carregamento dos dados
with st.spinner("Conectando ao Google Drive e processando bases..."):
    dados_marcas = carregar_dados_nuvem(URL_EXCEL)

if not dados_marcas:
    st.error("Nenhum dado foi carregado. Verifique as permissões de compartilhamento da planilha.")
    st.stop()

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
if st.sidebar.button("🔄 Forçar Atualização dos Dados"):
    st.cache_data.clear()
    st.rerun()

# 5. Corpo do Painel
st.title("📊 Painel de Controle de Estoques e Ruptura")
st.subheader(f"Análise Atualizada: {loja_selecionada_nome}")

abas_tela = st.tabs(list(dados_marcas.keys()))

for i, (nome_marca, df_completo) in enumerate(dados_marcas.items()):
    with abas_tela[i]:
        df_loja = df_completo[df_completo['PDV'] == pdv_selecionado]
        
        if df_loja.empty:
            st.warning(f"Sem registros de movimentação para este PDV na marca {nome_marca}.")
            continue
            
        # KPIs - Preço de Venda
        v_estoque_atual = df_loja['Valor_Estoque_Atual'].sum()
        v_estoque_min = df_loja['Valor_Estoque_Minimo'].sum()
        v_excesso_total = df_loja['Valor_Excesso'].sum()
        v_falta_total = df_loja['Valor_Falta'].sum()
        
        # KPIs - Preço de Custo
        v_custo_estoque_atual = df_loja['Valor_Custo_Estoque_Atual'].sum()
        v_custo_estoque_min = df_loja['Valor_Custo_Estoque_Minimo'].sum()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💰 Valor Estoque Atual (Venda)", f"R$ {v_estoque_atual:,.2f}")
        col2.metric("📉 Valor Estoque Mínimo (Venda)", f"R$ {v_estoque_min:,.2f}")
        col3.metric("️ Capital Preso (Excesso)", f"R$ {v_excesso_total:,.2f}", delta=f"{((v_excesso_total/v_estoque_atual)*100 if v_estoque_atual > 0 else 0):.1f}% do estoque", delta_color="inverse")
        col4.metric(" Risco de Ruptura (Falta)", f"R$ {v_falta_total:,.2f}", delta="Abaixo do Mínimo", delta_color="off")
        
        # Nova linha de KPIs - Custo
        st.markdown("---")
        st.subheader("💵 Análise de Custos")
        col5, col6 = st.columns(2)
        col5.metric("💵 Custo Total do Estoque Atual", f"R$ {v_custo_estoque_atual:,.2f}", help="Soma do preço de custo de todos os produtos em estoque")
        col6.metric("💵 Custo Total do Estoque Mínimo", f"R$ {v_custo_estoque_min:,.2f}", help="Soma do preço de custo do estoque mínimo necessário")
        
        # Verifica se há dados de custo
        tem_custo = v_custo_estoque_atual > 0
        
        if not tem_custo:
            st.info("ℹ️ **Atenção:** A planilha não contém coluna de 'Preço de Custo'. Os valores de custo estão zerados. Adicione uma coluna chamada 'Preço de Custo' ou 'Custo Unitário' na planilha para visualizar esta análise.")
        
        # Tabela de Custo por Curva (A, B, C, E)
        st.markdown("---")
        st.subheader("📊 Custo Total por Curva de Produto")
        
        if 'Classe' in df_loja.columns:
            # Agrupa por Classe e calcula o total de custo do estoque atual
            df_agrupado = df_loja.groupby('Classe').agg({
                'Valor_Custo_Estoque_Atual': 'sum',
                'SKU': 'count'
            }).reset_index()
            df_agrupado.columns = ['Curva', 'Custo Total', 'Qtd SKUs']
            
            # Ordena por Curva
            df_agrupado = df_agrupado.sort_values('Curva')
            
            # Formata para exibição
            df_exibicao = df_agrupado.copy()
            df_exibicao['Custo Total'] = df_exibicao['Custo Total'].apply(lambda x: f"R$ {x:,.2f}")
            
            # Exibe a tabela
            st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
            
            # Gráfico de custo por curva (apenas se houver custo)
            if tem_custo and not df_agrupado.empty:
                fig_custo = go.Figure()
                fig_custo.add_trace(go.Bar(
                    x=df_agrupado['Curva'], 
                    y=df_agrupado['Custo Total'], 
                    marker_color=['#00f2fe', '#ff4b4b', '#ffa500', '#90ee90'],
                    text=[f"R$ {v:,.2f}" for v in df_agrupado['Custo Total']],
                    textposition='auto'
                ))
                fig_custo.update_layout(
                    template='plotly_dark', 
                    plot_bgcolor='rgba(0,0,0,0)', 
                    paper_bgcolor='rgba(0,0,0,0)', 
                    height=350,
                    xaxis_title='Curva',
                    yaxis_title='Custo Total (R$)',
                    showlegend=False
                )
                st.plotly_chart(fig_custo, use_container_width=True)
        else:
            st.warning("Coluna 'Classe' não encontrada nos dados.")
        
        st.markdown("---")
        
        # Gráfico - Verifica se a coluna Categoria existe e se há dados
        if 'Categoria' in df_loja.columns:
            df_grafico = df_loja.groupby('Categoria')[['Valor_Estoque_Atual', 'Valor_Estoque_Minimo']].sum().reset_index()
            
            if not df_grafico.empty:
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_grafico['Categoria'], y=df_grafico['Valor_Estoque_Atual'], name='Estoque Atual (R$)', marker_color='#00f2fe'))
                fig.add_trace(go.Bar(x=df_grafico['Categoria'], y=df_grafico['Valor_Estoque_Minimo'], name='Estoque Mínimo (R$)', marker_color='#ff4b4b'))
                fig.update_layout(
                    barmode='group', 
                    template='plotly_dark', 
                    plot_bgcolor='rgba(0,0,0,0)', 
                    paper_bgcolor='rgba(0,0,0,0)', 
                    height=320,
                    xaxis_title='Categoria',
                    yaxis_title='Valor (R$)'
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Sem dados suficientes para gerar o gráfico de categorias.")
        else:
            st.warning("Coluna 'Categoria' não encontrada nos dados.")
        
        st.markdown("---")
        
        # Tabelas
        col_tab1, col_tab2 = st.columns(2)
        with col_tab1:
            st.write("### 🛑 Maiores Excessos Críticos (Dinheiro Parado)")
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