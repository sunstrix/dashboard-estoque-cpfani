# 📊 Dashboard de Performance de Estoque NSF - CPFANI

Painel de controle interativo para análise de estoques e ruptura de produtos das marcas **O Boticário**, **Eudora** e **Quem Disse, Berenice?** em 17 PDVs (Pontos de Venda).

Este dashboard fornece visualizações em tempo real dos níveis de estoque, identifica excessos críticos e produtos em risco de ruptura, permitindo tomadas de decisão rápidas e baseadas em dados.

---

## 🎯 Funcionalidades Principais

- ✅ **Análise por PDV**: Visualização individualizada para cada uma das 17 lojas
- ✅ **KPIs Financeiros**: Valor de estoque atual, estoque mínimo, capital preso e risco de ruptura
- ✅ **Custo Inteligente**: Regra automática que considera o maior valor entre preço de tabela e custo da planilha draft
- ✅ **Estoque de Segurança**: Importado de planilha separada com 3 abas (BOT, EUD, QDB)
- ✅ **Filtro de Excessos Críticos**: SKUs com estoque de segurança = 0 são automaticamente excluídos
- ✅ **Gráficos Interativos**: Comparativo visual entre marcas, categorias e curvas de produtos
- ✅ **Tabelas Detalhadas**: Listagem dos maiores excessos e produtos em falta
- ✅ **Dados em Tempo Real**: Leitura direta das planilhas do Google Sheets com cache de 1 hora
- ✅ **Interface Premium**: Design moderno em modo escuro com tema O Boticário (verde + dourado)
- ✅ **Filtros Dinâmicos**: Seleção rápida de loja/PDV e marca para análise focada
- ✅ **Logos das Marcas**: Identificação visual com logos oficiais (Boticário, Eudora, QDB)
- ✅ **Horário de Brasília**: Timestamp sempre exibido no fuso UTC-3

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.11+** - Linguagem principal
- **Streamlit** - Framework para dashboards web interativos
- **Pandas** - Manipulação e análise de dados
- **Plotly** - Visualizações gráficas interativas
- **OpenPyXL** - Leitura de arquivos Excel
- **Requests** - Download com retry automático de planilhas
- **Google Sheets** - Fonte de dados em nuvem (3 planilhas integradas)

---

## 📋 Pré-requisitos

Antes de começar, certifique-se de ter:

- Windows 10/11
- Conexão com a internet (para acessar as planilhas do Google Drive)
- Navegador web moderno (Chrome, Edge, Firefox)

**Nota:** O script de instalação automatizada (`instalar.bat`) cuidará de instalar o Python e todas as dependências necessárias.

---

## 🚀 Instalação Rápida

### Passo 1: Executar o Script de Instalação

1. Navegue até a pasta do projeto: `Desktop\dashboard-estoque-cpfani`
2. Clique com o botão direito no arquivo `instalar.bat` e selecione **"Executar como administrador"**
3. Aguarde a conclusão do processo. O script irá:
   - ✅ Verificar se o Python já está instalado
   - ✅ Baixar e instalar o Python 3.11.8 (se necessário)
   - ✅ Instalar todas as dependências do projeto
   - ✅ Gerar um arquivo de log detalhado (`instalar.log`)
4. Ao final, você verá a mensagem: **"Instalação concluída!"**

### Passo 2: Verificar a Instalação

Após a instalação, você pode verificar se tudo está correto abrindo o Prompt de Comando e executando:

```bash
python --version
pip list
```

Você deve ver o Python 3.11.8 e todas as bibliotecas listadas no `requirements.txt`.

---

## 📊 Configuração das Planilhas

O dashboard está integrado com **3 planilhas públicas** do Google Sheets:

### 1. Planilha Principal de Estoque

- **ID**: `1EDDyKie9UiugMLMowcPzHfViqzziFcSgxVPvZ2Rx3L0`
- **URL**: https://docs.google.com/spreadsheets/d/1EDDyKie9UiugMLMowcPzHfViqzziFcSgxVPvZ2Rx3L0/edit
- **Abas**: `BOTICARIO`, `EUDORA`, `QUEM_DISSE_BERENICE`

### 2. Planilha de Estoque de Segurança

- **ID**: `1uHonFnFM4p7bz4s7YpewhKHNs6fSEfw9rDMTKC7jtHE`
- **URL**: https://docs.google.com/spreadsheets/d/1uHonFnFM4p7bz4s7YpewhKHNs6fSEfw9rDMTKC7jtHE/edit
- **Abas**: `BOT` (O Boticário), `EUD` (Eudora), `QDB` (Quem Disse, Berenice?)
- **Colunas obrigatórias**: `PDV`, `SKU`, `ESTOQUE DE SEGURANCA`

### 3. Planilha Draft de Custos

- **ID**: `11Z21gFvJ9pm2xSlF3IweC7xcYZwAZWrjcWDnRe5LexY`
- **URL**: https://docs.google.com/spreadsheets/d/11Z21gFvJ9pm2xSlF3IweC7xcYZwAZWrjcWDnRe5LexY/edit
- **Coluna de Custo**: Coluna J (ou coluna nomeada como "CUSTO")
- **Colunas obrigatórias**: Loja (nome completo), SKU, Custo

---

## 💰 Regra de Custo Inteligente

O dashboard aplica automaticamente a seguinte regra para determinar o **Preço de Custo** de cada SKU:

| Situação | Preço de Tabela | Custo (Draft) | Resultado |
|----------|-----------------|---------------|-----------|
| Apenas Tabela | ✅ Tem valor | ❌ Vazio/0 | Usa Preço de Tabela |
| Apenas Draft | ❌ Vazio/0 | ✅ Tem valor | Usa Custo da Draft |
| Ambos | ✅ Tem valor | ✅ Tem valor | Usa o **MAIOR** valor |
| Nenhum | ❌ Vazio/0 | ❌ Vazio/0 | Resultado = 0 |

Esta regra garante que o custo considerado seja sempre o mais conservador (maior valor), evitando subavaliação do estoque.

---

## 🛡️ Regra de Excessos Críticos

**Importante:** SKUs que possuem **Estoque de Segurança = 0** são automaticamente **excluídos** da lista de Excessos Críticos. Isso evita que produtos sem meta de estoque mínimo apareçam erroneamente como "excesso".

---

## 🏪 Mapeamento de PDVs

O dashboard mapeia automaticamente os nomes completos das lojas da planilha draft para os códigos numéricos dos PDVs:

| Código | Loja | Nome na Planilha Draft |
|--------|------|------------------------|
| 4842 | Metrópole | Loja: 4842 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 5152 | Coração | Loja: 5152 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 6105 | Assai Anchieta | Loja: 6105 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 6106 | Direita | Loja: 6106 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 6110 | Arouche | Loja: 6110 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 8001 | Dom José | Loja: 8001 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 11576 | Davó | Loja: 11576 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 12055 | São Bento | Loja: 12055 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 12056 | Marechal | Loja: 12056 - S. P. ARON COSMETICOS EPP |
| 12605 | Coop | Loja: 12605 - N.S.F. COSMETICOS E PRESENTES LTDA. |
| 12645 | Light | Loja: 12645 - N. S. F. COSMETICOS E PRESENTES LTDA |
| 14120 | VD SBC | Loja: 14120 - ARPEL DISTRIBUIDORA DE COSMETICOS LTDA - EPP |
| 14353 | VD SP | Loja: 14353 - ARPEL DISTRIBUIDORA DE COSMETICOS LTDA - EPP |
| 20371 | Luz | Loja: 20371 - N. S. F. COSMÉTICOS E PRESENTES LTDA. |
| 21502 | Bem Barato | Loja: 21502 - N. S. F. COSMETICOS E PRESENTES LTD |
| 23000 | Outlet | Loja: 23000 - N. S. F. COSMETICOS E PRESENTES LTD |
| 23379 | Assai Piraporinha | Loja: 23379 - N. S. F. COSMETICOS E PRESENTES LTD |

---

## 📁 Estrutura do Projeto

```text
dashboard-estoque-cpfani/
│
├── app.py                      # Arquivo principal do dashboard
├── config.py                   # Configurações de conexão com Google Sheets
├── requirements.txt            # Lista de dependências Python
├── instalar.bat                # Script de instalação automatizada
├── logo_cp_fani.png            # Logo principal CP Fani
├── logo_boticario.png          # Logo O Boticário
├── logo_eudora.png             # Logo Eudora
├── logo_qdb.png                # Logo Quem Disse, Berenice?
├── .env                        # Variáveis de ambiente (não versionado)
├── .gitignore                  # Arquivos ignorados pelo Git
├── README.md                   # Este arquivo de documentação
└── instalar.log                # Log da instalação (gerado automaticamente)
```

---

## 🎨 Identidade Visual

O dashboard utiliza o tema **O Boticário** com as seguintes cores:

- **Verde Principal**: `#007A33` (cor oficial O Boticário)
- **Dourado**: `#D4AF37` (destaque para KPIs e títulos)
- **Fundo Escuro**: `#0e1117` (modo escuro premium)
- **Eudora**: `#a855f7` (roxo)
- **Quem Disse, Berenice?**: `#ff4b4b` (vermelho)

---

## ▶️ Executando o Dashboard

### Método 1: Via Linha de Comando (Recomendado)

1. Abra o Prompt de Comando ou PowerShell na pasta do projeto
2. Execute o comando:

```bash
streamlit run app.py
```

3. O dashboard abrirá automaticamente no navegador em `http://localhost:8501`

### Método 2: Script de Atalho

Você pode criar um arquivo `executar.bat` na raiz do projeto com o seguinte conteúdo para facilitar o início:

```bat
@echo off
cd /d "%~dp0"
streamlit run app.py
pause
```

---

## 🔒 Segurança e Privacidade

- ✅ **Dados Sensíveis**: O arquivo `.env` está no `.gitignore` e nunca será versionado
- ✅ **Planilhas Públicas**: As planilhas atuais são públicas, não exigindo autenticação
- ✅ **Cache de Dados**: Os dados são cacheados por 1 hora para melhorar performance
- ✅ **Retry Automático**: Downloads com até 3 tentativas automáticas em caso de falha
- ✅ **Tratamento de Erros**: Mensagens claras em caso de problemas de conexão

---

## 🔄 Atualização de Dados

Os dados são atualizados automaticamente a cada **1 hora** (cache do Streamlit).

Para forçar uma atualização imediata:

1. No dashboard, clique no botão **"🔄 Forçar Atualização dos Dados"** na barra lateral
2. Ou reinicie o aplicativo

---

## 🐛 Solução de Problemas

### Erro: "Nenhum dado foi carregado"

**Causa**: As planilhas não estão acessíveis ou as permissões estão incorretas.

**Solução**:
- Verifique se as 3 planilhas estão compartilhadas como públicas
- Confirme que os IDs das planilhas no `app.py` estão corretos
- Teste as URLs diretamente no navegador

### Erro: "Aba não encontrada"

**Causa**: Os nomes das abas nas planilhas não correspondem aos esperados.

**Solução**:
- Planilha principal: `BOTICARIO`, `EUDORA`, `QUEM_DISSE_BERENICE`
- Planilha de segurança: `BOT`, `EUD`, `QDB`
- Os nomes são *case-insensitive* (não diferenciam maiúsculas de minúsculas)

### Erro: "Python não é reconhecido"

**Causa**: O Python não foi instalado corretamente ou não está no PATH.

**Solução**:
- Execute o `instalar.bat` como administrador
- Reinicie o computador após a instalação
- Verifique se o Python está em: `C:\Program Files\Python311\`

### Erro: "Módulo não encontrado"

**Causa**: As dependências não foram instaladas.

**Solução**:
```bash
pip install -r requirements.txt
```

### Dashboard não abre no navegador

**Causa**: Porta 8501 pode estar em uso.

**Solução**:
```bash
streamlit run app.py --server.port 8502
```

### Logos não aparecem

**Causa**: Os arquivos de imagem não estão na raiz do projeto.

**Solução**:
- Certifique-se de que `logo_cp_fani.png`, `logo_boticario.png`, `logo_eudora.png` e `logo_qdb.png` estejam na pasta `Desktop\dashboard-estoque-cpfani\`

---

## 🤝 Como Contribuir

Contribuições são bem-vindas! Para contribuir:

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/NovaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona NovaFeature'`)
4. Push para a branch (`git push origin feature/NovaFeature`)
5. Abra um Pull Request

---

## 📝 Licença

Este projeto é proprietário e confidencial. Todos os direitos reservados à CPFANI.

---

## 📞 Suporte

Para dúvidas, problemas ou sugestões:

- **Desenvolvedor**: Alex Paulo
- **Projeto**: Dashboard de Performance de Estoque NSF
- **Versão**: 2.0.0
- **Última Atualização**: Junho de 2026

---

## 📊 PDVs Monitorados

O dashboard monitora os seguintes 17 PDVs:

| Código | Loja |
|--------|------|
| 4842 | Metrópole |
| 5152 | Coração |
| 6105 | Assai Anchieta |
| 6106 | Direita |
| 6110 | Arouche |
| 8001 | Dom José |
| 11576 | Davó |
| 12055 | São Bento |
| 12056 | Marechal |
| 12605 | Coop |
| 12645 | Light |
| 14120 | VD SBC |
| 14353 | VD SP |
| 20371 | Luz |
| 21502 | Bem Barato |
| 23000 | Outlet |
| 23379 | Assai Piraporinha |

---

**Desenvolvido com ❤️ para otimizar a gestão de estoques**