# 📊 Dashboard de Performance de Estoque NSF - CPFANI

Painel de controle interativo para análise de estoques e ruptura de produtos das marcas **O Boticário**, **Eudora** e **Quem Disse, Berenice?** em 17 PDVs (Pontos de Venda).

Este dashboard fornece visualizações em tempo real dos níveis de estoque, identifica excessos críticos (capital parado) e produtos em risco de ruptura, permitindo tomadas de decisão rápidas e baseadas em dados.

---

## 🎯 Funcionalidades Principais

- ✅ **Análise por PDV**: Visualização individualizada para cada uma das 17 lojas
- ✅ **KPIs Financeiros**: Valor de estoque atual, estoque mínimo, capital preso e risco de ruptura
- ✅ **Gráficos Interativos**: Comparativo visual entre estoque atual e mínimo por categoria
- ✅ **Tabelas Detalhadas**: Listagem dos maiores excessos e produtos em falta
- ✅ **Dados em Tempo Real**: Leitura direta da planilha do Google Sheets com cache de 1 hora
- ✅ **Interface Premium**: Design moderno em modo escuro com visual profissional
- ✅ **Filtros Dinâmicos**: Seleção rápida de loja/PDV para análise focada

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.11+** - Linguagem principal
- **Streamlit** - Framework para dashboards web interativos
- **Pandas** - Manipulação e análise de dados
- **Plotly** - Visualizações gráficas interativas
- **OpenPyXL** - Leitura de arquivos Excel
- **Google Sheets API** - Fonte de dados em nuvem

---

## 📋 Pré-requisitos

Antes de começar, certifique-se de ter:

- Windows 10/11
- Conexão com a internet (para acessar a planilha do Google Drive)
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

## 📊 Configuração da Planilha

O dashboard está configurado para ler dados de uma planilha pública do Google Sheets.

### Planilha Atual

- **ID da Planilha**: `1PbNYsNPp6ShErx0U3Ml_dJpN-0MPwoxz`
- **URL**: https://docs.google.com/spreadsheets/d/1PbNYsNPp6ShErx0U3Ml_dJpN-0MPwoxz/edit

### Estrutura da Planilha

A planilha deve conter **3 abas** com os seguintes nomes exatos:

1. **`BOTICARIO`** - Dados de O Boticário
2. **`EUDORA`** - Dados de Eudora
3. **`QUEM_DISSE_BERENICE`** - Dados de Quem Disse, Berenice?

### Colunas Obrigatórias em Cada Aba

Cada aba deve conter as seguintes colunas (na ordem desejada):

| Coluna | Tipo | Descrição |
|--------|------|-----------|
| `PDV` | Numérico | Código do Ponto de Venda (ex: 4842, 5152, etc.) |
| `SKU` | Texto | Código único do produto |
| `Descrição` | Texto | Nome/descrição do produto |
| `Categoria` | Texto | Categoria do produto (para agrupamento no gráfico) |
| `Classe` | Texto | Classificação ABC (A, B, C ou E) |
| `Estoque Atual` | Numérico | Quantidade atual em estoque |
| `Preço tabela` | Numérico | Preço unitário do produto |

### Regras de Estoque Mínimo por Classe

O dashboard calcula automaticamente o estoque mínimo com base na classificação:

- **Classe A**: Mínimo de 15 unidades
- **Classe B**: Mínimo de 10 unidades
- **Classe C**: Mínimo de 5 unidades
- **Classe E**: Mínimo de 2 unidades

### Permissões da Planilha

Como a planilha é **pública**, não é necessária configuração de credenciais. Basta garantir que ela esteja compartilhada como "Qualquer pessoa com o link pode visualizar".

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

## 📁 Estrutura do Projeto

```text
dashboard-estoque-cpfani/
│
├── app.py                      # Arquivo principal do dashboard
├── config.py                   # Configurações de conexão com Google Sheets (para planilhas privadas)
├── requirements.txt            # Lista de dependências Python
├── instalar.bat                # Script de instalação automatizada
├── .env                        # Variáveis de ambiente (não versionado)
├── .gitignore                  # Arquivos ignorados pelo Git
├── README.md                   # Este arquivo de documentação
└── instalar.log                # Log da instalação (gerado automaticamente)
```

---

## 🔒 Segurança e Privacidade

- ✅ **Dados Sensíveis**: O arquivo `.env` (que pode conter caminhos de credenciais) está no `.gitignore` e nunca será versionado
- ✅ **Credenciais do Google**: Arquivos como `credentials.json` e `service_account.json` são ignorados pelo Git
- ✅ **Planilha Pública**: A planilha atual é pública, não exigindo autenticação
- ✅ **Cache de Dados**: Os dados são cacheados por 1 hora para melhorar performance e reduzir requisições

### Migrando para Planilha Privada

Se você decidir usar uma planilha privada no futuro:

1. Crie uma Conta de Serviço no Google Cloud Console
2. Baixe o arquivo JSON de credenciais
3. Compartilhe a planilha com o e-mail da Conta de Serviço (como Leitor ou Editor)
4. Configure o caminho do arquivo JSON no `.env`:
   ```env
   GOOGLE_CREDENTIALS_PATH=caminho/para/seu/arquivo.json
   ```
5. O `config.py` já está preparado para essa configuração

---

## 🔄 Atualização de Dados

Os dados são atualizados automaticamente a cada **1 hora** (cache do Streamlit).

Para forçar uma atualização imediata:

1. No dashboard, clique no botão **"🔄 Forçar Atualização dos Dados"** na barra lateral
2. Ou reinicie o aplicativo

---

## 🐛 Solução de Problemas

### Erro: "Nenhum dado foi carregado"

**Causa**: A planilha não está acessível ou as permissões estão incorretas.

**Solução**:
- Verifique se a planilha está compartilhada como pública
- Confirme que o ID da planilha no `app.py` está correto
- Teste a URL diretamente no navegador: `https://docs.google.com/spreadsheets/d/{ID}/export?format=xlsx`

### Erro: "Aba não encontrada"

**Causa**: Os nomes das abas na planilha não correspondem aos esperados.

**Solução**:
- Verifique se as abas estão nomeadas exatamente como: `BOTICARIO`, `EUDORA`, `QUEM_DISSE_BERENICE`
- Os nomes são *case-sensitive* (diferenciam maiúsculas de minúsculas)

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
- **Versão**: 1.0.0
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