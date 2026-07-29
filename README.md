# 🤖 Automação e Classificação de Chamados de TI (MPMS)

Este projeto consiste em uma suíte de ferramentas desenvolvidas em Python para automatizar a extração, unificação, sincronização e classificação inteligente de chamados de suporte técnico (Manutenção de TI) provenientes de múltiplas plataformas (OTRS e CitSmart).

## 🚀 Funcionalidades

### 1. Web Scraping & Automação (RPA)
- **Extração Robusta & Híbrida (Selenium + Captura XHR):** Loga nos portais de suporte (OTRS e CitSmart) via Selenium. No CitSmart, utiliza interceptação de rede (XHR/Rede) para capturar a listagem de chamados em JSON diretamente da API interna do portal, processando as informações instantaneamente com velocidade imbatível.
- **BeautifulSoup & Requests (unidades_scraper.py):** Arquitetura nativa ultra-leve com `requests` e `BeautifulSoup` para varredura em tempo real de todas as ~280 páginas de promotorias do site institucional, obtendo o prédio físico exato com 100% de precisão.
- **Mapeamento de IP e Localidade por SCCM (WMI):** Integração avançada via WMI (Windows Management Instrumentation) consultando silenciosamente o servidor SCCM. Descobre o IP exato da máquina do usuário na rede para mapeamento hiper-preciso da localidade física usando ranges de sub-redes CIDR.
- **Cache Inteligente Completo (Descrições, Unidades e IPs):** No OTRS e no CitSmart, o robô memoriza os dados já obtidos em execuções anteriores. Isso previne centenas de consultas repetidas ao SCCM, Active Directory (LDAP) e cliques lentos no Selenium, reduzindo o tempo de resposta em chamados recorrentes a praticamente zero milissegundos.
- **Autolimpeza Preventiva de Disco:** Função automatizada que monitora e mantém no máximo os 10 arquivos mais recentes de cada etapa, evitando acúmulo de logs e planilhas mortas.
- **Unificação:** Consolida dados de sistemas legados e novos em um formato tabular padronizado.

### 2. Inteligência Artificial (NLP) e Machine Learning Contínuo
- **Classificação Automática:** Utiliza IA para ler a descrição do chamado e predizer a categoria (TAG) correta (ex: "IMPRESSORA", "REDE", "SOFTWARE").
- **Pipeline de NLP Especializado em TI:**
  - Limpeza de texto avançada com `spaCy` (remoção de stop words, pontuação).
  - Regras de negócio customizadas para preservar termos técnicos (ex: *ssd*, *memoriaram*, *enderecoip*) e numerações cruciais.
- **Arena de Algoritmos (GridSearchCV):** O sistema treina e compara múltiplos modelos (`LinearSVC`, `RandomForestClassifier`, `MultinomialNB`, `ComplementNB`) para eleger o que possui a melhor métrica de *F1-Weighted*.
- **Retreinamento Autônomo:** O sistema monitora a data de modificação da base de treino (`st_mtime`). Se novos chamados forem adicionados pelo usuário, a IA detecta a mudança e se retreina automaticamente na próxima execução.

### 3. Engenharia de Dados & Integração Segura com Excel
- **Sincronização *Append-Only*:** O sistema identifica chamados inéditos e os insere cirurgicamente no final da Planilha Master de produção, **sem sobrescrever** observações, andamentos ou edições manuais feitas pela equipe.
- **Tratamento de Anomalias:** Proteção contra vazamento de memória e erros de conversão do Pandas para o Excel (como o erro `65535` em células vazias).
- **Automação Visual Win32:** Uso nativo do COM (`pywin32`) para formatar a planilha Master (autofit de colunas, quebra de texto, pintura de linhas baseada em TAGs) de forma 100% invisível no background.

### 4. Execução Stealth (Invisível)
- O orquestrador roda via `pythonw.exe` com a flag `CREATE_NO_WINDOW`, garantindo processamento 100% em background, sem roubar o foco do usuário e sem disparar alertas indesejados.

### 5. Gestão Centralizada & Dashboard Modularizado
- **Arquitetura Modular em Camadas (`dashboard.py`):** O dashboard principal é estruturado como um orquestrador conciso (< 70 linhas) com separação completa em pastas:
  - `assets/css/styles.css`: Estilos globais e refinamentos de UI.
  - `src/components/`: Componentes reutilizáveis (cabeçalho/popover, status de logs).
  - `src/tabs/`: Módulos de páginas isolados (`chamados.py`, `redistribuicao.py`, `mapas.py`, `links_faqs.py`, `fiscalizacao.py`).
- **Painel Interativo Premium (Streamlit):** Interface gráfica web responsiva para acompanhamento dos chamados em tempo real, com ordenação inteligente de datas e filtros dinâmicos de Status, Unidade, Usuário e TAG de IA.
- **Deep Linking Centralizado:** Geração dinâmica de URLs diretas para os chamados tanto no OTRS quanto no CitSmart.
- **Persistência Inteligente (SQLite):** Todo o tráfego de dados gerado é consolidado num banco de dados local leve e ultra-rápido (`chamados.db`).

### 6. Módulo de Doação & Redistribuição de Máquinas
- **Inventário de Movimentações:** Aba dedicada no painel para visualização e análise de equipamentos destinados a doação, redistribuição, garantia ou baixados.
- **Gráficos Temporais e KPIs:** Métricas de acompanhamento de estoque e gráficos dinâmicos de distribuição por tipo e histórico por ano.
- **Gerador de Relatórios para Chamados (Rich Text HTML):** Ferramenta integrada na barra lateral que gera automaticamente textos formatados com tabelas estilizadas em HTML (Zebra Striping) a partir das movimentações de uma data específica.

### 7. Base de Conhecimento & FAQs do SharePoint
- **Sincronização em Lote via Playwright (`faq_scraper.py`):** Bot automatizado com navegação em modo headless via `Playwright` que varre as páginas institucionais de tutoriais do SharePoint, extrai a árvore de conteúdo em HTML limpo e armazena na tabela relacional `faqs` do SQLite.
- **Leitor Interativo e Modal no Dashboard (`st.dialog`):** Aba dedicada de FAQs e Tutoriais no dashboard com busca inteligente por palavra-chave e categorias.

### 8. Visualizador & Fiscalização de Contratos (Processos SAJ)
- **Integração em Tempo Real com OneDrive/SharePoint:** Leitura automatizada da planilha oficial de indicações de fiscais, portarias e processos SAJ sincronizada na nuvem.
- **Cards KPI & Resumo por Fiscal:** Indicadores de topo com contagem detalhada da carga de trabalho dos fiscais.
- **Filtros Dinâmicos e Exportação:** Filtro por fiscal responsável, busca textual e exportação filtrada para Excel.

---

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3.11+
- **Bibliotecas Principais:**
  - `selenium`: Navegação web automatizada para logins e sistemas dinâmicos.
  - `playwright`: Raspagem em lote em alta performance de tutoriais e artigos do SharePoint em headless mode.
  - `requests` & `beautifulsoup4`: Varredura estática ultra-veloz de portais institucionais públicos e sanitização de HTML.
  - `pandas`: Análise, manipulação e alinhamento inteligente de DataFrames.
  - `scikit-learn`: Treinamento pesado, tuning de hiperparâmetros e classificação.
  - `spacy`: Processamento de linguagem natural (NLP) e lematização.
  - `pywin32` e `WMI`: Automação nativa do Microsoft Excel e consultas profundas ao servidor SCCM para extração de IPs na rede.
  - `streamlit`: Construção do dashboard web moderno, rápido e interativo.
  - `sqlite3`: Banco de dados relacional embutido de altíssimo desempenho.

---

## 📂 Estrutura do Projeto

```text
automated-OTRS-and-CitSmart/
├── assets/
│   └── css/styles.css                # Estilos CSS globais da aplicação
├── src/
│   ├── components/                   # Componentes reutilizáveis do frontend
│   │   ├── header.py                 # Menu popover de navegação rápida
│   │   └── status_banner.py          # Checagem de status e leitor de logs do robô
│   ├── database.py                   # Interface DAO SQLite (UPSERTs, chamados, doações)
│   ├── tabs/                         # Módulos isolados por sistema/página
│   │   ├── chamados.py               # Aba 1: Painel Geral de Chamados
│   │   ├── redistribuicao.py         # Aba 2: Doação & Redistribuição
│   │   ├── mapas.py                  # Aba 3: Planta Baixa e Rotas Leaflet
│   │   ├── links_faqs.py             # Aba 4: FAQs, Tutoriais & Links Úteis
│   │   └── fiscalizacao.py           # Aba 5: Fiscalização de Contratos SAJ
│   ├── citsmart_scraper.py           # Bot de extração do CitSmart
│   ├── otrs_scraper.py               # Bot de extração do OTRS
│   ├── unidades_scraper.py           # Raspador de unidades/promotorias
│   ├── faq_scraper.py                # Bot Playwright para FAQs do SharePoint
│   ├── preprocess_chamados.py        # Limpeza e padronização de dados
│   ├── tag_classifier.py             # Classificador de IA com NLP (spaCy + Scikit-Learn)
│   └── sync_master.py                # Sincronizador Append-Only na Planilha Master
├── tests/                            # Suíte de testes unitários automatizados
│   ├── test_app.py                   # Testes de NLP, scrapers e regras de negócio
│   └── test_tabs_and_components.py   # Testes dos novos módulos do dashboard
├── dashboard.py                      # Orquestrador central conciso do Streamlit (< 70 linhas)
├── orquestrador.py                   # Script mestre do fluxo executado em background
└── salvar_senha.py                   # Utilitário de segurança e credenciais
```

---

## 📦 Como Instalar e Configurar

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/rezendepauloh/automacao-chamados-sti
   ```

2. **Instale os requisitos:**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   playwright install chromium
   python -m spacy download pt_core_news_sm
   ```

3. **Configure as credenciais criptografadas:**
   ```bash
   venv\Scripts\python.exe salvar_senha.py
   ```

4. **Executar o Dashboard Web:**
   ```bash
   streamlit run dashboard.py
   ```

---

## 🧪 Executando Testes Unitários

Para rodar todos os testes unitários integrados da aplicação e validar os módulos refatorados:

```bash
python -m unittest discover -s tests
```
