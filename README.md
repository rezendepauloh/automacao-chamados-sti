# 🤖 Automação e Classificação de Chamados de TI (MPMS)

Este projeto consiste em uma suíte de ferramentas desenvolvidas em Python para automatizar a extração, unificação, sincronização e classificação inteligente de chamados de suporte técnico (Manutenção de TI) provenientes de múltiplas plataformas (OTRS e CitSmart), agregando módulos de gestão de impressoras (PaperCut), mapa predial, escalas de plantão, conferência de portarias, vídeos de FAQs e notificações automatizadas.

---

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

### 4. Execução Stealth (Invisível) & Orquestração Unificada
- O orquestrador (`orquestrador.py`) roda via `pythonw.exe` com a flag `CREATE_NO_WINDOW`, garantindo processamento 100% em background.
- Fluxo sequencial automático: Coleta OTRS → Coleta CitSmart → Coleta PaperCut Impressoras → Pré-processamento → Classificação por IA → Sincronização de Portarias → Verificação de Alertas de Plantão.

### 5. Gestão Centralizada & Dashboard Modularizado
- **Arquitetura Modular em Camadas (`dashboard.py`):** O dashboard principal é estruturado como um orquestrador conciso com separação completa em pastas:
  - `assets/css/styles.css`: Estilos globais e refinamentos de UI.
  - `src/components/`: Componentes reutilizáveis (cabeçalho/popover, alertas, status de logs).
  - `src/tabs/`: Módulos de páginas isolados (`chamados.py`, `plantoes.py`, `portarias.py`, `notificacoes.py`, `redistribuicao.py`, `mapas.py`, `links_faqs.py`, `fiscalizacao.py`, `impressoras.py`).
- **Painel Interativo Premium (Streamlit):** Interface gráfica web responsiva para acompanhamento dos chamados em tempo real, com ordenação inteligente de datas e filtros dinâmicos de Status, Unidade, Usuário e TAG de IA.
- **Persistência Relacional (SQLite):** Dados consolidados no banco relacional `chamados.db`.

### 6. Módulo de Doação & Redistribuição de Máquinas
- **Inventário de Movimentações:** Aba dedicada no painel para visualização e análise de equipamentos destinados a doação, redistribuição, garantia ou baixados.
- **Gráficos Temporais e KPIs:** Métricas de acompanhamento de estoque e gráficos dinâmicos de distribuição por tipo e histórico por ano.
- **Gerador de Relatórios para Chamados (Rich Text HTML):** Ferramenta integrada na barra lateral que gera automaticamente textos formatados com tabelas estilizadas em HTML (Zebra Striping).

### 7. Base de Conhecimento, FAQs & Vídeos de Tutoriais
- **Sincronização em Lote via Playwright (`faq_scraper.py`):** Bot headless com `Playwright` que varre as páginas institucionais de tutoriais do SharePoint, extrai a árvore de conteúdo em HTML limpo e armazena na tabela relacional `faqs` do SQLite.
- **Visualizador de Vídeos FAQ Local/SharePoint:** Aba de vídeos com varredura recursiva de pastas, categorização automática por subpastas, reprodução em modal e botão de execução nativa via Windows (`os.startfile`) para suporte total a codecs de celular/Teams (H.265/HEVC).
- **Sub-Navegação em Abas Stylized:** Interface com navegação por abas customizadas via CSS e sincronização dinâmica da barra lateral de filtros conforme a aba ativa.

### 8. Conferência de Portarias dos Membros da Bancada
- **Integração com a API de Atos e Normas do MPMS:** Consulta automatizada para os servidores da bancada (*Paulo Rezende, Reginaldo Bandeira e Luiz Villalba*).
- **Sanitização Unicode & HTML:** Limpeza de tags HTML (`<strong>`), acentos e hífens Unicode quebrados (`\u0096`, `\u2013`), além de deduplicação inteligente.
- **Modal de Detalhes & Download de PDF:** Visualizador completo da ementa, diário oficial e download direto do PDF do anexo (`/download/{atocod}`).

### 9. Escala de Plantões da Bancada (Matutino & Semanal)
- **Calendário Interativo FullCalendar v6:** Exibição dinâmica das escalas em modo dark glassmorphism.
- **Coleta Autônoma (`plantoes_scraper.py`):** Bot de sincronização das escalas de Plantão Matutino (PGJ) e Plantão Semanal (SIMP).

### 10. Sistema Unificado de Notificações & Alertas Inteligentes
- **Notificações de Novas Portarias:** Gera alerta automaticamente no banco sempre que uma nova portaria é identificada pelo orquestrador.
- **Lembretes Antecipados de Plantão:**
  - *Plantão Matutino*: Notificação emitida 1 dia útil antes (se o plantão for na segunda-feira, a notificação é emitida na **sexta-feira**).
  - *Plantão Semanal*: Notificação emitida na **segunda-feira** da semana do plantão SIMP.
- **Alertas visuais Toast & Badge no Header:** Notificações em balão (`st.toast`) ao abrir o sistema e contador dinâmico de pendências no menu (`🔔 Central de Notificações (3)`).
- **Central de Gerenciamento (`notificacoes.py`):** Interface para filtrar por tipo/status, marcar como lida e redirecionar direto para a página referente.

### 11. Gestão de Impressoras & Dispositivos (PaperCut)
- **Coleta e Tratamento Autônomo (`papercut_scraper.py`):** Automação via Selenium que loga no sistema de gerenciamento de impressões PaperCut, navega até as listagens de impressoras e dispositivos multifuncionais (MFDs) e efetua a exportação dos relatórios em CSV.
- **Tratamento de Encodings e Limpeza Automática:** Processamento inteligente com detecção de encodings (`latin1`, `utf-8-sig`), leitura estruturada de delimitadores (`;`) e remoção automática de temporários baixados na pasta Downloads.
- **Visualização & Filtros no Dashboard (`impressoras.py`):** Interface dedicada no painel com cards KPI em tempo real (Total de Ativos, Filas, MFDs, Status OK e Alertas/Erros), busca textual e filtros dinâmicos laterais por Tipo, Status, Localização e Modelo, além de exportação para Excel.

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
  - `python-dotenv`: Gerenciamento seguro de variáveis de ambiente.

---

## 📂 Estrutura do Projeto

```text
automated-OTRS-and-CitSmart/
├── assets/
│   └── css/styles.css                # Estilos CSS globais da aplicação
├── src/
│   ├── components/                   # Componentes reutilizáveis do frontend
│   │   ├── header.py                 # Menu popover de navegação rápida e notificações
│   │   └── status_banner.py          # Checagem de status e leitor de logs do robô
│   ├── database.py                   # Interface DAO SQLite (chamados, doações, plantões, notificações, impressoras)
│   ├── tabs/                         # Módulos isolados por página
│   │   ├── chamados.py               # Aba 1: Painel Geral de Chamados
│   │   ├── plantoes.py               # Aba 2: Escala de Plantões da Bancada & FullCalendar
│   │   ├── portarias.py              # Aba 3: Portarias & Atos dos Membros da Bancada
│   │   ├── notificacoes.py           # Aba 4: Central de Notificações
│   │   ├── redistribuicao.py         # Aba 5: Doação & Redistribuição
│   │   ├── mapas.py                  # Aba 6: Planta Baixa e Rotas Leaflet
│   │   ├── links_faqs.py             # Aba 7: FAQs, Tutoriais (SharePoint) & Vídeos FAQ
│   │   ├── fiscalizacao.py           # Aba 8: Fiscalização de Contratos SAJ
│   │   └── impressoras.py            # Aba 9: Gestão de Impressoras & Dispositivos (PaperCut)
│   ├── citsmart_scraper.py           # Bot de extração do CitSmart
│   ├── otrs_scraper.py               # Bot de extração do OTRS
│   ├── papercut_scraper.py           # Bot de extração de impressoras e dispositivos do PaperCut
│   ├── plantoes_scraper.py           # Bot de extração dos plantões PGJ e SIMP
│   ├── sync_portarias.py             # Sincronizador de portarias para o orquestrador
│   ├── sync_plantoes_alerts.py       # Verificador de alertas de plantão para o orquestrador
│   ├── unidades_scraper.py           # Raspador de unidades/promotorias
│   ├── faq_scraper.py                # Bot Playwright para FAQs do SharePoint
│   ├── preprocess_chamados.py        # Limpeza e padronização de dados
│   ├── tag_classifier.py             # Classificador de IA com NLP (spaCy + Scikit-Learn)
│   ├── sync_master.py                # Sincronizador Append-Only na Planilha Master
│   └── salvar_senha.py               # Utilitário de segurança e credenciais (Keyring)
├── tests/                            # Suíte de testes unitários automatizados
│   ├── test_app.py                   # Testes de NLP, scrapers e regras de negócio
│   └── test_tabs_and_components.py   # Testes dos módulos do dashboard
├── .env.example                      # Template de variáveis de ambiente e endpoints
├── .env                              # Variáveis de ambiente locais (não commitado)
├── dashboard.py                      # Orquestrador central do Streamlit
└── orquestrador.py                   # Script mestre do fluxo executado em background
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

3. **Configure as Variáveis de Ambiente:**
   Copie o arquivo `.env.example` para `.env` e preencha as variáveis correspondentes:
   ```bash
   cp .env.example .env
   ```

4. **Configure as credenciais criptografadas do Windows (Keyring):**
   ```bash
   venv\Scripts\python.exe src/salvar_senha.py
   ```

5. **Executar a Orquestração em Segundo Plano:**
   ```bash
   venv\Scripts\python.exe orquestrador.py
   ```

6. **Executar o Dashboard Web:**
   ```bash
   streamlit run dashboard.py
   ```

---

## 🧪 Executando Testes Unitários

Para rodar todos os testes unitários integrados da aplicação e validar os módulos refatorados:

```bash
python -m unittest discover -s tests
```
