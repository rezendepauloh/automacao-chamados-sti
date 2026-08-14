# 🤖 Automação e Classificação de Chamados de TI (MPMS)

Este projeto consiste em uma suíte de ferramentas desenvolvidas em Python para automatizar a extração, unificação, sincronização e classificação inteligente de chamados de suporte técnico (Manutenção de TI) provenientes de múltiplas plataformas (OTRS, CitSmart e Central Telefônica OXE), agregando módulos de gestão de impressoras (PaperCut), mapa predial, escalas de plantão, conferência de portarias, vídeos de FAQs, notificações automatizadas e execução remota de scripts de automação PowerShell.

---

## 🚀 Funcionalidades

### 1. Web Scraping & Automação (RPA)

- **Extração Robusta & Híbrida (Selenium + Captura XHR):** Loga nos portais de suporte (OTRS e CitSmart) e na Central Telefônica OXE via Selenium. No CitSmart e no OXE, utiliza interceptação de rede (XHR/Rede) e requisições em lote paralelo (`Promise.all`) diretamente na API interna, processando milhares de registros em segundos.
- **BeautifulSoup & Requests (unidades_scraper.py):** Arquitetura nativa ultra-leve com `requests` e `BeautifulSoup` para varredura em tempo real de todas as ~280 páginas de promotorias do site institucional, obtendo o prédio físico exato com 100% de precisão.
- **Mapeamento de IP e Localidade por SCCM (WMI):** Integração avançada via WMI (Windows Management Instrumentation) consultando silenciosamente o servidor SCCM. Descobre o IP exato da máquina do usuário na rede para mapeamento hiper-preciso da localidade física usando ranges de sub-redes CIDR.
- **Cache Inteligente Completo (Descrições, Unidades e IPs):** No OTRS, CitSmart e OXE, o robô memoriza os dados já obtidos em execuções anteriores. Isso previne centenas de consultas repetidas ao SCCM, Active Directory (LDAP) e cliques lentos no Selenium, reduzindo o tempo de resposta a praticamente zero milissegundos.
- **Autolimpeza Preventiva de Disco & Rotação de Logs:**
  - _Arquivos_: Função automatizada `cleanup_old_files` que monitora e mantém no máximo os 10 arquivos mais recentes em `01 - Dados Brutos`, `02 - Dados tratados` e `03 - Dados prontos`.
  - _Logs_: Configuração nativa via `RotatingFileHandler` que limita os arquivos de log a 5 MB e preserva apenas os 3 históricos mais recentes.
- **Unificação:** Consolida dados de sistemas legados e novos em um formato tabular padronizado.

### 2. Inteligência Artificial (NLP) e Machine Learning Contínuo

- **Classificação Automática:** Utiliza IA para ler a descrição do chamado e predizer a categoria (TAG) correta (ex: "IMPRESSORA", "REDE", "SOFTWARE").
- **Pipeline de NLP Especializado em TI:**
  - Limpeza de texto avançada com `spaCy` (remoção de stop words, pontuação).
  - Regras de negócio customizadas para preservar termos técnicos (ex: _ssd_, _memoriaram_, _enderecoip_) e numerações cruciais.
- **Arena de Algoritmos (GridSearchCV):** O sistema treina e compara múltiplos modelos (`LinearSVC`, `RandomForestClassifier`, `MultinomialNB`, `ComplementNB`) para eleger o que possui a melhor métrica de _F1-Weighted_.
- **Retreinamento Autônomo:** O sistema monitora a data de modificação da base de treino (`st_mtime`). Se novos chamados forem adicionados pelo usuário, a IA detecta a mudança e se retreina automaticamente na próxima execução.

### 3. Engenharia de Dados & Integração Segura com Excel

- **Sincronização _Append-Only_:** O sistema identifica chamados inéditos e os insere cirurgicamente no final da Planilha Master de produção, **sem sobrescrever** observações, andamentos ou edições manuais feitas pela equipe.
- **Proteção contra Trava de Leitura (SharePoint / OneDrive):** Utilização de cópias voláteis com `tempfile` e `shutil.copy2` para contornar bloqueios de arquivo em uso no Windows (`PermissionError`) durante a leitura e sincronização em segundo plano de todas as planilhas do sistema (Doações/Redistribuição, Controle de Garantias e Fiscalização de Contratos SAJ).
- **Tratamento de Anomalias:** Proteção contra vazamento de memória e erros de conversão do Pandas para o Excel (como o erro `65535` em células vazias).
- **Automação Visual Win32:** Uso nativo do COM (`pywin32`) para formatar a planilha Master (autofit de colunas, quebra de texto, pintura de linhas baseada em TAGs) de forma 100% invisível no background.

### 4. Execução Stealth (Invisível) & Arquitetura Descentralizada

- **Orquestrador Mestre de Chamados (`orquestrador.py`):** Roda via `pythonw.exe` com a flag `CREATE_NO_WINDOW`, garantindo processamento 100% em background focado estritamente na esteira de chamados de TI (Coleta OTRS → Coleta CitSmart → Pré-processamento → Classificação de TAGs por IA).
- **Arquitetura Descentralizada por Módulos:** Cada aba/módulo independente (Portarias, Impressoras, Garantias, Fiscalização, Doações, Unidades) possui seu próprio sub-robô worker assíncrono disparado diretamente pela interface do Streamlit, permitindo livre uso do painel enquanto a coleta executa e exibindo logs em tempo real.

### 5. Gestão Centralizada, Dashboard Modular & Navegação por URL

- **Arquitetura Modular em Camadas (`dashboard.py`):** O dashboard principal é estruturado como um orquestrador conciso com separação completa em pastas:
  - `assets/css/styles.css`: Estilos globais e refinamentos de UI (com popover responsivo auto-ajustável de altura).
  - `src/components/`: Componentes reutilizáveis (cabeçalho/popover, alertas, status de logs, paginação `pagination.py`).
  - `src/tabs/`: Módulos de páginas isolados (`chamados.py`, `central_telefonica.py`, `plantoes.py`, `portarias.py`, `notificacoes.py`, `redistribuicao.py`, `garantia.py`, `mapas.py`, `links_faqs.py`, `fiscalizacao.py`, `impressoras.py`, `scripts_automacao.py`).

- **Navegação Persistente por URL (Query Parameters):** Sincronização bidirecional completa da página ativa e sub-abas via URL (`?tab=slug&subtab=slug`). Suporta F5 e compartilhamento de links diretos sem perder o foco do trabalho.
- **Painel Interativo Premium (Streamlit):** Interface gráfica web responsiva para acompanhamento dos chamados em tempo real, com ordenação inteligente de datas e filtros dinâmicos de Status, Unidade, Usuário e TAG de IA.
- **Persistência Relacional (SQLite):** Dados consolidados no banco relacional `chamados.db`.

### 6. Módulo da Central Telefônica (OXE Alcatel-Lucent)

- **Extração Acelerada via API REST (`oxe_scraper.py`):** Requisições em lote paralelo (`Promise.all`) via navegador que reduzem a coleta de 1.831 ramais de ~10 minutos para menos de 10 segundos. Captura todos os atributos de `Subscriber` e `Tsc_IP_subscriber` (Grupo de Captura, Categoria Rede Pública, Centro de Custo, IP, MAC Address).
- **Pré-processamento & Classificação (`preprocess_oxe.py`):** Normaliza a lista, formata o Nome Exibido, padroniza MAC Address e classifica a categoria do dispositivo (_Telefone IP Físico_, _Softphone_, _Analógico_, _Virtual_), persistindo na tabela relacional `central_telefonica` do SQLite.
- **Interface e Modal Ficha Técnica (`central_telefonica.py`):** Cards de métricas KPI, busca textual, filtros por Categoria e Tipo de Estação, além de modal interativo (`@st.dialog`) acionado pela seleção da linha com expansores categorizados e visualização do dicionário JSON da API.

### 7. Módulo de Doação & Redistribuição de Máquinas

- **Inventário de Movimentações:** Aba dedicada no painel para visualização e análise de equipamentos destinados a doação, redistribuição, garantia ou baixados.
- **Gráficos Temporais e KPIs:** Métricas de acompanhamento de estoque e gráficos dinâmicos de distribuição por tipo e histórico por ano.
- **Gerador de Relatórios para Chamados (Rich Text HTML):** Ferramenta integrada na barra lateral que gera automaticamente textos formatados com tabelas estilizadas em HTML (Zebra Striping).

### 8. Base de Conhecimento, FAQs, Vídeos & Galeria de Imagens FAQ

- **Sincronização em Lote via Playwright (`faq_scraper.py`):** Bot headless com `Playwright` que varre as páginas institucionais de tutoriais do SharePoint, extrai a árvore de conteúdo em HTML limpo e armazena na tabela relacional `faqs` do SQLite.
- **Visualizador de Vídeos FAQ Local/SharePoint:** Aba de vídeos com varredura recursiva de pastas, categorização automática por subpastas, reprodução em modal e botão de execução nativa via Windows (`os.startfile`) para suporte total a codecs de celular/Teams (H.265/HEVC).
- **Galeria de Imagens FAQ (Tutoriais):** Varredura recursiva em diretórios sincronizados por SharePoint/OneDrive (`.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.webp`) com agrupamento automático por Pasta/Tutorial. Possui visualizador em modal interativo (`@st.dialog`) em formato de carrossel de fotos (navegação `⬅️`/`➡️` persistente sem fechar o modal, limite visual de altura `max-height: 480px` com `object-fit: contain` e atalho para abertura no visualizador nativo do Windows via `os.startfile`).
- **Sub-Navegação Sincronizada por URL:** Interface com sub-abas sincronizadas por query parameters (`?tab=faq&subtab=sharepoint|videos|imagens|links`).

### 9. Conferência de Portarias dos Membros da Bancada

- **Integração com a API de Atos e Normas do MPMS:** Consulta automatizada para os servidores da bancada (_Paulo Rezende, Reginaldo Bandeira e Luiz Villalba_).
- **Sanitização Unicode & HTML:** Limpeza de tags HTML (`<strong>`), acentos e hífens Unicode quebrados (`\u0096`, `\u2013`), além de deduplicação inteligente.
- **Modal de Detalhes & Download de PDF:** Visualizador completo da ementa, diário oficial e download direto do PDF do anexo (`/download/{atocod}`).

### 10. Escala de Plantões da Bancada (Matutino & Semanal)

- **Calendário Interativo FullCalendar v6:** Exibição dinâmica das escalas em modo dark glassmorphism.
- **Coleta Autônoma (`plantoes_scraper.py`):** Bot de sincronização das escalas de Plantão Matutino (PGJ) e Plantão Semanal (SIMP).
- **Sub-Navegação Persistente:** Suporte a query parameters (`?tab=plantoes&subtab=agenda|matutino|semanal`).

### 11. Sistema Unificado de Notificações, Alertas & Paginação

- **Notificações de Novas Portarias:** Gera alerta automaticamente no banco sempre que uma nova portaria é identificada pelo orquestrador.
- **Lembretes Antecipados de Plantão:**
  - _Plantão Matutino_: Notificação emitida 1 dia útil antes (se o plantão for na segunda-feira, a notificação é emitida na **sexta-feira**).
  - _Plantão Semanal_: Notificação emitida na **segunda-feira** da semana do plantão SIMP.
- **Alertas visuais Toast & Badge no Header:** Notificações em balão (`st.toast`) ao abrir o sistema e contador dinâmico de pendências no menu (`🔔 Central de Notificações (3)`).
- **Central de Gerenciamento & Paginação (`notificacoes.py` + `pagination.py`):** Interface com busca/filtros por tipo e status, seletor dinâmico de itens por página no sidebar (`5, 10, 20, 50, 100`) e régua de navegação no rodapé.

### 12. Gestão de Impressoras & Dispositivos (PaperCut)

- **Coleta e Tratamento Autônomo (`papercut_scraper.py`):** Automação via Selenium que loga no sistema de gerenciamento de impressões PaperCut, navega até as listagens de impressoras e dispositivos multifuncionais (MFDs) e efetua a exportação dos relatórios em CSV.
- **Unificação Inteligente 360° & Desduplicação por Nome Canônico:** Fusão automática dos relatórios `printer_lists.csv` (Filas) e `devices_lists.csv` (MFDs) combinando IP de rede do dispositivo físico, nome do servidor de impressão, localização e modelo detalhado do equipamento em um registro único sem duplicatas.
- **Diagnóstico de Conectividade em Tempo Real (Ping Integrado):** Teste nativo de conectividade ICMP (`ping_host`) disponível diretamente no modal de detalhes da impressora (`@st.dialog`), exibindo alertas Toast e o log detalhado de saída do terminal em um accordion expansível.
- **Tratamento de Encodings e Limpeza Automática:** Processamento inteligente com detecção de encodings (`latin1`, `utf-8-sig`), leitura estruturada de delimitadores (`;`), rotação de 10 backups e remoção de temporários em Downloads.
- **Visualização & Filtros no Dashboard (`impressoras.py`):** Interface dedicada no painel com cards KPI em tempo real (Total de Ativos, Filas, MFDs, Status OK e Alertas/Erros), busca textual e filtros dinâmicos por Tipo, Status, Localização e Modelo.

### 13. Módulo de Scripts de Automação PowerShell (Background Worker)

- **Execução Remota de Rotinas de TI (`scripts_automacao.py`):**
  - _Analisador de Dispositivos_: Coleta inventário de hardware, BIOS, discos, drivers e programas de máquinas remotas via CIM/WSMan, gerando relatórios em HTML, PDF e Excel.
  - _Manutenção e Limpeza Remota_: Limpeza remota de temporários, Prefetch, Lixeira, cache de atualizações, Windows.old, Delivery Optimization, Crash Dumps e otimização/defrag de disco.
  - _Remoção de Perfis de Usuário_: Purga remota de contas e pastas de usuários inativos (`C:\Users`) via `Win32_UserProfile` e `StdRegProv`.
- **Background Task Persistence (Execução Assíncrona):** Dispara o script em uma thread/processo em segundo plano desacoplada do navegador. Permite que o usuário navegue por outras abas do dashboard ou pressione **F5** sem interromper o script no Windows.
- **Detecção Dinâmica do PowerShell Engine:** Detecta automaticamente a presença do `pwsh.exe` (PowerShell Core 7+) no sistema; caso contrário, faz o fallback seguro para o `powershell.exe` (Windows PowerShell 5.1).
- **Auto-Fix de Credenciais DPAPI (`cred_admin.xml`):** Identifica falhas de criptografia DPAPI e regenera automaticamente os arquivos de credenciais usando as credenciais administrativas do `SCCM_ADMIN_USER` salvas no Keyring do Windows.

### 14. Catálogo Unificado de Unidades, Ramais (PDF / Intranet) & Monitoramento de Robôs

- **Extrator de Ramais Telefônicos da Intranet (`ramais_scraper.py`):** Autenticação automatizada via `requests.Session()` com credenciais do sistema (`USERNAME` / `PASSWORD`), busca dinâmica das URLs e download em memória dos PDFs oficiais de ramais (Comarcas do Interior e Capital/PGJ).
- **Processamento Inteligente de PDF (`pdfplumber`):** Varredura linha a linha e extração estruturada de tabelas relacionando comarcas, prédios, setores, membros e seus respectivos números telefônicos na tabela relacional `ramais_mpms` do SQLite.
- **Relação Unificada com Gestão Interativa (`unidades.py`):** Consolidação da lista oficial do Portal Web com as Unidades Manuais locais (badges de origem `📌 Manual` e `🌐 Portal Web`).
- **Modal Interativo (`@st.dialog`) & Edição por Seleção de Linha:** Ao clicar em qualquer linha da tabela (`on_select="rerun"`):
  - _Registro Manual_: Abre formulário preenchido permitindo edição em tempo real de todos os atributos com salvamento ou exclusão direta.
  - _Registro do Portal_: Exibe a ficha completa formatada em modo de leitura.
- **Acompanhamento de Robôs em Segundo Plano (Accordions & Logs):** Indicadores no sidebar com botões desabilitados durante a execução, acompanhamento do progresso através de `st.expander` com leitor de logs em tempo real e notificação em balão `st.toast` ao concluir.

### 15. Componentes Globais Reutilizáveis (Subtabs & Calendário Master)

### 15. Componentes Globais Reutilizáveis (Subtabs & Calendário Master)

- **Sub-Navegação por Abas Nativas (`src/components/subtabs.py`):** Componente padronizado com isolamento CSS que simula abas nativas para rádios do Streamlit, garantindo sincronização imediata dos estados com os query parameters da URL (`?subtab=slug`).
- **Motor Centralizado de Calendário Master (`src/components/calendar.py`):** Função `render_master_calendar` que encapsula o FullCalendar v6 com modal dinâmico inteligente, adaptação automática de temas claro/escuro (incluindo o popover do "+X mais"), estilização vermelha `#ff4b4b` para abas ativas e exibição completa de chamados técnicos.
- **Fechamento Automático de Chamados Ausentes (`close_missing_tickets_by_base`):** Mecanismo de sincronização relacional no SQLite que identifica chamados encerrados nos portais de origem e atualiza seu status para `'Fechado'`, com trava de segurança por volume mínimo (`active_ids >= 3`).
- **Conformidade com a API Moderna do Streamlit:** Migração global de parâmetros legados de largura para `width='stretch'` e componentes de HTML customizados para `st.components.v1.html(...)`.

---

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3.11+ & PowerShell 5.1+ / 7+ (pwsh)
- **Bibliotecas Principais:**
  - `selenium`: Navegação web automatizada para logins e sistemas dinâmicos.
  - `playwright`: Raspagem em lote em alta performance de tutoriais e artigos do SharePoint em headless mode.
  - `pdfplumber`: Leitura profunda e extração tabular de documentos PDF em memória.
  - `requests` & `beautifulsoup4`: Varredura estática ultra-veloz de portais institucionais públicos e sanitização de HTML.
  - `pandas`: Análise, manipulação e alinhamento inteligente de DataFrames.
  - `scikit-learn`: Treinamento pesado, tuning de hiperparâmetros e classificação.
  - `spacy`: Processamento de linguagem natural (NLP) e lematização.
  - `pywin32`, `keyring` & `WMI`: Automação nativa do Microsoft Excel, cofre de senhas do Windows e consultas profundas ao servidor SCCM/CIM.
  - `streamlit`: Construção do dashboard web moderno, rápido e interativo.
  - `sqlite3`: Banco de dados relacional embutido de altíssimo desempenho.
  - `python-dotenv`: Gerenciamento seguro de variáveis de ambiente.

---

## 📂 Estrutura do Projeto

```text
automated-OTRS-and-CitSmart/
├── assets/
│   └── css/styles.css                # Estilos CSS globais da aplicação (com ajuste responsivo de menus e abas subtabs)
├── debug_logs/
│   ├── citsmart/                     # Logs e screenshots do CitSmart
│   ├── otrs/                         # Logs e screenshots do OTRS
│   ├── oxe/                          # Logs e screenshots da Central Telefônica OXE
│   ├── papercut/                     # Logs do PaperCut
│   ├── plantoes/                     # Logs de escalas de plantão
│   ├── preprocessamento/             # Logs de tratamento de dados
│   └── scripts/                      # Logs centralizados de execução dos scripts de automação PowerShell
├── debug/                            # Logs de execução em background dos scrapers (unidades, ramais)
├── src/
│   ├── database/                     # Camada modular de banco de dados SQLite
│   │   ├── __init__.py               # Reexportação pública unificada da DAO
│   │   ├── connection.py             # Conexão central com chamados.db
│   │   ├── tickets_db.py            # Operações de chamados e fechamento por base (close_missing_tickets_by_base)
│   │   ├── ramais_db.py             # Tabela e consultas de ramais telefônicos
│   │   ├── map_db.py                # Tabelas do mapa predial, salas e rotas
│   │   ├── plantoes_db.py           # Tabelas e persistência de escalas de plantão
│   │   ├── notifications_db.py     # Gerenciamento de alertas e notificações
│   │   ├── impressoras_db.py       # Tabela de dispositivos do PaperCut
│   │   ├── central_db.py           # Tabela de ramais do OXE Alcatel
│   │   ├── garantia_db.py          # Tabela de controle de garantias
│   │   ├── events_db.py            # Tabela de eventos manuais do calendário
│   │   ├── unidades_db.py          # Tabela de unidades e promotorias
│   │   ├── fiscalizacao_db.py      # Tabela de fiscalização de contratos SAJ
│   │   └── donations_db.py         # Tabela de doações e redistribuição
│   ├── scrapers/                     # Scrapers e bots automatizados de coleta
│   │   ├── citsmart_scraper.py       # Bot de extração do CitSmart
│   │   ├── otrs_scraper.py           # Bot de extração do OTRS
│   │   ├── oxe_scraper.py            # Bot de extração em lote paralelo (Promise.all) do OXE
│   │   ├── papercut_scraper.py       # Bot de extração de impressoras do PaperCut
│   │   ├── plantoes_scraper.py       # Bot de extração dos plantões PGJ e SIMP
│   │   ├── ramais_scraper.py         # Extrator PDF de ramais da Intranet
│   │   ├── unidades_scraper.py       # Raspador de unidades no portal do MPMS
│   │   └── faq_scraper.py            # Bot Playwright para FAQs do SharePoint
│   ├── syncs/                        # Workers assíncronos de sincronização em segundo plano
│   │   ├── sync_donations.py         # Sincronizador da planilha de doações/redistribuição
│   │   ├── sync_fiscalizacao.py      # Sincronizador da planilha de fiscalização SAJ
│   │   ├── sync_garantia.py          # Sincronizador da planilha de garantias
│   │   ├── sync_master.py            # Sincronizador Append-Only na Planilha Master
│   │   ├── sync_portarias.py         # Sincronizador da API de portarias
│   │   └── sync_plantoes_alerts.py   # Verificador de alertas de plantão
│   ├── components/                   # Componentes reutilizáveis do frontend
│   │   ├── header.py                 # Menu popover de navegação rápida com query parameters e notificações
│   │   ├── subtabs.py                # Componente de sub-abas sincronizadas com query parameters na URL
│   │   ├── calendar.py               # Componente mestre FullCalendar v6 com modal dinâmico inteligente
│   │   ├── pagination.py             # Componente de paginação e seletor de registros por página
│   │   └── status_banner.py          # Checagem de status e leitor de logs do robô
│   ├── tabs/                         # Módulos isolados por página
│   │   ├── chamados.py               # Aba 1: Painel Geral de Chamados (?tab=chamados&subtab=...)
│   │   ├── central_telefonica.py     # Aba 2: Central Telefônica OXE & Modal Ficha Técnica (?tab=central-telefonica)
│   │   ├── plantoes.py               # Aba 3: Escala de Plantões da Bancada & FullCalendar (?tab=plantoes&subtab=...)
│   │   ├── portarias.py              # Aba 4: Portarias & Atos dos Membros da Bancada (?tab=portarias)
│   │   ├── notificacoes.py           # Aba 5: Central de Notificações com Paginação (?tab=notificacoes)
│   │   ├── redistribuicao.py         # Aba 6: Doação & Redistribuição (?tab=redistribuicao)
│   │   ├── mapas.py                  # Aba 7: Planta Baixa e Rotas Leaflet (?tab=mapa)
│   │   ├── links_faqs.py             # Aba 8: FAQs, Tutoriais (SharePoint), Vídeos FAQ & Galeria de Imagens FAQ (?tab=faq&subtab=...)
│   │   ├── fiscalizacao.py           # Aba 9: Fiscalização de Contratos SAJ (?tab=fiscalizacao&subtab=...)
│   │   ├── impressoras.py            # Aba 10: Gestão de Impressoras & Dispositivos PaperCut (?tab=impressoras&subtab=...)
│   │   ├── scripts_automacao.py      # Aba 11: Scripts de Automação PowerShell (?tab=scripts-automacao&subtab=...)
│   │   ├── calendario_geral.py       # Aba 12: Calendário Geral Unificado (?tab=calendario-geral)
│   │   └── unidades.py               # Aba 13: Catálogo Unificado de Unidades & Lista de Ramais Telefônicos (?tab=unidades&subtab=...)
│   ├── preprocess_chamados.py        # Limpeza e padronização de chamados
│   ├── preprocess_oxe.py             # Pré-processamento e classificação de ramais do OXE
│   ├── tag_classifier.py             # Classificador de IA com NLP (spaCy + Scikit-Learn)
│   └── salvar_senha.py               # Utilitário de segurança e credenciais (Keyring)
├── tests/                            # Suíte de testes unitários automatizados
│   ├── test_app.py                   # Testes de NLP, scrapers, banco e fechamento por base
│   └── test_tabs_and_components.py   # Testes dos módulos do dashboard
├── .env.example                      # Template de variáveis de ambiente e caminhos dos scripts
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
   Copie o arquivo `.env.example` para `.env` e preencha as variáveis e caminhos dos scripts PowerShell correspondentes:

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
