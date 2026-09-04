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
  - `src/tabs/`: Módulos de páginas isolados (`chamados.py`, `unidades.py`, `central_telefonica.py`, `plantoes.py`, `calendario_geral.py`, `mapas.py`, `redistribuicao.py`, `fiscalizacao.py`, `viagens.py`, `garantia.py`, `links_faqs.py`, `portarias.py`, `impressoras.py`, `scripts_automacao.py`, `notificacoes.py`, `configuracoes.py`).

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

- **Integração com a API de Atos e Normas do MPMS:** Consulta automatizada para os servidores da bancada.
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

### 13. Módulo de Scripts de Automação PowerShell & Relatórios Interativos

- **Execução Remota de Rotinas de TI (`scripts_automacao.py`):**
  - _Analisador de Dispositivos_: Coleta inventário de hardware, BIOS, discos, drivers e programas de máquinas remotas via CIM/WSMan, gerando relatórios ricos em HTML, PDF e Excel.
  - _Manutenção e Limpeza Remota_: Limpeza remota de temporários, Prefetch, Lixeira, cache de atualizações, Windows.old, Delivery Optimization, Crash Dumps e otimização/defrag de disco.
  - _Remoção de Perfis de Usuário_: Purga remota de contas e pastas de usuários inativos (`C:\Users`) via `Win32_UserProfile` e `StdRegProv`.
- **Relatórios HTML Dinâmicos e Interativos (`GeradorHtml.ps1`):**
  - _Design Moderno e Responsivo_: Identificação visual completa no cabeçalho com modelo e nome da máquina.
  - _Ordenação Interativa de Tabelas (`<th>`)_: Clique em qualquer coluna para ordenar dados numéricos, datas (formato pt-BR) e texto.
  - _Filtros de Pesquisa por Tabela (`.table-filter`)_: Inputs de busca individuais para seções de alta densidade (Programas, Serviços, Processos e Drivers).
  - _ScrollSpy & Sub-links de Gráficos_: Menu lateral com classe `.active` dinâmica e links diretos para seções de tabelas e para cada gráfico individual.
  - _Seções Accordion Expansíveis (`.section-card`)_: Todos os cards podem ser recolhidos ou expandidos ao clicar no título ou no ícone `▼`.
  - _Scrollbars Personalizadas_: Estilização elegante de barras de rolagem para Chrome, Edge, Safari (`::-webkit-scrollbar`) e Firefox (`scrollbar-width`).
  - _Gráficos Interativos_: Visualizações completas com Chart.js (slots de RAM, discos, espaço por perfil, programas por ano, tipo de inicialização de serviços, consumo de RAM por processo e origem de drivers) e fallback offline em CSS.
- **Download Direto para a Máquina do Usuário (Pronto para Docker):**
  - Botões de download instantâneo para relatórios **HTML (`.html`)**, **PDF (`.pdf`)** e **Excel (`.xlsx`)** na interface web Streamlit, permitindo que o usuário salve os arquivos diretamente na sua máquina local mesmo se a aplicação rodar em um contêiner Docker/Linux no Red Hat.
- **Background Task Persistence (Execução Assíncrona):** Dispara o script em uma thread/processo em segundo plano desacoplada do navegador. Permite que o usuário navegue por outras abas do dashboard ou pressione **F5** sem interromper o script no Windows.
- **Detecção Dinâmica do PowerShell Engine:** Detecta automaticamente a presença do `pwsh.exe` (PowerShell Core 7+) no sistema; caso contrário, faz o fallback seguro para o `powershell.exe` (Windows PowerShell 5.1).
- **Auto-Fix de Credenciais DPAPI (`cred_admin.xml`):** Identifica falhas de criptografia DPAPI e regenera automaticamente os arquivos de credenciais usando as credenciais administrativas do `SCCM_ADMIN_USER` salvas no Keyring do Windows.


### 14. Módulo de Viagens da Bancada STI

- **Cronograma Integrado de Viagens (`viagens.py`):**
  - _Calendário Interativo (FullCalendar v6)_: Visualização das viagens da equipe por períodos (multi-dias com ajuste automático de término inclusivo), detalhando localidade, técnicos escalados, número de chamados e datas.
  - _Modal de Detalhes Dinâmico_: Exibição instantânea dos dados ao clicar no evento (Destino, Técnico, Chamado, Saída e Retorno).
  - _Tabela Filtrável e Paginada_: Listagem tabular com busca em tempo real por técnico, destino e chamados, com exportação para Excel (`.xlsx`) e seletor dinâmico de itens por página.
  - _Sincronização com o SharePoint_: Download direto HTTP e fallback para automação com Selenium ou caminho local no OneDrive corporativo, com persistência na tabela `viagens`.
  - _Integração com o Calendário Geral_: Camada de ativação/desativação dedicada (`Viagens da Bancada`) integrada com a pesquisa global unificada.

### 15. Disparador Local Windows via Protocol Handler (`bancada://`)

- **Execução Desacoplada do Navegador:**
  - Registro de protocolo URI customizado no Windows (`bancada://run?tool=...&target=...`) acionado diretamente por hiperlinks na interface web Streamlit.
  - Instalador rápido em lote (`instalar_disparador_windows.cmd`) e registro de chaves (`instalar_protocolo_bancada.reg`).
  - Lançador nativo em PowerShell (`bancada-launcher.ps1`) com elevação de privilégios (`Start-Process -Verb RunAs`) executado em janela oculta.
  - Detecção inteligente do interpretador: seleciona automaticamente o PowerShell 7+ (`pwsh.exe`) se disponível, com fallback transparente para o Windows PowerShell 5.1 (`powershell.exe`).
  - Abertura automática de relatórios gerados no navegador padrão sem bloqueios de permissão do container.

### 16. Painel Central de Configurações & Cofre de Credenciais Criptografado

- **Substituição Segura do `.env` (`configuracoes.py` & `settings_db.py`):**
  - Gerenciamento unificado de todas as variáveis do sistema diretamente pelo painel web, organizado por abas temáticas (*Rede / AD*, *SCCM*, *PaperCut*, *Telefonia OXE*, *SharePoint & Planilhas*, *Portais & URLs*, *Inteligência Artificial*).
  - Criptografia simétrica com **Fernet (AES-128)** e derivação de chave de máquina via **PBKDF2HMAC** para todos os campos sensíveis (senhas e tokens de API).
  - Fallback automático para variáveis de ambiente `.env` e Keyring, garantindo compatibilidade reversa e portabilidade para PostgreSQL / Red Hat.

### 17. Catálogo Unificado de Unidades, Ramais (PDF / Intranet) & Monitoramento de Robôs

- **Extrator de Ramais Telefônicos da Intranet (`ramais_scraper.py`):** Autenticação automatizada via `requests.Session()` com credenciais do sistema (`USERNAME` / `PASSWORD`), busca dinâmica das URLs e download em memória dos PDFs oficiais de ramais (Comarcas do Interior e Capital/PGJ).
- **Processamento Inteligente de PDF (`pdfplumber`):** Varredura linha a linha e extração estruturada de tabelas relacionando comarcas, prédios, setores, membros e seus respectivos números telefônicos na tabela relacional `ramais_mpms` do SQLite.
- **Relação Unificada com Gestão Interativa (`unidades.py`):** Consolidação da lista oficial do Portal Web com as Unidades Manuais locais (badges de origem `📌 Manual` e `🌐 Portal Web`).
- **Modal Interativo (`@st.dialog`) & Edição por Seleção de Linha:** Ao clicar em qualquer linha da tabela (`on_select="rerun"`):
  - _Registro Manual_: Abre formulário preenchido permitindo edição em tempo real de todos os atributos com salvamento ou exclusão direta.
  - _Registro do Portal_: Exibe a ficha completa formatada em modo de leitura.
- **Acompanhamento de Robôs em Segundo Plano (Accordions & Logs):** Indicadores no sidebar com botões desabilitados durante a execução, acompanhamento do progresso através de `st.expander` com leitor de logs em tempo real e notificação em balão `st.toast` ao concluir.

### 18. Componentes Globais Reutilizáveis (Subtabs & Calendário Master)

- **Sub-Navegação por Abas Nativas (`src/components/subtabs.py`):** Componente padronizado com isolamento CSS que simula abas nativas para rádios do Streamlit, garantindo sincronização imediata dos estados com os query parameters da URL (`?subtab=slug`).
- **Motor Centralizado de Calendário Master (`src/components/calendar.py`):** Função `render_master_calendar` que encapsula o FullCalendar v6 com modal dinâmico inteligente, adaptação automática de temas claro/escuro (incluindo o popover do "+X mais"), estilização vermelha `#ff4b4b` para abas ativas e exibição completa de chamados técnicos, plantões, garantias, portarias e viagens.
- **Fechamento Automático de Chamados Ausentes (`close_missing_tickets_by_base`):** Mecanismo de sincronização relacional no SQLite que identifica chamados encerrados nos portais de origem e atualiza seu status para `'Fechado'`, com trava de segurança por volume mínimo (`active_ids >= 3`).
- **Conformidade com a API Moderna do Streamlit:** Migração global de parâmetros legados de largura para `width='stretch'` e componentes de HTML customizados para `st.components.v1.html(...)`.

### 19. Notificações Inteligentes no WhatsApp (Evolution API Docker)

- **Container Dedicado da Evolution API (`docker-compose.yml`):**
  - Integração nativa com a instância `bancada_evolution_api` em Docker (porta `8080`), conectada ao PostgreSQL da bancada.
  - Conexão do telefone institucional da Bancada STI (`+55 67 98478-2034`).
- **Pareamento Visual por QR Code no Navegador:**
  - Aba de configurações `📱 WhatsApp & Alertas (Evolution)` com geração de QR Code base64 dinâmico na tela do Streamlit.
  - Status em tempo real da conexão (`🟢 Conectado`, `🟡 Aguardando Leitura`, `⚪ Desconectado`), permitindo parear o celular funcional ou desconectar com 1 clique.
- **Disparos Automáticos D-1 às 12:00 (Dias Úteis):**
  - _Regra de Dias Úteis_: Eventos de Terça a Sexta avisam no dia anterior útil ($D-1$ às 12:00); eventos de Sábado, Domingo ou Segunda são disparados na **Sexta-feira anterior às 12:00**.
  - _Plantões Matutinos e Semanais_: Avisos direcionados e personalizados pelo primeiro nome do técnico escalado.
  - _Viagens da Bancada_: Alertas com destino, datas de saída/retorno e chamado vinculado para todos os servidores da bancada que irão viajar.
  - _Novas Portarias_: Avisos instantâneos com número do ato e resumo da ementa sempre que os integrantes forem citados no Diário Oficial.
- **Fuzzy Matching Inteligente & Destinatários Exclusivos:**
  - Mapeamento flexível de nomes (`member_matcher.py`) que reconhece apelidos e abreviações das planilhas.
  - Envio restrito e seguro exclusivamente para os servidores autorizados.
  - Registro de auditoria contra envios duplicados (`whatsapp_disparos_log`).

### 20. Sistema de Agendamentos & Cron Jobs em Segundo Plano

- **Motor Autônomo em Background (`src/services/cron_scheduler.py`):**
  - Thread daemon nativa inicializada automaticamente pelo entrypoint do container (`init.py`).
  - Execução contínua sem depender do usuário estar com o navegador aberto.
  - Verificação a cada 30 segundos dos intervalos definidos no banco relacional.
- **Gestão Visual Completa (`⚙️ Configurações > ⏰ Agendamentos & Cron Jobs`):**
  - _Modo Recorrente_: Configuração de intervalos em minutos, horas ou dias (ex: Portarias a cada 2 horas, Viagens a cada 4 horas).
  - _Modo Horário Fixo Diário_: Configuração de horário exato (ex: `12:00` para os alertas WhatsApp) com filtro opcional de dias úteis.
  - _Controle Individual_: Chave liga/desliga para cada rotina de sincronização e scrapers.
  - _Disparo Imediato_: Botão **`🚀 Executar Agora`** para acionamento sob demanda em thread isolada.
- **Auditoria & Histórico (`cron_logs`):**
  - Tabela com rastreamento de início, término, duração em segundos, status (`🟢 Sucesso`, `🔴 Erro`, `⏳ Executando`) e mensagem de retorno de cada execução automática.

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
│   └── css/styles.css                # Estilos CSS globais da aplicação (com ajuste responsivo de UI)
├── debug_logs/                       # Logs organizados por módulo (otrs, citsmart, oxe, papercut, plantoes, viagens, etc.)
├── models/                           # Modelos de Machine Learning e classificadores treinados serializados (.joblib)
├── uploads/                          # Diretório de arquivos enviados, anexos e mídias temporárias
├── 01 - Dados Brutos/                # Planilhas e relatórios brutos baixados pelos robôs
├── 02 - Dados tratados/              # Dados intermediários limpos e normalizados
├── 03 - Dados prontos/               # Bases consolidadas prontas para consumo
├── src/
│   ├── components/                   # Componentes reutilizáveis do frontend (header, subtabs, calendar, pagination, status_banner)
│   ├── database/                     # Camada modular de banco de dados SQLite/Postgres (12 módulos relacionais e conexões)
│   ├── js/                           # Arquivos estáticos de suporte ao cliente JS (server-info.js)
│   ├── protocol_handler/             # Disparador nativo local Windows (bancada://, instalador .cmd, registro .reg, launcher .ps1)
│   ├── scrapers/                     # Bots e raspadores automatizados (otrs, citsmart, oxe, papercut, plantoes, ramais, unidades, faq)
│   ├── scripts_powershell/           # Scripts de automação remota modularizados (analisador, manutencao, perfis)
│   ├── services/                     # Serviços de background e integrações (cron_scheduler.py, evolution_client.py, member_matcher.py)
│   ├── syncs/                        # Workers assíncronos de sincronização (doações, fiscalização, garantia, plantões, portarias, viagens, whatsapp)
│   ├── tabs/                         # Módulos de páginas isolados por funcionalidade no dashboard Streamlit (16 abas completas)
│   ├── config.py                     # Configurações centralizadas, logging e carregador do cofre/ambiente
│   ├── crypto_utils.py               # Utilitários de criptografia Fernet (AES-128) e derivação PBKDF2 para senhas
│   ├── manual_entries.py             # Base de dados de entradas e cadastros manuais de unidades
│   ├── preprocess_chamados.py        # Limpeza e padronização dos chamados de TI
│   ├── preprocess_oxe.py             # Tratamento e normalização dos ramais do OXE
│   ├── salvar_senha.py               # Utilitário interativo de credenciais e cofre de senhas (Keyring)
│   ├── tag_classifier.py             # Classificador de IA com NLP (spaCy + Scikit-Learn)
│   └── terminal.py                   # Formatador de logs com cores ANSI para o terminal
├── tests/                            # Suíte de testes unitários automatizados (NLP, scrapers, banco e tabs)
├── .env.example                      # Template de variáveis de ambiente
├── .env                              # Variáveis de ambiente locais (não commitado)
├── .dockerignore                     # Filtro de arquivos excluídos do contexto da imagem Docker
├── 00-iniciar.sh                     # Lançador principal Linux / WSL para Docker
├── 00-iniciar.cmd                    # Lançador principal Windows (duplo clique) para Docker
├── sistema-bancada.desktop           # Atalho para o menu de aplicativos desktop do Linux / Ubuntu
├── Dockerfile                        # Especificação da imagem Docker containerizada
├── Dockerfile-MP-RedHat              # Especificação de imagem Docker para ambiente Red Hat Enterprise Linux
├── docker-compose.yml                # Orquestrador Docker dos serviços web, postgres e evolution-api
├── requirements.txt                  # Lista de dependências Python (pip)
├── chamados.db                       # Banco de dados SQLite relacional local
├── init.py                           # Entrypoint do container (inicia o daemon de cron e o Streamlit)
├── dashboard.py                      # Orquestrador central do Streamlit Dashboard
└── orquestrador.py                   # Script mestre de coleta e classificação de chamados executado em background
```

---

## 📦 Como Executar a Aplicação (100% Docker)

O sistema opera de forma totalmente containerizada através do **Docker Compose**, eliminando a necessidade de gerenciar ambientes virtuais (`venv`) ou instalar bibliotecas Python no sistema operacional host.

---

### 🪟 1. Execução no Windows (via Docker no WSL)

1. **Clone o repositório:**
   ```cmd
   git clone https://github.com/rezendepauloh/automacao-chamados-sti.git
   cd automacao-chamados-sti
   ```

2. **Configure as Variáveis de Ambiente:**
   Copie `.env.example` para `.env`:
   ```cmd
   copy .env.example .env
   ```

3. **Inicie o sistema via Atalho Batch:**
   Basta dar duplo clique em `00-iniciar.cmd` ou executar no prompt:
   ```cmd
   00-iniciar.cmd
   ```
   > O script repassa a execução automaticamente para o Docker dentro do WSL, exibindo o banner estilizado, o QR Code de rede e os logs em tempo real. Pressionar <kbd>CTRL</kbd> + <kbd>C</kbd> encerra os containers automaticamente.

---

### 🐧 2. Execução no WSL Linux (Ubuntu)

1. **Clone o repositório no WSL:**
   ```bash
   git clone https://github.com/rezendepauloh/automacao-chamados-sti.git
   cd automacao-chamados-sti
   ```

2. **Conceda permissões de execução e confie no lançador Desktop:**
   ```bash
   chmod +x 00-iniciar.sh sistema-bancada.desktop
   gio set sistema-bancada.desktop metadata::trusted true 2>/dev/null || true
   ```

3. **Configure as Variáveis de Ambiente:**
   ```bash
   cp .env.example .env
   ```

4. **Inicie o sistema via Script Shell (`00-iniciar.sh`):**
   O script verifica se a imagem já foi construída (subindo instantaneamente sem rebuilds), transmite os logs e encerra todos os containers com <kbd>CTRL</kbd> + <kbd>C</kbd>:
   ```bash
   ./00-iniciar.sh
   ```

5. **Parâmetros Utilitários de Inicialização:**
   - **Configurar/Atualizar Senhas Seguras (Keyring):**
     ```bash
     ./00-iniciar.sh --config-senhas
     # ou abreviado:
     ./00-iniciar.sh -s
     ```
     > 🔑 **O que faz:** Abre o assistente interativo `salvar_senha.py` dentro do container Docker. Ele verifica se já existem senhas salvas para OTRS, CitSmart, SCCM Admin, PaperCut ou OXE, permitindo atualizar ou manter as credenciais no cofre persistente (`~/.local/share/python_keyring`). As senhas são digitadas com caracteres ocultos por segurança.

   - **Forçar Reconstrução da Imagem Docker:**
     ```bash
     ./00-iniciar.sh --build
     # ou abreviado:
     ./00-iniciar.sh -b
     ```
     > 🛠️ **O que faz:** Reconstrói a imagem Docker do zero para incorporar alterações no `Dockerfile` ou no `requirements.txt`. Pode ser combinado com o configurador de senhas (ex: `./00-iniciar.sh --build -s`).

6. **(Opcional) Atalho para o Menu de Aplicativos do Linux / Desktop:**
   ```bash
   cp sistema-bancada.desktop ~/.local/share/applications/
   ```

7. **Gerenciamento manual via Docker Compose:**
   ```bash
   # Construir a imagem
   docker compose build

   # Subir os containers em segundo plano
   docker compose up -d web

   # Executar configurador de senhas manualmente no container
   docker compose run --rm web python src/salvar_senha.py

   # Ver logs em tempo real
   docker compose logs -f web

   # Parar os containers
   docker compose down
   ```

---

### 🐳 3. Estrutura de Containers (WSL / Debian)

O projeto conta com arquitetura pronta para execução universal em containers orquestrada pelo **Docker Compose**:

- **`web` (`bancada_streamlit_app`):** Container principal baseado em `Dockerfile` (`python:3.11-slim`) com Playwright/Chromium headless, healthcheck ativo, fuso horário brasileiro (`America/Campo_Grande`), bypass de SSL corporativo e Live-Reload mapeado para `src/`, `assets/`, `dashboard.py` e `init.py`.
- **`db` (`bancada_postgres_db`):** Banco relacional PostgreSQL 15 Alpine (`postgres:15-alpine`), garantindo persistência estruturada de dados através do volume dedicado `postgres_data`.
- **`evolution-api` (`bancada_evolution_api`):** API integrada para WhatsApp (`evoapicloud/evolution-api`), conectada ao PostgreSQL para gerenciar sessões, QR Code e envio assíncrono de mensagens e alertas.
- **Volumes Persistentes:** Banco de dados SQLite (`chamados.db`), instâncias Evolution (`evolution_instances`, `evolution_store`), cofre de senhas do Linux (`~/.local/share/python_keyring`), pastas de dados (`01`, `02`, `03`, `models`, `debug_logs`, `uploads`).

---

## 🧪 Executando Testes Unitários

Para rodar todos os testes unitários integrados da aplicação dentro do container Docker:

```bash
docker compose exec web python -m unittest discover -s tests
```



