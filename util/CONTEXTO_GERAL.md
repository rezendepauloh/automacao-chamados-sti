# 🧭 Contexto Geral do Sistema Bancada — Automação de Chamados STI

> **Documento de Contexto Vivo & Roteiro Evolutivo**  
> **Última Atualização:** 02/10/2026  
> **Finalidade:** Servir de referência central unificada para alinhar o estado atual da arquitetura, módulos entregues, decisões de design e nortear as próximas etapas de desenvolvimento e resolução de bugs.

---

## 🏛️ 1. Visão Geral do Sistema

O **Sistema Bancada STI** é uma plataforma corporativa desenvolvida para a equipe de Tecnologia da Informação do Ministério Público do Estado de Mato Grosso do Sul (MPMS). O sistema automatiza a extração, unificação, tratamento inteligente e visualização de chamados técnicos, integrando múltiplas fontes de dados (OTRS, CitSmart, Central Telefônica OXE, Active Directory, SCCM, PaperCut e SIMP).

### 🛠️ Stack Tecnológica Central
- **Interface / Dashboard:** Python 3.11+, Streamlit (arquitetura multi-abas modernas, componentes modulares `subtabs`, `metric_cards`, `calendar`).
- **Persistência Relacional:** SQLite local com modo WAL (`chamados.db`) e suporte a PostgreSQL corporativo via Docker.
- **Segurança de Credenciais:** Criptografia simétrica Fernet (AES-128-CBC + HMAC-SHA256) em `crypto_utils.py` com cofre persistente `python_keyring`.
- **Containers & Orquestração:** Docker Compose (`web`, `db` PostgreSQL e `evolution-api` para WhatsApp).
- **Integração Windows/WSL:** Protocol Handler local `bancada://` acionando scripts PowerShell nativos (`bancada-launcher.ps1`, `sccm_sync.ps1`, etc.).

---

## 🧩 2. Módulos e Funcionalidades Implementadas

### A. Chamados de TI (OTRS, CitSmart & Central OXE)
- **Raspagem Automatizada:** Coletores dedicados (`otrs_scraper.py`, `citsmart_scraper.py` e `oxe_scraper.py`) com persistência relacional.
- **Limpeza & NLP:** Remoção de saudações/assinaturas, fechamento de chamados ausentes (`close_missing_tickets_by_base`) e desvio inteligente de apoio remoto (ex: Costa Rica -> Ricardo Brandão II).
- **Classificação por IA:** Pipeline de Machine Learning (TF-IDF + Naive Bayes / SVM) para categorização automática por TAGs de atendimento.
- **Tabela Interativa & Modal (`@st.dialog`):** Ficha detalhada do chamado com resumo executivo, dados do solicitante, localização física, histórico de notas e edição em tempo real de título e localidade.

### B. Gestão de Ativos & Infraestrutura
- **Active Directory (LDAP):** Organograma em árvore das Unidades Organizacionais (OUs), contas de usuários com crachá RFID PaperCut (`pager`), inventário de computadores/servidores e ações remotas (RDP, C$, Ping).
- **Inventário MECM/SCCM:** Cache relacional de dispositivos (`SMS_R_System`), usuários e coleções, com disparador de controle remoto oficial (`CmRcViewer.exe`) e sincronizador PowerShell.
- **Telefonia OXE & Ramais:** Inventário de ramais analógicos/IP, linhas diretas, entroncamentos e busca rápida por usuário/setor.
- **Impressoras PaperCut:** Monitoramento de servidores de impressão, status de filas e contadores de páginas.

### C. Apoio Operacional & Logística
- **Escalas de Plantão (SIMP):** Plantão matutino diário e semanal de aviso, com detecção e destaque dos técnicos da bancada.
- **Contratos & Garantias:** Acompanhamento de garantias vigentes de equipamentos de TI.
- **Fiscalização & Portarias:** Monitoramento de publicações oficiais e designações de fiscais técnicos.
- **Controle de Viagens:** Agendamentos de viagens técnicas para comarcas do interior.
- **Calendário Master (FullCalendar v6):** Calendário integrado exibindo chamados, plantões, viagens e garantias com exportação RFC 5545 `.ics`.
- **Notificações no WhatsApp:** Disparo de avisos de plantão e alertas via Evolution API.

### D. Qualidade & Testes Automatizados
- **Suíte Unificada (`tests/run_all.py`):** 29 testes automatizados cobrindo banco de dados, criptografia, serviços, componentes de interface, scripts PowerShell e integrações, com runner ANSI colorido e execução via `./00-iniciar.sh --tests` ou pelo menu (opção `6`).

---

## 🎯 3. Foco Atual & Próximas Etapas

### ✅ Concluído: Resolução Definitiva do Bug de IP de Origem e Hostname nos Modais
- **Status:** **Resolvido e Validado (100% dos testes aprovados)**.
- **Entregas Realizadas:**
  1. **Sanitização de Valores Nulos na UI ([`src/tabs/chamados.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/tabs/chamados.py)):**
     - Criado helper local `_sanitize_val()` que intercepta `float('nan')`, `"nan"`, `None`, `null` ou strings vazias, exibindo `N/A` de forma limpa.
     - Adicionado enriquecimento dinâmico em tempo real caso `ip_origem` ou `hostname` estejam ausentes no chamado aberto no modal.
     - Inclusão dos botões de ação remota instantânea (`bancada://run?tool=cmrc` e `bancada://run?tool=rdp`) quando a máquina ou IP são identificados.
  2. **Lookup e Fallback no Cache Relacional do SCCM ([`src/database/sccm_db.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/database/sccm_db.py) e [`src/config.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/config.py)):**
     - Criada a função [`get_device_by_user()`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/database/sccm_db.py) buscando por `last_logon_user` em `sccm_cache_devices`, priorizando máquinas ativas e IPs de rede interna `10.x`.
     - Integrado lookup transparente no início de [`fetch_sccm_data()`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/config.py), permitindo obter os dados instantaneamente mesmo sem conexão direta via WMI/CIM ou em ambientes de contêiner.
     - Criada a rotina [`update_ticket_device_info()`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/database/tickets_db.py) para persistir o IP/Hostname enriquecido no SQLite.
  3. **Rotina de Salvamento no Banco e Scrapers ([`src/database/tickets_db.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/database/tickets_db.py), [`src/preprocess_chamados.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/preprocess_chamados.py), [`citsmart_scraper.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/scrapers/citsmart_scraper.py) e [`otrs_scraper.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/scrapers/otrs_scraper.py)):**
     - [`save_tickets_to_db()`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/database/tickets_db.py) higieniza todas as entradas antes de gravar, enriquecendo chamados novos ou atualizados via cache do SCCM e preservando dados já existentes.
     - Scrapers e scripts de pré-processamento agora contam com `.fillna("")` e higienização de cache para que nenhum valor `NaN` seja gerado ou gravado nos arquivos intermediários ou no banco de dados.
  4. **Testes Automatizados ([`tests/unit/test_database.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/tests/unit/test_database.py)):**
     - Novos testes unitários adicionados (`test_sccm_get_device_by_user` e `test_tickets_save_nan_sanitization_and_enrichment`), elevando a suíte para 29 testes executados com 100% de sucesso.

### ✅ Concluído: Resolução Definitiva do Bug de Localidades com 'nan' e 'Não encontrado no AD'
- **Status:** **Resolvido e Validado (100% dos 30 testes aprovados)**.
- **Entregas Realizadas:**
  1. **Classificação e Composição de Localidade ([`src/tag_classifier.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/tag_classifier.py)):**
     - Função `_clean_loc()` criada no fallback de localidade física para interceptar `nan`, `none`, `null`, `n/d` e variações de gênero como `"Não encontrado no AD"` e `"Não encontrada no AD"`.
     - Impede a concatenação incorreta `"nan - Não encontrado no AD"`, definindo como `"Não identificada"` ou aproveitando a parte válida existente.
  2. **Sanitização no Carregamento e Persistência do Banco ([`src/database/tickets_db.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/database/tickets_db.py)):**
     - Em `save_tickets_to_db()`, valores nulos ou erros de busca do AD são tratados antes do INSERT/UPDATE.
     - Em `load_data()`, filtros tratam `cidade_predio`, `unidade` e `localidade_fisica`, garantindo que strings residuais nunca alcancem os DataFrames e a interface gráfica.
  3. **Interface dos Modais e Edição Manual ([`src/tabs/chamados.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/tabs/chamados.py)):**
     - Sanitização em `show_ticket_details()` para exibição limpa de `Localidade: Não identificada` quando ausente.
     - Os inputs do expander "Editar Localização Manual" agora usam `_sanitize_val()`, impedindo que os campos de texto abram preenchidos com o texto literal `"nan"`.
  4. **Scripts de Limpeza e Migração para Produção ([`util/limpeza_banco/limpar_localidades_nan.sql`](automated-OTRS-and-CitSmart/util/limpeza_banco/limpar_localidades_nan.sql) e [`util/limpeza_banco/limpar_localidades_nan.py`](automated-OTRS-and-CitSmart/util/limpeza_banco/limpar_localidades_nan.py)):**
     - Criados scripts SQL e Python prontos para execução em ambientes de homologação e produção, higienizando chamados existentes no banco `chamados.db`.
  6. **Padronização de Prédios de Campo Grande sem Sub-setores ([`src/manual_entries.py`](automated-OTRS-and-CitSmart/src/manual_entries.py)):**
     - Faixas de IP da PGJ (ex: `10.111.144.0/24` e `10.111.145.0/24`) foram padronizadas para `"Campo Grande - PGJ"` em vez de acrescentar sufixos de setores (`" - CI"` ou `" - STI"`), mantendo a localidade física limpa e coerente com a lista de prédios.
     - Atualizados os scripts de limpeza (`limpar_localidades_nan.sql` e `limpar_localidades_nan.py`) para normalizar chamados legados com sufixos duplicados.
  7. **Padronização de Comarcas do Interior ([`src/tag_classifier.py`](automated-OTRS-and-CitSmart/src/tag_classifier.py)):**
     - A lógica de fallback agora atribui diretamente a comarca/cidade limpa à `Localidade física` (ex: `Terenos`, `Paranaíba`, `Cassilândia`, `Ivinhema`) sem concatenar o nome da promotoria interna (ex: `1ª PJ de Terenos`), deixando a especificação da PJ exclusivamente no campo e filtro `Unidade`.

### ✅ Concluído: Geração Inteligente de Títulos para Chamados sem Título (CitSmart)
- **Status:** **Resolvido e Validado (100% dos 33 testes aprovados)**.
- **Entregas Realizadas:**
  1. **Análise de Descrição e Síntese de Título ([`src/tag_classifier.py`](automated-OTRS-and-CitSmart/src/tag_classifier.py)):**
     - Implementada a função [`generate_synthetic_title(tag, description)`](automated-OTRS-and-CitSmart/src/tag_classifier.py), que combina a `TAG` prevista pelo modelo Scikit-Learn (SVM / ComplementNB) com processamento de linguagem natural (NLP).
     - Remove tags HTML, entidades, saudações compostas ("Bom dia Prezados", "Olá tudo bem"), e preâmbulos burocráticos ("Por determinação do Promotor...", "solicito a...", "venho por meio deste solicitar...").
     - Isola a oração substantiva do problema ou pedido, compondo títulos objetivos como `[IMPRESSORA] Instalação da impressora PRT-5394 no setor` ou `[INSTALAÇÃO HARDWARE] Instalação do Microcomputador Pat. 80700`.
  2. **Preservação de Títulos Nativos e Edições Manuais:**
     - [`generate_missing_titles(df)`](automated-OTRS-and-CitSmart/src/tag_classifier.py) atua exclusivamente em registros onde a coluna `Título` está vazia ou nula, garantindo que títulos nativos do OTRS permaneçam intactos.
     - [`save_tickets_to_db()`](automated-OTRS-and-CitSmart/src/database/tickets_db.py) verifica se o chamado já possui título (ou edição manual salva via UI modal) para nunca sobrescrevê-lo com dados em branco.
  3. **Migração e Atualização da Base:**
     - Integrado ao utilitário [`util/limpeza_banco/limpar_localidades_nan.py`](automated-OTRS-and-CitSmart/util/limpeza_banco/limpar_localidades_nan.py) para que todas as bases (inclusive em produção) possam preencher os títulos vazios de forma retroativa.

### ✅ Concluído: Blindagem Estrita das Filas de Manutenção (OTRS & CitSmart)
- **Status:** **Resolvido e Validado (100% dos 33 testes aprovados)**.
- **Entregas Realizadas:**
  1. **Navegação Direta OTRS ([`src/scrapers/otrs_scraper.py`](automated-OTRS-and-CitSmart/src/scrapers/otrs_scraper.py)):**
     - O scraper navega diretamente para `QueueID=11` (Manutenção da STI) com fallback seguro para `QueueID=0`, garantindo a exclusividade dos chamados da fila.
  2. **Validação Estrita via DOM e Filtro de Grupo no CitSmart ([`src/scrapers/citsmart_scraper.py`](automated-OTRS-and-CitSmart/src/scrapers/citsmart_scraper.py)):**
     - Identificado que o backend do CitSmart disparava respostas XHR com 100 itens sem a chave de grupo preenchida no JSON.
     - Implementado cruzamento bidirecional: o scraper lê os IDs renderizados na tabela DOM (`#table tbody tr`) da fila selecionada (`copilot_novo` / `[N1] Manutenção`) e filtra estritamente os tickets capturados, descartando qualquer chamado excedente de outras equipes.
     - Descarte imediato de tickets automáticos (`Monitoramento Adm MPMS` e `Adm Ticket Por Email`) tanto no scraper quanto em [`save_tickets_to_db()`](automated-OTRS-and-CitSmart/src/database/tickets_db.py).
  3. **Higienização e Reorganização de Utilitários ([`util/`](automated-OTRS-and-CitSmart/util/)):**
     - Scripts organizados em subpastas (`limpeza_banco/`, `migracoes/`, `diagnosticos/`, `arquivos/`, `docs/`) com documentações em Markdown padronizadas para execução via Bash (Linux / WSL / Red Hat).
     - Banco SQLite expurgado, cravando com precisão a volumetria da equipe (52/53 OTRS + 19 CitSmart).

---

### ✅ Concluído: Otimização da Aba SCCM & Normalização Inteligente do Windows 11
- **Status:** **Resolvido e Validado (100% dos 36 testes aprovados)**.
- **Entregas Realizadas:**
  1. **Migração dos Filtros, Pesquisas e Paginação para a Barra Lateral ([`src/tabs/sccm.py`](automated-OTRS-and-CitSmart/src/tabs/sccm.py)):**
     - Inputs de pesquisa, seleção e o seletor universal de itens por página (`render_items_per_page_selector`) contextualizados dinamicamente no `st.sidebar` para todas as sub-abas (`💻 Dispositivos`, `👤 Usuários`, `📁 Coleções de Dispositivos`, `👥 Coleções de Usuários` e `🛡️ Configurações & Conformidade`).
     - Integração de paginação completa via [`src/components/pagination.py`](automated-OTRS-and-CitSmart/src/components/pagination.py) com régua de navegação (`render_pagination_controls`), suporte a `10, 25, 50, 100, "Todos"` registros e persistência de seleção no modal de ações rápidas.
     - A área principal agora destaca com máxima visibilidade os cards de métricas (KPIs), as tabelas de dados completas e as ações remotas instantâneas (`bancada://run?tool=cmrc`, `rdp`, `explorer`, `ping`).
  2. **Detecção e Normalização Técnica do Windows 11 por Build ([`src/database/sccm_db.py`](automated-OTRS-and-CitSmart/src/database/sccm_db.py)):**
     - Criadas as funções `extract_windows_build(os_version)` e `normalize_os_name(os_name, os_version)`.
     - Tratamento da limitação da classe `SMS_R_System` do SCCM (que reporta "Microsoft Windows NT Workstation 10.0" por compartilhamento do kernel NT 10.0):
       * `Build >= 22000` (ex: 26100, 26200, 22631, 22621) normalizado para `"Windows 11"`.
       * `Build < 22000` (ex: 19045) normalizado para `"Windows 10"`.
       * Servidores normalizados para `"Windows Server 2022"`, `"Windows Server 2019"`, `"Windows Server 2016"`, `"Windows Server 2025"` etc.
     - Implementada a rotina de migração idempotente `migrate_sccm_os_normalization()` executada no startup das tabelas, higienizando as mais de 2.450 estações no cache relacional local.
     - `save_sccm_devices()` e `get_sccm_devices_df()` blindados com a normalização automática.
  3. **Expansão do Modal de Ficha Técnica & Hardware do Computador ([`src/tabs/sccm.py`](automated-OTRS-and-CitSmart/src/tabs/sccm.py)):**
     - Modal reconfigurado para `@st.dialog(..., width="large")`, proporcionando uma visualização executiva ampla e limpa.
     - Painel dividido em 3 colunas de especificações:
       * **Hardware & Componentes:** Fabricante, Modelo da Máquina, Processador (CPU), Memória RAM e Discos/Armazenamento (com capacidade e espaço livre).
       * **Sistema & Agente:** Sistema Operacional normalizado, Versão de Build, Versão do Cliente SCCM, Último Check-in e Resource ID.
       * **Rede & Domínio:** IP(s), MAC Address, Site AD e Domínio corporativo.
     - Inclusão do botão de **Diagnóstico Remoto da Bancada** (`bancada://run?tool=analisador&host={name}`) somado ao CmRcViewer, RDP, Explorer C$ e Teste Ping.
     - Enriquecimento do script de sincronização [`src/scripts_powershell/sccm_sync.ps1`](automated-OTRS-and-CitSmart/src/scripts_powershell/sccm_sync.ps1) para consultar `SMS_G_System_COMPUTER_SYSTEM`, `SMS_G_System_PROCESSOR` e `SMS_G_System_LOGICAL_DISK`, persistindo nas novas colunas de hardware em [`src/database/sccm_db.py`](automated-OTRS-and-CitSmart/src/database/sccm_db.py).
   4. **Aliases de Modelos de Hardware e Filtro Dinâmico por Modelo ([`src/database/sccm_db.py`](automated-OTRS-and-CitSmart/src/database/sccm_db.py) & [`src/tabs/sccm.py`](automated-OTRS-and-CitSmart/src/tabs/sccm.py)):**
      - Centralização do dicionário documentado `HARDWARE_MODEL_ALIASES` para conversão de Machine Types / MTM da Lenovo/Dell (ex: `11DUSD3R00` -> `"Lenovo ThinkCentre M70q Gen 2"`, `12TES8R800` -> `"Lenovo ThinkCentre M70q Gen 5"`, `20W1S6CB00` -> `"Lenovo ThinkPad T14 Gen 1"`, etc.).
      - Função `normalize_hardware_model()` com matching exato e por prefixo de MTM, associada à rotina de migração em startup `migrate_sccm_model_aliases()`.
      - Novo filtro `<select>` dinâmico na barra lateral (`st.sidebar.selectbox` "⚙️ Modelo de Hardware") alimentado automaticamente com os modelos existentes no banco.
      - Busca inteligente no sidebar e SQL (`get_sccm_devices_df()`): pesquisa simultânea por Hostname, Usuário, IP, Modelo comercial, Fabricante e código MTM original (inclusive dentro do `raw_json`).
      - No modal amplo, modelos mapeados mostram o nome comercial amigável acompanhado do código de fábrica discreto (ex: *Lenovo ThinkCentre M70q Gen 2 (11DUSD3R00)*).
   5. **Padronização de Datas e Horas no Padrão Brasileiro (DD/MM/AAAA HH:MM:SS) ([`src/tabs/sccm.py`](automated-OTRS-and-CitSmart/src/tabs/sccm.py)):**
      - Criação da função auxiliar `format_sccm_datetime()`, que realiza o parse resiliente de timestamps WMI/CIM do SCCM (ex: `20260929080059.657000+***`) e ISO-8601 para a máscara brasileira `DD/MM/AAAA HH:MM:SS`.
      - Aplicada nas tabelas de **Coleções de Dispositivos e Usuários** (coluna *Última Atualização*), na sub-aba **Conformidade** (coluna *Última Atividade*) e no modal de **Ficha Técnica** (*Último Check-in*).
   6. **Validação e Testes Automatizados ([`tests/unit/test_database.py`](automated-OTRS-and-CitSmart/tests/unit/test_database.py) & [`tests/components/test_components.py`](automated-OTRS-and-CitSmart/tests/components/test_components.py)):**
      - Adicionados os testes `test_sccm_os_normalization_and_build_detection`, `test_sccm_hardware_model_aliases_and_filter` e `test_format_sccm_datetime`.
      - Suíte unificada de testes (`tests/run_all.py`) operando em 100% verde (**36 testes aprovados**).

### ✅ Concluído: Estabilização Operacional, Housekeeping e Resiliência de Credenciais
- **Status:** **Resolvido e Validado (100% dos 39 testes aprovados)**.
- **Entregas Realizadas:**
  1. **Resolução da Exceção no Cron Daemon (`'dict' object has no attribute 'to_dict'`) ([`src/services/cron_scheduler.py`](automated-OTRS-and-CitSmart/src/services/cron_scheduler.py)):**
     - O erro ocorria durante o loop contínuo do daemon devido à incompatibilidade do método `.to_dict()` ao iterar sobre DataFrames em ambientes com mocks de testes ou estruturas nativas.
     - Implementada conversão polimórfica defensiva (`hasattr(row, "to_dict")` com fallback para `dict(row)`).
     - Atualizado o mock universal [`tests/test_helpers.py`](automated-OTRS-and-CitSmart/tests/test_helpers.py) com a classe `MockSeries(dict)` implementando `.to_dict()` e compatibilidade total com o pandas real.
     - Protegidos também os pontos de conversão de linhas em [`src/tabs/central_telefonica.py`](automated-OTRS-and-CitSmart/src/tabs/central_telefonica.py), [`src/tabs/sccm.py`](automated-OTRS-and-CitSmart/src/tabs/sccm.py), [`src/tabs/unidades.py`](automated-OTRS-and-CitSmart/src/tabs/unidades.py) e [`src/scrapers/papercut_scraper.py`](automated-OTRS-and-CitSmart/src/scrapers/papercut_scraper.py).
  2. **Tratamento Resiliente de Credenciais e Portabilidade entre Ambientes ([`src/crypto_utils.py`](automated-OTRS-and-CitSmart/src/crypto_utils.py)):**
     - Adicionada validação estrita de integridade para chaves Fernet (`_is_valid_fernet_key()`), gerando automaticamente nova chave mestra caso o arquivo `.secret.key` esteja corrompido ou `APP_SECRET_KEY` seja inválida, prevenindo travamentos fatais (`ValueError`).
     - Implementada detecção de tokens Fernet incompatíveis (`is_fernet_token()`), gerando logs orientativos (`logger.warning`) ao invés de repassar strings cifradas inválidas para tentativas de autenticação no AD/OTRS/CitSmart ao migrar a base para outra máquina sem a chave original.
     - Expandida a suíte em [`tests/unit/test_crypto.py`](automated-OTRS-and-CitSmart/tests/unit/test_crypto.py) com testes de chaves corrompidas e tokens legados.
  3. **Housekeeping de Exceções Silenciosas e Logs Estruturados:**
     - Higienizados blocos `except: pass` e `except Exception: pass` em [`src/tabs/active_directory.py`](automated-OTRS-and-CitSmart/src/tabs/active_directory.py) (exportação de usuários e máquinas para Excel via openpyxl), [`src/tabs/chamados.py`](automated-OTRS-and-CitSmart/src/tabs/chamados.py) (persistência de dados de rede de máquinas em chamados) e [`src/database/cron_db.py`](automated-OTRS-and-CitSmart/src/database/cron_db.py) (fallback de tarefas padrão).
  4. **Expansão de Testes Automatizados:**
     - Testes adicionados para `BancadaCronDaemon` em [`tests/services/test_services.py`](automated-OTRS-and-CitSmart/tests/services/test_services.py).
     - Suíte unificada de testes (`python3 tests/run_all.py`) rodando em **100% de aprovação (39 testes)** em ~196ms.
  5. **Restauração dos Controles de Edição e Malha no Mapa ([`src/tabs/mapas.py`](automated-OTRS-and-CitSmart/src/tabs/mapas.py)):**
     - Identificado que na migração para abas modulares, os botões Leaflet no canto superior direito (`topright`), abaixo do botão de fullscreen, não estavam sendo instanciados.
     - Restaurados os botões:
       - `MeshToggleControl` (`👁️`): Alterna a visibilidade da malha de caminhos e nós de pathfinding (`debugLayer`).
       - `DevModeControl` (`🛠️`): Alterna o modo de edição do mapa, ativando cursor `crosshair`, elementos arrastáveis (`draggable`), popups de edição e criação de nós e pins ao clicar.
       - `SaveControl` (`💾`): Botão de salvar alterações que aparece dinamicamente no modo dev e sincroniza via POST HTTP (`http://localhost:8099/save_config`) com o SQLite e o JSON físico.
     - Restauradas as funções globais de interação (`saveConfigToDb`, `saveDevElement`, `connectToLastNode`, `removeNode`, `removePin`, `removeEdge`, `updateNode`, `updatePin`, `setLastNode`).
  6. **Suporte à Coluna "Chamado Diária" em Viagens ([`src/tabs/viagens.py`](automated-OTRS-and-CitSmart/src/tabs/viagens.py)):**
     - Adicionada a coluna `chamado_diaria TEXT` na tabela `viagens` em [`src/database/viagens_db.py`](automated-OTRS-and-CitSmart/src/database/viagens_db.py) com rotina de migração defensiva (`PRAGMA table_info` + `ALTER TABLE`).
     - Leitura e normalização resiliente da coluna a partir da planilha Excel oficial (`Chamado Diária`, `Chamado Diaria`, `Diária`).
     - Inclusão da coluna na tabela interativa da aba Viagens com `st.column_config.TextColumn("💵 Chamado Diária")`, no filtro de busca textual, na exportação Excel, nos `extendedProps` do calendário master ([`src/components/calendar.py`](automated-OTRS-and-CitSmart/src/components/calendar.py) e [`src/tabs/calendario_geral.py`](automated-OTRS-and-CitSmart/src/tabs/calendario_geral.py)), na exportação `.ics` ([`src/services/ics_export.py`](automated-OTRS-and-CitSmart/src/services/ics_export.py)) e nas mensagens do WhatsApp ([`src/syncs/sync_whatsapp_scheduler.py`](automated-OTRS-and-CitSmart/src/syncs/sync_whatsapp_scheduler.py)).
     - Ajustado o modal de configuração de viagens para `width="large"`.
     - Suíte de testes expandida para 40 testes (100% de aprovação).
  7. **Ordenação Cronológica das Datas de Viagens ([`src/tabs/viagens.py`](automated-OTRS-and-CitSmart/src/tabs/viagens.py)):**
     - Corrigido o comportamento de clique nos títulos "Saída" e "Retorno" na tabela de viagens, que antes ordenavam alfabeticamente por serem strings em formato brasileiro (`DD/MM/YYYY`).
     - Convertidas as colunas para objetos `datetime` nativos (`data_saida` e `data_retorno`) a partir de `saida_iso`/`retorno_iso` com fallback resiliente para parsing de `saida_br`/`retorno_br`.
     - Configurado no `st.dataframe` o componente oficial `st.column_config.DateColumn("📅 Saída", format="DD/MM/YYYY")` e `st.column_config.DateColumn("🏁 Retorno", format="DD/MM/YYYY")`, garantindo ordenação estritamente cronológica crescente/decrescente com visual amigável brasileiro e exportação Excel padronizada.
  8. **Ordenação Cronológica das Tabelas do SCCM ([`src/tabs/sccm.py`](automated-OTRS-and-CitSmart/src/tabs/sccm.py)):**
     - Corrigido o comportamento de ordenação por clique nos cabeçalhos "Última Atualização" (sub-abas de Coleções de Dispositivos e Usuários) e "Última Atividade" (sub-aba de Configurações & Conformidade).
     - Criada a função helper `parse_sccm_datetime(val)` que interpreta datas CIM/WMI (ex: `YYYYMMDDHHmmss.microsec+tz`), ISO-8601 e variações comuns para objetos `datetime`/`Timestamp` nativos.
     - As tabelas passaram a utilizar a coluna formatada via `st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm:ss")`, permitindo ordenar estritamente por ordem cronológica (ano, mês, dia, hora) ao invés da ordenação léxica de string.
     - Atualizados os testes unitários em [`tests/components/test_components.py`](automated-OTRS-and-CitSmart/tests/components/test_components.py) cobrindo tanto `format_sccm_datetime` quanto `parse_sccm_datetime`.

   9. **Estabilização do Leitor de FAQs, Mídia Autenticada e Suíte de Testes ([`src/tabs/links_faqs.py`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py)):**
      - **Diagnóstico da Quebra de Imagens (HTTP 401 Unauthorized):** As capturas de tela dos tutoriais do SharePoint Online exigem cookies de sessão corporativa (`rtFa`, `FedAuth`, etc.). Ao carregar a página no Streamlit em `localhost:8501`, o navegador aplica restrições estritas de SameSite/CORS para requisições cross-origin `<img>`, bloqueando o envio dos cookies e gerando erro 401.
      - **Extração Automática de Cookies & Organização em Pastas por Tutorial:**
        * Criado mecanismo que extrai diretamente do perfil de usuário do Firefox (`/mnt/c/Users/paulogoncalves/.../cookies.sqlite`) os cookies autenticados de `ministeriopublicoms.sharepoint.com` e os armazena com segurança em `uploads/faq/sharepoint_cookies.json`.
        * **Estrutura de Pastas Padronizada por Tutorial:** Em vez de salvar arquivos soltos na raiz, tanto imagens quanto vídeos possuem suas subpastas dedicadas (`uploads/faq/imagens/<slug_tutorial>/` e `uploads/faq/videos/<slug_tutorial>/`), garantindo uma organização uniforme e espelhada.
        * Implementada migração automática e defensiva: arquivos previamente baixados na raiz foram realocados para as respectivas pastas de tutoriais.
        * A **Galeria de Imagens** e a **Aba de Vídeos** agrupam e exibem mídias por pasta de tutorial com contadores dinâmicos, carrossel integrado e atalhos para abertura na intranet.
        * Implementadas as funções [`ensure_sharepoint_image_cached()`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py), [`ensure_sharepoint_video_cached()`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py), [`get_image_as_base64()`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py) e [`get_video_as_base64()`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py), que injetam Data URIs base64 diretamente no HTML do leitor de FAQs, eliminando 401s e garantindo leitura e reprodução de vídeos 100% offline.
        * **Eliminação do Espaço em Branco nos Vídeos do SharePoint:** Resolvido o problema de placeholders vazios (`aria-busy="true"` com `min-height: 694px` do `DocumentEmbedWebPart`) através do mapeamento das chamadas da API REST do SharePoint (`CanvasContent1`), enriquecendo o conteúdo HTML no banco relacional local e substituindo os placeholders por players nativos `<video controls>` estilizados e responsivos (`.sp-video-card`).
        * **Tratamento Inteligente de Vídeos Pesados (>75MB):** Vídeos de alta definição com centenas de megabytes (como o da bancada do *Lenovo T14* com 514MB) sobrecarregavam a memória do navegador se injetados diretamente em base64 via DOM. Para esses casos, o sistema exibe um card destacado com atalho duplo: abertura instantânea no player local do sistema operacional / VLC (`bancada://run?tool=vlc&target=...`) e atalho corporativo para o Microsoft Stream Online, enquanto vídeos menores (como o de *HP Elitedesk 800 G5* de 41MB e da impressora de 8MB) continuam tocando diretamente no modal via HTML5 player.
        * **Suporte a `tool=vlc` no Protocol Handler ([`src/protocol_handler/bancada-launcher.ps1`](automated-OTRS-and-CitSmart/src/protocol_handler/bancada-launcher.ps1)):** Isenção do parâmetro obrigatório `host` para as ferramentas `vlc` e `player`, tradução automática de caminhos do container Docker (`/app/...`) para caminhos de rede UNC do Windows/WSL (`\\wsl.localhost\Ubuntu-26.04\...`), detecção automática do executável do VLC (`vlc.exe`) ou fallback direto para o player padrão de mídia do Windows.
        * **Correção de Mapeamento de URLs do SharePoint:** Tratada a divergência de links codificados em URL (`%20` vs espaço) na rotina de extração dos webparts e corrigido o link do vídeo *HP Elitedesk 800 G5*, que estava apontando para um link genérico inexistente no SharePoint.
        * **Generalização e Normalização Universal de URLs de Mídia:** Implementada a função [`normalize_sharepoint_url()`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py) que resolve caminhos relativos e codifica devidamente caracteres especiais e espaços (`quote/unquote`) tanto em imagens (`<img>`, `.imagePlugin`) quanto em vídeos (`.sp-video-embed`, `controldata`).
        * **Limpeza Automática de Avatares e Placeholders:** Decomposição de imagens de perfil de autores (`userphoto.jpg`), bibliotecas CDN de fontes/ícones do Office e remoção de containers vazios residuais nas 36 páginas do FAQ.
        * **Sincronização em Lote Enriquecida:** O botão **`🔄 Sincronizar Mídias dos FAQs`** na barra lateral utiliza agora o motor do BeautifulSoup para escanear recursivamente todas as tags e metadados dos 27 tutoriais que contêm mídias (172 imagens únicas e 22 vídeos), garantindo download resiliente por tutorial.
        * **Generalização das Subabas Standalone de Vídeos e Imagens:**
          - **Aba de Vídeos FAQ (`🎥 Vídeos FAQ (Tutoriais)`):**
            * Substituído o fluxo legado lento de download por Selenium pela rotina ultrarrápida [`ensure_sharepoint_video_cached()`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py) usando os cookies corporativos do SharePoint.
            * Suporte automático a vídeos grandes (>75MB): o modal detecta o tamanho do arquivo e renderiza botão direto com o disparador Windows `bancada://run?tool=vlc&target=...`, preservando a memória e evitando travamento do navegador.
            * Em cada card de vídeo da grade, quando o vídeo já existe localmente, é exibido o botão **`▶️ VLC ↗`** com link direto `bancada://` para reprodução instantânea no VLC Player no Windows, além do botão de assistir no SharePoint Stream.
            * Mapeamento resiliente de vídeos locais mesmo quando salvos em diferentes subpastas organizadas.
          - **Aba de Imagens FAQ (`🖼️ Imagens FAQ (Galeria)`):**
            * Estrutura de pastas da galeria agora mapeada diretamente a partir da biblioteca oficial **`Imagens FAQ`** do SharePoint (`catalog_by_cat` do SQLite + pastas em `uploads/faq/imagens/`), desvinculando-a da tabela de artigos/páginas do FAQ (`df_faqs`).
            * Cada card de pasta agora aponta especificamente para a respectiva subpasta dentro da biblioteca no SharePoint (`.../Shared Documents/Tutoriais-FAQs/Imagens FAQ/<pasta>`), corrigindo o redirecionamento errôneo que antes levava para páginas de artigos (`SitePages/*.aspx`).
            * Unificação do catálogo do SQLite (`faq_imagens`, 260 registros catalogados em 21 pastas reais da biblioteca) com as imagens locais, garantindo contadores precisos e visualização completa no carrossel de fotos.
            * No carrossel do modal de imagens, suporte a download sob demanda com autenticação por cookies para imagens ainda não baixadas localmente, além de renderização direta com Base64 Data URI eliminando bloqueios de CORS e 401s.
        * **Sincronização 100% Dinâmica via API REST do SharePoint Online ([`sync_sharepoint_catalog_via_api`](automated-OTRS-and-CitSmart/src/tabs/links_faqs.py)):**
          - Implementado extrator que consulta diretamente as APIs REST nativas do SharePoint com os cookies de sessão corporativa:
            * `Site Pages` (`_api/web/lists/getbytitle('Site Pages')/items`): 40 artigos e tutoriais com conteúdo HTML Canvas recuperados e sincronizados em `faqs`.
            * `Vídeos FAQ` (`_api/web/GetFolderByServerRelativeUrl(...)/Folders & /Files`): 69 vídeos organizados em 23 pastas sincronizados em `faq_videos`.
            * `Imagens FAQ` (recursivo multi-nível): 260 imagens sincronizadas em `faq_imagens`.
          - **Worker Assíncrono em Background ([`src/syncs/sync_sharepoint_catalog.py`](automated-OTRS-and-CitSmart/src/syncs/sync_sharepoint_catalog.py)):**
            * Execução não-bloqueante via subprocess e lock file (`sharepoint_catalog_sync.lock`).
            * Cabeçalho ASCII estilizado no terminal e no Docker (`╔════ WORKER - SINCRONIZAÇÃO DE CATÁLOGO SHAREPOINT ════╗`).
            * Log dedicado gravado em `debug_logs/faq/sync_sharepoint_catalog.log`.
            * Accordion dinâmico (`render_log_expander`) em tempo real no topo da página de FAQ com auto-refresh a cada 3s e notificação toast ao concluir.
        * **Melhorias e Correções de Bugs no Active Directory e SCCM:**
          - **Resolução do Loop do Modal no SCCM ([`src/tabs/sccm.py`](automated-OTRS-and-CitSmart/src/tabs/sccm.py)):** Corrigido o disparo contínuo de `modal_device_details` adicionando controle de estado em `st.session_state["last_selected_sccm_dev"]` e botão explícito de fechamento que zera a seleção, evitando que o modal reabra indevidamente a cada rerun.
          - **Sincronização Assíncrona do Active Directory ([`src/syncs/sync_ad_catalog.py`](automated-OTRS-and-CitSmart/src/syncs/sync_ad_catalog.py)):** Criado worker dedicado não-bloqueante para a consulta LDAP corporativa com cabeçalho ASCII no terminal/Docker, log em `debug_logs/ad/sync_ad.log`, accordion dinâmico em tempo real (`render_log_expander`) e integração ao agendador do cron (`sync_ad_catalog`).
          - **Ponte Cruzada Bidirecional AD ↔ SCCM:**
            * Da Ficha do Computador no AD (`show_computer_details_dialog`): botão direto *"🔍 Ver Inventário Completo de Hardware no SCCM"* com preenchimento automático de busca via URL.
            * Da Ficha do Usuário no AD (`show_user_details_dialog`): botão direto *"💻 Buscar Dispositivos do Usuário no SCCM"* para auditar as estações utilizadas pelo servidor.
            * Da Ficha Técnica no SCCM (`modal_device_details`): botões rápidos para abrir a conta da máquina no AD e o cadastro do usuário logado no AD com 1 clique.
        * **Navegação Híbrida no Menu Principal ([`src/components/header.py`](automated-OTRS-and-CitSmart/src/components/header.py)):**
          - Cada item do menu superior (hambúrguer) agora possui layout em duas colunas (`[5, 1]`):
            * **Lado esquerdo:** Botão principal de navegação interna que atualiza a página mantendo a mesma aba e os estados da sessão.
            * **Canto direito:** Botão com ícone `↗` (`st.link_button` com link relativo `?tab=<slug>`), disparando a abertura da página diretamente em uma **nova aba (`target="_blank"`)**.
          - Estilização refinada em [`assets/css/styles.css`](automated-OTRS-and-CitSmart/assets/css/styles.css) com largura equilibrada e microinterações de hover.
        * **Suíte Completa de 46 Testes Automatizados 100% Verde (`python3 tests/run_all.py`).**

---

    10. **Implementação Concluída: Módulo de Férias da Bancada ([`src/tabs/ferias.py`](automated-OTRS-and-CitSmart/src/tabs/ferias.py)):**
       - **Fonte de Dados & Extração da Planilha do SharePoint:**
         * Integrada a planilha oficial `Previsão de Férias-Manutencao.xlsx` mantida na biblioteca `Shared Documents/DIT-Manutenção` do SharePoint.
         * Parser robusto em [`src/database/ferias_db.py`](automated-OTRS-and-CitSmart/src/database/ferias_db.py) com suporte aos exercícios anuais (2024, 2025, 2026, 2027) e às modalidades: *Férias Regulamentares*, *Licenças & Compensações* e *Recesso Forense*.
         * Extração de períodos contínuos e listas de datas (`DD a DD`, `DD/MM a DD/MM`, `DD-DD/MM`, `3-7/ago`, `30 e 31`, etc.).
       - **Persistência Relacional SQLite:**
         * Criada tabela dedicada `ferias_bancada` no banco `chamados.db` com cálculo automático de dias, tipo de escala e paleta de cores institucional.
       - **Worker de Sincronização em Segundo Plano ([`src/syncs/sync_ferias.py`](automated-OTRS-and-CitSmart/src/syncs/sync_ferias.py)):**
         * Execução assíncrona não-bloqueante com lock file (`ferias_sync.lock`), log dedicado em `debug_logs/ferias/sync_ferias.log` e accordion em tempo real com auto-refresh (`render_log_expander`).
         * Suporte a download HTTP autenticado via cookies do SharePoint, leitura de cópia local sincronizada do OneDrive e contingência via Selenium.
         * Integrado ao daemon agendador de tarefas periódicas ([`src/services/cron_scheduler.py`](automated-OTRS-and-CitSmart/src/services/cron_scheduler.py) e [`src/database/cron_db.py`](automated-OTRS-and-CitSmart/src/database/cron_db.py)).
       - **Interface Streamlit & Sub-abas com URL State (`?tab=ferias&subtab=...`):**
         * **📊 Planilha & Escala:** Tabela completa com filtros por exercício/ano, membro e modalidade, paginação dinâmica (`paginate_items`), exportação CSV e KPIs resumidos (períodos, total de dias, membros e alerta expansível com detecção automática de sobreposições de ausências da equipe).
         * **📅 Calendário de Férias:** Calendário interativo FullCalendar renderizado via [`render_master_calendar()`](automated-OTRS-and-CitSmart/src/components/calendar.py) com cores temáticas por integrante e modal inteligente com ficha completa de detalhes.
         * **Calendário Geral Integrado:** Adicionada camada de Férias da Bancada no [`src/tabs/calendario_geral.py`](automated-OTRS-and-CitSmart/src/tabs/calendario_geral.py) com controle de exibição via checkbox na sidebar.
       - **Navegação Híbrida no Header ([`src/components/header.py`](automated-OTRS-and-CitSmart/src/components/header.py)):**
         * Mapeamento de rota `"🏖️ Férias da Bancada": "ferias"` com navegação interna no menu hambúrguer e atalho `↗` para abertura em nova aba (`target="_blank"`).
       - **Suíte de Testes 100% Verde:**
         * Implementados testes unitários em [`tests/unit/test_ferias.py`](automated-OTRS-and-CitSmart/tests/unit/test_ferias.py).
         * Suíte completa agora com **50 testes automatizados passando com 100% de sucesso** (`python3 tests/run_all.py`).

---

## 📋 5. Próximas Etapas e Melhorias Planejadas

---

### 5.2 Outras Melhorias Planejadas
1. **Sincronização Periódica Automática do Cache do SCCM:**
   - Agendamento da rotina via daemon interno de cron (`cron_scheduler.py`).
2. **Métricas de Acurácia de Localização:**
   - Painel de taxa de correspondência de chamados direcionados por IP vs. NLP textual.




