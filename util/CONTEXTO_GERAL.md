# 🧭 Contexto Geral do Sistema Bancada — Automação de Chamados STI

> **Documento de Contexto Vivo & Roteiro Evolutivo**  
> **Última Atualização:** 25/09/2026  
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
  4. **Scripts de Limpeza e Migração para Produção ([`util/limpar_localidades_nan.sql`](file:///home/paulo/PythonProjects/automacao-chamados-sti/util/limpar_localidades_nan.sql) e [`util/limpar_localidades_nan.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/util/limpar_localidades_nan.py)):**
     - Criados scripts SQL e Python prontos para execução em ambientes de homologação e produção, higienizando chamados existentes no banco `chamados.db`.
  6. **Padronização de Prédios de Campo Grande sem Sub-setores ([`src/manual_entries.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/manual_entries.py)):**
     - Faixas de IP da PGJ (ex: `10.111.144.0/24` e `10.111.145.0/24`) foram padronizadas para `"Campo Grande - PGJ"` em vez de acrescentar sufixos de setores (`" - CI"` ou `" - STI"`), mantendo a localidade física limpa e coerente com a lista de prédios.
     - Atualizados os scripts de limpeza (`limpar_localidades_nan.sql` e `limpar_localidades_nan.py`) para normalizar chamados legados com sufixos duplicados.
  7. **Padronização de Comarcas do Interior ([`src/tag_classifier.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/tag_classifier.py)):**
     - A lógica de fallback agora atribui diretamente a comarca/cidade limpa à `Localidade física` (ex: `Terenos`, `Paranaíba`, `Cassilândia`, `Ivinhema`) sem concatenar o nome da promotoria interna (ex: `1ª PJ de Terenos`), deixando a especificação da PJ exclusivamente no campo e filtro `Unidade`.

### ✅ Concluído: Geração Inteligente de Títulos para Chamados sem Título (CitSmart)
- **Status:** **Resolvido e Validado (100% dos 33 testes aprovados)**.
- **Entregas Realizadas:**
  1. **Análise de Descrição e Síntese de Título ([`src/tag_classifier.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/tag_classifier.py)):**
     - Implementada a função [`generate_synthetic_title(tag, description)`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/tag_classifier.py), que combina a `TAG` prevista pelo modelo Scikit-Learn (SVM / ComplementNB) com processamento de linguagem natural (NLP).
     - Remove tags HTML, entidades, saudações compostas ("Bom dia Prezados", "Olá tudo bem"), e preâmbulos burocráticos ("Por determinação do Promotor...", "solicito a...", "venho por meio deste solicitar...").
     - Isola a oração substantiva do problema ou pedido, compondo títulos objetivos como `[IMPRESSORA] Instalação da impressora PRT-5394 no setor` ou `[INSTALAÇÃO HARDWARE] Instalação do Microcomputador Pat. 80700`.
  2. **Preservação de Títulos Nativos e Edições Manuais:**
     - [`generate_missing_titles(df)`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/tag_classifier.py) atua exclusivamente em registros onde a coluna `Título` está vazia ou nula, garantindo que títulos nativos do OTRS permaneçam intactos.
     - [`save_tickets_to_db()`](file:///home/paulo/PythonProjects/automacao-chamados-sti/src/database/tickets_db.py) verifica se o chamado já possui título (ou edição manual salva via UI modal) para nunca sobrescrevê-lo com dados em branco.
  3. **Migração e Atualização da Base:**
     - Integrado ao utilitário [`util/limpar_localidades_nan.py`](file:///home/paulo/PythonProjects/automacao-chamados-sti/util/limpar_localidades_nan.py) para que todas as bases (inclusive em produção) possam preencher os títulos vazios de forma retroativa.

### 📋 Próximas Etapas e Melhorias Planejadas
1. **Sincronização Periódica do Cache do SCCM:**
   - Agendamento de rotina periódica no daemon cron interno (`cron_scheduler.py`) para atualizar automaticamente `sccm_cache_devices` com novas estações e logons.
2. **Histórico de Máquinas do Usuário:**
   - Possibilidade de exibir no modal se o solicitante possui mais de uma estação mapeada (ex: notebook corporativo + desktop da mesa).
3. **Métricas de Acurácia de Localização:**
   - Painel analítico exibindo a taxa de correspondência de chamados direcionados por IP vs. NLP textual.

