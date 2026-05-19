# 🤖 Automação e Classificação de Chamados de TI (MPMS)

Este projeto consiste em uma suite de ferramentas desenvolvidas em Python para automatizar a extração, unificação, sincronização e classificação inteligente de chamados de suporte técnico (Manutenção de TI) provenientes de múltiplas plataformas (OTRS e CitSmart).

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

### 5. Gestão Centralizada (Dashboard Web e Banco de Dados)
- **Painel Interativo Premium (Streamlit):** Interface gráfica web lindíssima e responsiva (Hot-Reloading) para acompanhamento dos chamados em tempo real, com ordenação inteligente de datas e filtros dinâmicos de Status, Unidade, Usuário e TAG de IA.
- **Deep Linking Centralizado (Acesso em Um Clique):** Geração dinâmica de URLs diretas para os chamados tanto no OTRS quanto no CitSmart. Exibido de forma ultra-elegante no Streamlit através de `st.column_config.LinkColumn` (coluna "Link Direto") e pelo botão nativo `st.link_button` no modal de detalhes, permitindo abrir o chamado de origem em uma nova aba instantaneamente.
- **Gestão de Visibilidade Dinâmica:** Controle avançado via UI (ex: toggle elegante na barra lateral para mostrar ou ocultar o IP de rede dos usuários sob demanda).
- **Persistência Inteligente (SQLite):** Todo o tráfego de dados gerado é consolidado num banco de dados local leve e ultra-rápido (`chamados.db`), que monitora e gerencia os estados lógicos de "Aberto" e "Fechado" autonomamente conforme chamados novos chegam ou desaparecem da fila de triagem.


---

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3.11+
- **Bibliotecas Principais:**
  - `selenium`: Navegação web automatizada para logins e sistemas dinâmicos.
  - `requests` & `beautifulsoup4`: Varredura estática ultra-veloz de portais institucionais públicos.
  - `pandas`: Análise, manipulação e alinhamento inteligente de DataFrames.
  - `scikit-learn`: Treinamento pesado, tuning de hiperparâmetros e classificação.
  - `spacy`: Processamento de linguagem natural (NLP) e lematização.
  - `pywin32` e `WMI`: Automação nativa do Microsoft Excel e consultas profundas ao servidor SCCM para extração de IPs na rede.
  - `streamlit`: Construção do dashboard web moderno, rápido e interativo.
  - `sqlite3`: Banco de dados relacional embutido de altíssimo desempenho para retenção do ciclo de vida dos chamados.

---

## 📂 Estrutura do Projeto

- `salvar_senha.py`: Utilitário de segurança para salvar e criptografar as credenciais de acesso localmente, evitando senhas expostas no código.
- `citsmart_scraper.py`: Bot para extração do sistema LowCode/CitSmart.
- `otrs_scraper.py`: Bot para extração do sistema legado OTRS.
- `unidades_scraper.py`: Scraper que atualiza a lista de unidades/promotorias do site oficial do MPMS em tempo real com `BeautifulSoup`.
- `preprocess_chamados.py`: Limpeza, padronização, remoção de assinaturas/saudações e unificação das bases brutas.
- `tag_classifier.py`: O "cérebro" da IA. Limpa o texto com NLP, treina os modelos, avalia métricas e classifica os novos chamados.
- `sync_master.py`: Compara os chamados novos com a base de produção, sincroniza os estados "Aberto/Fechado" com o banco de dados SQLite e realiza a inserção *Append-Only* com formatações Win32 na Planilha Master.
- `database.py`: Interface dedicada de controle (DAO) do banco de dados relacional SQLite, responsável por Inserções (UPSERT) dinâmicas e transições de status do ciclo de vida dos chamados.
- `dashboard.py`: Aplicação Web (Frontend em Streamlit) para que os usuários e gestores leiam o banco de dados, apliquem filtros em tempo real e visualizem chamados de forma responsiva.
- `manual_entries.py`: Hub central de regras estáticas (ranges de IP CIDR para rastreamento de prédios geográficos, Regex de NLP e mapeamentos institucionais).
- `config.py`: Central de configurações unificadas, conexões seguras ao AD (LDAP), queries WMI de rede ao SCCM para resolução de IPs e utilitários globais do sistema.
- `orquestrador.py`: Script mestre do fluxo, executado 100% em background pelo Agendador de Tarefas do Windows.
- `test_app.py`: Suíte de testes automatizados unitários para validação de NLP, IA, IP Ranges, limpeza de disco e rotinas críticas do pipeline.

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
   python -m spacy download pt_core_news_sm
   ```

3. **Configure as credenciais criptografadas:**
   Execute o script utilitário para registrar suas chaves e logins com segurança:
   ```bash
   venv\Scripts\python.exe salvar_senha.py
   ```

## 🧪 Executando Testes Unitários

Para garantir que as modificações de fluxo ou de dados não introduziram regressões no pipeline, execute os testes unitários integrados:
```bash
venv\Scripts\python.exe -m unittest test_app.py
```
