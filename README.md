# 🤖 Automação e Classificação de Chamados de TI (MPMS)

Este projeto consiste em uma suite de ferramentas desenvolvidas em Python para automatizar a extração, unificação, sincronização e classificação inteligente de chamados de suporte técnico (Manutenção de TI) provenientes de múltiplas plataformas (OTRS e CitSmart).

## 🚀 Funcionalidades

### 1. Web Scraping & Automação (RPA)
- **Extração Robusta & Híbrida (Selenium + Captura XHR):** Loga nos portais de suporte (OTRS e CitSmart) via Selenium. No CitSmart, utiliza interceptação de rede (XHR/Rede) para capturar a listagem de chamados em JSON diretamente da API interna do portal, processando as informações instantaneamente com velocidade imbatível.
- **BeautifulSoup & Requests (unidades_scraper.py):** Totalmente migrado de Selenium para uma arquitetura nativa ultra-leve com `requests` e `BeautifulSoup`. Realiza uma varredura em tempo real de todas as ~280 páginas de promotorias e procuradorias do site do MPMS em menos de 4 minutos, obtendo o prédio físico exato de cada promotoria com 100% de atualização e frescor para as equipes de suporte presencial.
- **Cache Inteligente de Descrições e Unidades:** Tanto no OTRS quanto no CitSmart, o robô carrega as descrições e unidades mapeadas na execução anterior. Isso reduz cliques desnecessários de navegação para chamados recorrentes e pula dezenas de buscas lentas no Active Directory (LDAP), reduzindo tempos de resposta.
- **Autolimpeza Preventiva de Disco:** Função automatizada que monitora e mantém no máximo os 10 arquivos brutos, tratados e classificados mais recentes de cada etapa (OTRS, CitSmart, Unificados e Tagged), evitando acúmulo desnecessário de dados e protegendo arquivos críticos como a base de dados de IA (`Chamados_Treino.xlsx`).
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

---

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3.11+
- **Bibliotecas Principais:**
  - `selenium`: Navegação web automatizada para logins e sistemas dinâmicos.
  - `requests` & `beautifulsoup4`: Varredura estática ultra-veloz de portais institucionais públicos.
  - `pandas`: Análise, manipulação e alinhamento inteligente de DataFrames.
  - `scikit-learn`: Treinamento pesado, tuning de hiperparâmetros e classificação.
  - `spacy`: Processamento de linguagem natural (NLP) e lematização.
  - `pywin32`: Automação nativa e formatação do Microsoft Excel.

---

## 📂 Estrutura do Projeto

- `salvar_senha.py`: Utilitário de segurança para salvar e criptografar as credenciais de acesso localmente, evitando senhas expostas no código.
- `citsmart_scraper.py`: Bot para extração do sistema LowCode/CitSmart.
- `otrs_scraper.py`: Bot para extração do sistema legado OTRS.
- `unidades_scraper.py`: Scraper que atualiza a lista de unidades/promotorias do site oficial do MPMS em tempo real com `BeautifulSoup`.
- `preprocess_chamados.py`: Limpeza, padronização, remoção de assinaturas/saudações e unificação das bases brutas.
- `tag_classifier.py`: O "cérebro" da IA. Limpa o texto com NLP, treina os modelos, avalia métricas e classifica os novos chamados.
- `sync_master.py`: O maestro da integração. Compara os chamados novos com a base de produção, faz a inserção segura (*Append-Only*) na Planilha Master e aplica a formatação visual (Win32 COM) de forma invisível.
- `config.py`: Central de configurações, variáveis de ambiente, gerenciamento de caminhos e funções de infraestrutura unificadas (inicializador de navegadores, consulta segura ao AD, gravação formatada de tabelas no Excel, rotinas de autolimpeza de arquivos e logs).
- `orquestrador.py`: Script principal executado em background pelo Agendador de Tarefas do Windows.
- `test_app.py`: Suíte de testes automatizados unitários para validação de algoritmos de processamento de texto, IA, limpeza de disco e raspagem nativa de unidades.

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
