# 🤖 Automação e Classificação de Chamados de TI (MPMS)

Este projeto consiste em uma suite de ferramentas desenvolvidas em Python para automatizar a extração, unificação e classificação inteligente de chamados de suporte técnico (Manutenção de TI) provenientes de múltiplas plataformas (OTRS e CitSmart).

## 🚀 Funcionalidades

### 1. Web Scraping & Automação (RPA)

- **Selenium WebDriver:** Loga automaticamente nos portais de suporte (OTRS e CitSmart).
- **Extração Robusta:** Lida com paginação dinâmica, máscaras de carregamento (loading masks) e autenticação LDAP.
- **Unificação:** Consolida dados de sistemas legados e novos em um formato padronizado.

### 2. Inteligência Artificial (NLP)

- **Classificação Automática:** Utiliza Machine Learning (`scikit-learn`) para ler a descrição do chamado e predizer a categoria (TAG) correta (ex: "IMPRESSORA", "REDE", "SOFTWARE").
- **Pipeline de NLP:**
  - Limpeza de texto (remoção de stop words, pontuação).
  - Vetorização `TF-IDF`.
  - Modelo `LinearSVC` (Support Vector Machine) otimizado via `GridSearchCV`.

### 3. Engenharia de Dados & Excel

- **Manipulação Avançada:** Uso intensivo de `Pandas` para tratamento de dados.
- **Integração COM (Win32):** Automação do Microsoft Excel para formatação visual, ajuste de colunas (autofit) e sincronização entre planilhas sem corromper a formatação original.

## 🔒 Segurança e Privacidade

Este projeto foi desenhado respeitando normas de segurança institucional:

- **Gestão de Credenciais:** As senhas não ficam no código. É utilizada a biblioteca `keyring` para armazenar e consultar senhas diretamente no Gerenciador de Credenciais criptografado do Windows.
- **Dados Sensíveis:** Scripts com dados reais de unidades e servidores, bem como arquivos de log e planilhas geradas, são estritamente ignorados via `.gitignore`.
- **Execução Stealth (Invisível):** O orquestrador roda via `pythonw.exe` com a flag `CREATE_NO_WINDOW`, garantindo processamento 100% em background, sem roubo de foco do usuário e sem disparar alertas de antivírus.

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3.11+
- **Bibliotecas Principais:**
  - `selenium`: Navegação web automatizada.
  - `pandas`: Análise e manipulação de dados.
  - `scikit-learn`: Treinamento do modelo de classificação.
  - `spacy`: Processamento de linguagem natural (NLP).
  - `pywin32`: Automação nativa do Windows/Excel.

## 📂 Estrutura do Projeto

- `salvar_senha.py`: Utilitário para salvar e criptografar as credenciais de acesso localmente.
- `citsmart_scraper.py`: Bot para extração do sistema LowCode/CitSmart.
- `otrs_scraper.py`: Bot para extração do sistema OTRS.
- `unidades_scraper.py`: Scraper que atualiza a lista de unidades/promotorias do site oficial.
- `preprocess_chamados.py`: Limpeza, padronização e unificação das bases.
- `tag_classifier.py`: O "cérebro" do projeto. Treina a IA e classifica novos chamados.
- `config.py`: Central de configurações e caminhos.
- `orquestrador.py`: Script principal executado em background pelo Agendador de Tarefas do Windows para acionar todos os robôs em sequência.

## 📦 Como Instalar e Configurar

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/rezendepauloh/automacao-chamados-sti](https://github.com/rezendepauloh/automacao-chamados-sti)
   ```
