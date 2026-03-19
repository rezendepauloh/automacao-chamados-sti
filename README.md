# 🤖 Automação e Classificação de Chamados de TI (MPMS)

Este projeto consiste em uma suite de ferramentas desenvolvidas em Python para automatizar a extração, unificação, sincronização e classificação inteligente de chamados de suporte técnico (Manutenção de TI) provenientes de múltiplas plataformas (OTRS e CitSmart).

## 🚀 Funcionalidades

### 1. Web Scraping & Automação (RPA)
- **Selenium WebDriver:** Loga automaticamente nos portais de suporte (OTRS e CitSmart).
- **Extração Robusta:** Lida com paginação dinâmica, máscaras de carregamento (*loading masks*) e autenticação LDAP.
- **Unificação:** Consolida dados de sistemas legados e novos em um formato tabular padronizado.

### 2. Inteligência Artificial (NLP) e Machine Learning Contínuo
- **Classificação Automática:** Utiliza IA para ler a descrição do chamado e predizer a categoria (TAG) correta (ex: "IMPRESSORA", "REDE", "SOFTWARE").
- **Pipeline de NLP Especializado em TI:** - Limpeza de texto avançada com `spaCy` (remoção de stop words, pontuação).
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
  - `selenium`: Navegação web automatizada.
  - `pandas`: Análise, manipulação e alinhamento inteligente de DataFrames.
  - `scikit-learn`: Treinamento pesado, tuning de hiperparâmetros e classificação.
  - `spacy`: Processamento de linguagem natural (NLP) e lematização.
  - `pywin32`: Automação nativa e formatação do Microsoft Excel.

---

## 📂 Estrutura do Projeto

- `salvar_senha.py`: Utilitário de segurança para salvar e criptografar as credenciais de acesso localmente, evitando senhas expostas no código.
- `citsmart_scraper.py`: Bot para extração do sistema LowCode/CitSmart.
- `otrs_scraper.py`: Bot para extração do sistema legado OTRS.
- `unidades_scraper.py`: Scraper que atualiza a lista de unidades/promotorias do site oficial do MPMS.
- `preprocess_chamados.py`: Limpeza, padronização, remoção de assinaturas/saudações e unificação das bases brutas.
- `tag_classifier.py`: O "cérebro" da IA. Limpa o texto com NLP, treina os modelos, avalia métricas e classifica os novos chamados.
- `sync_master.py`: O maestro da integração. Compara os chamados novos com a base de produção, faz a inserção segura (*Append-Only*) na Planilha Master e aplica a formatação visual (Win32 COM) de forma invisível.
- `config.py`: Central de configurações, variáveis de ambiente e mapeamento de caminhos.
- `orquestrador.py`: Script principal executado em background pelo Agendador de Tarefas do Windows.

## 📦 Como Instalar e Configurar

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/rezendepauloh/automacao-chamados-sti](https://github.com/rezendepauloh/automacao-chamados-sti)
   ```
