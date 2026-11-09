# 📚 Documentação Técnica — Sistema Bancada STI

> **Central unificada de documentação de engenharia, arquitetura, manuais operacionais e guias práticos do sistema.**  
> Plataforma integrada de Esteira de Chamados, Triagem com IA, Gestão de Bancada & Patrimônio, Plantões Corporativos e Automação Remota para a **Tecnologia da Informação do Ministério Público do Estado de Mato Grosso do Sul (MPMS)**.

---

## 🗂️ Estrutura e Mapa de Navegação da Documentação

```text
Docs/
├── README.md                                         # Este índice central unificado
│
├── 🏛️ arquitetura/                                   # Decisões fundamentais de engenharia e modelagem
│   ├── dossie_arquitetura_sistema.md                 # Dossiê de contexto global do ecossistema Bancada STI
│   ├── modelagem_banco_de_dados_hibrido.md           # Camada DAO híbrida (PostgreSQL e SQLite), schemas e pools
│   ├── pipeline_ia_classificacao_tags.md             # spaCy batch, modelo campeão LinearSVC e rotina GridSearchCV
│   ├── seguranca_criptografia_e_cofre.md             # Criptografia simétrica Fernet, cofre e fallback gracioso
│   └── automacao_remota_powershell_e_wmi.md          # Protocolo bancada://, scripts UTF-8 BOM e SCCM/WMI
│
├── 📖 guias/                                         # Manuais práticos de operação, instalação e utilitários
│   ├── instalacao_e_deploy_docker.md                 # Deploy oficial via Docker Compose (Streamlit + Postgres)
│   ├── esteira_chamados_e_scrapers.md                # Orquestrador, scraper HTTP veloz OTRS e scraper CitSmart
│   ├── gestao_bancada_redistribuicao_termos.md       # Máquinas, baixas, doações, termos e HTML para OTRS
│   ├── agendador_cron_daemon_e_tarefas.md            # BancadaCronDaemon, tarefas periódicas e controle na UI
│   ├── manutencao_e_limpeza_banco.md                 # Limpeza de chamados automáticos, tabelas de apoio e NaN
│   ├── migracao_dados_e_correcao_datas.md            # Migrações SQLite -> Postgres, datas ISO e unidades manuais
│   └── conversao_arquivos_e_validacao.md             # Conversor PDF->PNG (PyMuPDF) e validador JSON FullCalendar
│
├── 🧪 testes/                                        # Suíte de testes automatizados e integridade
│   └── suite_de_testes_automatizados.md              # Manual do runner master run_all.py (90 testes no Docker)
│
└── 🚨 issues/                                        # Registro de roadmap, pendências e melhorias
    ├── backlog_e_historico_de_issues.md              # Histórico de entregas recentes e roadmap futuro
    └── historico_sprints_legadas.md                  # Registro detalhado das entregas das sprints anteriores (#1 a #12)
```

---

## 🚀 Acessos Rápidos Mais Utilizados

### 1. Operação Técnica Diária
- **Como executar a esteira de chamados:** [Guia da Esteira de Chamados & Scrapers](guias/esteira_chamados_e_scrapers.md)
- **Como gerar relatórios e termos de bancada para o OTRS:** [Guia de Gestão de Bancada & Redistribuição](guias/gestao_bancada_redistribuicao_termos.md)
- **Como configurar rotinas periódicas de background:** [Guia do Cron Daemon](guias/agendador_cron_daemon_e_tarefas.md)
- **Como converter anexos PDF em PNG:** [Guia de Utilitários e Conversão](guias/conversao_arquivos_e_validacao.md)

### 2. Engenharia e Arquitetura de Software
- **Visão macro e decisões técnicas do sistema:** [Dossiê Global de Arquitetura](arquitetura/dossie_arquitetura_sistema.md)
- **Como funciona a camada de banco híbrido (PostgreSQL/SQLite):** [Modelagem de Banco de Dados](arquitetura/modelagem_banco_de_dados_hibrido.md)
- **Como a IA pré-processa e classifica os chamados:** [Pipeline de Inteligência Artificial & TAGs](arquitetura/pipeline_ia_classificacao_tags.md)
- **Políticas de criptografia e fallback seguro de senhas:** [Segurança & Cofre de Credenciais](arquitetura/seguranca_criptografia_e_cofre.md)
- **Integração do protocolo remoto e scripts PowerShell:** [Automação Remota PowerShell & WMI](arquitetura/automacao_remota_powershell_e_wmi.md)

### 3. Instalação, Infraestrutura & Qualidade
- **Como rodar a suíte de testes (90 testes com 100% de sucesso):** [Manual de Testes Automatizados](testes/suite_de_testes_automatizados.md)
- **Como subir o ambiente completo via Docker Compose:** [Guia de Instalação e Deploy Docker](guias/instalacao_e_deploy_docker.md)
- **Como migrar a base SQLite para o PostgreSQL:** [Guia de Migração de Dados](guias/migracao_dados_e_correcao_datas.md)
- **Como realizar expurgo e limpeza de dados legados:** [Guia de Manutenção e Limpeza do Banco](guias/manutencao_e_limpeza_banco.md)
- **Acompanhamento do roadmap e entregas:** [Backlog e Histórico de Issues](issues/backlog_e_historico_de_issues.md)

