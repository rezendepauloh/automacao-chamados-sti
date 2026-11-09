# 🛠️ Utilitários do Sistema Bancada STI (`util/`)

Este diretório reúne scripts de apoio operacional, rotinas de migração de banco de dados, expurgo de registros órfãos, ferramentas de diagnóstico e benchmarks de performance.

> 📚 **Documentação Técnica Completa:**  
> Os manuais operacionais e guias detalhados de cada utilitário estão unificados em [**`Docs/guias/`**](../Docs/guias).

---

## 📁 Estrutura de Diretórios e Utilitários

```text
util/
├── README.md                       # Este guia de referência rápida dos utilitários
├── comparar_scrapers.py            # Benchmark de performance: OTRS HTTP Fast vs Selenium
├── schema_sqlite.sql               # Esquema DDL de tabelas para banco SQLite local
│
├── 🧹 limpeza_banco/               # Scripts e queries de higienização de dados
│   ├── limpar_chamados.py          # Expurgador de chamados gerados por robôs/e-mails automáticos
│   ├── limpar_localidades_nan.py   # Sanitizador de localidades físicas e remoção de strings 'nan'
│   ├── limpar_localidades_nan.sql  # Query SQL equivalente para execução direta via psql ou sqlite3
│   ├── limpar_impressoras.py       # Limpeza e reset da tabela de impressoras (PaperCut)
│   └── limpar_plantoes.py          # Limpeza e reset das tabelas de plantão matutino e semanal
│
├── 🚚 migracoes/                   # Scripts de migração de esquema e normalização de dados
│   ├── migrar_sqlite_para_postgres.py # Migrador universal: transfere todas as tabelas SQLite -> PostgreSQL
│   ├── fix_db_dates.py             # Normaliza datas de chamados para o formato ISO internacional
│   ├── fix_db_invert_dates.py      # Inversão cirúrgica de dia e mês para datas gravadas no futuro
│   └── migrate_manual_entries.py   # Migra cadastros de prédios do código Python para o banco
│
├── 🔍 diagnosticos/                # Ferramentas de auditoria e validação
│   ├── inspect_db_dates.py         # Relatório de diagnóstico dos formatos de data gravados no banco
│   └── test_json_events.py         # Validador de serialização JSON dos eventos do FullCalendar
│
├── 📄 arquivos/                    # Processamento de arquivos e documentos
│   └── converter_pdf.py            # Conversor de páginas de PDFs em imagens PNG de alta resolução
│
└── 🌐 rede_firewall/               # Automação de conectividade de rede
    ├── configurar_portproxy_firewall.ps1 # Configura redirecionamento de portas e regras do Windows Firewall
    └── README.md                   # Guia de configuração de portas e firewall para WSL2
```

---

## 🚀 Como Executar os Scripts

### 🐳 Execução Canônica via Contêiner Docker (Recomendado)
A partir da raiz do repositório, execute qualquer utilitário dentro do contêiner da aplicação:

```bash
# 1. Migração completa do SQLite para PostgreSQL:
docker exec -e POSTGRES_HOST=db bancada_streamlit_app python util/migracoes/migrar_sqlite_para_postgres.py

# 2. Benchmark comparativo entre scrapers do OTRS:
docker exec bancada_streamlit_app python util/comparar_scrapers.py

# 3. Limpeza de chamados automáticos:
docker exec bancada_streamlit_app python util/limpeza_banco/limpar_chamados.py

# 4. Correção e normalização de datas:
docker exec bancada_streamlit_app python util/migracoes/fix_db_dates.py

# 5. Conversão de PDFs de uploads para PNG:
docker exec bancada_streamlit_app python util/arquivos/converter_pdf.py
```

### 💻 Execução Direta no Terminal (Host Linux / WSL)
Todos os utilitários contam com resolução dinâmica de caminhos (`sys.path`), podendo ser executados diretamente:

```bash
python3 util/migracoes/fix_db_dates.py
python3 util/diagnosticos/inspect_db_dates.py
```

---

## 📖 Guias Relacionados na Central de Documentação

- **Manutenção e Limpeza:** [`Docs/guias/manutencao_e_limpeza_banco.md`](../Docs/guias/manutencao_e_limpeza_banco.md)
- **Migração e Normalização de Datas:** [`Docs/guias/migracao_dados_e_correcao_datas.md`](../Docs/guias/migracao_dados_e_correcao_datas.md)
- **Conversão de Mídias e Validação:** [`Docs/guias/conversao_arquivos_e_validacao.md`](../Docs/guias/conversao_arquivos_e_validacao.md)
- **Esteira de Chamados e Scrapers:** [`Docs/guias/esteira_chamados_e_scrapers.md`](../Docs/guias/esteira_chamados_e_scrapers.md)
