# 🛠️ Utilitários do Sistema Bancada STI (`util/`)

Este diretório reúne scripts de apoio administrativo, migrações pontuais, rotinas de limpeza de banco e ferramentas de diagnóstico para o sistema de chamados.

Todos os scripts são desenvolvidos e preparados para execução nativa em ambientes **Linux**, **WSL** (Ubuntu 24.04/26.04) e **Red Hat Enterprise Linux (RHEL)** via terminal **Bash** ou diretamente nos contêineres Docker.

---

## 📁 Estrutura de Diretórios

```text
util/
├── CONTEXTO_GERAL.md               # Documentação viva de arquitetura e evolução do projeto
├── README.md                       # Este guia explicativo dos diretórios e scripts
│
├── limpeza_banco/                  # Scripts e queries de higienização de tabelas SQLite
│   ├── limpar_chamados.py          # Expurgador de chamados gerados por usuários automáticos
│   ├── limpar_localidades_nan.py   # Sanitizador de localidades físicas, nulos e reconstrução
│   ├── limpar_localidades_nan.sql  # Script SQL puro equivalente para execução via CLI sqlite3
│   ├── limpar_impressoras.py       # Limpa registros e resíduos de cabeçalho da tabela de impressoras
│   └── limpar_plantoes.py          # Limpa tabelas de plantões matutino e semanal
│
├── migracoes/                      # Scripts para ajuste e migração de dados legados
│   ├── fix_db_dates.py             # Normaliza datas de chamados (ISO, reversão de dia/mês inválidos)
│   ├── fix_db_invert_dates.py      # Inversão cirúrgica de dia e mês para datas no futuro
│   └── migrate_manual_entries.py   # Migra cadastros de prédios de código Python para SQLite
│
├── diagnosticos/                   # Scripts de inspeção, auditoria e testes pontuais
│   ├── inspect_db_dates.py         # Relatório de formatos de data gravados no SQLite
│   └── test_json_events.py         # Valida serialização de eventos JSON para o FullCalendar
│
├── arquivos/                       # Ferramentas auxiliares de conversão e mídia
│   └── converter_pdf.py            # Conversor de páginas de PDFs em PNGs de alta definição
│
└── docs/                           # Documentações e guias explicativos de cada rotina
    ├── limpar_chamados_automaticos.md # Instruções de expurgo de chamados de monitoramento
    ├── limpar_localidade_nan.md       # Dossiê técnico da normalização de localidades físicas
    ├── limpar_tabelas_apoio.md        # Limpeza e reset de impressoras e plantões
    ├── correcao_datas.md              # Normalização e correção cirúrgica de datas
    ├── migracao_unidades_manuais.md   # Migração de dicionários de prédios para SQLite
    ├── validacao_eventos_calendario.md# Validação de eventos JSON para o FullCalendar
    └── conversor_pdf_para_png.md      # Conversão de páginas de PDFs em PNGs de alta definição
```

---

## 🚀 Como Executar os Scripts (Bash / WSL / Linux / RHEL)

Todos os scripts utilizam resolução dinâmica de caminho e podem ser executados a partir da raiz do projeto:

### 1. Limpeza de Banco (`limpeza_banco/`)
- **Remover chamados de monitoramento / robôs de e-mail:**
  ```bash
  python3 util/limpeza_banco/limpar_chamados.py
  ```
- **Higienizar localidades 'nan' e reconstruir títulos e prédios:**
  ```bash
  python3 util/limpeza_banco/limpar_localidades_nan.py
  # Ou via CLI sqlite3:
  sqlite3 chamados.db < util/limpeza_banco/limpar_localidades_nan.sql
  ```
- **Limpar registros de impressoras ou plantões:**
  ```bash
  python3 util/limpeza_banco/limpar_impressoras.py
  python3 util/limpeza_banco/limpar_plantoes.py
  ```

### 2. Migrações e Ajustes de Dados (`migracoes/`)
- **Corrigir datas no futuro ou formatos fora do padrão ISO:**
  ```bash
  python3 util/migracoes/fix_db_dates.py
  python3 util/migracoes/fix_db_invert_dates.py
  ```
- **Migrar dicionários de prédios/unidades para a tabela SQLite:**
  ```bash
  python3 util/migracoes/migrate_manual_entries.py
  ```

### 3. Diagnósticos (`diagnosticos/`)
- **Inspecionar formatos de datas no banco:**
  ```bash
  python3 util/migracoes/inspect_db_dates.py
  ```
- **Testar geração do payload de eventos do calendário:**
  ```bash
  python3 util/diagnosticos/test_json_events.py
  ```

### 4. Execução em Contêiner Docker
Se estiver rodando a aplicação em contêineres Docker, execute qualquer utilitário com:
```bash
docker exec -i $(docker ps -q -f name=bancada_streamlit_app) python util/limpeza_banco/limpar_chamados.py
```
