# Guia de Migração de Dados e Correção de Datas

> **Scripts em:** `util/migracoes/`  
> **Banco:** PostgreSQL & SQLite

---

## 🎯 1. Visão Geral

Este guia orienta a execução dos scripts de migração estrutural, normalização de datas ISO e conversão entre motores de banco de dados.

---

## 🐘 2. Migração do SQLite para PostgreSQL (`migrar_sqlite_para_postgres.py`)

Para transferir integralmente os dados do arquivo SQLite legado (`data/chamados.db`) para o PostgreSQL rodando no contêiner `bancada_postgres_db`:

### Como Executar:
```bash
docker exec -e POSTGRES_HOST=db bancada_streamlit_app python util/migracoes/migrar_sqlite_para_postgres.py
```

### O que o script faz:
1. Conecta-se ao SQLite e ao PostgreSQL.
2. Inspeciona e replica o schema de mais de 35 tabelas.
3. Transfere os dados linha a linha com batch commits.
4. Ajusta as sequences primárias do PostgreSQL para evitar conflitos de ID nas próximas inserções.

---

## 📅 3. Correção e Normalização de Datas (`fix_db_dates.py`)

Chamados legados importados de planilhas Excel antigas podem conter datas no formato brasileiro `DD/MM/YYYY HH:MM:SS` ou com dias e meses invertidos acidentalmente.

### Diagnóstico de Datas (`inspect_db_dates.py`):
```bash
docker exec bancada_streamlit_app python util/diagnosticos/inspect_db_dates.py
```

### Correção Automática (`fix_db_dates.py`):
```bash
docker exec bancada_streamlit_app python util/migracoes/fix_db_dates.py
```
- Converte qualquer data no padrão brasileiro para o padrão ISO internacional (`YYYY-MM-DD HH:MM:SS`).
- Detecta datas no futuro e corrige inversões onde o dia foi gravado no lugar do mês.

---

## 🏢 4. Migração de Cadastros Manuais de Prédios (`migrate_manual_entries.py`)

Transfere dados que estavam definidos estaticamente no código Python para a tabela relacional `unidades_manuais`, permitindo edição via interface gráfica:

```bash
docker exec bancada_streamlit_app python util/migracoes/migrate_manual_entries.py
```

