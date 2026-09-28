# 🏢 Migração de Cadastros Manuais de Prédios e Unidades

Este documento detalha o processo de migração dos registros estáticos de prédios e unidades mapeados no código Python ([`src/manual_entries.py`](file:///home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart/src/manual_entries.py)) para a tabela relacional `unidades_manuais` no SQLite (`chamados.db`).

---

## 🎯 1. Objetivo da Migração

Inicialmente, os prédios de Campo Grande e comarcas do interior com seus respectivos titulares, siglas, telefones e URLs eram mantidos estaticamente em listas de dicionários dentro do código-fonte.

A migração para a tabela `unidades_manuais`:
1. Permite edição, inclusão e exclusão direta através da interface gráfica (Streamlit) sem necessidade de alterar código.
2. Unifica a busca de localização física e informações de contato em consultas SQL rápidas.

---

## 🚀 2. Como Executar a Migração (`migrate_manual_entries.py`)

O script [`util/migracoes/migrate_manual_entries.py`](file:///home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart/util/migracoes/migrate_manual_entries.py) lê todos os registros de `manual_entries.py` e os insere/atualiza no banco SQLite.

### Executar via Bash:
```bash
python3 util/migracoes/migrate_manual_entries.py
```

### Executar via Docker:
```bash
docker exec -i $(docker ps -q -f name=bancada_streamlit_app) python util/migracoes/migrate_manual_entries.py
```

---

## 🔍 3. Verificação no Banco de Dados

Para validar a quantidade de unidades manuais migradas:
```bash
sqlite3 chamados.db "SELECT count(*) FROM unidades_manuais;"
```
