# 🧹 Limpeza e Reset de Tabelas de Apoio (Impressoras e Plantões)

Este documento orienta a execução dos scripts de limpeza das tabelas `impressoras`, `plantoes_matutino` e `plantoes_semanal` no SQLite (`chamados.db`).

Esses utilitários são úteis quando há falha de importação de CSVs corrompidos do PaperCut ou sincronização truncada de plantões do SIMP, permitindo resetar as tabelas antes de uma nova coleta.

---

## 🖨️ 1. Limpeza da Tabela de Impressoras (`limpar_impressoras.py`)

O script [`util/limpeza_banco/limpar_impressoras.py`](automated-OTRS-and-CitSmart/util/limpeza_banco/limpar_impressoras.py) apaga todos os registros da tabela `impressoras` para permitir uma reimportação limpa a partir do PaperCut.

### Como Executar via Bash:
```bash
python3 util/limpeza_banco/limpar_impressoras.py
```

### Executar via Docker:
```bash
docker exec -i $(docker ps -q -f name=bancada_streamlit_app) python util/limpeza_banco/limpar_impressoras.py
```

### Executar via SQL direto (CLI sqlite3):
```bash
sqlite3 chamados.db "DELETE FROM impressoras;"
```

---

## 📅 2. Limpeza das Tabelas de Plantões (`limpar_plantoes.py`)

O script [`util/limpeza_banco/limpar_plantoes.py`](automated-OTRS-and-CitSmart/util/limpeza_banco/limpar_plantoes.py) remove todos os registros das tabelas `plantoes_matutino` e `plantoes_semanal`.

### Como Executar via Bash:
```bash
python3 util/limpeza_banco/limpar_plantoes.py
```

### Executar via Docker:
```bash
docker exec -i $(docker ps -q -f name=bancada_streamlit_app) python util/limpeza_banco/limpar_plantoes.py
```

### Executar via SQL direto:
```bash
sqlite3 chamados.db "DELETE FROM plantoes_matutino; DELETE FROM plantoes_semanal;"
```
