# Guia de Manutenção e Limpeza do Banco de Dados

> **Scripts em:** `util/limpeza_banco/`  
> **Suporte:** PostgreSQL e SQLite

---

## 🎯 1. Visão Geral

Com o uso contínuo do sistema, certas tabelas podem acumular ruídos ou registros automáticos que afetam relatórios e o treinamento da inteligência artificial. Este guia reúne os utilitários de limpeza e sanitização do banco de dados.

---

## 🧹 2. Remoção de Chamados Automáticos (`limpar_chamados.py`)

No OTRS, alguns e-mails e alertas automáticos geram chamados espúrios de robôs internos (ex: *"Monitoramento Adm Mpms"*, *"Adm Ticket Por Email"*).

O script `util/limpeza_banco/limpar_chamados.py` purga esses registros da base para evitar distorções nas métricas da equipe.

### Como Executar:
```bash
docker exec bancada_streamlit_app python util/limpeza_banco/limpar_chamados.py
```

---

## 🖨️ 3. Limpeza de Tabelas de Apoio (Impressoras e Plantões)

Em casos de importação corrompida de CSVs do PaperCut ou sincronização truncada de plantões, é recomendado resetar as tabelas para permitir uma nova carga limpa:

### Resetar Tabela de Impressoras (`limpar_impressoras.py`):
```bash
docker exec bancada_streamlit_app python util/limpeza_banco/limpar_impressoras.py
```

### Resetar via SQL Direto no PostgreSQL:
```bash
docker exec -i bancada_postgres_db psql -U bancada_user -d bancada_mpms -c "TRUNCATE impressoras CASCADE;"
```

---

## 🧼 4. Sanitização de Valores `NaN` e Localidades Nulas

Se houver registros antigos com string `"nan"` ou `"não encontrado"` gravados nas colunas `unidade`, `cidade_predio` ou `localidade_fisica`, execute a rotina de saneamento para normalizá-los como `NULL` ou vazio:

```sql
UPDATE chamados 
SET cidade_predio = '' 
WHERE LOWER(cidade_predio) IN ('nan', 'none', 'não encontrado', 'nao encontrado');

UPDATE chamados 
SET unidade = '' 
WHERE LOWER(unidade) IN ('nan', 'none');
```

