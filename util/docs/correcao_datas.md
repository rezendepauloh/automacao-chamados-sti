# 📅 Normalização e Correção de Datas de Chamados

Este documento orienta o uso dos scripts responsáveis por diagnosticar e reparar datas mal formatadas ou invertidas (dia/mês trocados) na tabela `chamados` do SQLite (`chamados.db`).

---

## 🔍 1. Diagnóstico de Formatos de Data (`inspect_db_dates.py`)

O script [`util/diagnosticos/inspect_db_dates.py`](file:///home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart/util/diagnosticos/inspect_db_dates.py) inspeciona as primeiras e últimas 20 datas gravadas no banco e totaliza a quantidade de registros em formato ISO (`YYYY-MM-DD`), formato brasileiro (`DD/MM/YYYY`) ou outros/inválidos.

### Como Executar via Bash:
```bash
python3 util/diagnosticos/inspect_db_dates.py
```

---

## 🛠️ 2. Correção de Datas Inválidas e Formato Brasileiro (`fix_db_dates.py`)

O script [`util/migracoes/fix_db_dates.py`](file:///home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart/util/migracoes/fix_db_dates.py):
1. Converte qualquer data no padrão brasileiro `DD/MM/YYYY HH:MM:SS` para o padrão ISO `YYYY-MM-DD HH:MM:SS`.
2. Detecta datas no futuro e corrige inversões acidentais onde o dia foi gravado no lugar do mês.

### Como Executar via Bash:
```bash
python3 util/migracoes/fix_db_dates.py
```

---

## 🔄 3. Inversão Cirúrgica de Dia e Mês no Futuro (`fix_db_invert_dates.py`)

O script [`util/migracoes/fix_db_invert_dates.py`](file:///home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart/util/migracoes/fix_db_invert_dates.py) atua de forma cirúrgica apenas nas datas em formato ISO `YYYY-MM-DD` que resultaram em datas futuras devido à troca de mês/dia (ex: `2026-12-08` quando deveria ser `2026-08-12`).

### Como Executar via Bash:
```bash
python3 util/migracoes/fix_db_invert_dates.py
```
