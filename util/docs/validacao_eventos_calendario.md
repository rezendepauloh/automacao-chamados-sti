# 📅 Validação de Eventos do Calendário FullCalendar

Este documento orienta o teste e validação da estrutura JSON que alimenta o calendário master interativo da Bancada STI (baseado em **FullCalendar v6**).

---

## 🎯 1. Objetivo do Teste (`test_json_events.py`)

A aba de Calendário Geral agrega em uma única linha do tempo:
- Chamados de TI (com data de criação ISO, autor, localidade e histórico de notas);
- Escalas de plantão matutino e semanal;
- Férias, garantias e viagens técnicas.

O script [`util/diagnosticos/test_json_events.py`](file:///home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart/util/diagnosticos/test_json_events.py) simula a carga de chamados do banco de dados SQLite e valida se todos os atributos esperados pelo frontend JavaScript do FullCalendar (como `id`, `title`, `start`, `backgroundColor`, `extendedProps`) estão devidamente formatados e sem erros de serialização JSON.

---

## 🚀 2. Como Executar via Bash

Execute a partir da raiz do projeto:
```bash
python3 util/diagnosticos/test_json_events.py
```

### Executar dentro do contêiner Docker:
```bash
docker exec -i $(docker ps -q -f name=bancada_streamlit_app) python util/diagnosticos/test_json_events.py
```

O script imprimirá um resumo com a quantidade de eventos gerados com sucesso e exibirá a prévia do primeiro evento JSON formatado.
