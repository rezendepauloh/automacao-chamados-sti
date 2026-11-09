# Guia da Suíte de Testes Automatizados e Integridade

> **Runner Mestre:** [`tests/run_all.py`](../../tests/run_all.py)  
> **Helper Universal de Mocks:** [`tests/test_helpers.py`](../../tests/test_helpers.py)

---

## 🎯 1. Filosofia de Testes

Para garantir estabilidade contínua e evitar regressões nas operações críticas da TI do MPMS:
- **Critério de Aceite Estrito:** 100% de sucesso (0 falhas e 0 erros).
- **Ambiente Canônico:** Os testes devem ser executados prioritariamente dentro do contêiner Docker `bancada_streamlit_app`.
- **Isolamento e Mocks Transparentes:** `tests/test_helpers.py` provê mocks universais seguros para dependências que possam não estar presentes fora do Docker, permitindo testes isolados e determinísticos.

---

## 🚀 2. Como Executar a Suíte Completa

Para rodar todos os testes automatizados com relatório formatado e estatísticas de tempo:

```bash
docker exec bancada_streamlit_app python tests/run_all.py
```

### Exemplo de Saída Esperada:
```text
╔═══════════════════════════════════════════════════════════════╗
║      SISTEMA BANCADA STI — SUÍTE DE TESTES AUTOMATIZADOS      ║
╚═══════════════════════════════════════════════════════════════╝

▶ Executando Suíte: 💾 Banco de Dados, Criptografia & Regras Unitárias (/app/tests/unit)
...
=================================================================
RESUMO GERAL DA VALIDAÇÃO DO SISTEMA BANCADA STI:
• Total de Testes Executados : 90
• Sucesso Total              : 90 (100%)
• Falhas                     : 0
• Erros                      : 0
• Duração Total              : ~25s
=================================================================

🎉 PARABÉNS! TODOS OS TESTES PASSARAM COM 100% DE SUCESSO!
```

---

## 📂 3. Organização dos Módulos de Teste

A suíte está dividida em 5 categorias funcionais:

### 1. `tests/unit/` (Testes Unitários de Regras de Negócio e DAOs)
- `test_ml_champion_pipeline.py`: Testa o modelo campeão de IA (`LinearSVC`), hiperparâmetros do TF-IDF, pré-processamento spaCy em lote, enriquecimento de features e registro no Cron.
- `test_crypto.py`: Criptografia simétrica Fernet, decriptografia, fallback com log para chaves incompatíveis e mascaramento seguro de senhas.
- `test_database.py`: Operações CRUD, índices de banco, sanitização de `NaN` e enriquecimento de dados.
- `test_auth.py`: Validação de tokens de autenticação, RBAC e permissões de perfil.
- `test_ferias.py`: Parser de períodos de férias, detecção de sobreposição de datas e regras de plantão.
- `test_senhas.py`: CRUD do cofre de senhas de sistemas e auditoria.
- `test_postgres_compatibility.py`: Compatibilidade cruzada de queries SQL entre SQLite e PostgreSQL.

### 2. `tests/components/` (Componentes Streamlit e UI)
- `test_components.py`: Valida geração de relatórios HTML de preparação para OTRS, cálculo de rotas Dijkstra, formatação de datas e banners de status.

### 3. `tests/scripts/` (Scripts do Sistema e PowerShell)
- `test_scripts.py`: Valida existência de scripts, queries do SCCM e conformidade obrigatória com encoding UTF-8 com BOM (`utf-8-sig`) nos arquivos `.ps1`.

### 4. `tests/services/` (Serviços e Background Workers)
- `test_cron_scheduler.py`: Valida ciclo de vida da thread do `BancadaCronDaemon`, agendamentos e persistência no banco.

### 5. `tests/integration/` (Testes de Integração e Ambiente)
- `test_integration.py`: Valida sincronia entre chaves do `.env.example` e `config.py`, além de exportações de calendários no formato ICS.

---

## 🧪 4. Como Executar um Teste Unitário Específico

Caso esteja desenvolvendo ou refatorando um único módulo:

```bash
docker exec bancada_streamlit_app python -m unittest tests/unit/test_ml_champion_pipeline.py
docker exec bancada_streamlit_app python -m unittest tests/unit/test_crypto.py
```

