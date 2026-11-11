# Modelagem do Banco de Dados Híbrido (PostgreSQL & SQLite)

> **Módulo:** `src/database/`  
> **Gerenciador de Conexão:** [`src/database/connection.py`](../../src/database/connection.py)

---

## 🎯 1. Filosofia da Camada DAO Híbrida

O Sistema Bancada STI adota uma arquitetura de banco de dados híbrida para conciliar robustez de nível empresarial com facilidade de desenvolvimento local:

1. **PostgreSQL 15 (Padrão em Produção/Docker):**
   - Roda no contêiner dedicado `bancada_postgres_db`.
   - Utiliza pool de conexões gerenciado via `psycopg2.pool.SimpleConnectionPool` para alta concorrência sem estouro de conexões.
   - Suporta concorrência real entre o dashboard Streamlit, o Cron Daemon e o Orquestrador.
2. **SQLite 3 (Fallback & Contingência Local):**
   - Arquivo `data/chamados.db`, criado automaticamente sob demanda apenas quando `DB_TYPE=sqlite`.
   - Utilizado de forma transparente caso o PostgreSQL esteja inacessível ou se `DB_TYPE=sqlite` for explicitamente configurado no `.env`.
   - O wrapper `get_db_connection()` detecta o driver e retorna uma conexão compatível.

> **Bootstrap Automático em Novas Instalações:** Na inicialização do contêiner (`init.py`), a função `initialize_database()` (`src/database/bootstrap.py`) aguarda o PostgreSQL ficar saudável e cria de forma idempotente todas as tabelas do sistema, configurações padrão (`.env`), agendamentos de cron e índices de performance.

---

## 🏛️ 2. Gerenciamento de Conexão (`connection.py`)

A função mestre `get_db_connection()` encapsula o ciclo de vida:

```python
from src.database.connection import get_db_connection

with get_db_connection() as conn:
    cur = conn.cursor()
    cur.execute("SELECT count(*) FROM chamados")
    total = cur.fetchone()[0]
```

### Configurações de Ambiente (`.env`):
- `DB_TYPE`: `postgres` (padrão) ou `sqlite`.
- `POSTGRES_HOST`: `db` (dentro da rede Docker Compose) ou `localhost`.
- `POSTGRES_PORT`: `5432`.
- `POSTGRES_DB`: `bancada_mpms`.
- `POSTGRES_USER`: `bancada_user`.
- `POSTGRES_PASSWORD`: Senha cifrada via Fernet ou plana no `.env`.

---

## 📊 3. Principais Tabelas do Esquema

O banco gerencia mais de 35 tabelas relacionais organizadas por domínio de negócio:

### 3.1 Tabela Central de Chamados (`chamados`)
Armazena o histórico e estado de todos os chamados consolidados da TI:
- `id` (VARCHAR/TEXT, PK): Identificador unificado do chamado (ex: `113099`).
- `ticket_number` (VARCHAR): Número amigável de protocolo.
- `sistema_origem` (VARCHAR): `OTRS` ou `CitSmart`.
- `titulo` (TEXT): Título/assunto da solicitação.
- `descricao` (TEXT): Conteúdo descritivo inicial.
- `comentarios` (TEXT): JSON/texto com notas técnicas live e interações.
- `cliente` (VARCHAR): Nome do solicitante/servidor.
- `unidade` (VARCHAR): Unidade/setor institucional solicitante.
- `cidade_predio` (VARCHAR): Localidade física resolvida.
- `tag` (VARCHAR): Categoria predita pelo modelo campeão de IA ou ajustada manualmente.
- `estado` (VARCHAR): `Aberto`, `Fechado com Sucesso`, `Pendente`, etc.
- `data_criacao` (TIMESTAMP): Data ISO de abertura.
- `data_fechamento` (TIMESTAMP): Data ISO de encerramento.

### 3.2 Tabela de Agendamentos do Cron (`cron_schedules`)
Controla o agendador em segundo plano (`BancadaCronDaemon`):
- `task_id` (VARCHAR, PK): Identificador único da rotina (ex: `retrain_ia_gridsearch`, `sync_sccm`).
- `nome` (VARCHAR): Rótulo amigável exibido na aba de configurações.
- `categoria` (VARCHAR): Agrupador funcional (`Inteligência Artificial`, `Inventário`, `Plantões`).
- `intervalo_valor` (INTEGER): Valor do ciclo de disparo.
- `intervalo_unidade` (VARCHAR): `minutos`, `horas`, `dias`.
- `ativo` (BOOLEAN): Estado liga/desliga da rotina.
- `ultima_execucao` (TIMESTAMP): Registro de término do último ciclo.
- `status_execucao` (VARCHAR): `Sucesso`, `Executando`, `Falha`.
- `detalhes_execucao` (TEXT): Log textual com sumário da última execução.

### 3.3 Tabelas de Apoio e Cadastros Corporativos
- `unidades_manuais`: Cadastro de prédios, comarcas, siglas e contatos do interior.
- `ramais_mpms`: Catálogo de ramais telefônicos e titulares de cada setor.
- `plantoes_matutino` e `plantoes_semanal`: Escalas oficiais de sobreaviso da TI.
- `sccm_workstations`: Cache de estações inventariadas, endereços IP, build do Windows e último login.
- `garantia_equipamentos` e `garantia_chamados`: Controle de contratos de garantia e chamados abertos em fabricantes.
- `redistribuicao_maquinas`: Controle de estações separadas para doação ou descarte patrimonial.
- `system_settings`: Tabela de chave/valor para parâmetros gerais do sistema.

---

## ⚡ 4. Índices de Otimização de Performance

O script `ensure_database_indexes()` em `src/database/tickets_db.py` valida e garante índices essenciais em tempo de inicialização:

```sql
CREATE INDEX IF NOT EXISTS idx_chamados_tag ON chamados (tag);
CREATE INDEX IF NOT EXISTS idx_chamados_estado ON chamados (estado);
CREATE INDEX IF NOT EXISTS idx_chamados_data_criacao ON chamados (data_criacao);
CREATE INDEX IF NOT EXISTS idx_chamados_sistema_origem ON chamados (sistema_origem);
CREATE INDEX IF NOT EXISTS idx_chamados_cidade_predio ON chamados (cidade_predio);
```

---

## 🔄 5. Scripts de Migração (`util/migracoes/`)

Para mover bases e atualizar schemas com segurança:
- `util/migracoes/migrar_sqlite_para_postgres.py`: Migrador universal que lê todas as tabelas do `data/chamados.db` SQLite e insere em lote no PostgreSQL com reconstrução de sequências e índices.
- `util/migracoes/fix_db_dates.py`: Repara datas legadas no formato brasileiro ou com dia/mês invertidos.
- `util/migracoes/migrate_manual_entries.py`: Migra dados estáticos de código para a tabela relacional `unidades_manuais`.

