import os
import sqlite3
import logging
import warnings
from pathlib import Path

# Suprime aviso informativo do pandas ao usar conexão DBAPI2 nativa (psycopg2) sem SQLAlchemy
warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

logger = logging.getLogger(__name__)

DB_PATH = Path("chamados.db")
DB_TYPE = os.getenv("DB_TYPE", "sqlite").lower()

# Pool global de conexões para PostgreSQL (Thread-safe)
_pg_pool = None


import re

try:
    from psycopg2.extensions import cursor as _pg_base_cursor
except Exception:
    _pg_base_cursor = object

class PostgreSQLCompatibleCursor(_pg_base_cursor):
    """
    Cursor PostgreSQL adaptativo que traduz automaticamente sintaxes SQLite comuns:
    1. Placeholders '?' são substituídos por '%s'.
    2. 'INTEGER PRIMARY KEY AUTOINCREMENT' é convertido para 'SERIAL PRIMARY KEY'.
    3. 'PRAGMA table_info(tabela)' é emulado consultando information_schema.columns.
    4. Emulação de 'lastrowid' para INSERTs em tabelas com chaves SERIAL / IDENTITY.
    5. Retorno de 'self' em execute/executemany para method-chaining fluído.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._custom_lastrowid = None

    @property
    def lastrowid(self):
        if self._custom_lastrowid is not None:
            return self._custom_lastrowid
        try:
            return super().lastrowid
        except Exception:
            return None

    @staticmethod
    def _translate_sqlite_query(query: str, has_vars: bool = False) -> str:
        if 'INTEGER PRIMARY KEY AUTOINCREMENT' in query:
            query = query.replace('INTEGER PRIMARY KEY AUTOINCREMENT', 'SERIAL PRIMARY KEY')

        if 'ON CONFLICT' in query:
            # Em CREATE TABLE no PostgreSQL, UNIQUE(...) ON CONFLICT REPLACE / IGNORE não é sintaxe ANSI (apenas UNIQUE(...))
            query = re.sub(r'\s+ON\s+CONFLICT\s+(?:REPLACE|IGNORE)', '', query, flags=re.I)

        if re.search(r'INSERT\s+OR\s+IGNORE\s+INTO', query, re.I):
            query = re.sub(r'INSERT\s+OR\s+IGNORE\s+INTO', 'INSERT INTO', query, flags=re.I)
            if 'ON CONFLICT' not in query.upper():
                stripped = query.rstrip()
                if stripped.endswith(';'):
                    query = stripped[:-1] + ' ON CONFLICT DO NOTHING;'
                else:
                    query = stripped + ' ON CONFLICT DO NOTHING'

        if re.search(r'INSERT\s+OR\s+REPLACE\s+INTO', query, re.I):
            query = re.sub(r'INSERT\s+OR\s+REPLACE\s+INTO', 'INSERT INTO', query, flags=re.I)

        if 'COLLATE NOCASE' in query:
            query = re.sub(r'\s+COLLATE\s+NOCASE', '', query, flags=re.I)

        if has_vars and '?' in query:
            query = query.replace('?', '%s')

        return query

    def execute(self, query, vars=None):
        self._custom_lastrowid = None
        if 'PRAGMA table_info(' in query:
            m = re.search(r'PRAGMA\s+table_info\((.+?)\)', query, re.I)
            if m:
                tbl = m.group(1).strip('"\' ')
                return super().execute('''
                    SELECT ordinal_position, column_name, data_type, 
                           CASE WHEN is_nullable = 'NO' THEN 1 ELSE 0 END,
                           column_default, 0
                    FROM information_schema.columns 
                    WHERE table_name = %s
                    ORDER BY ordinal_position
                ''', (tbl,))

        query = self._translate_sqlite_query(query, has_vars=(vars is not None))

        # Detecta tabela em INSERT para emular cursor.lastrowid via currval / pg_get_serial_sequence
        insert_match = re.match(r'^\s*INSERT\s+INTO\s+([a-zA-Z0-9_]+)', query, re.I)
        target_table = insert_match.group(1).lower() if insert_match else None

        super().execute(query, vars)

        # Emulação de lastrowid apenas para tabelas que utilizam chave primária SERIAL
        SERIAL_TABLES = {
            'ad_cache_group_members', 'ad_sync_meta', 'comentarios', 'eventos_manuais',
            'faq_videos', 'faqs', 'ferias_bancada', 'garantia', 'garantia_contratos',
            'notificacoes', 'plantoes_matutino', 'plantoes_semanal', 'ramais_mpms',
            'senhas_cofre', 'unidades', 'cron_logs', 'faq_imagens', 'garantia_chamados',
            'garantia_agendamentos', 'impressoras', 'viagens', 'unidades_manuais',
            'whatsapp_destinatarios', 'whatsapp_disparos_log'
        }
        if target_table and target_table in SERIAL_TABLES:
            try:
                seq_name = f"{target_table}_id_seq"
                with self.connection.cursor() as seq_cur:
                    seq_cur.execute(f"SAVEPOINT _sp_lastrowid; SELECT currval('{seq_name}'::regclass);")
                    val = seq_cur.fetchone()
                    if val and val[0] is not None:
                        self._custom_lastrowid = int(val[0])
                    seq_cur.execute("RELEASE SAVEPOINT _sp_lastrowid;")
            except Exception:
                try:
                    with self.connection.cursor() as sp_cur:
                        sp_cur.execute("ROLLBACK TO SAVEPOINT _sp_lastrowid;")
                except Exception:
                    pass

        return self

    def executemany(self, query, vars_list):
        self._custom_lastrowid = None
        query = self._translate_sqlite_query(query, has_vars=bool(vars_list))
        super().executemany(query, vars_list)
        return self


def get_pg_pool():
    """Retorna o pool de conexões ThreadedConnectionPool para PostgreSQL (singleton lazy)."""
    global _pg_pool
    if _pg_pool is None:
        try:
            from psycopg2.pool import ThreadedConnectionPool
            pg_host = os.getenv("POSTGRES_HOST", "localhost")
            pg_port = os.getenv("POSTGRES_PORT", "5432")
            pg_db = os.getenv("POSTGRES_DB", "chamados")
            pg_user = os.getenv("POSTGRES_USER", "postgres")
            pg_pass = os.getenv("POSTGRES_PASSWORD", "postgres")
            _pg_pool = ThreadedConnectionPool(
                minconn=1,
                maxconn=20,
                host=pg_host,
                port=pg_port,
                dbname=pg_db,
                user=pg_user,
                password=pg_pass,
                cursor_factory=PostgreSQLCompatibleCursor
            )
            logger.info("🔌 Pool de conexões PostgreSQL inicializado com sucesso (1-20 conexões).")
        except Exception as e:
            logger.warning(f"⚠️ Não foi possível inicializar pool PostgreSQL: {e}. Usando fallback direto.")
            _pg_pool = None
    return _pg_pool


class PooledPostgresConnection:
    """Wrapper para conexão do pool PostgreSQL que devolve a conexão ao pool quando close() for chamado."""
    def __init__(self, raw_conn, pool):
        self._raw_conn = raw_conn
        self._pool = pool
        self._closed = False

    def close(self):
        if not self._closed:
            self._closed = True
            try:
                # Limpa qualquer transação abortada ou pendente antes de devolver ao pool
                if getattr(self._raw_conn, "closed", 0) == 0:
                    try:
                        self._raw_conn.rollback()
                    except Exception:
                        pass
                self._pool.putconn(self._raw_conn)
            except Exception:
                pass

    def __enter__(self):
        return self._raw_conn.__enter__()

    def __exit__(self, exc_type, exc_val, exc_tb):
        return self._raw_conn.__exit__(exc_type, exc_val, exc_tb)

    def __getattr__(self, name):
        return getattr(self._raw_conn, name)


def save_dataframe_to_db(df, table_name: str, conn, if_exists: str = "replace"):
    """
    Persiste um DataFrame pandas em uma tabela do banco de dados de forma compatível
    tanto com SQLite (sqlite3) quanto com PostgreSQL (psycopg2/PooledPostgresConnection),
    sem depender de SQLAlchemy ou sqlite_master.
    """
    import pandas as pd
    if isinstance(conn, sqlite3.Connection) and hasattr(df, "to_sql"):
        try:
            df.to_sql(table_name, conn, if_exists=if_exists, index=False)
            return
        except Exception:
            pass

    cursor = conn.cursor()
    cols = [str(c) for c in df.columns]
    if not cols:
        return

    escaped_cols = [f'"{c}"' for c in cols]

    if if_exists == "replace":
        cursor.execute(f'DROP TABLE IF EXISTS "{table_name}"')
        col_defs = []
        for c in cols:
            dtype_str = str(getattr(df[c], "dtype", "object")).lower() if hasattr(df, "__getitem__") else "object"
            if "int" in dtype_str:
                pg_type = "BIGINT"
            elif "float" in dtype_str:
                pg_type = "DOUBLE PRECISION"
            else:
                pg_type = "TEXT"
            col_defs.append(f'"{c}" {pg_type}')
        cursor.execute(f'CREATE TABLE "{table_name}" ({", ".join(col_defs)})')
    elif if_exists == "append":
        col_defs = [f'"{c}" TEXT' for c in cols]
        cursor.execute(f'CREATE TABLE IF NOT EXISTS "{table_name}" ({", ".join(col_defs)})')

    if len(df) > 0:
        rows = []
        for _, r in df.iterrows():
            row_vals = []
            for c in cols:
                v = r[c] if c in r else None
                if v is None or (hasattr(pd, "isna") and pd.isna(v)):
                    row_vals.append(None)
                elif hasattr(v, "isoformat") and not isinstance(v, str):
                    row_vals.append(v.strftime("%Y-%m-%d %H:%M:%S") if hasattr(v, "strftime") else str(v))
                elif hasattr(v, "item"):
                    row_vals.append(v.item())
                else:
                    row_vals.append(v)
            rows.append(tuple(row_vals))

        placeholders = ", ".join(["?"] * len(cols))
        insert_sql = f'INSERT INTO "{table_name}" ({", ".join(escaped_cols)}) VALUES ({placeholders})'
        cursor.executemany(insert_sql, rows)


def get_connection():
    """
    Retorna uma conexão com o banco de dados.
    Suporta modo híbrido: 'sqlite' (padrão local chamados.db) ou 'postgres' (servidor PostgreSQL).
    Para PostgreSQL, utiliza o Connection Pool gerenciado com fallback direto se necessário.
    """
    db_mode = os.getenv("DB_TYPE", DB_TYPE).lower()
    if db_mode in ["postgres", "postgresql"]:
        pool = get_pg_pool()
        if pool:
            try:
                raw_conn = pool.getconn()
                return PooledPostgresConnection(raw_conn, pool)
            except Exception as e_pool:
                logger.warning(f"Falha ao obter conexão do pool Postgres: {e_pool}. Conectando diretamente.")
        
        import psycopg2
        pg_host = os.getenv("POSTGRES_HOST", "localhost")
        pg_port = os.getenv("POSTGRES_PORT", "5432")
        pg_db = os.getenv("POSTGRES_DB", "chamados")
        pg_user = os.getenv("POSTGRES_USER", "postgres")
        pg_pass = os.getenv("POSTGRES_PASSWORD", "postgres")
        return psycopg2.connect(
            host=pg_host,
            port=pg_port,
            dbname=pg_db,
            user=pg_user,
            password=pg_pass,
            cursor_factory=PostgreSQLCompatibleCursor
        )
    else:
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        try:
            conn.execute("PRAGMA journal_mode = WAL;")
            conn.execute("PRAGMA busy_timeout = 30000;")
            conn.execute("PRAGMA synchronous = NORMAL;")
        except Exception:
            pass
        return conn


def ensure_database_indexes():
    """
    Cria índices estratégicos caso não existam para tabelas de grande volume
    (chamados, usuários AD, computadores AD, dispositivos SCCM, central telefônica e unidades).
    Melhora substancialmente o desempenho de buscas e paginação.
    """
    conn = get_connection()
    cursor = conn.cursor()
    indexes = [
        # Índices na tabela de chamados
        ("idx_chamados_id", "chamados", "(id)"),
        ("idx_chamados_status", "chamados", "(status)"),
        ("idx_chamados_usuario", "chamados", "(usuario)"),
        ("idx_chamados_data_criacao", "chamados", "(data_criacao)"),
        ("idx_comentarios_chamado_id", "comentarios", "(chamado_id)"),
        # Índices no Active Directory
        ("idx_ad_users_sam", "ad_cache_users", "(sam_account_name)"),
        ("idx_ad_users_active", "ad_cache_users", "(is_active)"),
        ("idx_ad_users_dept", "ad_cache_users", "(department)"),
        ("idx_ad_computers_name", "ad_cache_computers", "(name)"),
        ("idx_ad_computers_active", "ad_cache_computers", "(is_active)"),
        # Índices no SCCM
        ("idx_sccm_devices_res_id", "sccm_cache_devices", "(resource_id)"),
        ("idx_sccm_devices_name", "sccm_cache_devices", "(name)"),
        ("idx_sccm_devices_user", "sccm_cache_devices", "(last_logon_user)"),
        ("idx_sccm_devices_active", "sccm_cache_devices", "(client_active)"),
        ("idx_sccm_users_name", "sccm_cache_users", "(user_name)"),
        ("idx_sccm_cols_id", "sccm_cache_collections", "(collection_id)"),
        # Índices na telefonia e unidades
        ("idx_central_ramal", "central_telefonica", "(ramal)"),
        ("idx_unidades_cidade", "unidades", "(cidade)"),
        ("idx_unidades_tipo", "unidades", "(tipo)"),
        ("idx_ramais_num", "ramais_telefonicos", "(ramal)")
    ]

    for idx_name, table, cols in indexes:
        try:
            cursor.execute(f"CREATE INDEX IF NOT EXISTS {idx_name} ON {table} {cols};")
        except Exception:
            # Se a tabela ainda não foi criada, não interrompe
            pass

    try:
        conn.commit()
    except Exception:
        pass
    finally:
        try:
            cursor.close()
        except Exception:
            pass
        conn.close()
