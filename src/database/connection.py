import os
import sqlite3
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

DB_PATH = Path("chamados.db")
DB_TYPE = os.getenv("DB_TYPE", "sqlite").lower()

# Pool global de conexões para PostgreSQL (Thread-safe)
_pg_pool = None


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
                password=pg_pass
            )
            logger.info("🔌 Pool de conexões PostgreSQL inicializado com sucesso (1-20 conexões).")
        except Exception as e:
            logger.warning(f"⚠️ Não foi possível inicializar pool PostgreSQL: {e}. Usando fallback direto.")
            _pg_pool = None
    return _pg_pool


def get_connection():
    """
    Retorna uma conexão com o banco de dados.
    Suporta modo híbrido: 'sqlite' (padrão local chamados.db) ou 'postgres' (servidor PostgreSQL).
    Para PostgreSQL, utiliza o Connection Pool gerenciado com fallback direto se necessário.
    """
    if DB_TYPE in ["postgres", "postgresql"]:
        pool = get_pg_pool()
        if pool:
            try:
                conn = pool.getconn()
                # Encapsula o close para devolver a conexão ao pool quando conn.close() for chamado
                orig_close = conn.close
                def _pooled_close():
                    try:
                        pool.putconn(conn)
                    except Exception:
                        pass
                conn.close = _pooled_close
                return conn
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
            password=pg_pass
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
