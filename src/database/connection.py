import os
import sqlite3
from pathlib import Path

DB_PATH = Path("chamados.db")
DB_TYPE = os.getenv("DB_TYPE", "sqlite").lower()

def get_connection():
    """
    Retorna uma conexão com o banco de dados.
    Suporta modo híbrido: 'sqlite' (padrão local chamados.db) ou 'postgres' (servidor PostgreSQL).
    """
    if DB_TYPE in ["postgres", "postgresql"]:
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
