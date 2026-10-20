"""
Script de Migração Completa: SQLite (chamados.db) -> PostgreSQL (Docker / RedHat MPMS).
Copia com segurança todas as 42 tabelas do SQLite para o banco PostgreSQL de produção.
Ajusta tipos, sequências de chaves primárias e cria índices de performance.
"""

import os
import sys
import time
import sqlite3
import psycopg2
from psycopg2.extras import execute_batch
from pathlib import Path

# Adiciona o diretório raiz ao path
ROOT_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT_DIR))

SQLITE_PATH = ROOT_DIR / "chamados.db"

PG_HOST = os.getenv("POSTGRES_HOST", "db")
PG_PORT = int(os.getenv("POSTGRES_PORT", "5432"))
PG_DB = os.getenv("POSTGRES_DB", "chamados")
PG_USER = os.getenv("POSTGRES_USER", "bancada_user")
PG_PASS = os.getenv("POSTGRES_PASSWORD", "bancada_pass")


def map_sqlite_type_to_pg(sqlite_type: str, col_name: str = "") -> str:
    """Converte tipos de dados do SQLite para tipos equivalentes robustos no PostgreSQL."""
    t = (sqlite_type or "TEXT").upper().strip()
    low_col = col_name.lower()
    if low_col in ["is_active", "ativo", "lida", "tag_manual", "dados_manuais", "enabled"]:
        return "INTEGER"
    if "tamanho" in low_col or "bytes" in low_col or "size" in low_col:
        return "BIGINT"
    if "INT" in t:
        return "BIGINT" if "time" in low_col or "id" in low_col else "INTEGER"
    if "REAL" in t or "FLOA" in t or "DOUB" in t:
        return "DOUBLE PRECISION"
    if "BLOB" in t:
        return "BYTEA"
    if "BOOL" in t:
        return "INTEGER"
    return "TEXT"


def migrate():
    print("=" * 70)
    print("🚀 INICIANDO MIGRAÇÃO: SQLite (chamados.db) -> PostgreSQL")
    print(f"   Origem: {SQLITE_PATH}")
    print(f"   Destino: postgresql://{PG_USER}@{PG_HOST}:{PG_PORT}/{PG_DB}")
    print("=" * 70)

    if not SQLITE_PATH.exists():
        print(f"❌ Arquivo {SQLITE_PATH} não encontrado!")
        sys.exit(1)

    # Conectar ao SQLite
    sqlite_conn = sqlite3.connect(SQLITE_PATH)
    sqlite_cur = sqliteconn_cur = sqlite_conn.cursor()

    # Conectar ao PostgreSQL
    try:
        pg_conn = psycopg2.connect(
            host=PG_HOST,
            port=PG_PORT,
            dbname=PG_DB,
            user=PG_USER,
            password=PG_PASS
        )
        pg_conn.autocommit = False
        pg_cur = pg_conn.cursor()
    except Exception as e:
        print(f"❌ Erro ao conectar ao PostgreSQL: {e}")
        sys.exit(1)

    # Obter tabelas do SQLite
    sqlite_cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
    tables = [row[0] for row in sqlite_cur.fetchall()]
    print(f"📋 {len(tables)} tabelas encontradas no SQLite para migração.\n")

    total_migrated_rows = 0
    start_time = time.time()

    for table in tables:
        print(f"📦 Migrando tabela: [{table}]...")
        
        # Obter colunas e PK
        sqlite_cur.execute(f"PRAGMA table_info(\"{table}\")")
        columns_info = sqlite_cur.fetchall()
        # columns_info: [(cid, name, type, notnull, dflt_value, pk), ...]
        
        col_defs = []
        pk_cols = []
        col_names = []
        
        for col in columns_info:
            c_name = col[1]
            c_type = col[2]
            c_pk = col[5]
            col_names.append(c_name)
            
            pg_type = map_sqlite_type_to_pg(c_type, c_name)
            
            # Se for chave primária auto-increment
            if c_pk > 0:
                pk_cols.append(f'"{c_name}"')
                if "INT" in (c_type or "").upper():
                    # Usamos INTEGER/BIGINT e criaremos sequência posteriormente
                    col_defs.append(f'"{c_name}" {pg_type}')
                else:
                    col_defs.append(f'"{c_name}" {pg_type}')
            else:
                col_defs.append(f'"{c_name}" {pg_type}')

        pk_clause = f", PRIMARY KEY ({', '.join(pk_cols)})" if pk_cols else ""
        create_table_sql = f'CREATE TABLE IF NOT EXISTS "{table}" (\n  ' + ',\n  '.join(col_defs) + pk_clause + '\n);'

        # Cria a tabela no Postgres se não existir
        try:
            pg_cur.execute(create_table_sql)
            pg_conn.commit()
        except Exception as e:
            pg_conn.rollback()
            print(f"   ⚠️ Aviso na criação de [{table}]: {e}. Continuando com tabela existente.")

        # Limpa dados anteriores da tabela no Postgres antes de transferir
        try:
            pg_cur.execute(f'TRUNCATE TABLE "{table}" CASCADE;')
            pg_conn.commit()
        except Exception:
            pg_conn.rollback()
            try:
                pg_cur.execute(f'DELETE FROM "{table}";')
                pg_conn.commit()
            except Exception as e_del:
                pg_conn.rollback()
                print(f"   ⚠️ Falha ao truncar tabela [{table}]: {e_del}")

        # Lê os dados do SQLite
        escaped_cols = [f'"{c}"' for c in col_names]
        select_sql = f'SELECT {", ".join(escaped_cols)} FROM "{table}"'
        sqlite_cur.execute(select_sql)
        rows = sqlite_cur.fetchall()

        if rows:
            placeholders = ", ".join(["%s"] * len(col_names))
            insert_sql = f'INSERT INTO "{table}" ({", ".join(escaped_cols)}) VALUES ({placeholders})'
            
            try:
                execute_batch(pg_cur, insert_sql, rows, page_size=2000)
                pg_conn.commit()
                print(f"   ✅ {len(rows)} registros transferidos com sucesso.")
                total_migrated_rows += len(rows)
            except Exception as e_ins:
                pg_conn.rollback()
                print(f"   ❌ Erro ao inserir dados em [{table}]: {e_ins}")
                # Tenta linha a linha para isolar
                success_count = 0
                for row in rows:
                    try:
                        pg_cur.execute(insert_sql, row)
                        pg_conn.commit()
                        success_count += 1
                    except Exception:
                        pg_conn.rollback()
                print(f"   ⚠️ Inserção linha a linha recuperou {success_count}/{len(rows)} linhas.")
                total_migrated_rows += success_count
        else:
            print("   ℹ️ Tabela vazia no SQLite.")

        # Se houver PK integer, atualiza/cria sequência para auto-increment
        for col in columns_info:
            c_name = col[1]
            c_type = col[2]
            c_pk = col[5]
            if c_pk == 1 and "INT" in (c_type or "").upper():
                seq_name = f"{table}_{c_name}_seq"
                try:
                    pg_cur.execute(f'CREATE SEQUENCE IF NOT EXISTS "{seq_name}";')
                    pg_cur.execute(f'ALTER TABLE "{table}" ALTER COLUMN "{c_name}" SET DEFAULT nextval(\'"{seq_name}"\');')
                    pg_cur.execute(f'SELECT COALESCE(MAX("{c_name}"), 0) + 1 FROM "{table}";')
                    max_id = pg_cur.fetchone()[0]
                    pg_cur.execute(f'SELECT setval(\'"{seq_name}"\', %s, false);', (max_id,))
                    pg_conn.commit()
                except Exception as e_seq:
                    pg_conn.rollback()
                    # Ignora se sequência já estiver vinculada
                    pass

    # Criação dos índices estratégicos
    print("\n🔍 Criando índices estratégicos no PostgreSQL...")
    from src.database.connection import ensure_database_indexes
    # Força DB_TYPE como postgres para criar no pg
    os.environ["DB_TYPE"] = "postgres"
    try:
        ensure_database_indexes()
        print("✅ Índices criados e verificados com sucesso!")
    except Exception as e_idx:
        print(f"⚠️ Erro ao criar índices: {e_idx}")

    duration = time.time() - start_time
    print("=" * 70)
    print(f"🎉 MIGRAÇÃO CONCLUÍDA EM {duration:.2f} segundos!")
    print(f"📊 Total de registros transferidos para PostgreSQL: {total_migrated_rows}")
    print("=" * 70)

    sqlite_conn.close()
    pg_conn.close()


if __name__ == "__main__":
    migrate()
