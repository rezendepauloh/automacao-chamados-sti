# -*- coding: utf-8 -*-
"""
Testes unitários de compatibilidade híbrida SQLite <-> PostgreSQL (src.database.*).
Valida tradução de sintaxe SQL, cursor adaptativo PostgreSQLCompatibleCursor,
devolução limpa de conexões ao pool (rollback automático), persistência de DataFrames
sem depender de sqlite_master (save_dataframe_to_db), preenchimento de data_criacao e
deduplicação em notificacoes, plantoes, ramais_config, mapa_config, settings e AD cache.
"""

import os
import sys
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.database.connection import (
    PostgreSQLCompatibleCursor,
    PooledPostgresConnection,
    save_dataframe_to_db,
)
from src.database.notifications_db import (
    setup_notifications_table,
    add_notification,
    get_notifications,
)
from src.database.plantoes_db import (
    setup_plantoes_tables,
    save_plantoes_matutino,
    save_plantoes_semanal,
    get_plantoes_matutino,
    get_plantoes_semanal,
)
from src.database.ad_db import (
    setup_ad_tables,
    save_ad_cache,
    get_ad_users_df,
    get_ad_computers_df,
)
from src.database.settings_db import (
    setup_settings_table,
    set_setting,
    get_setting,
)
from src.database.map_db import (
    setup_map_tables,
    save_map_config,
    get_map_config,
    get_map_pins,
)
from src.database.ramais_db import (
    setup_ramais_config_table,
    save_ramais_config,
    get_ramais_config,
)


class _NonSqliteDBAPIWrapper:
    """Simula uma conexão DBAPI2 que não herda de sqlite3.Connection (ex.: PooledPostgresConnection)."""

    def __init__(self, sqlite_conn: sqlite3.Connection):
        self._conn = sqlite_conn

    def cursor(self):
        return self._conn.cursor()

    def commit(self):
        return self._conn.commit()

    def rollback(self):
        return self._conn.rollback()

    def close(self):
        return self._conn.close()


class TestPostgresCompatibility(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_pg_compat.db"

        self.patchers = [
            patch("src.database.connection.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.notifications_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.plantoes_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.ad_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.settings_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.map_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.ramais_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.connection.DB_TYPE", "sqlite"),
            patch("src.database.notifications_db.DB_TYPE", "sqlite"),
            patch("src.database.ad_db.DB_TYPE", "sqlite"),
            patch("src.database.settings_db.DB_TYPE", "sqlite"),
        ]
        for p in self.patchers:
            p.start()

    def tearDown(self):
        for p in self.patchers:
            p.stop()
        self.temp_dir.cleanup()

    def test_translate_sqlite_query_syntax(self):
        """Valida tradução de sintaxe SQLite para PostgreSQL no cursor adaptativo."""
        translate = PostgreSQLCompatibleCursor._translate_sqlite_query

        # 1. Placeholders ? -> %s quando há variáveis
        q1 = translate("SELECT * FROM chamados WHERE id = ? AND status = ?", has_vars=True)
        self.assertEqual(q1, "SELECT * FROM chamados WHERE id = %s AND status = %s")

        # Sem variáveis, não altera ? literais
        q1_no_vars = translate("SELECT '?' FROM chamados", has_vars=False)
        self.assertEqual(q1_no_vars, "SELECT '?' FROM chamados")

        # 2. INTEGER PRIMARY KEY AUTOINCREMENT -> SERIAL PRIMARY KEY
        q2 = translate("CREATE TABLE t (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT)")
        self.assertIn("SERIAL PRIMARY KEY", q2)
        self.assertNotIn("AUTOINCREMENT", q2)

        # 3. INSERT OR IGNORE INTO -> INSERT INTO ... ON CONFLICT DO NOTHING
        q3 = translate("INSERT OR IGNORE INTO faqs (titulo, url) VALUES (?, ?)", has_vars=True)
        self.assertEqual(q3, "INSERT INTO faqs (titulo, url) VALUES (%s, %s) ON CONFLICT DO NOTHING")

        # 4. Remoção de COLLATE NOCASE e ON CONFLICT REPLACE em DDL
        q4 = translate("SELECT * FROM users WHERE name = ? COLLATE NOCASE", has_vars=True)
        self.assertNotIn("COLLATE NOCASE", q4)
        self.assertIn("%s", q4)

        q5 = translate("CREATE TABLE x (k TEXT UNIQUE ON CONFLICT REPLACE)")
        self.assertNotIn("ON CONFLICT REPLACE", q5)

    def test_pooled_postgres_connection_close_rolls_back_aborted_tx(self):
        """Garante que PooledPostgresConnection.close() executa rollback antes de devolver ao pool."""
        raw_conn = MagicMock()
        raw_conn.closed = 0
        pool = MagicMock()

        conn_wrapper = PooledPostgresConnection(raw_conn, pool)
        conn_wrapper.close()

        raw_conn.rollback.assert_called_once()
        pool.putconn.assert_called_once_with(raw_conn)

        # Segunda chamada a close() deve ser idempotente
        conn_wrapper.close()
        self.assertEqual(raw_conn.rollback.call_count, 1)
        self.assertEqual(pool.putconn.call_count, 1)

    def test_save_dataframe_to_db_without_sqlite_master(self):
        """Valida que save_dataframe_to_db persiste DataFrames em conexões não-SQLite sem consultar sqlite_master."""
        raw_sqlite = sqlite3.connect(self.db_path)
        wrapped_conn = _NonSqliteDBAPIWrapper(raw_sqlite)

        df_initial = pd.DataFrame([
            {"contrato": "CT-001", "fiscal": "Paulo Henrique", "qtd": 3},
            {"contrato": "CT-002", "fiscal": "Maria Souza", "qtd": 5},
        ])
        save_dataframe_to_db(df_initial, "fiscalizacao_teste", wrapped_conn, if_exists="replace")
        wrapped_conn.commit()

        df_loaded = pd.read_sql_query('SELECT * FROM "fiscalizacao_teste" ORDER BY contrato', raw_sqlite)
        self.assertEqual(len(df_loaded), 2)
        self.assertEqual(df_loaded.iloc[0]["fiscal"], "Paulo Henrique")
        self.assertEqual(int(df_loaded.iloc[1]["qtd"]), 5)

        # Testa substituição (if_exists="replace") e append
        df_replace = pd.DataFrame([{"contrato": "CT-003", "fiscal": "Carlos", "qtd": 1}])
        save_dataframe_to_db(df_replace, "fiscalizacao_teste", wrapped_conn, if_exists="replace")
        save_dataframe_to_db(df_initial, "fiscalizacao_teste", wrapped_conn, if_exists="append")
        wrapped_conn.commit()

        df_final = pd.read_sql_query('SELECT * FROM "fiscalizacao_teste"', raw_sqlite)
        self.assertEqual(len(df_final), 3)
        raw_sqlite.close()

    def test_notifications_timestamp_and_deduplication(self):
        """Garante que add_notification preenche data_criacao (nunca None) e não duplica alertas idênticos."""
        setup_notifications_table()

        first = add_notification(
            tipo="Portaria",
            titulo="Nova Portaria nº 1028/2013",
            mensagem="Servidores: Paulo Henrique Gonçalves Rezende.",
            data_evento="12/08/2013",
            link_pagina="https://www.mpms.mp.br/portarias",
        )
        self.assertTrue(first)

        # Tentativa duplicada deve retornar False e não criar novo registro
        duplicate = add_notification(
            tipo="Portaria",
            titulo="Nova Portaria nº 1028/2013",
            mensagem="Servidores: Paulo Henrique Gonçalves Rezende.",
            data_evento="12/08/2013",
            link_pagina="https://www.mpms.mp.br/portarias",
        )
        self.assertFalse(duplicate)

        df = get_notifications(limit=10)
        self.assertEqual(len(df), 1)
        data_criacao = df.iloc[0]["data_criacao"]
        self.assertIsNotNone(data_criacao)
        self.assertNotEqual(str(data_criacao).strip().lower(), "none")
        self.assertNotEqual(str(data_criacao).strip(), "")

    def test_plantoes_upsert_without_duplication(self):
        """Garante que save_plantoes_matutino e save_plantoes_semanal fazem ON CONFLICT DO UPDATE sem duplicar."""
        setup_plantoes_tables()

        rec_mat = [{
            "ano": 2026,
            "data_iso": "2026-05-10",
            "dia_semana": "Domingo",
            "servidor": "Paulo Henrique",
            "telefone": "67 98888-0000",
        }]
        save_plantoes_matutino(rec_mat)
        # Atualiza telefone do mesmo plantão
        rec_mat[0]["telefone"] = "67 99999-1111"
        save_plantoes_matutino(rec_mat)

        df_mat = get_plantoes_matutino(2026)
        self.assertEqual(len(df_mat), 1)
        self.assertEqual(df_mat.iloc[0]["telefone"], "67 99999-1111")

        rec_sem = [{
            "ano": 2026,
            "mes": "Maio",
            "periodo_str": "04/05/2026 a 11/05/2026",
            "data_inicio": "2026-05-04 19:01:00",
            "data_fim": "2026-05-11 11:59:00",
            "service_desk": "Analista A",
            "manutencao": "Paulo Henrique",
            "infraestrutura": "Analista B",
            "desenvolvimento": "Analista C",
        }]
        save_plantoes_semanal(rec_sem)
        rec_sem[0]["service_desk"] = "Analista Atualizado"
        save_plantoes_semanal(rec_sem)

        df_sem = get_plantoes_semanal(2026)
        self.assertEqual(len(df_sem), 1)
        self.assertEqual(df_sem.iloc[0]["service_desk"], "Analista Atualizado")

    def test_ad_cache_integer_flags_and_filtering(self):
        """Valida que is_active é persistido e filtrado como inteiro (1/0), compatível com coluna INTEGER no Postgres."""
        setup_ad_tables()
        users = [
            {"sam_account_name": "ativo.user", "display_name": "Usuário Ativo", "is_active": True},
            {"sam_account_name": "inativo.user", "display_name": "Usuário Inativo", "is_active": False},
        ]
        computers = [
            {"name": "MPE-ATIVO", "operating_system": "Windows 11", "is_active": True},
            {"name": "MPE-INATIVO", "operating_system": "Windows 10", "is_active": False},
        ]
        save_ad_cache(ous=[], users=users, groups=[], memberships=[], computers=computers)

        df_u_active = get_ad_users_df(status_filter="Ativos")
        self.assertEqual(len(df_u_active), 1)
        self.assertEqual(df_u_active.iloc[0]["sam_account_name"], "ativo.user")

        df_u_inactive = get_ad_users_df(status_filter="Desativados")
        self.assertEqual(len(df_u_inactive), 1)
        self.assertEqual(df_u_inactive.iloc[0]["sam_account_name"], "inativo.user")

        df_c_active = get_ad_computers_df(status_filter="Ativos")
        self.assertEqual(len(df_c_active), 1)
        self.assertEqual(df_c_active.iloc[0]["name"], "MPE-ATIVO")

    def test_settings_map_and_ramais_config_upserts(self):
        """Valida upserts compatíveis com PostgreSQL (ON CONFLICT DO UPDATE) em settings, mapa e ramais."""
        setup_settings_table()
        set_setting("TEST_KEY", "valor_1", is_secret=False)
        self.assertEqual(get_setting("TEST_KEY"), "valor_1")
        set_setting("TEST_KEY", "valor_2", is_secret=False)
        self.assertEqual(get_setting("TEST_KEY"), "valor_2")

        setup_map_tables()
        save_map_config({
            "predios": [{"id": "sede", "nome": "PGJ Sede"}],
            "pins": [{"id": "pin1", "predio_id": "sede", "pavimento_id": 1, "sala": "STI", "x": 10, "y": 20, "descricao": "v1"}],
        })
        save_map_config({
            "predios": [{"id": "sede", "nome": "PGJ Sede Atualizada"}],
            "pins": [{"id": "pin1", "predio_id": "sede", "pavimento_id": 1, "sala": "STI", "x": 15, "y": 25, "descricao": "v2"}],
        })
        cfg = get_map_config()
        self.assertEqual(cfg["predios"][0]["nome"], "PGJ Sede Atualizada")
        pins = get_map_pins(predio_id="sede", pavimento_id=1)
        self.assertEqual(len(pins), 1)
        self.assertEqual(pins[0]["descricao"], "v2")

        setup_ramais_config_table()
        save_ramais_config("https://mpms.mp.br/int_v1.pdf", "https://mpms.mp.br/cap_v1.pdf")
        save_ramais_config("https://mpms.mp.br/int_v2.pdf", "https://mpms.mp.br/cap_v2.pdf")
        r_cfg = get_ramais_config()
        self.assertEqual(r_cfg["Interior"], "https://mpms.mp.br/int_v2.pdf")
        self.assertEqual(r_cfg["Capital / PGJ"], "https://mpms.mp.br/cap_v2.pdf")


if __name__ == "__main__":
    unittest.main()
