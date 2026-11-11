# -*- coding: utf-8 -*-
"""
Testes unitários do bootstrap de diretórios e banco de dados para novas instalações.
"""

import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.config import ensure_project_directories
from src.database import (
    ad_db,
    bootstrap,
    connection,
    cron_db,
    donations_db,
    events_db,
    ferias_db,
    garantia_db,
    impressoras_db,
    map_db,
    notifications_db,
    plantoes_db,
    ramais_db,
    sccm_db,
    senhas_db,
    settings_db,
    tickets_db,
    unidades_db,
    viagens_db,
    whatsapp_db,
)
from src.database.connection import PostgreSQLCompatibleCursor


class TestBootstrapInstallation(unittest.TestCase):
    """Valida a criação automática de pastas, .env e schema completo do banco."""

    def test_ensure_project_directories_creates_full_tree(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            ensure_project_directories(base_dir=root)

            expected = [
                root / "data",
                root / "data" / "01 - Dados Brutos",
                root / "data" / "02 - Dados tratados",
                root / "data" / "03 - Dados prontos",
                root / "models",
                root / "uploads",
                root / "uploads" / "faq" / "videos",
                root / "uploads" / "faq" / "imagens",
                root / "uploads" / "temp",
                root / "src" / "js",
                root / "debug_logs",
                root / "debug_logs" / "orquestrador",
                root / "debug_logs" / "otrs",
                root / "debug_logs" / "citsmart",
                root / "debug_logs" / "tag",
            ]
            for d in expected:
                self.assertTrue(d.is_dir(), f"Diretório ausente após bootstrap: {d}")

    def test_initialize_database_creates_all_tables_and_is_idempotent(self):
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "fresh_install.db"
            mods = [
                ad_db, bootstrap, connection, cron_db, donations_db, events_db,
                ferias_db, garantia_db, impressoras_db, map_db, notifications_db,
                plantoes_db, ramais_db, sccm_db, senhas_db, settings_db, tickets_db,
                unidades_db, viagens_db, whatsapp_db,
            ]
            patchers = [
                patch.object(m, "get_connection", side_effect=lambda: sqlite3.connect(db_path))
                for m in mods if hasattr(m, "get_connection")
            ]
            for mod in [connection, ad_db, settings_db, notifications_db, tickets_db, sccm_db, cron_db]:
                if hasattr(mod, "DB_TYPE"):
                    patchers.append(patch.object(mod, "DB_TYPE", "sqlite"))

            for p in patchers:
                p.start()
            try:
                first = bootstrap.initialize_database(wait_retries=1)
                self.assertEqual(first["errors"], [])
                self.assertEqual(len(first["initialized"]), 22)

                # Execução repetida deve ser idempotente
                second = bootstrap.initialize_database(wait_retries=1)
                self.assertEqual(second["errors"], [])

                conn = sqlite3.connect(db_path)
                cur = conn.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
                tables = {row[0] for row in cur.fetchall()}
                cur.execute("SELECT COUNT(*) FROM cron_schedules")
                cron_count = cur.fetchone()[0]
                cur.execute("SELECT COUNT(*) FROM system_settings")
                settings_count = cur.fetchone()[0]
                conn.close()

                for req_tbl in {
                    "chamados", "comentarios", "system_settings", "cron_schedules", "cron_logs",
                    "faqs", "faq_videos", "faq_imagens", "ad_cache_users", "sccm_cache_devices",
                    "plantoes_matutino", "unidades", "ramais_mpms", "senhas_cofre", "notificacoes",
                }:
                    self.assertIn(req_tbl, tables)

                self.assertGreater(cron_count, 0)
                self.assertGreater(settings_count, 0)
            finally:
                for p in patchers:
                    p.stop()

    def test_postgres_ddl_translates_datetime_to_timestamp(self):
        ddl = "CREATE TABLE IF NOT EXISTS t (id INTEGER PRIMARY KEY AUTOINCREMENT, dt DATETIME DEFAULT CURRENT_TIMESTAMP)"
        translated = PostgreSQLCompatibleCursor._translate_sqlite_query(ddl)
        self.assertIn("SERIAL PRIMARY KEY", translated)
        self.assertIn("TIMESTAMP DEFAULT CURRENT_TIMESTAMP", translated)
        self.assertNotIn("DATETIME", translated)

    def test_launchers_prepare_env_and_directories(self):
        sh_text = (ROOT_DIR / "00-iniciar.sh").read_text(encoding="utf-8")
        cmd_text = (ROOT_DIR / "00-iniciar.cmd").read_text(encoding="utf-8")

        self.assertIn(".env.example", sh_text)
        self.assertIn("ensure_host_directories", sh_text)
        self.assertNotIn("touch data/chamados.db", sh_text)

        self.assertIn(".env.example", cmd_text)
        self.assertIn("01 - Dados Brutos", cmd_text)


if __name__ == "__main__":
    unittest.main()

