# -*- coding: utf-8 -*-
"""
Testes unitários da organização centralizada de diretórios em data/.
Garante que caminhos de dados brutos, tratados, prontos, SQLite fallback,
calendário ICS e inventário SCCM apontam para data/.
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.config import (
    BASE_DIR,
    DATA_DIR,
    INPUT_DIR_BRUTOS,
    OUTPUT_DIR_TRATADOS,
    OUTPUT_DIR_PRONTO,
    SQLITE_DB_PATH,
    CALENDAR_ICS_PATH,
    BACKUP_PATH_OTRS,
    TEMP_PATH_OTRS,
)
from src.database.connection import DB_PATH
from src.services.ics_export import update_published_ics_file


class TestDataDirectoryPaths(unittest.TestCase):
    """Valida a centralização dos diretórios e arquivos de dados em data/."""

    def test_config_data_directories_point_to_data_root(self):
        self.assertEqual(DATA_DIR, BASE_DIR / "data")
        self.assertEqual(INPUT_DIR_BRUTOS, DATA_DIR / "01 - Dados Brutos")
        self.assertEqual(OUTPUT_DIR_TRATADOS, DATA_DIR / "02 - Dados tratados")
        self.assertEqual(OUTPUT_DIR_PRONTO, DATA_DIR / "03 - Dados prontos")
        self.assertEqual(SQLITE_DB_PATH, DATA_DIR / "chamados.db")
        self.assertEqual(CALENDAR_ICS_PATH, DATA_DIR / "calendario.ics")
        self.assertEqual(BACKUP_PATH_OTRS.parent, DATA_DIR)
        self.assertEqual(TEMP_PATH_OTRS.parent, DATA_DIR)

    def test_database_connection_sqlite_fallback_points_to_data(self):
        self.assertEqual(Path(DB_PATH).resolve(), (BASE_DIR / "data" / "chamados.db").resolve())

    def test_update_published_ics_writes_to_data_and_not_root(self):
        written_paths = []

        def fake_write_text(self_path, data, encoding="utf-8"):
            written_paths.append(self_path.resolve())
            return len(data)

        with patch("src.services.ics_export.fetch_all_unified_calendar_events", return_value=[]):
            with patch("pathlib.Path.mkdir"):
                with patch("pathlib.Path.write_text", new=fake_write_text):
                    update_published_ics_file()

        self.assertIn((BASE_DIR / "data" / "calendario.ics").resolve(), written_paths)
        self.assertIn((BASE_DIR / "src" / "js" / "calendario.ics").resolve(), written_paths)
        self.assertNotIn((BASE_DIR / "calendario.ics").resolve(), written_paths)

    def test_sccm_launcher_and_service_use_data_brutos(self):
        launcher = BASE_DIR / "src" / "protocol_handler" / "bancada-launcher.ps1"
        content = launcher.read_text(encoding="utf-8-sig")
        self.assertIn(r"data\01 - Dados Brutos\sccm_inventory.json", content)
        self.assertNotIn(r"\automacao-chamados-sti\01 - Dados Brutos", content)
        self.assertNotIn(r"\automated-OTRS-and-CitSmart\01 - Dados Brutos", content)

        sccm_service = (BASE_DIR / "src" / "services" / "sccm_service.py").read_text(encoding="utf-8")
        self.assertIn("data/01 - Dados Brutos/sccm_inventory.json", sccm_service)


if __name__ == "__main__":
    unittest.main()

