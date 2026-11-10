import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path

from src.config import (
    get_spreadsheet_setting,
    is_sharepoint_url,
    resolve_spreadsheet_path,
    _cfg,
    BASE_DIR,
    USER_HOME
)

class TestSpreadsheetConfig(unittest.TestCase):
    """
    Testes unitários para validação da esteira de resolução de caminhos e URLs de planilhas.
    Garante o cumprimento estrito das prioridades:
    1. Banco de Dados (system_settings)
    2. Fallback de Ambiente (.env)
    3. Valor default
    """

    def test_priority_1_db_takes_precedence_over_env(self):
        """Valida que o valor persistido no banco de dados tem prioridade absoluta sobre o .env."""
        with patch("src.database.settings_db.get_setting", return_value="https://sharepoint.com/planilha_db.xlsx"):
            with patch.dict("os.environ", {"DONATIONS_EXCEL_RELATIVE_PATH": "caminho_env_antigo.xlsx"}):
                res = get_spreadsheet_setting("DONATIONS_EXCEL_RELATIVE_PATH", default="default.xlsx")
                self.assertEqual(res, "https://sharepoint.com/planilha_db.xlsx")

    def test_priority_2_env_fallback_when_db_is_empty(self):
        """Valida que, se o banco estiver vazio ou retornar None/vazio, recorre ao .env."""
        with patch("src.database.settings_db.get_setting", return_value=""):
            with patch.dict("os.environ", {"DONATIONS_EXCEL_RELATIVE_PATH": "https://sharepoint.com/planilha_env.xlsx"}):
                res = get_spreadsheet_setting("DONATIONS_EXCEL_RELATIVE_PATH", default="default.xlsx")
                self.assertEqual(res, "https://sharepoint.com/planilha_env.xlsx")

    def test_priority_3_default_fallback_when_both_empty(self):
        """Valida que, se banco e .env forem vazios, o valor default é retornado."""
        with patch("src.database.settings_db.get_setting", return_value=""):
            with patch.dict("os.environ", {"DONATIONS_EXCEL_RELATIVE_PATH": ""}):
                with patch("dotenv.dotenv_values", return_value={}):
                    res = get_spreadsheet_setting("DONATIONS_EXCEL_RELATIVE_PATH", default="https://default.com/planilha.xlsx")
                    self.assertEqual(res, "https://default.com/planilha.xlsx")

    def test_is_sharepoint_url(self):
        """Valida o detector de URLs remotas do SharePoint."""
        self.assertTrue(is_sharepoint_url("https://ministeriopublicoms.sharepoint.com/:x:/s/..."))
        self.assertTrue(is_sharepoint_url("http://sharepoint.local/doc.xlsx"))
        self.assertTrue(is_sharepoint_url("  HTTPS://sharepoint.com/doc.xlsx  "))
        
        # Não são URLs
        self.assertFalse(is_sharepoint_url("OneDrive - MPMS\\Planilha.xlsx"))
        self.assertFalse(is_sharepoint_url(r"C:\Users\test\Planilha.xlsx"))
        self.assertFalse(is_sharepoint_url("uploads/planilha.xlsx"))
        self.assertFalse(is_sharepoint_url(""))
        self.assertFalse(is_sharepoint_url(None))

    def test_resolve_spreadsheet_path_url_safety(self):
        """Garante que URLs remotas NUNCA são concatenadas com USER_HOME (o que corromperia o Path)."""
        url = "https://ministeriopublicoms.sharepoint.com/:x:/s/dit-manutencao/test.xlsx"
        with patch("src.config.get_spreadsheet_setting", return_value=url):
            resolved = resolve_spreadsheet_path("DONATIONS_EXCEL_RELATIVE_PATH", "doacoes_redistribuicao.xlsx")
            self.assertIsInstance(resolved, Path)
            # Deve apontar para o cache local em uploads/
            expected = BASE_DIR / "uploads" / "doacoes_redistribuicao.xlsx"
            self.assertEqual(resolved, expected)
            self.assertNotIn("https:", str(resolved))

    def test_resolve_spreadsheet_path_local_relative(self):
        """Garante que caminhos locais relativos são resolvidos a partir de USER_HOME."""
        rel_path = "Documentos/Planilha_Local.xlsx"
        with patch("src.config.get_spreadsheet_setting", return_value=rel_path):
            resolved = resolve_spreadsheet_path("DONATIONS_EXCEL_RELATIVE_PATH", "doacoes_redistribuicao.xlsx")
            self.assertIsInstance(resolved, Path)
            expected = USER_HOME / rel_path
            self.assertEqual(resolved, expected)

    def test_resolve_spreadsheet_path_local_absolute(self):
        """Garante que caminhos absolutos locais são preservados sem concatenação duplicada."""
        abs_path = "/tmp/planilha_absoluta.xlsx"
        with patch("src.config.get_spreadsheet_setting", return_value=abs_path):
            resolved = resolve_spreadsheet_path("DONATIONS_EXCEL_RELATIVE_PATH")
            self.assertEqual(resolved, Path(abs_path))

if __name__ == "__main__":
    unittest.main()

