import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path
import json

from src.scrapers.sharepoint_downloader import (
    build_download_url,
    get_cached_cookies,
    save_cached_cookies,
    download_via_requests,
    download_sharepoint_file
)

class TestSharepointDownloader(unittest.TestCase):
    """Testes unitários para o módulo acelerado de download do SharePoint."""

    def test_build_download_url(self):
        """Testa a formatação segura de URLs com parâmetro de download."""
        self.assertEqual(
            build_download_url("https://sharepoint.com/:x:/s/doc?e=123"),
            "https://sharepoint.com/:x:/s/doc?e=123&download=1"
        )
        self.assertEqual(
            build_download_url("https://sharepoint.com/:x:/s/doc"),
            "https://sharepoint.com/:x:/s/doc?download=1"
        )
        # Não duplica
        self.assertEqual(
            build_download_url("https://sharepoint.com/:x:/s/doc?download=1"),
            "https://sharepoint.com/:x:/s/doc?download=1"
        )

    @patch("src.database.settings_db.get_setting")
    def test_get_cached_cookies(self, mock_get):
        """Testa o carregamento de cookies em cache do banco de dados."""
        mock_get.return_value = json.dumps([{"name": "FedAuth", "value": "secret123"}])
        cookies = get_cached_cookies()
        self.assertIsNotNone(cookies)
        self.assertEqual(len(cookies), 1)
        self.assertEqual(cookies[0]["name"], "FedAuth")

        # Caso vazio ou inválido
        mock_get.return_value = ""
        self.assertIsNone(get_cached_cookies())

        mock_get.return_value = "invalid-json"
        self.assertIsNone(get_cached_cookies())

    @patch("src.database.settings_db.set_setting", return_value=True)
    def test_save_cached_cookies(self, mock_set):
        """Testa o salvamento seguro de cookies no banco."""
        res = save_cached_cookies([{"name": "rtFa", "value": "val456"}])
        self.assertTrue(res)
        mock_set.assert_called_once()
        key_arg = mock_set.call_args.kwargs.get("key") or (mock_set.call_args[0][0] if mock_set.call_args[0] else None)
        self.assertEqual(key_arg, "SHAREPOINT_SESSION_COOKIES")

    @patch("requests.Session.get")
    def test_download_via_requests_success(self, mock_get):
        """Testa download bem-sucedido via HTTP requests com cabeçalho PK de Excel."""
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.content = b"PK\x03\x04conteudo_falso_excel"
        mock_get.return_value = mock_resp

        out_path = Path("/tmp/test_excel_ok.xlsx")
        cookies = [{"name": "FedAuth", "value": "token"}]
        res = download_via_requests("https://sharepoint.com/doc", out_path, cookies)

        self.assertEqual(res, out_path)
        self.assertTrue(out_path.exists())
        self.assertEqual(out_path.read_bytes(), b"PK\x03\x04conteudo_falso_excel")
        out_path.unlink(missing_ok=True)

    @patch("requests.Session.get")
    def test_download_via_requests_rejection_html(self, mock_get):
        """Testa rejeição caso o SharePoint retorne tela de login HTML."""
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.content = b"<html><head><title>Sign in</title></head></html>"
        mock_get.return_value = mock_resp

        out_path = Path("/tmp/test_excel_fail.xlsx")
        cookies = [{"name": "FedAuth", "value": "expired"}]
        res = download_via_requests("https://sharepoint.com/doc", out_path, cookies)

        self.assertIsNone(res)
        self.assertFalse(out_path.exists())

    @patch("src.scrapers.sharepoint_downloader.get_cached_cookies")
    @patch("src.scrapers.sharepoint_downloader.download_via_requests")
    def test_download_sharepoint_file_tier1_fast_path(self, mock_dl, mock_cookies):
        """Testa que se o cache for válido, ele nunca chama o Playwright (Tier 1 Fast Path)."""
        mock_cookies.return_value = [{"name": "FedAuth", "value": "valid"}]
        expected_path = Path("/tmp/result.xlsx")
        mock_dl.return_value = expected_path

        with patch("src.scrapers.sharepoint_downloader.renew_cookies_with_playwright") as mock_pw:
            res = download_sharepoint_file("https://sharepoint.com/doc", expected_path)
            self.assertEqual(res, expected_path)
            mock_pw.assert_not_called()

    @patch("src.scrapers.sharepoint_downloader.get_cached_cookies", return_value=None)
    @patch("src.scrapers.sharepoint_downloader.renew_cookies_with_playwright")
    @patch("src.scrapers.sharepoint_downloader.download_via_requests")
    def test_download_sharepoint_file_tier2_fallback(self, mock_dl, mock_pw, mock_cookies):
        """Testa que se o cache estiver ausente, ele aciona a renovação via Playwright (Tier 2)."""
        mock_pw.return_value = [{"name": "FedAuth", "value": "fresh"}]
        expected_path = Path("/tmp/result.xlsx")
        mock_dl.return_value = expected_path

        res = download_sharepoint_file("https://sharepoint.com/doc", expected_path)
        self.assertEqual(res, expected_path)
        mock_pw.assert_called_once()

if __name__ == "__main__":
    unittest.main()
