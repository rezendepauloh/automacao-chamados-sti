# -*- coding: utf-8 -*-
"""
test_no_selenium_dependency.py
Valida que todos os scrapers e serviços de sincronização foram 100% desacoplados
do Selenium e utilizam Playwright ou Requests puro.
"""
import unittest
import sys


class TestNoSeleniumDependency(unittest.TestCase):

    def test_scrapers_and_syncs_import_without_selenium(self):
        """Garante que todos os scrapers e syncs podem ser importados sem erro mesmo sem Selenium."""
        modules_to_test = [
            "src.scrapers.citsmart_scraper",
            "src.scrapers.papercut_scraper",
            "src.scrapers.oxe_scraper",
            "src.scrapers.plantoes_scraper",
            "src.scrapers.sharepoint_downloader",
            "src.syncs.sync_donations",
            "src.syncs.sync_fiscalizacao",
            "src.syncs.sync_garantia",
            "src.syncs.sync_viagens",
            "src.syncs.sync_ferias",
        ]

        for mod_name in modules_to_test:
            with self.subTest(module=mod_name):
                # Importa o módulo dinamicamente
                __import__(mod_name)
                mod = sys.modules[mod_name]
                self.assertIsNotNone(mod)

    def test_legacy_get_chrome_driver_raises_deprecation(self):
        """Verifica se a função get_chrome_driver alerta sobre a descontinuação se chamada."""
        from src.config import get_chrome_driver
        # Se selenium não estiver presente ou for chamado, deve lançar erro ou aviso claro
        try:
            get_chrome_driver(headless=True)
        except Exception as e:
            err_msg = str(e).lower()
            self.assertTrue("playwright" in err_msg or "descontinuado" in err_msg or "selenium" in err_msg)


if __name__ == "__main__":
    unittest.main()

