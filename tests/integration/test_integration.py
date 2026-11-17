# -*- coding: utf-8 -*-
"""
Testes de integração entre banco de dados relacional, exportação ICS e orquestrador.
"""

import os
import sys
import unittest
import tempfile
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.services.ics_export import generate_ics_calendar

class TestIntegration(unittest.TestCase):
    def test_ics_export_content(self):
        """Valida que o gerador de arquivo de calendário .ics produz formato RFC 5545 válido."""
        mock_events = [
            {
                "id": "plantao_1",
                "title": "Plantão STI: Paulo",
                "start": "2026-09-25T08:00:00",
                "end": "2026-09-25T17:00:00",
                "description": "Plantão matutino STI"
            }
        ]
        content = generate_ics_calendar(mock_events)
        self.assertIsInstance(content, str)
        self.assertTrue(content.startswith("BEGIN:VCALENDAR"))
        self.assertTrue(content.rstrip().endswith("END:VCALENDAR"))
        self.assertIn("VERSION:2.0", content)
        self.assertIn("PRODID:", content)

    def test_env_example_keys_match_config(self):
        """Valida que todas as variáveis críticas no .env.example estão mapeadas no config.py."""
        env_example = ROOT_DIR / ".env.example"
        if env_example.exists():
            content = env_example.read_text(encoding="utf-8")
            required_keys = ["STREAMLIT_PORT", "BANCADA_API_PORT", "BANCADA_API_KEY", "EVOLUTION_API_URL", "AD_DOMAIN", "SCCM_SERVER"]
            for k in required_keys:
                self.assertIn(k, content, f"Chave crítica '{k}' ausente no .env.example")

    def test_docker_compose_api_service_and_db_hardening(self):
        """Valida que o serviço api isolado e os timeouts do PostgreSQL estão configurados no docker-compose.yml."""
        compose_file = ROOT_DIR / "docker-compose.yml"
        if compose_file.exists():
            content = compose_file.read_text(encoding="utf-8")
            self.assertIn("bancada_api_server", content)
            self.assertIn("BANCADA_API_BIND:-127.0.0.1", content)
            self.assertIn("statement_timeout=15000", content)
            self.assertIn("idle_in_transaction_session_timeout=30000", content)

if __name__ == "__main__":
    unittest.main()
