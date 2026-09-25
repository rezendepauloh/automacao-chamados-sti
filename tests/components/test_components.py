# -*- coding: utf-8 -*-
"""
Testes unitários de componentes visuais e abas (src.components.* e src.tabs.*).
Valida metric_cards, subtabs, status_banner, mapas dijkstra e utilitários de interface.
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.components.status_banner import check_orquestrador_running, read_last_log_lines
from src.tabs.chamados import summarize_ticket_locally
from src.tabs.fiscalizacao import _formatar_texto_portaria
from src.tabs.mapas import calculate_dijkstra_route, get_image_base64
from src.tabs.plantoes import is_bancada_member

class TestComponentsAndTabs(unittest.TestCase):
    def test_status_banner_functions(self):
        """Valida leitura de logs e verificação do status do orquestrador."""
        is_running = check_orquestrador_running()
        self.assertIsInstance(is_running, bool)

        logs = read_last_log_lines(5)
        self.assertIsInstance(logs, str)

    def test_summarize_ticket_locally(self):
        """Valida gerador de resumo conciso do corpo do chamado."""
        desc = "Bom dia. Gostaria de solicitar formatação do computador devido à lentidão extrema."
        summary = summarize_ticket_locally(desc, "", max_sentences=1)
        self.assertIsInstance(summary, str)
        self.assertTrue(len(summary) > 0)

    def test_formatar_texto_portaria(self):
        """Valida realce de servidores designados em texto de portarias."""
        texto = "Designar o servidor Paulo Henrique Gonçalves Rezende para comissão técnica."
        destacado = _formatar_texto_portaria(texto, nomes_destacar=["Paulo Henrique Gonçalves Rezende"])
        self.assertIn("🟢 Paulo Henrique Gonçalves Rezende", destacado)

    def test_calculate_dijkstra_route_empty(self):
        """Valida que grafo vazio não falha no cálculo de rotas do mapa."""
        caminhos_vazios = {"nós": [], "arestas": []}
        start = {"pavimento_id": 0, "x": 0, "y": 0}
        end = {"pavimento_id": 0, "x": 10, "y": 10}
        rota = calculate_dijkstra_route(caminhos_vazios, start, end)
        self.assertEqual(rota, [])

    def test_get_image_base64_nonexistent(self):
        """Valida retorno seguro ao buscar imagem inexistente para planta de mapa."""
        b64 = get_image_base64(Path("arquivo_inexistente_123.png"))
        self.assertEqual(b64, "")

    def test_is_bancada_member(self):
        """Valida membros oficiais da bancada técnica."""
        self.assertTrue(is_bancada_member("Paulo Henrique Gonçalves Rezende"))
        self.assertTrue(is_bancada_member("Reginaldo da Silva Bandeira"))
        self.assertTrue(is_bancada_member("Luiz Leonardo Villalba"))
        self.assertFalse(is_bancada_member("Usuário Qualquer"))

if __name__ == "__main__":
    unittest.main()
