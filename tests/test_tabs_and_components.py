# -*- coding: utf-8 -*-
import sys
import unittest
from pathlib import Path
import pandas as pd

# Adiciona a raiz do projeto e a pasta src ao sys.path
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

from src.components.status_banner import check_orquestrador_running, read_last_log_lines
from src.database import load_data
from src.tabs.chamados import summarize_ticket_locally
from src.tabs.fiscalizacao import _formatar_texto_portaria
from src.tabs.mapas import calculate_dijkstra_route, get_image_base64


class TestComponentsAndTabs(unittest.TestCase):
    """Testes unitários para as novas abas e componentes refatorados."""

    def test_status_banner_functions(self):
        """Testa as funções de checagem do orquestrador e leitura de logs."""
        is_running = check_orquestrador_running()
        self.assertIsInstance(is_running, bool)

        logs = read_last_log_lines(5)
        self.assertIsInstance(logs, str)

    def test_load_data_returns_dataframe(self):
        """Testa se load_data retorna um DataFrame pandas."""
        df = load_data()
        self.assertIsInstance(df, pd.DataFrame)

    def test_summarize_ticket_locally(self):
        """Testa o resumidor local de chamados NLP."""
        desc = "Bom dia. Gostaria de solicitar a formatação do computador devido à extrema lentidão e travamentos no sistema."
        summary = summarize_ticket_locally(desc, "", max_sentences=1)
        self.assertIsInstance(summary, str)
        self.assertTrue(len(summary) > 0)

    def test_formatar_texto_portaria(self):
        """Testa a formatação e destaque de portarias do MPMS."""
        texto_bruto = "Designar o servidor Paulo Henrique Gonçalves Rezende para atuar como fiscal titular.\n\nProcuradoria-Geral de Justiça"
        destacado = _formatar_texto_portaria(texto_bruto, nomes_destacar=["Paulo Henrique Gonçalves Rezende"])
        self.assertIn("🟢 Paulo Henrique Gonçalves Rezende", destacado)
        self.assertNotIn("Procuradoria-Geral de Justiça", destacado)

    def test_calculate_dijkstra_route_empty(self):
        """Testa se o cálculo de rota Dijkstra lida com mapas vazios de forma segura."""
        caminhos_vazios = {"nós": [], "arestas": []}
        start_pin = {"pavimento_id": 0, "x": 10, "y": 10}
        end_pin = {"pavimento_id": 0, "x": 50, "y": 50}
        route = calculate_dijkstra_route(caminhos_vazios, start_pin, end_pin)
        self.assertEqual(route, [])

    def test_get_image_base64_nonexistent(self):
        """Testa fallback da conversão de imagem inexistente para base64."""
        b64 = get_image_base64(Path("caminho_inexistente.png"))
        self.assertEqual(b64, "")


if __name__ == "__main__":
    unittest.main()
