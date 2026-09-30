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
from src.tabs.sccm import format_sccm_datetime, parse_sccm_datetime
from src.tabs.links_faqs import (
    parse_sharepoint_content,
    format_file_size,
    scan_video_faqs,
    scan_image_faqs
)

class TestComponentsAndTabs(unittest.TestCase):
    def test_format_sccm_datetime(self):
        """Valida conversão de datas WMI e ISO do SCCM para formato brasileiro DD/MM/AAAA HH:MM:SS e datetime nativo."""
        # WMI CIM DateTime (YYYYMMDDHHmmss.microsec+tz)
        wmi_val = "20260929080059.657000+***"
        self.assertEqual(format_sccm_datetime(wmi_val), "29/09/2026 08:00:59")
        parsed_wmi = parse_sccm_datetime(wmi_val)
        self.assertIsNotNone(parsed_wmi)
        self.assertEqual(parsed_wmi.year, 2026)
        self.assertEqual(parsed_wmi.month, 9)
        self.assertEqual(parsed_wmi.day, 29)

        # ISO 8601
        iso_val = "2026-09-29T16:30:00"
        self.assertEqual(format_sccm_datetime(iso_val), "29/09/2026 16:30:00")
        parsed_iso = parse_sccm_datetime(iso_val)
        self.assertIsNotNone(parsed_iso)
        self.assertEqual(parsed_iso.hour, 16)
        self.assertEqual(parsed_iso.minute, 30)

        # Valores vazios ou nulos
        self.assertEqual(format_sccm_datetime(None), "-")
        self.assertEqual(format_sccm_datetime(""), "-")
        self.assertIsNone(parse_sccm_datetime(None))
        self.assertIsNone(parse_sccm_datetime(""))
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

    def test_faq_format_file_size(self):
        """Valida a conversão correta de bytes para KB, MB e GB."""
        self.assertEqual(format_file_size(500), "0.5 KB")
        self.assertEqual(format_file_size(1024 * 1024 * 5), "5.0 MB")
        self.assertEqual(format_file_size(1024 * 1024 * 1024 * 2), "2.00 GB")

    def test_faq_parse_sharepoint_content_image_plugin(self):
        """Valida substituição de div.imagePlugin por figure e link com abertura no SharePoint."""
        raw_html = '''
        <p>Texto anterior</p>
        <div class="imagePlugin hasCaption" data-captiontext="Legenda de Teste" data-imageurl="/sites/teste/foto.jpg">
            <div>interior</div>
        </div>
        <p>Texto posterior</p>
        '''
        parsed = parse_sharepoint_content(raw_html)
        self.assertIn("https://ministeriopublicoms.sharepoint.com/sites/teste/foto.jpg", parsed)
        self.assertIn("Legenda de Teste", parsed)
        self.assertIn("sp-img-card", parsed)
        self.assertNotIn("imagePlugin", parsed)

    def test_faq_parse_sharepoint_content_clean_author_and_icons(self):
        """Valida remoção do cabeçalho de autoria e tags de ícones quebrados <i>."""
        raw_html = '''
        <div>Paulo Henrique Gonçalves Rezende Published 10/05/2024</div>
        <p>Passo 1: Conecte o cabo <i class="ms-Icon ms-Icon--Plug"></i></p>
        '''
        parsed = parse_sharepoint_content(raw_html)
        self.assertNotIn("Paulo Henrique", parsed)
        self.assertNotIn("<i", parsed)
        self.assertIn("Passo 1: Conecte o cabo", parsed)

    def test_faq_parse_sharepoint_content_video_controldata(self):
        """Valida que controldata com vídeo SharePoint é convertido em tag <video>."""
        raw_html = '''
        <div data-sp-controldata='{"properties": {"file": "/sites/dit/video.mp4"}}'>
            Vídeo incorporado
        </div>
        '''
        parsed = parse_sharepoint_content(raw_html)
        self.assertIn("<video", parsed)
        self.assertIn("https://ministeriopublicoms.sharepoint.com/sites/dit/video.mp4", parsed)

    def test_faq_parse_sharepoint_content_sp_video_embed(self):
        """Valida conversão de div.sp-video-embed em card visual de vídeo."""
        raw_html = '''
        <div class="sp-video-embed" data-videourl="https://ministeriopublicoms.sharepoint.com/sites/dit/teste.mp4" data-videotitle="Tutorial Teste">
        </div>
        '''
        parsed = parse_sharepoint_content(raw_html)
        self.assertTrue("<video" in parsed or "sp-img-card" in parsed or "sp-video-card" in parsed)
        self.assertIn("Tutorial Teste", parsed)

    def test_faq_scan_directories_empty_or_nonexistent(self):
        """Valida comportamento seguro ao varrer pastas inexistentes."""
        nonexistent = Path("/caminho/completamente/inexistente/xyz_123")
        self.assertEqual(scan_video_faqs(nonexistent), [])
        self.assertEqual(scan_image_faqs(nonexistent), [])

if __name__ == "__main__":
    unittest.main()
