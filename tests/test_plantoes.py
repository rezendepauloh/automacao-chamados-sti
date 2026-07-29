# -*- coding: utf-8 -*-
import sys
import unittest
from pathlib import Path
import pandas as pd

# Adiciona a raiz do projeto e a pasta src ao sys.path
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

from src.plantoes_scraper import parse_data_matutino_string, parse_simp_periodo
from src.database import (
    setup_plantoes_tables, 
    save_plantoes_matutino, 
    save_plantoes_semanal, 
    get_plantoes_matutino, 
    get_plantoes_semanal
)
from src.tabs.plantoes import is_bancada_member


class TestPlantoes(unittest.TestCase):
    """Testes unitários para o módulo de plantões matutino e semanal SIMP."""

    def test_is_bancada_member(self):
        """Testa se a filtragem identifica corretamente os 3 integrantes da bancada."""
        self.assertTrue(is_bancada_member("Paulo Henrique Gonçalves Rezende"))
        self.assertTrue(is_bancada_member("Reginaldo da Silva Bandeira"))
        self.assertTrue(is_bancada_member("Luiz Leonardo Villalba"))
        self.assertFalse(is_bancada_member("Fulano de Tal"))

    def test_parse_data_matutino_string(self):
        """Testa a conversão de datas textuais da planilha de plantão matutino."""
        dt_iso, dia_sem = parse_data_matutino_string("Segunda-feira - 12 de Janeiro", 2026)
        self.assertEqual(dt_iso, "2026-01-12")
        self.assertEqual(dia_sem, "Segunda-feira")

        dt_iso2, dia_sem2 = parse_data_matutino_string("Terça-feira - 03 de Fevereiro", 2026)
        self.assertEqual(dt_iso2, "2026-02-03")
        self.assertEqual(dia_sem2, "Terça-feira")

    def test_parse_simp_periodo(self):
        """Testa o parsing de strings de período do portal SIMP."""
        periodo = "06/07/2026 19:01 a 13/07/2026 11:59"
        dt_ini, dt_fim = parse_simp_periodo(periodo, 2026)
        self.assertEqual(dt_ini, "2026-07-06 19:01:00")
        self.assertEqual(dt_fim, "2026-07-13 11:59:00")

    def test_database_plantoes_crud(self):
        """Testa se a gravação e leitura de plantões no SQLite funcionam perfeitamente."""
        setup_plantoes_tables()
        
        rec_mat = [{
            "ano": 2026, "data_iso": "2026-01-20", "dia_semana": "Terça-feira",
            "servidor": "Reginaldo da Silva Bandeira", "telefone": "+55 67 991455446"
        }]
        save_plantoes_matutino(rec_mat)
        df_mat = get_plantoes_matutino(2026)
        self.assertFalse(df_mat.empty)
        self.assertIn("Reginaldo da Silva Bandeira", df_mat['servidor'].values)

        rec_sem = [{
            "ano": 2026, "mes": "Julho", "periodo_str": "06/07/2026 19:01 a 13/07/2026 11:59",
            "data_inicio": "2026-07-06 19:01:00", "data_fim": "2026-07-13 11:59:00",
            "service_desk": "Anderson Miranda",
            "manutencao": "Paulo Henrique Gonçalves Rezende",
            "infraestrutura": "Joabe Guimarães",
            "desenvolvimento": "Albert Einstein"
        }]
        save_plantoes_semanal(rec_sem)


        df_sem = get_plantoes_semanal(2026)
        self.assertFalse(df_sem.empty)
        self.assertIn("Paulo Henrique Gonçalves Rezende", df_sem['manutencao'].values)


if __name__ == "__main__":
    unittest.main()
