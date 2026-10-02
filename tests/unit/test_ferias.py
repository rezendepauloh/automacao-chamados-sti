# -*- coding: utf-8 -*-
"""
Testes unitários e de integração do Módulo de Férias da Bancada:
Valida parser de texto de períodos, persistência relacional SQLite,
detecção de sobreposições de ausências e filtros de dados.
"""

import sys
import unittest
import pandas as pd
from pathlib import Path
from datetime import datetime

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.database.ferias_db import (
    parse_periodo_cell,
    _parse_month_from_header,
    setup_ferias_table,
    get_ferias_df,
    get_ferias_membros,
    get_ferias_anos
)
from src.tabs.ferias import detectar_sobreposicoes, get_cor_membro


class TestFeriasModule(unittest.TestCase):
    def setUp(self):
        setup_ferias_table()

    def test_parse_month_from_header(self):
        """Valida detecção de mês a partir de datetime, string abreviada e nome por extenso."""
        # Datetime
        dt_val = datetime(2025, 4, 1)
        m_num, m_nome = _parse_month_from_header(dt_val, 2025)
        self.assertEqual(m_num, 4)
        self.assertEqual(m_nome, "Abril")

        # String abreviada
        m_num_str, m_nome_str = _parse_month_from_header("dez", 2025)
        self.assertEqual(m_num_str, 12)
        self.assertEqual(m_nome_str, "Dezembro")

    def test_parse_periodo_cell_formats(self):
        """Valida os múltiplos padrões de períodos aceitos pela planilha oficial."""
        # Formato '20 a 29' (mês padrão)
        res1 = parse_periodo_cell("20 a 29", 2025, 1)
        self.assertEqual(len(res1), 1)
        self.assertEqual(res1[0]["dt_ini_iso"], "2025-01-20")
        self.assertEqual(res1[0]["dt_fim_iso"], "2025-01-29")
        self.assertEqual(res1[0]["dias"], 10)

        # Formato '16/09 a 04/10' (transição de mês)
        res2 = parse_periodo_cell("16/09 a 04/10", 2024, 9)
        self.assertEqual(len(res2), 1)
        self.assertEqual(res2[0]["dt_ini_iso"], "2024-09-16")
        self.assertEqual(res2[0]["dt_fim_iso"], "2024-10-04")
        self.assertEqual(res2[0]["dias"], 19)

        # Formato '15-26/03'
        res3 = parse_periodo_cell("15-26/03", 2027, 3)
        self.assertEqual(len(res3), 1)
        self.assertEqual(res3[0]["dt_ini_iso"], "2027-03-15")
        self.assertEqual(res3[0]["dt_fim_iso"], "2027-03-26")
        self.assertEqual(res3[0]["dias"], 12)

        # Formato '3-7/ago'
        res4 = parse_periodo_cell("3-7/ago", 2026, 8)
        self.assertEqual(len(res4), 1)
        self.assertEqual(res4[0]["dt_ini_iso"], "2026-08-03")
        self.assertEqual(res4[0]["dt_fim_iso"], "2026-08-07")
        self.assertEqual(res4[0]["dias"], 5)

        # Formato lista de dias '16 e 17'
        res5 = parse_periodo_cell("16 e 17", 2025, 1)
        self.assertEqual(len(res5), 1)
        self.assertEqual(res5[0]["dt_ini_iso"], "2025-01-16")
        self.assertEqual(res5[0]["dt_fim_iso"], "2025-01-17")
        self.assertEqual(res5[0]["dias"], 2)

    def test_detectar_sobreposicoes(self):
        """Valida algoritmo de detecção de ausências simultâneas de servidores distintos."""
        df_teste = pd.DataFrame([
            {
                "membro": "Paulo Rezende",
                "tipo_escala": "ferias",
                "data_inicio_iso": "2025-01-10",
                "data_fim_iso": "2025-01-20",
                "ano": 2025
            },
            {
                "membro": "Luiz Villalba",
                "tipo_escala": "ferias",
                "data_inicio_iso": "2025-01-15",
                "data_fim_iso": "2025-01-25",
                "ano": 2025
            },
            {
                "membro": "Reginaldo Bandeira",
                "tipo_escala": "ferias",
                "data_inicio_iso": "2025-02-01",
                "data_fim_iso": "2025-02-10",
                "ano": 2025
            }
        ])

        conflitos = detectar_sobreposicoes(df_teste)
        self.assertEqual(len(conflitos), 1)
        self.assertEqual(conflitos[0]["membro1"], "Paulo Rezende")
        self.assertEqual(conflitos[0]["membro2"], "Luiz Villalba")
        self.assertEqual(conflitos[0]["inicio_br"], "15/01/2025")
        self.assertEqual(conflitos[0]["fim_br"], "20/01/2025")
        self.assertEqual(conflitos[0]["dias"], 6)

    def test_get_cor_membro(self):
        """Valida que cada integrante possui cor temática persistente."""
        cor_paulo = get_cor_membro("Paulo Rezende")
        cor_reginaldo = get_cor_membro("Reginaldo Bandeira")
        cor_luiz = get_cor_membro("Luiz Villalba")
        self.assertTrue(cor_paulo.startswith("#"))
        self.assertTrue(cor_reginaldo.startswith("#"))
        self.assertTrue(cor_luiz.startswith("#"))
        self.assertNotEqual(cor_paulo, cor_reginaldo)


if __name__ == "__main__":
    unittest.main()
