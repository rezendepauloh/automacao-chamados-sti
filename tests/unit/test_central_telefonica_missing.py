# -*- coding: utf-8 -*-
"""
Testes unitários para detecção de números de ramais não cadastrados (lacunas sequenciais)
na aba Central Telefônica (OXE) — `src/tabs/central_telefonica.py`.
"""

import sys
import unittest
from pathlib import Path
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers  # noqa: F401
from src.tabs.central_telefonica import (
    MISSING_EXTENSION_CATEGORY,
    MISSING_EXTENSION_DISPLAY_NAME,
    MISSING_EXTENSION_STATUS,
    MAX_CUSTOM_RANGE_SPAN,
    OFFICIAL_EXTENSION_RANGES,
    build_official_ranges_status_df,
    find_missing_extensions,
    format_gap_ranges,
    get_unit_for_extension,
    merge_registered_and_missing,
    summarize_missing_by_range,
)


class TestCentralTelefonicaMissingExtensions(unittest.TestCase):
    def test_detects_sequential_gap_example_from_user(self):
        """Valida o cenário clássico: 2001..2009 pulando para 2015 -> retorna 2010..2014."""
        ramais = [str(n) for n in range(2001, 2010)] + ["2015", "2016"]
        df = pd.DataFrame({"ramal": ramais})

        missing = find_missing_extensions(df)
        self.assertEqual(missing["ramal"].tolist(), ["2010", "2011", "2012", "2013", "2014"])
        self.assertTrue((missing["faixa"] == "2000–2099").all())
        self.assertTrue((missing["ramal_anterior"] == "2009").all())
        self.assertTrue((missing["ramal_proximo"] == "2015").all())
        self.assertTrue((missing["categoria_dispositivo"] == MISSING_EXTENSION_CATEGORY).all())
        self.assertTrue((missing["nome_exibido"] == MISSING_EXTENSION_DISPLAY_NAME).all())
        self.assertTrue((missing["status_equipamento"] == MISSING_EXTENSION_STATUS).all())
        self.assertTrue(missing["is_missing_extension"].all())

    def test_does_not_generate_giant_gaps_between_distinct_hundreds(self):
        """Não deve preencher milhares de números entre faixas distantes (ex: 2000s, 8900s, 171000s)."""
        df = pd.DataFrame({
            "ramal": ["9", "*2252", "2000", "2002", "8900", "8901", "8903", "171000", "171002"]
        })

        missing = find_missing_extensions(df)
        self.assertEqual(missing["ramal"].tolist(), ["2001", "8902", "171001"])
        self.assertEqual(missing["faixa"].tolist(), ["2000–2099", "8900–8999", "171000–171099"])

    def test_empty_and_contiguous_dataframes_return_empty_missing(self):
        """DataFrame vazio ou 100% contíguo não deve gerar lacunas."""
        self.assertTrue(find_missing_extensions(pd.DataFrame()).empty)
        self.assertTrue(find_missing_extensions(pd.DataFrame({"outro": ["1", "3"]})).empty)

        df_full = pd.DataFrame({"ramal": ["2100", "2101", "2102", "2103"]})
        self.assertTrue(find_missing_extensions(df_full).empty)

        # Bloco com apenas 1 ramal não gera lacuna interna
        df_single = pd.DataFrame({"ramal": ["3099"]})
        self.assertTrue(find_missing_extensions(df_single).empty)

    def test_custom_range_includes_boundary_endpoints_and_handles_reversed_bounds(self):
        """Modo faixa personalizada deve incluir pontas livres e tolerar limites invertidos."""
        df = pd.DataFrame({"ramal": ["2001", "2002", "2004"]})

        missing = find_missing_extensions(df, custom_range=(2005, 2000))
        self.assertEqual(missing["ramal"].tolist(), ["2000", "2003", "2005"])
        self.assertTrue((missing["faixa"] == "2000–2005").all())

        # Vizinho de 2000: anterior '-', próximo '2001'
        row_2000 = missing[missing["ramal"] == "2000"].iloc[0]
        self.assertEqual(row_2000["ramal_anterior"], "-")
        self.assertEqual(row_2000["ramal_proximo"], "2001")

        # Vizinho de 2005: anterior '2004', próximo '-'
        row_2005 = missing[missing["ramal"] == "2005"].iloc[0]
        self.assertEqual(row_2005["ramal_anterior"], "2004")
        self.assertEqual(row_2005["ramal_proximo"], "-")

    def test_custom_range_respects_safety_cap(self):
        """Faixas personalizadas gigantes devem ser limitadas por MAX_CUSTOM_RANGE_SPAN."""
        df = pd.DataFrame({"ramal": []})
        missing = find_missing_extensions(df, custom_range=(1000, 1000 + MAX_CUSTOM_RANGE_SPAN + 500))
        self.assertEqual(len(missing), MAX_CUSTOM_RANGE_SPAN)

    def test_format_gap_ranges_compacts_consecutive_numbers(self):
        """Agrupa sequências consecutivas em intervalos legíveis."""
        self.assertEqual(format_gap_ranges([]), "-")
        self.assertEqual(format_gap_ranges([2007]), "2007")
        self.assertEqual(
            format_gap_ranges([2010, 2011, 2012, 2013, 2014, 2020, 2025, 2026]),
            "2010–2014, 2020, 2025–2026",
        )

    def test_summarize_missing_by_range(self):
        """Valida o resumo agregado por faixa (centena)."""
        df = pd.DataFrame({
            "ramal": ["2000", "2001", "2005", "2200", "2202"]
        })
        missing = find_missing_extensions(df)
        summary = summarize_missing_by_range(df, missing)

        self.assertEqual(len(summary), 2)
        row_2000 = summary[summary["Faixa"] == "2000–2099"].iloc[0]
        self.assertEqual(row_2000["Intervalo Ativo"], "2000–2005")
        self.assertEqual(row_2000["Cadastrados"], 3)
        self.assertEqual(row_2000["Disponíveis (Sem Cadastro)"], 3)
        self.assertEqual(row_2000["Primeiro Livre"], "2002")
        self.assertEqual(row_2000["Números Disponíveis (Lacunas)"], "2002–2004")

        row_2200 = summary[summary["Faixa"] == "2200–2299"].iloc[0]
        self.assertEqual(row_2200["Disponíveis (Sem Cadastro)"], 1)
        self.assertEqual(row_2200["Primeiro Livre"], "2201")
        self.assertEqual(row_2200["Números Disponíveis (Lacunas)"], "2201")

    def test_merge_registered_and_missing_preserves_numeric_order(self):
        """Mesclagem deve inserir os ramais livres exatamente na posição numérica correta."""
        df = pd.DataFrame({
            "ramal": ["*399", "2001", "2004"],
            "nome_exibido": ["Especial", "João", "Maria"],
            "mac_address": ["-", "AA:BB:CC:DD:EE:01", "AA:BB:CC:DD:EE:02"],
        })
        missing = find_missing_extensions(df)
        merged = merge_registered_and_missing(df, missing)

        self.assertEqual(merged["ramal"].tolist(), ["2001", "2002", "2003", "2004", "*399"])
        self.assertEqual(merged["is_missing_extension"].tolist(), [False, True, True, False, False])

    def test_official_ranges_and_unit_lookup(self):
        """Valida a tabela oficial de ranges por unidade e o mapeamento de ramal -> unidade."""
        self.assertEqual(len(OFFICIAL_EXTENSION_RANGES), 20)
        self.assertEqual(get_unit_for_extension(4250), "PGJ 1")
        self.assertEqual(get_unit_for_extension(2007), "PGJ 2")
        self.assertEqual(get_unit_for_extension(8948), "PGJ 3")
        self.assertEqual(get_unit_for_extension(9150), "PGJ 4")
        self.assertEqual(get_unit_for_extension(2450), "PJCHA")
        self.assertEqual(get_unit_for_extension(2850), "PJCGR")
        self.assertEqual(get_unit_for_extension(4050), "PJESP")
        self.assertEqual(get_unit_for_extension(8750), "CORUMBÁ")
        self.assertEqual(get_unit_for_extension(8850), "TRÊS LAGOAS")
        self.assertEqual(get_unit_for_extension(1111), "-")

    def test_build_official_ranges_status_df_computes_occupancy_and_first_free(self):
        """Valida o cruzamento da tabela oficial de ranges com os ramais cadastrados na OXE."""
        # Simula faixa 2000..2099 quase lotada (faltando 2007 e 2032) e faixa 2100..2199 lotada
        ramais_2000 = [str(n) for n in range(2000, 2100) if n not in (2007, 2032)]
        ramais_2100 = [str(n) for n in range(2100, 2200)]
        df = pd.DataFrame({"ramal": ramais_2000 + ramais_2100})

        status_df = build_official_ranges_status_df(df)
        self.assertEqual(len(status_df), len(OFFICIAL_EXTENSION_RANGES))

        row_2000 = status_df[status_df["Faixa de Ramais"] == "2000–2099"].iloc[0]
        self.assertEqual(row_2000["Unidade"], "PGJ 2")
        self.assertEqual(row_2000["Início"], "67 3318-2000")
        self.assertEqual(row_2000["Fim"], "67 3318-2099")
        self.assertEqual(row_2000["Cadastrados"], 98)
        self.assertEqual(row_2000["Disponíveis"], 2)
        self.assertEqual(row_2000["Ocupação (%)"], 98.0)
        self.assertEqual(row_2000["Primeiro Livre"], "2007")
        self.assertEqual(row_2000["Números Disponíveis"], "2007, 2032")

        row_2100 = status_df[status_df["Faixa de Ramais"] == "2100–2199"].iloc[0]
        self.assertEqual(row_2100["Cadastrados"], 100)
        self.assertEqual(row_2100["Disponíveis"], 0)
        self.assertEqual(row_2100["Ocupação (%)"], 100.0)
        self.assertEqual(row_2100["Primeiro Livre"], "Lotada")

        # Faixa vazia (ex: PGJ 4 9000..9099)
        row_9000 = status_df[status_df["Faixa de Ramais"] == "9000–9099"].iloc[0]
        self.assertEqual(row_9000["Unidade"], "PGJ 4")
        self.assertEqual(row_9000["Cadastrados"], 0)
        self.assertEqual(row_9000["Disponíveis"], 100)
        self.assertEqual(row_9000["Primeiro Livre"], "9000")

        # Verifica unidade oficial no resumo de lacunas
        missing = find_missing_extensions(df)
        summary = summarize_missing_by_range(df, missing)
        self.assertEqual(summary.iloc[0]["Unidade Oficial"], "PGJ 2")


if __name__ == "__main__":
    unittest.main()
