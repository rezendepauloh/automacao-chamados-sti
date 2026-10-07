# -*- coding: utf-8 -*-
"""
Testes unitários e de integração para os mecanismos de Caching do Streamlit (@st.cache_data).
Valida funções cacheadas do Active Directory, SCCM, Central Telefônica, Unidades e Mapas,
além de validar a mecânica de invalidação via st.cache_data.clear().
"""

import sys
import unittest
from pathlib import Path
import pandas as pd
import streamlit as st

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.tabs.active_directory import (
    _cached_get_ad_ous,
    _cached_get_ad_ou_stats,
    _cached_get_all_ou_entities_compact,
    _cached_get_ad_users_df,
    _cached_get_ad_computers_df,
    _cached_get_ad_groups_df,
    _cached_get_ad_departments,
    _cached_get_ad_operating_systems
)
from src.tabs.sccm import (
    _cached_get_sccm_devices_df,
    _cached_get_sccm_users_df,
    _cached_get_sccm_collections_df
)
from src.tabs.central_telefonica import _cached_get_central_telefonica_df
from src.tabs.unidades import _cached_get_unidades_df, _cached_get_ramais_df
from src.tabs.mapas import get_image_dimensions, get_image_base64


class TestCacheSuites(unittest.TestCase):
    def setUp(self):
        # Garante cache limpo antes de cada teste
        st.cache_data.clear()

    def test_ad_cache_functions_return_types(self):
        """Valida que todas as funções cacheadas do Active Directory retornam tipos consistentes."""
        ous = _cached_get_ad_ous()
        self.assertIsInstance(ous, pd.DataFrame)

        ou_stats = _cached_get_ad_ou_stats()
        self.assertIsInstance(ou_stats, dict)

        entities = _cached_get_all_ou_entities_compact()
        self.assertIsInstance(entities, dict)
        self.assertIn("users", entities)
        self.assertIn("comps", entities)

        users = _cached_get_ad_users_df(status_filter="Todos")
        self.assertIsInstance(users, pd.DataFrame)

        computers = _cached_get_ad_computers_df(status_filter="Todos")
        self.assertIsInstance(computers, pd.DataFrame)

        groups = _cached_get_ad_groups_df()
        self.assertIsInstance(groups, pd.DataFrame)

        depts = _cached_get_ad_departments()
        self.assertIsInstance(depts, list)

        os_list = _cached_get_ad_operating_systems()
        self.assertIsInstance(os_list, list)

    def test_sccm_cache_functions_return_types(self):
        """Valida que as funções cacheadas do módulo SCCM retornam DataFrames íntegros."""
        devices = _cached_get_sccm_devices_df()
        self.assertIsInstance(devices, pd.DataFrame)

        users = _cached_get_sccm_users_df()
        self.assertIsInstance(users, pd.DataFrame)

        cols_dev = _cached_get_sccm_collections_df(col_type="Dispositivos")
        self.assertIsInstance(cols_dev, pd.DataFrame)

        cols_usr = _cached_get_sccm_collections_df(col_type="Usuários")
        self.assertIsInstance(cols_usr, pd.DataFrame)

    def test_central_telefonica_cache_function(self):
        """Valida retorno do cache de telefonia Alcatel/OXE."""
        df_oxe = _cached_get_central_telefonica_df()
        self.assertIsInstance(df_oxe, pd.DataFrame)

    def test_unidades_and_ramais_cache_functions(self):
        """Valida retorno das funções cacheadas de Unidades do MPMS e Ramais."""
        df_unidades = _cached_get_unidades_df()
        self.assertIsInstance(df_unidades, pd.DataFrame)

        df_ramais = _cached_get_ramais_df()
        self.assertIsInstance(df_ramais, pd.DataFrame)

    def test_mapas_cache_image_functions(self):
        """Valida que helpers cacheados de imagem tratam caminhos inexistentes com fallback."""
        fake_path = Path("/caminho/falso/planta_inexistente.png")
        dims = get_image_dimensions(fake_path)
        self.assertEqual(dims, (1000, 1000))

        b64 = get_image_base64(fake_path)
        self.assertEqual(b64, "")

    def test_cache_invalidation_lifecycle(self):
        """Valida que st.cache_data.clear() reseta e reexecuta com segurança sem lançar exceções."""
        # 1. Carrega inicial (alimenta o cache)
        ous_1 = _cached_get_ad_ous()
        devs_1 = _cached_get_sccm_devices_df()

        # 2. Invalidação manual simulando fim de sincronização
        st.cache_data.clear()

        # 3. Nova leitura logo após o clear
        ous_2 = _cached_get_ad_ous()
        devs_2 = _cached_get_sccm_devices_df()

        self.assertIsInstance(ous_2, pd.DataFrame)
        self.assertIsInstance(devs_2, pd.DataFrame)
        self.assertEqual(len(ous_1), len(ous_2))
        self.assertEqual(len(devs_1), len(devs_2))


if __name__ == "__main__":
    unittest.main()
