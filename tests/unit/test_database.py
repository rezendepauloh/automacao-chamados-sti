# -*- coding: utf-8 -*-
"""
Testes unitários de persistência e banco de dados SQLite (src.database.*).
Valida criação de tabelas, inserção, deduplicação, updates e consultas do SCCM, Plantões e Notificações
em banco isolado em memória ou temporário.
"""

import os
import sys
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.database.sccm_db import (
    setup_sccm_tables,
    save_sccm_devices,
    save_sccm_users,
    save_sccm_collections,
    get_sccm_devices_df,
    get_sccm_collections_df,
    get_sccm_users_df,
    get_device_by_user
)
from src.database.tickets_db import (
    setup_database,
    save_tickets_to_db,
    update_ticket_device_info,
    load_data
)
from src.database.plantoes_db import (
    setup_plantoes_tables,
    save_plantoes_matutino,
    save_plantoes_semanal,
    get_plantoes_matutino,
    get_plantoes_semanal
)

class TestDatabaseModule(unittest.TestCase):
    def setUp(self):
        # Cria banco de dados SQLite temporário e isolado
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_chamados.db"
        
        # Patch na conexão global para abrir conexão no banco temporário
        self.patchers = [
            patch("src.database.connection.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.sccm_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.tickets_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
            patch("src.database.plantoes_db.get_connection", side_effect=lambda: sqlite3.connect(self.db_path)),
        ]
        for p in self.patchers:
            p.start()

    def tearDown(self):
        for p in self.patchers:
            p.stop()
        self.temp_dir.cleanup()

    def test_sccm_tables_and_crud(self):
        """Valida criação de tabelas e inserção/consulta de dispositivos do SCCM."""
        setup_sccm_tables()

        # Inserção de estações de trabalho
        mock_devices = [
            {
                "ResourceID": "16777220",
                "Name": "DESKTOP-BANCADA01",
                "LastLogonUserName": "paulo_admin",
                "IPAddresses": ["192.168.1.100"],
                "MACAddresses": ["00:11:22:33:44:55"],
                "Manufacturer": "Dell Inc.",
                "Model": "OptiPlex 7090",
                "OperatingSystemNameandVersion": "Microsoft Windows 11 Enterprise",
                "Build": "22631",
                "ClientVersion": "5.00.9106.1000",
                "Active": 1,
                "ADSiteName": "Default-First-Site-Name",
                "DistinguishedName": "CN=DESKTOP-BANCADA01,OU=Estacoes,DC=mpe,DC=local"
            }
        ]
        
        saved_count = save_sccm_devices(mock_devices)
        self.assertEqual(saved_count, 1)

        # Consulta com filtro de busca
        df_dev = get_sccm_devices_df(search_text="DESKTOP-BANCADA01")
        self.assertFalse(df_dev.empty)
        self.assertEqual(df_dev['name'].values[0], "DESKTOP-BANCADA01")

    def test_sccm_collections_and_users(self):
        """Valida inserção e listagem de coleções e usuários do SCCM."""
        setup_sccm_tables()

        mock_collections = [
            {
                "CollectionID": "SMS00001",
                "Name": "All Systems",
                "CollectionType": 2,
                "MemberCount": 2840,
                "Comment": "Coleção raiz de computadores"
            }
        ]
        save_sccm_collections(mock_collections)
        df_cols = get_sccm_collections_df()
        self.assertFalse(df_cols.empty)

        mock_users = [
            {
                "ResourceID": "20971521",
                "UserName": "reginaldo_silva",
                "FullUserName": "Reginaldo da Silva Bandeira",
                "WindowsNTDomain": "MPE",
                "DistinguishedName": "CN=Reginaldo Bandeira,OU=STI,DC=mpe,DC=local"
            }
        ]
        user_saved = save_sccm_users(mock_users)
        self.assertEqual(user_saved, 1)

    def test_plantoes_crud(self):
        """Valida gravação e consulta de escalas de plantão matutino e semanal."""
        setup_plantoes_tables()

        rec_mat = [{
            "ano": 2026,
            "data_iso": "2026-03-15",
            "dia_semana": "Domingo",
            "servidor": "Paulo Henrique Gonçalves Rezende",
            "telefone": "+55 67 99999-9999"
        }]
        save_plantoes_matutino(rec_mat)
        df_mat = get_plantoes_matutino(2026)
        self.assertFalse(df_mat.empty)
        self.assertIn("Paulo Henrique Gonçalves Rezende", df_mat['servidor'].values)

        # Plantão Semanal
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

    def test_parse_data_matutino_string(self):
        """Testa a conversão de datas textuais da planilha de plantão matutino."""
        from src.scrapers.plantoes_scraper import parse_data_matutino_string, parse_simp_periodo
        dt_iso, dia_sem = parse_data_matutino_string("Segunda-feira - 12 de Janeiro", 2026)
        self.assertEqual(dt_iso, "2026-01-12")
        self.assertEqual(dia_sem, "Segunda-feira")

        periodo = "06/07/2026 19:01 a 13/07/2026 11:59"
        dt_ini, dt_fim = parse_simp_periodo(periodo, 2026)
        self.assertEqual(dt_ini, "2026-07-06 19:01:00")
        self.assertEqual(dt_fim, "2026-07-13 11:59:00")

    def test_sccm_get_device_by_user(self):
        """Valida a busca exata e enriquecimento por usuário no cache do SCCM."""
        setup_sccm_tables()
        mock_devices = [
            {
                "ResourceID": "99901",
                "Name": "MPE-80370",
                "LastLogonUserName": "nairaoliveira",
                "IPAddresses": ["10.111.131.15", "fe80::1"],
                "ClientActiveStatus": 1
            }
        ]
        save_sccm_devices(mock_devices)

        dev = get_device_by_user("nairaoliveira")
        self.assertIsNotNone(dev)
        self.assertEqual(dev["hostname"], "MPE-80370")
        self.assertEqual(dev["ip"], "10.111.131.15")

        # Testa com domínio e maiúsculas
        dev_domain = get_device_by_user("MPE\\NAIRAOLIVEIRA")
        self.assertIsNotNone(dev_domain)
        self.assertEqual(dev_domain["hostname"], "MPE-80370")

        # Testa usuário inexistente
        self.assertIsNone(get_device_by_user("usuario_nao_existe_xyz"))

    def test_tickets_save_nan_sanitization_and_enrichment(self):
        """Garante que valores NaN não sejam gravados e que ocorra enriquecimento via SCCM."""
        import pandas as pd

        setup_database()
        setup_sccm_tables()

        # Cadastra estação de teste para enriquecimento
        save_sccm_devices([{
            "ResourceID": "99902",
            "Name": "MPE-70001",
            "LastLogonUserName": "joaosilva",
            "IPAddresses": ["10.111.64.99"],
            "ClientActiveStatus": 1
        }])

        df_chamados = pd.DataFrame([{
            "Chamado#": "117387",
            "Data Criação": "2026-09-25 10:00:00",
            "Título": float('nan'),
            "Cidade - Prédio": "Campo Grande - Sede",
            "Unidade": "STI",
            "Localidade física": "Campo Grande - Sede",
            "Nome do Usuário": "João da Silva",
            "ID do Cliente": "joaosilva",
            "Descrição": "Problema no monitor",
            "TAG": "MONITOR",
            "IP_Origem": float('nan'),
            "Hostname": "nan",
            "Base": "CitSmart",
            "Link": None,
            "Comentários": "[]"
        }])

        save_tickets_to_db(df_chamados)

        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT id, ip_origem, hostname, titulo FROM chamados WHERE id = '117387'")
        row = c.fetchone()
        conn.close()

        self.assertIsNotNone(row)
        cid, ip, host, titulo = row

        # IP e Hostname devem ter sido enriquecidos via cache SCCM
        self.assertEqual(ip, "10.111.64.99")
        self.assertEqual(host, "MPE-70001")
        # Título não deve ser 'nan' literal
        self.assertEqual(titulo, "")

        # Testa update_ticket_device_info
        update_ticket_device_info("117387", "10.111.64.100", "MPE-70002")
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT ip_origem, hostname FROM chamados WHERE id = '117387'")
        row_up = c.fetchone()
        conn.close()
        self.assertEqual(row_up[0], "10.111.64.100")
        self.assertEqual(row_up[1], "MPE-70002")

if __name__ == "__main__":
    unittest.main()
