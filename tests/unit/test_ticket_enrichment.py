# -*- coding: utf-8 -*-
"""
Testes unitários para enriquecimento de chamados, resolução de sAMAccountName,
descoberta em cascata de IP/Hostname e formatação brasileira de datas.
"""

import os
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.database.ad_db import save_ad_cache, setup_ad_tables
from src.database.sccm_db import save_sccm_devices, save_sccm_users, setup_sccm_tables
from src.database.tickets_db import (
    get_ticket_by_id,
    save_tickets_to_db,
    setup_database,
    update_ticket_client_id,
    update_ticket_device_info,
)
from src.services.ticket_enrichment import (
    extract_network_info_from_text,
    format_br_datetime,
    resolve_ticket_device_and_user,
    resolve_user_sam_account_name,
    sanitize_network_value,
)


class TestTicketEnrichment(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_enrichment.db"
        self.patcher_env = patch.dict(os.environ, {"DB_TYPE": "sqlite"})
        self.patcher_db = patch("src.database.connection.DB_PATH", self.db_path)
        self.patcher_db_type = patch("src.database.connection.DB_TYPE", "sqlite")
        self.patcher_ad_type = patch("src.database.ad_db.DB_TYPE", "sqlite")
        self.patcher_sccm_type = patch("src.database.sccm_db.DB_TYPE", "sqlite")
        self.patcher_tkt_type = patch("src.database.tickets_db.DB_TYPE", "sqlite")
        self.patcher_tkt_path = patch("src.database.tickets_db.DB_PATH", self.db_path)

        self.patcher_env.start()
        self.patcher_db.start()
        self.patcher_db_type.start()
        self.patcher_ad_type.start()
        self.patcher_sccm_type.start()
        self.patcher_tkt_type.start()
        self.patcher_tkt_path.start()

        setup_database()
        setup_ad_tables()
        setup_sccm_tables()

    def tearDown(self):
        self.patcher_tkt_path.stop()
        self.patcher_tkt_type.stop()
        self.patcher_sccm_type.stop()
        self.patcher_ad_type.stop()
        self.patcher_db_type.stop()
        self.patcher_db.stop()
        self.patcher_env.stop()
        self.temp_dir.cleanup()

    def test_format_br_datetime_variants(self):
        """Valida conversão de datas ISO, BR e inválidas para dd/mm/aaaa hh:mm:ss."""
        self.assertEqual(format_br_datetime("2026-11-19 09:30:15"), "19/11/2026 09:30:15")
        self.assertEqual(format_br_datetime("2026-11-19T09:30:15.456Z"), "19/11/2026 09:30:15")
        # Garante dayfirst=True para datas brasileiras (05/11 = 5 de novembro, nunca 11 de maio)
        self.assertEqual(format_br_datetime("05/11/2026 14:20"), "05/11/2026 14:20:00")
        self.assertEqual(format_br_datetime("05/11/2026 14:20:45"), "05/11/2026 14:20:45")
        self.assertEqual(format_br_datetime(None), "N/A")
        self.assertEqual(format_br_datetime(""), "N/A")
        self.assertEqual(format_br_datetime("nan"), "N/A")

    def test_extract_network_info_from_text(self):
        """Valida extração de IP, Hostname corporativo e Patrimônio evitando falsos positivos."""
        txt1 = "Computador MPE-58063 está sem rede. IP de origem: 10.111.64.55. Build 10.0.22631.3880."
        res1 = extract_network_info_from_text(txt1)
        self.assertEqual(res1["ip"], "10.111.64.55")
        self.assertEqual(res1["hostname"], "MPE-58063")

        # Não deve confundir 'Campo Grande - PGJ' com hostname
        txt2 = "Localidade: Campo Grande - PGJ. Patrimônio: 43188. IP 172.20.5.151 travando."
        res2 = extract_network_info_from_text(txt2)
        self.assertEqual(res2["ip"], "172.20.5.151")
        self.assertEqual(res2["hostname"], "")
        self.assertEqual(res2["patrimonio"], "43188")

        # Deve ignorar 127.0.0.1 e versões
        txt3 = "Erro na versão v10.0.190.1 ou localhost 127.0.0.1 na estação PGJ-NT-0042"
        res3 = extract_network_info_from_text(txt3)
        self.assertEqual(res3["ip"], "")
        self.assertEqual(res3["hostname"], "PGJ-NT-0042")

    def test_sanitize_network_value(self):
        """Descarta sentinelas de erro/timeout para garantir fallback limpo na UI."""
        for invalid in ("", None, "nan", "N/A", "Timeout", "Erro", "Acesso Negado"):
            self.assertEqual(sanitize_network_value(invalid), "")
        self.assertEqual(sanitize_network_value("10.111.1.10"), "10.111.1.10")

    def test_resolve_sam_account_name_from_ad_and_sccm(self):
        """Resolve sAMAccountName a partir do nome completo via ad_cache_users e sccm_cache_users."""
        save_ad_cache(
            ous=[],
            users=[{
                "sam_account_name": "bruna.santos",
                "display_name": "Bruna Luciano de Freitas Santos",
                "mail": "bruna.santos@mpms.mp.br",
                "dn": "CN=Bruna Luciano de Freitas Santos,OU=Usuarios,DC=mpms,DC=local",
                "is_active": True,
            }],
            groups=[],
            memberships=[],
            computers=[],
        )

        # Casing diferente ("De" vs "de")
        res = resolve_user_sam_account_name(
            usuario="Bruna Luciano De Freitas Santos",
            id_cliente="",
            allow_ldap=False,
        )
        self.assertEqual(res["sam_account_name"], "bruna.santos")

        # Normaliza prefixo de domínio / e-mail quando id_cliente já vem preenchido
        res_dom = resolve_user_sam_account_name(usuario="", id_cliente="MPE\\bruna.santos", allow_ldap=False)
        self.assertEqual(res_dom["sam_account_name"], "bruna.santos")

        # Fallback via sccm_cache_users quando não existir no AD cache
        save_sccm_users([{
            "ResourceID": "88001",
            "UserName": "carlos.almeida",
            "FullUserName": "Carlos Eduardo Almeida",
            "WindowsNTDomain": "MPE",
        }])
        res_sccm = resolve_user_sam_account_name(
            usuario="Carlos Eduardo Almeida",
            id_cliente="",
            allow_ldap=False,
        )
        self.assertEqual(res_sccm["sam_account_name"], "carlos.almeida")

    def test_cascade_resolution_citsmart_125729_style(self):
        """Simula o caso #125729: id_cliente/IP/Hostname vazios enriquecidos via AD + SCCM."""
        save_ad_cache(
            ous=[],
            users=[{
                "sam_account_name": "bruna.santos",
                "display_name": "Bruna Luciano de Freitas Santos",
                "dn": "CN=Bruna,DC=mpms,DC=local",
                "is_active": True,
            }],
            groups=[],
            memberships=[],
            computers=[],
        )
        save_sccm_devices([{
            "ResourceID": "50123",
            "Name": "PGJ-55123",
            "LastLogonUserName": "bruna.santos",
            "IPAddresses": ["fe80::1", "10.111.20.45"],
            "ClientActiveStatus": 1,
        }])

        res = resolve_ticket_device_and_user(
            {
                "id": "125729",
                "usuario": "Bruna Luciano De Freitas Santos",
                "id_cliente": "",
                "titulo": "Meu computador está reiniciando sozinho do nada",
                "descricao": "Reiniciando durante o uso.",
                "ip_origem": "",
                "hostname": "",
            },
            allow_network_fallbacks=False,
        )
        self.assertEqual(res["sam_account_name"], "bruna.santos")
        self.assertEqual(res["ip_origem"], "10.111.20.45")
        self.assertEqual(res["hostname"], "PGJ-55123")

    def test_cascade_cross_lookups_and_fallbacks(self):
        """Valida cruzamento por texto, patrimônio, AD computers, histórico e fallback N/A."""
        save_sccm_devices([{
            "ResourceID": "60001",
            "Name": "MPE-77440",
            "LastLogonUserName": "outro.usuario",
            "IPAddresses": ["10.111.88.12"],
            "ClientActiveStatus": 1,
        }])

        # 1. Descobre Hostname nos comentários e busca IP no SCCM
        res_comm = resolve_ticket_device_and_user(
            {"id": "1", "usuario": "Sem Match", "descricao": "Monitor piscando", "ip_origem": "", "hostname": ""},
            comments=[{"data": "19/11/2026 10:00", "autor": "Tecnico", "texto": "Verificado na máquina MPE-77440."}],
            allow_network_fallbacks=False,
        )
        self.assertEqual(res_comm["hostname"], "MPE-77440")
        self.assertEqual(res_comm["ip_origem"], "10.111.88.12")

        # 2. Descobre por Patrimônio na descrição
        res_pat = resolve_ticket_device_and_user(
            {"id": "2", "usuario": "Sem Match", "descricao": "Equipamento patrimônio 77440 com defeito"},
            allow_network_fallbacks=False,
        )
        self.assertEqual(res_pat["hostname"], "MPE-77440")
        self.assertEqual(res_pat["ip_origem"], "10.111.88.12")

        # 3. Fallback via ad_cache_computers (managed_by)
        save_ad_cache(
            ous=[],
            users=[{
                "sam_account_name": "maria.lima",
                "display_name": "Maria Clara Lima",
                "dn": "CN=Maria Clara Lima,DC=mpms,DC=local",
                "is_active": True,
            }],
            groups=[],
            memberships=[],
            computers=[{
                "name": "PGJ-90123",
                "dns_hostname": "PGJ-90123.mpms.local",
                "managed_by": "CN=Maria Clara Lima,DC=mpms,DC=local",
                "is_active": True,
            }],
        )
        res_ad_comp = resolve_ticket_device_and_user(
            {"id": "3", "usuario": "Maria Clara Lima", "descricao": "Sem rede"},
            allow_network_fallbacks=False,
        )
        self.assertEqual(res_ad_comp["sam_account_name"], "maria.lima")
        self.assertEqual(res_ad_comp["hostname"], "PGJ-90123")

        # 4. Fallback via histórico de chamados anteriores do mesmo usuário
        df_hist = pd.DataFrame([{
            "Chamado#": "90001",
            "Data Criação": "2026-11-10 08:00:00",
            "Título": "Chamado antigo",
            "Cidade - Prédio": "Campo Grande",
            "Unidade": "STI",
            "Localidade física": "Campo Grande",
            "Nome do Usuário": "Joana Prado",
            "ID do Cliente": "joana.prado",
            "Descrição": "Troca de mouse",
            "TAG": "SUPORTE",
            "IP_Origem": "10.111.99.77",
            "Hostname": "MPE-99077",
            "Base": "OTRS",
            "Link": "https://central.mpms.mp.br/otrs/90001",
            "Comentários": "[]",
        }])
        save_tickets_to_db(df_hist)
        res_hist = resolve_ticket_device_and_user(
            {"id": "90002", "usuario": "Joana Prado", "id_cliente": "", "descricao": "Novo chamado"},
            allow_network_fallbacks=False,
        )
        self.assertEqual(res_hist["ip_origem"], "10.111.99.77")
        self.assertEqual(res_hist["hostname"], "MPE-99077")

        # 5. Último fallback retorna vazio (para exibir "N/A") mesmo com sentinelas inválidos
        res_none = resolve_ticket_device_and_user(
            {"id": "99999", "usuario": "Desconhecido Total", "ip_origem": "Timeout", "hostname": "Erro"},
            allow_network_fallbacks=False,
        )
        self.assertEqual(res_none["sam_account_name"], "")
        self.assertEqual(res_none["ip_origem"], "")
        self.assertEqual(res_none["hostname"], "")

    def test_passive_enrichment_preserves_data_atualizacao(self):
        """Atualizações passivas do modal não devem sobrescrever data_atualizacao."""
        df = pd.DataFrame([{
            "Chamado#": "125729",
            "Data Criação": "2026-11-18 10:15:00",
            "Título": "Teste Preservação",
            "Cidade - Prédio": "Campo Grande",
            "Unidade": "PGJ",
            "Localidade física": "Campo Grande - PGJ",
            "Nome do Usuário": "Usuario Sem Cache",
            "ID do Cliente": "",
            "Descrição": "Sem info",
            "TAG": "MONITOR",
            "IP_Origem": "",
            "Hostname": "",
            "Base": "CitSmart",
            "Link": "",
            "Comentários": "[]",
        }])
        save_tickets_to_db(df)
        before = get_ticket_by_id("125729")
        orig_upd = before["data_atualizacao"]

        update_ticket_client_id("125729", "usuario.cache", touch_updated_at=False)
        update_ticket_device_info("125729", "10.111.1.2", "PGJ-11122", touch_updated_at=False)

        after = get_ticket_by_id("125729")
        self.assertEqual(after["id_cliente"], "usuario.cache")
        self.assertEqual(after["ip_origem"], "10.111.1.2")
        self.assertEqual(after["hostname"], "PGJ-11122")
        self.assertEqual(after["data_atualizacao"], orig_upd)


if __name__ == "__main__":
    unittest.main()
