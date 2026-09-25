# -*- coding: utf-8 -*-
"""
Testes unitários dos serviços de negócio (src.services.*).
Valida sccm_service, formatação de queries CIM/WMI, importação de JSON do inventário
e member_matcher.
"""

import os
import sys
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.services.sccm_service import import_sccm_inventory_json
from src.services.member_matcher import resolve_bancada_member

class TestServices(unittest.TestCase):
    def test_import_sccm_inventory_json_valid_file(self):
        """Valida a importação de inventário gerado pelo disparador Windows (bancada://)."""
        temp_dir = tempfile.TemporaryDirectory()
        json_path = Path(temp_dir.name) / "sccm_inventory.json"
        
        sample_data = {
            "generated_at": "2026-09-25T12:00:00",
            "devices": [
                {
                    "ResourceID": "1001",
                    "Name": "NOTEBOOK-STI01",
                    "LastLogonUserName": "paulo_admin",
                    "IPAddresses": "192.168.10.50",
                    "MACAddresses": "AA:BB:CC:DD:EE:FF",
                    "OperatingSystemNameandVersion": "Microsoft Windows 11",
                    "Build": "22631",
                    "ClientVersion": "5.0.9000",
                    "Active": 1,
                    "ADSiteName": "PGJ",
                    "DistinguishedName": "CN=NOTEBOOK-STI01,OU=STI,DC=mpe,DC=local"
                }
            ],
            "users": [
                {
                    "ResourceID": "2001",
                    "UserName": "paulo_admin",
                    "FullUserName": "Paulo Henrique Gonçalves Rezende",
                    "WindowsNTDomain": "MPE",
                    "DistinguishedName": "CN=paulo_admin,OU=Users,DC=mpe,DC=local"
                }
            ],
            "collections": [
                {
                    "CollectionID": "SMS00001",
                    "Name": "All Systems",
                    "CollectionType": "2",
                    "MemberCount": 1500,
                    "Comment": "Todas as estações"
                }
            ]
        }
        
        json_path.write_text(json.dumps(sample_data), encoding="utf-8")

        with patch("src.services.sccm_service.save_sccm_devices", return_value=1) as m_dev, \
             patch("src.services.sccm_service.save_sccm_users", return_value=1) as m_usr, \
             patch("src.services.sccm_service.save_sccm_collections", return_value=1) as m_col:
            
            res = import_sccm_inventory_json(str(json_path))
            self.assertEqual(res["devices"], 1)
            self.assertEqual(res["users"], 1)
            self.assertEqual(res["collections"], 1)
            m_dev.assert_called_once()
            m_usr.assert_called_once()
            m_col.assert_called_once()

        temp_dir.cleanup()

    def test_import_sccm_inventory_json_missing_file(self):
        """Valida que arquivo ausente não causa exceção e retorna contadores zerados."""
        res = import_sccm_inventory_json("/caminho/completamente/inexistente/sccm.json")
        self.assertEqual(res, {"devices": 0, "users": 0, "collections": 0})

    def test_member_matcher_names(self):
        """Valida o comparador/normalizador de nomes de técnicos e servidores da bancada."""
        m1 = resolve_bancada_member("Paulo Henrique Gonçalves Rezende")
        self.assertIsNotNone(m1)
        self.assertEqual(m1["primeiro_nome"], "Paulo")

        m2 = resolve_bancada_member("reginaldo bandeira")
        self.assertIsNotNone(m2)
        self.assertEqual(m2["primeiro_nome"], "Reginaldo")

        m3 = resolve_bancada_member("Carlos Eduardo Desconhecido")
        self.assertIsNone(m3)

if __name__ == "__main__":
    unittest.main()
