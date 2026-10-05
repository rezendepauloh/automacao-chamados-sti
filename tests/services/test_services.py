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

    def test_cron_daemon_should_run_and_resilient_loop(self):
        """Valida que o BancadaCronDaemon avalia agendamentos corretamente e não quebra com dicts nativos."""
        from src.services.cron_scheduler import BancadaCronDaemon
        from datetime import datetime

        daemon = BancadaCronDaemon()
        now = datetime(2026, 9, 30, 12, 0, 0) # Quarta-feira

        # Tarefa ativa de horário fixo no horário atual
        task_active = {
            "task_id": "test_task_1",
            "nome": "Tarefa Teste",
            "ativo": 1,
            "tipo_agendamento": "horario_fixo",
            "horario_fixo": "12:00",
            "apenas_dias_uteis": 1,
            "ultima_execucao": None
        }
        self.assertTrue(daemon._should_run(task_active, now))

        # Tarefa inativa
        task_inactive = dict(task_active, ativo=0)
        self.assertFalse(daemon._should_run(task_inactive, now))

        # Tarefa de fim de semana agendada apenas para dias úteis
        sunday = datetime(2026, 10, 4, 12, 0, 0) # Domingo
        self.assertFalse(daemon._should_run(task_active, sunday))

    def test_cron_tasks_ad_and_sccm_registered(self):
        """Valida que as tarefas de Active Directory e SCCM estão registradas nos cron jobs e executáveis."""
        from src.database.cron_db import get_cron_schedules, setup_cron_tables
        from src.services.cron_scheduler import execute_task_by_id

        setup_cron_tables()
        df = get_cron_schedules()
        task_ids = [r["task_id"] for _, r in df.iterrows()]
        self.assertIn("sync_ad_catalog", task_ids)
        self.assertIn("sync_sccm", task_ids)

        # Valida que o despachador execute_task_by_id reconhece sync_sccm e sync_ad_catalog
        with patch("src.syncs.sync_ad_catalog.run_ad_sync", return_value={"success": True, "total_ous": 5, "total_users": 10, "total_computers": 2, "total_groups": 3}):
            msg_ad = execute_task_by_id("sync_ad_catalog")
            self.assertIn("Active Directory sincronizado com sucesso", msg_ad)

        with patch("src.services.sccm_service.sync_all_sccm", return_value={"devices": 15, "users": 10, "collections": 5}):
            msg_sccm = execute_task_by_id("sync_sccm")
            self.assertIn("Inventário SCCM sincronizado com sucesso", msg_sccm)


if __name__ == "__main__":
    unittest.main()

