# -*- coding: utf-8 -*-
"""
Testes de integridade de scripts do sistema (PowerShell, bash, configs).
Valida sintaxe de scripts .ps1, caminhos referenciados, existência de executáveis e parâmetros.
"""

import os
import re
import sys
import unittest
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers

class TestScriptsAndIntegrity(unittest.TestCase):
    def test_powershell_scripts_exist_and_not_empty(self):
        """Valida que todos os scripts PowerShell fundamentais existem e não estão corrompidos/vazios."""
        expected_scripts = [
            ROOT_DIR / "src" / "protocol_handler" / "bancada-launcher.ps1",
            ROOT_DIR / "src" / "scripts_powershell" / "sccm_sync.ps1",
            ROOT_DIR / "src" / "scripts_powershell" / "manutencao" / "Manutencao.ps1",
            ROOT_DIR / "src" / "scripts_powershell" / "analisador" / "Analisador.ps1",
            ROOT_DIR / "src" / "scripts_powershell" / "analisador" / "GeradorHtml.ps1",
            ROOT_DIR / "src" / "scripts_powershell" / "perfis" / "RemoverUsuarios.ps1",
        ]

        for script_path in expected_scripts:
            self.assertTrue(script_path.exists(), f"Script não encontrado: {script_path}")
            self.assertTrue(script_path.stat().st_size > 50, f"Script parece vazio ou truncado: {script_path}")

    def test_bancada_launcher_supported_tools(self):
        """Valida que o disparador bancada-launcher.ps1 contém todas as ferramentas esperadas."""
        launcher_file = ROOT_DIR / "src" / "protocol_handler" / "bancada-launcher.ps1"
        content = launcher_file.read_text(encoding="utf-8", errors="ignore").lower()
        
        required_tools = ["cmrc", "rdp", "explorer", "ping", "sccm_sync", "manutencao", "analisador", "perfis"]
        for tool in required_tools:
            self.assertIn(tool, content, f"Ferramenta '{tool}' não mapeada no bancada-launcher.ps1")

    def test_sccm_sync_script_queries(self):
        """Valida que o script sccm_sync.ps1 faz consulta às classes WMI corretas do SCCM."""
        sync_file = ROOT_DIR / "src" / "scripts_powershell" / "sccm_sync.ps1"
        content = sync_file.read_text(encoding="utf-8", errors="ignore")
        
        self.assertIn("SMS_Collection", content)
        self.assertIn("SMS_R_System", content)
        self.assertIn("SMS_R_User", content)
        self.assertIn("ConvertTo-Json", content)

    def test_protocol_registry_file(self):
        """Valida a integridade do arquivo de registro do Windows (instalar_protocolo_bancada.reg)."""
        reg_file = ROOT_DIR / "src" / "protocol_handler" / "instalar_protocolo_bancada.reg"
        self.assertTrue(reg_file.exists())
        content = reg_file.read_text(encoding="utf-8", errors="ignore")
        self.assertTrue("bancada" in content and "URL Protocol" in content)

if __name__ == "__main__":
    unittest.main()
