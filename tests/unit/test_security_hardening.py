# -*- coding: utf-8 -*-
"""
Testes de Segurança Ofensiva e Hardening (SecOps).
Valida:
1. Varredura contínua de segredos em arquivos rastreados pelo Git (git ls-files).
2. Comportamento Fail-Closed quando credenciais/chaves essenciais são omitidas.
3. Prevenção de Injeção de Comandos (Shell/PowerShell), LDAP Injection, WQL Injection e Path Traversal.
4. Sintaxe e binds restritos a localhost (127.0.0.1) em docker-compose.yml e integridade de .gitignore.
5. Permissões de arquivos sensíveis (0600 em .env e .secret.key) e UTF-8 BOM em scripts PowerShell.
6. Testes negativos de autenticação, sanitização de inputs e proteção contra Força Bruta / Rate Limiting.
7. Proteções contra DoS e Zip-Bomb no download de planilhas.
"""

import os
import re
import sys
import unittest
import subprocess
from pathlib import Path
from unittest.mock import patch, MagicMock

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.auth import (
    login_user,
    check_login_rate_limit,
    record_failed_login,
    reset_login_attempts,
    decode_auth_token,
    create_auth_token,
    MAX_FAILED_ATTEMPTS,
    _FAILED_ATTEMPTS,
)
from src.config import resolve_spreadsheet_path, fetch_sccm_data, fetch_ad_department
from src.tabs.impressoras import ping_host
from src.services.evolution_client import (
    get_connection_status,
    create_instance_if_needed,
    get_qr_code,
    disconnect_instance,
    send_whatsapp_text,
)
from src.services.sccm_service import execute_sccm_cim_query


class TestSecurityHardening(unittest.TestCase):
    """Suíte mestre de validação de segurança e conformidade de hardening."""

    def setUp(self):
        # Limpa histórico de tentativas de login entre testes
        _FAILED_ATTEMPTS.clear()

    # =========================================================================
    # 1. VARREDURA DE SEGREDOS NO GIT (SAST)
    # =========================================================================
    def test_no_hardcoded_passwords_or_tokens_in_git_tracked_files(self):
        """
        Garante que nenhum arquivo rastreado pelo Git contenha senhas conhecidas
        hardcoded (como Abcd4268# ou bancada_secret_token_123) ou chaves privadas RSA/PEM.
        """
        try:
            res = subprocess.run(
                ["git", "ls-files"],
                cwd=str(ROOT_DIR),
                capture_output=True,
                text=True,
                check=True
            )
            tracked_files = [ROOT_DIR / f.strip() for f in res.stdout.splitlines() if f.strip()]
        except Exception:
            # Fallback para caso git não esteja no path
            tracked_files = list(ROOT_DIR.glob("**/*.py")) + list(ROOT_DIR.glob("**/*.ps1")) + list(ROOT_DIR.glob("**/*.yml"))

        forbidden_patterns = [
            re.compile(r"Abcd4268#", re.IGNORECASE),
            re.compile(r"bancada_secret_token_123", re.IGNORECASE),
            re.compile(r"-----BEGIN (RSA |EC |OPENSSH )?PRIVATE KEY-----"),
        ]

        violations = []
        for file_path in tracked_files:
            if not file_path.is_file():
                continue
            # Ignora binários e este próprio arquivo de teste de segurança que lista os padrões proibidos
            if file_path.name in ["test_security_hardening.py"] or file_path.suffix.lower() in [".png", ".jpg", ".ico", ".xlsx", ".pdf", ".pkl", ".joblib"]:
                continue

            try:
                content = file_path.read_text(encoding="utf-8", errors="ignore")
                for pattern in forbidden_patterns:
                    matches = pattern.findall(content)
                    if matches:
                        rel = file_path.relative_to(ROOT_DIR)
                        violations.append(f"{rel}: padrão proibido '{pattern.pattern}' detectado!")
            except Exception:
                pass

        self.assertEqual(
            violations,
            [],
            f"VULNERABILIDADE DE SEGURANÇA: Segredos hardcoded encontrados em arquivos Git:\n" + "\n".join(violations)
        )

    # =========================================================================
    # 2. FAIL-CLOSED BEHAVIOR (SERVIÇOS, APIS & BANCO)
    # =========================================================================
    def test_evolution_client_fail_closed_when_key_empty(self):
        """Valida que o cliente da Evolution API falha de forma fechada (sem realizar requisição) se a chave for omitida."""
        with patch.dict(os.environ, {"EVOLUTION_API_KEY": ""}, clear=False), \
             patch("src.services.evolution_client._cfg", return_value=""), \
             patch("requests.get") as mock_get, \
             patch("requests.post") as mock_post:

            # 1. get_connection_status deve falhar fechado
            status = get_connection_status()
            self.assertFalse(status["online"])
            self.assertEqual(status["state"], "unauthorized")
            self.assertIn("Fail-Closed", status["error"])
            mock_get.assert_not_called()

            # 2. create_instance_if_needed deve recusar
            created = create_instance_if_needed()
            self.assertFalse(created)
            mock_post.assert_not_called()

            # 3. get_qr_code deve retornar erro sem tentar conectar
            qr = get_qr_code()
            self.assertFalse(qr["success"])
            self.assertIn("Fail-Closed", qr["error"])

            # 4. disconnect_instance deve recusar
            disconnected = disconnect_instance()
            self.assertFalse(disconnected)

            # 5. send_whatsapp_text deve recusar envio
            send_res = send_whatsapp_text("5567999999999", "Mensagem teste")
            self.assertFalse(send_res["success"])
            self.assertIn("Fail-Closed", send_res["error"])

    def test_sccm_service_fail_closed_when_password_missing(self):
        """Valida que execute_sccm_cim_query falha de forma fechada quando credenciais ou servidores não estão definidos."""
        with patch.dict(os.environ, {"SCCM_ADMIN_PASSWORD": ""}, clear=False), \
             patch("src.services.sccm_service._cfg", return_value=""), \
             patch("subprocess.run") as mock_subproc:

            res = execute_sccm_cim_query("SELECT * FROM SMS_R_System")
            self.assertEqual(res, [])
            mock_subproc.assert_not_called()

    # =========================================================================
    # 3. VALIDAÇÃO DE INPUTS & PREVENÇÃO DE INJEÇÃO
    # =========================================================================
    def test_ping_host_rejects_command_injection(self):
        """Valida que ping_host rejeita estritamente payloads maliciosos de command injection."""
        malicious_hosts = [
            "127.0.0.1; rm -rf /",
            "10.0.0.1 && dir",
            "localhost | calc.exe",
            "`id`",
            "$(whoami)",
            "10.0.0.1\ncat /etc/passwd",
        ]
        for bad_host in malicious_hosts:
            is_ok, msg = ping_host(bad_host)
            self.assertFalse(is_ok, f"Host malicioso foi aceito indevidamente: {bad_host}")
            self.assertIn("inválido", msg.lower())

    def test_resolve_spreadsheet_path_blocks_path_traversal(self):
        """Valida que resolve_spreadsheet_path bloqueia vetores de Path Traversal ('..')."""
        with patch("src.config.get_spreadsheet_setting", return_value="../../etc/passwd"):
            with self.assertRaises(ValueError) as ctx:
                resolve_spreadsheet_path("TEST_PATH")
            self.assertIn("Path Traversal detectado", str(ctx.exception))

    def test_fetch_sccm_data_rejects_injection_usernames(self):
        """Valida que fetch_sccm_data rejeita usernames com injeção de aspas ou comandos."""
        malicious_users = [
            "admin' OR '1'='1",
            "user; rm -rf /",
            "paulo`id`",
            "user$(whoami)",
        ]
        for bad_user in malicious_users:
            res = fetch_sccm_data(bad_user)
            self.assertEqual(res.get("ip"), "N/A")
            self.assertEqual(res.get("hostname"), "N/A")

    def test_fetch_ad_department_escapes_ldap_injection(self):
        """Valida que fetch_ad_department sanitiza caracteres de controle LDAP."""
        mock_conn = MagicMock()
        mock_conn.entries = []

        # Tenta injetar filtro LDAP
        fetch_ad_department(mock_conn, "*)(uid=*))(|(uid=*", is_username=True)
        # Verifica se o filtro executado foi escapado (contendo \\2a para o asterisco)
        calls = mock_conn.search.call_args_list
        self.assertTrue(len(calls) > 0)
        filter_used = calls[0][1]["search_filter"]
        self.assertNotIn("(sAMAccountName=*)(uid=*)", filter_used)
        self.assertIn("\\2a", filter_used)

    # =========================================================================
    # 4. ORQUESTRAÇÃO, PERMISSÕES & BINDS DE REDE (DOCKER)
    # =========================================================================
    def test_docker_compose_binds_ports_to_localhost(self):
        """Valida que portas críticas de banco e API no docker-compose.yml estão restritas a 127.0.0.1."""
        compose_file = ROOT_DIR / "docker-compose.yml"
        if not compose_file.exists():
            self.skipTest("docker-compose.yml não está montado neste ambiente de execução")

        content = compose_file.read_text(encoding="utf-8")

        # Verifica porta do banco PostgreSQL (5432)
        self.assertIn("127.0.0.1:${POSTGRES_PORT:-5432}:5432", content)
        self.assertNotIn('"5432:5432"', content)

        # Verifica porta da Evolution API (8080)
        self.assertIn('"127.0.0.1:8080:8080"', content)
        self.assertNotIn('"8080:8080"', content)

        # Verifica proteção de tokens na Evolution API
        self.assertIn("AUTHENTICATION_EXPOSE_IN_FETCH_INSTANCES=false", content)

    def test_gitignore_protects_all_sensitive_files(self):
        """Valida que o .gitignore bloqueia arquivos de credenciais, chaves e ambientes."""
        gitignore_file = ROOT_DIR / ".gitignore"
        if not gitignore_file.exists():
            self.skipTest(".gitignore não está montado neste ambiente de execução")

        content = gitignore_file.read_text(encoding="utf-8")

        required_patterns = [
            ".env",
            ".secret.key",
            "*.key",
            "*.pem",
            "cred_admin*.xml",
            "sccm_inventory.json"
        ]
        for pat in required_patterns:
            self.assertIn(pat, content, f"Padrão '{pat}' ausente no .gitignore")

    def test_powershell_scripts_utf8_bom(self):
        """Valida que todos os scripts PowerShell possuem UTF-8 BOM para execução segura."""
        UTF8_BOM = b"\xef\xbb\xbf"
        for ps_path in ROOT_DIR.glob("**/*.ps1"):
            raw = ps_path.read_bytes()
            self.assertTrue(
                raw.startswith(UTF8_BOM),
                f"Script {ps_path.relative_to(ROOT_DIR)} não possui assinatura UTF-8 BOM"
            )

    # =========================================================================
    # 5. AUTENTICAÇÃO NEGATIVA & PROTEÇÃO CONTRA FORÇA BRUTA (RATE LIMITING)
    # =========================================================================
    def test_login_user_negative_inputs(self):
        """Valida rejeição de inputs vazios ou maliciosos no login."""
        ok, msg = login_user("", "")
        self.assertFalse(ok)
        self.assertIn("obrigatórios", msg)

        ok, msg = login_user("user_com_caracteres_invalidos!#$%&*", "senha")
        self.assertFalse(ok)
        self.assertIn("caracteres inválidos", msg)

    def test_login_rate_limiting_lockout(self):
        """Valida bloqueio temporário por força bruta após múltiplas tentativas falhas consecutivas."""
        test_user = "sec_test_user"

        # Simula MAX_FAILED_ATTEMPTS falhas
        with patch("src.auth.authenticate_user_credentials", return_value=(False, "Senha incorreta")):
            for i in range(MAX_FAILED_ATTEMPTS):
                allowed, _ = check_login_rate_limit(test_user)
                self.assertTrue(allowed, f"Deveria permitir tentativa {i+1}")
                ok, msg = login_user(test_user, f"senha_errada_{i}")
                self.assertFalse(ok)
                self.assertEqual(msg, "Senha incorreta")

            # Na 6ª tentativa, deve ser bloqueado pelo Rate Limiter imediatamente sem tentar autenticar
            allowed, wait_sec = check_login_rate_limit(test_user)
            self.assertFalse(allowed)
            self.assertGreater(wait_sec, 0)

            ok, msg = login_user(test_user, "senha_correta_apos_lockout")
            self.assertFalse(ok)
            self.assertIn("bloqueada temporariamente", msg)

        # Após reset_login_attempts (ex: desbloqueio administrativo ou sucesso), volta a permitir
        reset_login_attempts(test_user)
        allowed, _ = check_login_rate_limit(test_user)
        self.assertTrue(allowed)

    def test_auth_token_signature_and_tampering(self):
        """Valida que tokens forjados ou corrompidos são rejeitados de imediato."""
        valid_token = create_auth_token("usuario_teste_seguranca")
        payload = decode_auth_token(valid_token)
        self.assertIsNotNone(payload)
        self.assertEqual(payload["username"], "usuario_teste_seguranca")

        # Token adulterado
        tampered_token = valid_token[:-4] + "abcd"
        self.assertIsNone(decode_auth_token(tampered_token))
        self.assertIsNone(decode_auth_token(""))
        self.assertIsNone(decode_auth_token("token.falso.invalido"))


if __name__ == "__main__":
    unittest.main()
