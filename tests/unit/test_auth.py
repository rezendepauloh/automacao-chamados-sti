import os
import unittest
import time
from unittest.mock import patch

from src.auth import (
    ADMIN_USERS,
    load_admin_users,
    create_auth_token,
    decode_auth_token,
    get_auth_secret,
    is_admin,
    can_edit,
    SESSION_DURATION_SECONDS,
)

class TestAuthAndRBAC(unittest.TestCase):
    """Testes unitários para o sistema de autenticação, Tokens Assinados e controle de acesso (RBAC)."""

    def test_admin_users_configuration(self):
        """Verifica se os administradores do .env estão configurados com privilégios e contas de admin."""
        self.assertTrue(len(ADMIN_USERS) > 0)
        for username, user_data in ADMIN_USERS.items():
            self.assertEqual(user_data["role"], "admin")
            self.assertTrue(user_data.get("admin_sys"))
            self.assertTrue(user_data.get("admin_ad"))

    def test_token_creation_and_decoding_admin(self):
        """Testa geração de token seguro assinado para usuário admin e valida payload retornado."""
        first_admin = list(ADMIN_USERS.keys())[0]
        token = create_auth_token(first_admin)
        self.assertIsInstance(token, str)
        self.assertTrue(len(token) > 20)

        payload = decode_auth_token(token)
        self.assertIsNotNone(payload)
        self.assertEqual(payload["username"], first_admin)
        self.assertEqual(payload["role"], "admin")
        self.assertEqual(payload["admin_sys"], ADMIN_USERS[first_admin]["admin_sys"])
        self.assertEqual(payload["admin_ad"], ADMIN_USERS[first_admin]["admin_ad"])
        self.assertEqual(payload["display_name"], ADMIN_USERS[first_admin]["display_name"])

    def test_token_creation_and_decoding_viewer(self):
        """Testa geração de token para terceirizado/colaborador comum (deve receber role viewer)."""
        token = create_auth_token("colaborador_externo")
        payload = decode_auth_token(token)
        self.assertIsNotNone(payload)
        self.assertEqual(payload["username"], "colaborador_externo")
        self.assertEqual(payload["role"], "viewer")
        self.assertIsNone(payload.get("admin_sys"))
        self.assertIsNone(payload.get("admin_ad"))

    def test_invalid_and_tampered_token(self):
        """Testa validação contra tokens forjados ou corrompidos."""
        self.assertIsNone(decode_auth_token(""))
        self.assertIsNone(decode_auth_token(None))
        self.assertIsNone(decode_auth_token("token.invalido.123"))

        # Token adulterado
        token = create_auth_token("admin_test_user")
        adulterado = token[:-5] + "XXXXX"
        self.assertIsNone(decode_auth_token(adulterado))

    def test_expired_token(self):
        """Testa rejeição de token expirado com max_age curto."""
        import itsdangerous
        from src.auth import _get_serializer
        s = _get_serializer()
        # Token gerado com timestamp artificialmente no passado
        token = s.dumps({"username": "admin_test_user", "role": "admin"})
        # Validando com max_age=-1 (já expirado)
        with self.assertRaises(itsdangerous.SignatureExpired):
            s.loads(token, max_age=-1)

    @patch("src.auth.get_current_user")
    def test_is_admin_and_can_edit_helpers(self, mock_get_user):
        """Testa as funções auxiliares is_admin e can_edit para diferentes perfis."""
        # Cenário 1: Administrador
        mock_get_user.return_value = {"username": "admin_test_user", "role": "admin"}
        self.assertTrue(is_admin())
        self.assertTrue(can_edit())

        # Cenário 2: Colaborador / Viewer
        mock_get_user.return_value = {"username": "usuario_comum", "role": "viewer"}
        self.assertFalse(is_admin())
        self.assertFalse(can_edit())

        # Cenário 3: Não logado
        mock_get_user.return_value = None
        self.assertFalse(is_admin())
        self.assertFalse(can_edit())

    def test_load_admin_users_from_env_json(self):
        """Valida que load_admin_users carrega corretamente JSON customizado do .env."""
        domain = os.getenv("AD_EMAIL").strip().lstrip("@")
        custom_json = f'{{"novoadmin": {{"display_name": "Novo Admin", "role": "admin", "admin_sys": "novo_adm", "admin_ad": "novo_ad", "email": "novo@{domain}"}}}}'
        with patch.dict("os.environ", {"ADMIN_USERS_JSON": custom_json}):
            loaded = load_admin_users()
            self.assertIn("novoadmin", loaded)
            self.assertEqual(loaded["novoadmin"]["role"], "admin")
            self.assertEqual(loaded["novoadmin"]["admin_sys"], "novo_adm")

if __name__ == "__main__":
    unittest.main()
