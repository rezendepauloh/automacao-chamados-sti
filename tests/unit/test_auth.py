import unittest
import time
from unittest.mock import patch

from src.auth import (
    ADMIN_USERS,
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
        """Verifica se os 3 analistas da Bancada estão configurados com privilégios e contas de admin."""
        self.assertIn("paulogoncalves", ADMIN_USERS)
        self.assertIn("reginaldosb", ADMIN_USERS)
        self.assertIn("luizvillalba", ADMIN_USERS)

        paulo = ADMIN_USERS["paulogoncalves"]
        self.assertEqual(paulo["role"], "admin")
        self.assertEqual(paulo["admin_sys"], "paulo_admin")
        self.assertEqual(paulo["admin_ad"], "paulo_admin_ad")

        reginaldo = ADMIN_USERS["reginaldosb"]
        self.assertEqual(reginaldo["role"], "admin")
        self.assertEqual(reginaldo["admin_sys"], "reginaldo_admin")
        self.assertEqual(reginaldo["admin_ad"], "reginaldo_admin_ad")

        villalba = ADMIN_USERS["luizvillalba"]
        self.assertEqual(villalba["role"], "admin")
        self.assertEqual(villalba["admin_sys"], "villalba_admin")
        self.assertEqual(villalba["admin_ad"], "villalba_admin_ad")

    def test_token_creation_and_decoding_admin(self):
        """Testa geração de token seguro assinado para usuário admin e valida payload retornado."""
        token = create_auth_token("paulogoncalves")
        self.assertIsInstance(token, str)
        self.assertTrue(len(token) > 20)

        payload = decode_auth_token(token)
        self.assertIsNotNone(payload)
        self.assertEqual(payload["username"], "paulogoncalves")
        self.assertEqual(payload["role"], "admin")
        self.assertEqual(payload["admin_sys"], "paulo_admin")
        self.assertEqual(payload["admin_ad"], "paulo_admin_ad")
        self.assertIn("Paulo Henrique", payload["display_name"])

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
        token = create_auth_token("paulogoncalves")
        adulterado = token[:-5] + "XXXXX"
        self.assertIsNone(decode_auth_token(adulterado))

    def test_expired_token(self):
        """Testa rejeição de token expirado com max_age curto."""
        import itsdangerous
        from src.auth import _get_serializer
        s = _get_serializer()
        # Token gerado com timestamp artificialmente no passado
        token = s.dumps({"username": "paulogoncalves", "role": "admin"})
        # Validando com max_age=-1 (já expirado)
        with self.assertRaises(itsdangerous.SignatureExpired):
            s.loads(token, max_age=-1)

    @patch("src.auth.get_current_user")
    def test_is_admin_and_can_edit_helpers(self, mock_get_user):
        """Testa as funções auxiliares is_admin e can_edit para diferentes perfis."""
        # Cenário 1: Administrador
        mock_get_user.return_value = {"username": "paulogoncalves", "role": "admin"}
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

if __name__ == "__main__":
    unittest.main()
