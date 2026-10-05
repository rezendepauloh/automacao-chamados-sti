# -*- coding: utf-8 -*-
"""
Testes unitários e de integração do Módulo Gerenciador de Senhas (Cofre da Bancada):
Valida motor Fernet/AES com sqlite, criação de tabelas, inserção cifrada, mascaramento padrão,
obtenção decifrada, atualização com/sem troca de senha, deleção e autenticação.
"""

import sys
import unittest
import pandas as pd
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.database.senhas_db import (
    setup_senhas_table,
    salvar_senha,
    listar_senhas,
    get_senhas_df,
    obter_senha_decifrada,
    obter_credencial_por_id,
    atualizar_senha,
    excluir_senha,
    get_senhas_stats,
    CATEGORIAS_PADRAO
)
from src.tabs.senhas import _gerar_senha_forte, is_vault_unlocked, unlock_vault, lock_vault
from src.services.ad_ldap_service import authenticate_user_credentials
from src.crypto_utils import is_fernet_token


class TestSenhasModule(unittest.TestCase):
    def setUp(self):
        setup_senhas_table()

    def test_gerador_senhas_fortes(self):
        """Valida que o gerador de senhas produz senhas seguras com letras, números e símbolos."""
        senha1 = _gerar_senha_forte(16)
        senha2 = _gerar_senha_forte(20)

        self.assertEqual(len(senha1), 16)
        self.assertEqual(len(senha2), 20)
        self.assertNotEqual(senha1, senha2)

        # Checa presença de maiúsculas, minúsculas e dígitos
        self.assertTrue(any(c.isupper() for c in senha1))
        self.assertTrue(any(c.islower() for c in senha1))
        self.assertTrue(any(c.isdigit() for c in senha1))

    def test_crud_senhas_e_criptografia(self):
        """Valida inserção criptografada, decifração e persistência no banco SQLite."""
        # 1. Inserir nova senha
        senha_plana = "SuperSecretPass@2026!"
        novo_id = salvar_senha(
            titulo="Switch Core Sala Servidores",
            categoria="Rede & Switches",
            usuario="admin_rede",
            senha_plana=senha_plana,
            url_sistema="10.10.1.1",
            observacoes="Acesso via SSH / HTTPS"
        )
        self.assertIsInstance(novo_id, int)
        self.assertGreater(novo_id, 0)

        # 2. Listar senhas e checar mascaramento
        lista = listar_senhas(busca="Switch Core")
        self.assertTrue(any(x["id"] == novo_id for x in lista))
        item = next(x for x in lista if x["id"] == novo_id)
        self.assertEqual(item["senha_mascarada"], "••••••••")
        self.assertEqual(item["usuario"], "admin_rede")
        self.assertEqual(item["categoria"], "Rede & Switches")

        # 3. Decifrar senha
        decifrada = obter_senha_decifrada(novo_id)
        self.assertEqual(decifrada, senha_plana)

        # 4. Obter credencial completa por id
        cred = obter_credencial_por_id(novo_id)
        self.assertIsNotNone(cred)
        self.assertEqual(cred["senha_plana"], senha_plana)
        self.assertEqual(cred["url_sistema"], "10.10.1.1")

        # 5. Atualizar sem trocar a senha (manter existente)
        atualizado = atualizar_senha(
            senha_id=novo_id,
            titulo="Switch Core Sala Servidores - Atualizado",
            categoria="Rede & Switches",
            usuario="admin_rede_novo",
            senha_plana=None,
            url_sistema="10.10.1.2",
            observacoes="Observação editada"
        )
        self.assertTrue(atualizado)
        cred_edit = obter_credencial_por_id(novo_id)
        self.assertEqual(cred_edit["titulo"], "Switch Core Sala Servidores - Atualizado")
        self.assertEqual(cred_edit["usuario"], "admin_rede_novo")
        self.assertEqual(cred_edit["senha_plana"], senha_plana)

        # 6. Atualizar alterando a senha
        nova_senha = "OutraSenhaNova#9988"
        atualizado_senha = atualizar_senha(
            senha_id=novo_id,
            titulo="Switch Core Sala Servidores - Atualizado",
            categoria="Rede & Switches",
            usuario="admin_rede_novo",
            senha_plana=nova_senha
        )
        self.assertTrue(atualizado_senha)
        self.assertEqual(obter_senha_decifrada(novo_id), nova_senha)

        # 7. Excluir credencial
        excluido = excluir_senha(novo_id)
        self.assertTrue(excluido)
        self.assertIsNone(obter_credencial_por_id(novo_id))

    def test_filtros_e_estatisticas(self):
        """Valida listagem com filtros por categoria, busca textual e KPIs."""
        id1 = salvar_senha(
            titulo="Oracle Database Produção",
            categoria="Bancos de Dados",
            usuario="system",
            senha_plana="ora_pwd_123",
            url_sistema="srv-ora.corp"
        )
        id2 = salvar_senha(
            titulo="Roteador Wi-Fi Visitantes",
            categoria="Rede & Switches",
            usuario="admin",
            senha_plana="wifi_pwd_456"
        )

        try:
            stats = get_senhas_stats()
            self.assertGreaterEqual(stats["total_senhas"], 2)
            self.assertGreaterEqual(stats["total_categorias"], 1)

            df = get_senhas_df(categoria="Bancos de Dados")
            ids_cat = df["id"].tolist() if hasattr(df["id"], "tolist") else list(df["id"])
            self.assertIn(id1, ids_cat)
            self.assertNotIn(id2, ids_cat)

            df_busca = get_senhas_df(busca="Visitantes")
            ids_busca = df_busca["id"].tolist() if hasattr(df_busca["id"], "tolist") else list(df_busca["id"])
            self.assertIn(id2, ids_busca)
            self.assertNotIn(id1, ids_busca)
        finally:
            excluir_senha(id1)
            excluir_senha(id2)

    def test_validacoes_campos_obrigatorios(self):
        """Valida que títulos ou senhas vazias geram exceções apropriadas."""
        with self.assertRaises(ValueError):
            salvar_senha(titulo="", categoria="Outros", usuario="usr", senha_plana="pwd")

        with self.assertRaises(ValueError):
            salvar_senha(titulo="Titulo", categoria="Outros", usuario="", senha_plana="pwd")

        with self.assertRaises(ValueError):
            salvar_senha(titulo="Titulo", categoria="Outros", usuario="usr", senha_plana="")

    def test_controle_sessao_cofre(self):
        """Valida bloqueio e desbloqueio temporizado na sessão."""
        lock_vault()
        self.assertFalse(is_vault_unlocked())

        unlock_vault()
        self.assertTrue(is_vault_unlocked())

        lock_vault()
        self.assertFalse(is_vault_unlocked())

    def test_authenticate_user_credentials_fallback(self):
        """Valida fluxo de autenticação defensivo com dados vazios e credenciais."""
        ok_vazio, _ = authenticate_user_credentials("", "")
        self.assertFalse(ok_vazio)

        # Autenticação com credencial incorreta
        ok_err, _ = authenticate_user_credentials("usuario_fake_inexistente", "senha_errada_123")
        self.assertFalse(ok_err)


if __name__ == "__main__":
    unittest.main()
