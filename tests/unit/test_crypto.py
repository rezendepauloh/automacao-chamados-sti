# -*- coding: utf-8 -*-
"""
Testes unitários de criptografia e cofre de credenciais (src.crypto_utils).
Valida criptografia simétrica Fernet, decriptografia, fallback de texto plano e máscara de segredos.
"""

import os
import sys
import unittest
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import tests.test_helpers
from src.crypto_utils import encrypt_value, decrypt_value, mask_secret

class TestCryptoUtils(unittest.TestCase):
    def test_encrypt_and_decrypt_roundtrip(self):
        """Valida se uma senha é criptografada e decriptografada de volta com 100% de exatidão."""
        original = "MinhaSenhaSuperSecreta@2026#Bancada"
        encrypted = encrypt_value(original)
        
        self.assertNotEqual(original, encrypted)
        self.assertTrue(len(encrypted) > len(original))
        
        decrypted = decrypt_value(encrypted)
        self.assertEqual(original, decrypted)

    def test_encrypt_empty_or_none(self):
        """Valida comportamento seguro ao receber strings vazias ou nulas."""
        self.assertEqual(encrypt_value(""), "")
        self.assertEqual(encrypt_value(None), "")
        self.assertEqual(decrypt_value(""), "")
        self.assertEqual(decrypt_value(None), "")

    def test_decrypt_plaintext_fallback(self):
        """Valida que textos já em formato plano (legados) não quebram e são retornados como estão."""
        raw_text = "senha_legada_sem_criptografia"
        result = decrypt_value(raw_text)
        self.assertEqual(result, raw_text)

    def test_decrypt_incompatible_or_corrupted_token_fallback(self):
        """Valida que tokens cifrados com outra chave não geram exceção fatal e caem no fallback gracioso com log de aviso."""
        from cryptography.fernet import Fernet
        outra_chave = Fernet.generate_key()
        f_outro = Fernet(outra_chave)
        token_estranho = f_outro.encrypt(b"senha_ambiente_antigo").decode("utf-8")
        
        # Não deve lançar exceção, deve devolver o token para fallback e registrar warning
        resultado = decrypt_value(token_estranho)
        self.assertEqual(resultado, token_estranho)

    def test_is_valid_fernet_key_and_token(self):
        """Valida detecção de integridade de chave Fernet e identificação heurística de tokens."""
        from src.crypto_utils import _is_valid_fernet_key, is_fernet_token
        from cryptography.fernet import Fernet
        
        valid_key = Fernet.generate_key()
        self.assertTrue(_is_valid_fernet_key(valid_key))
        self.assertFalse(_is_valid_fernet_key(b"chave_curta_invalida"))
        self.assertFalse(_is_valid_fernet_key(None))
        self.assertFalse(_is_valid_fernet_key(b""))

        token = Fernet(valid_key).encrypt(b"teste").decode("utf-8")
        self.assertTrue(is_fernet_token(token))
        self.assertFalse(is_fernet_token("senha_plana_123"))
        self.assertFalse(is_fernet_token(None))

    def test_mask_secret(self):
        """Valida mascaramento visual seguro para exibição em interfaces/logs."""
        self.assertEqual(mask_secret(""), "")
        self.assertEqual(mask_secret("1234"), "••••")
        
        long_pass = "Abcd4268#MasterPassword"
        masked = mask_secret(long_pass, visible_chars=4)
        self.assertTrue(masked.startswith("Abcd"))
        self.assertTrue(masked.endswith("word"))
        self.assertIn("••••••••", masked)

if __name__ == "__main__":
    unittest.main()

