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
