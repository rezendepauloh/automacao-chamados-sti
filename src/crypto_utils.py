import os
from pathlib import Path
from cryptography.fernet import Fernet

import logging

logger = logging.getLogger(__name__)

# Arquivo para armazenar a chave mestra caso não exista no ambiente
_KEY_FILE = Path(__file__).parent.parent / ".secret.key"

def _is_valid_fernet_key(key_bytes: bytes) -> bool:
    """Verifica se os bytes fornecidos formam uma chave Fernet válida."""
    if not key_bytes or len(key_bytes) != 44:
        return False
    try:
        Fernet(key_bytes)
        return True
    except Exception:
        return False

def _get_or_create_key() -> bytes:
    """
    Obtém a chave Fernet a partir da variável de ambiente APP_SECRET_KEY
    ou a partir do arquivo local .secret.key. Se não existir ou for inválida, gera e salva uma nova.
    """
    env_key = os.getenv("APP_SECRET_KEY")
    if env_key and env_key.strip():
        k = env_key.strip().encode()
        if _is_valid_fernet_key(k):
            return k
        logger.warning("⚠️ APP_SECRET_KEY fornecida no ambiente é inválida para Fernet. Recorrendo à chave local.")

    if _KEY_FILE.exists():
        try:
            key_bytes = _KEY_FILE.read_bytes().strip()
            if _is_valid_fernet_key(key_bytes):
                return key_bytes
            logger.warning("⚠️ Arquivo .secret.key local continha chave corrompida. Uma nova chave será gerada.")
        except Exception as e:
            logger.warning(f"⚠️ Falha ao ler .secret.key: {e}")

    # Gera uma nova chave Fernet válida
    new_key = Fernet.generate_key()
    try:
        _KEY_FILE.write_bytes(new_key)
        # Tenta restringir permissões em sistemas POSIX
        if os.name == 'posix':
            os.chmod(_KEY_FILE, 0o600)
    except Exception as e:
        logger.error(f"Não foi possível salvar .secret.key no disco: {e}")

    return new_key

def encrypt_value(plain_text: str) -> str:
    """Criptografa uma string usando Fernet (AES-128-CBC + HMAC-SHA256). Retorna string cifrada em Base64."""
    if not plain_text:
        return ""
    try:
        key = _get_or_create_key()
        f = Fernet(key)
        return f.encrypt(plain_text.encode('utf-8')).decode('utf-8')
    except Exception as e:
        # Em caso de falha severa, não falha silenciosamente
        raise RuntimeError(f"Erro ao criptografar dado sensível: {e}")

def is_fernet_token(text: str) -> bool:
    """Verifica heuristicamente se a string tem a estrutura de um token Fernet (começa com gAAAAA...)."""
    if not text or not isinstance(text, str):
        return False
    return text.startswith("gAAAAA") and len(text) > 50

def decrypt_value(cipher_text: str) -> str:
    """
    Decriptografa uma string cifrada via Fernet.
    Se não estiver cifrada ou falhar por chave incompatível (ex: troca de máquina/ambiente),
    retorna o texto original como fallback e registra aviso claro nos logs.
    """
    if not cipher_text:
        return ""
    try:
        key = _get_or_create_key()
        f = Fernet(key)
        return f.decrypt(cipher_text.encode('utf-8')).decode('utf-8')
    except Exception as e:
        if is_fernet_token(cipher_text):
            logger.warning(
                "⚠️ A credencial cifrada não pôde ser decriptografada com a chave Fernet atual. "
                "Possível troca de ambiente ou chave mestra ausente/incompatível. "
                "Redefina a senha correspondente na aba de Configurações."
            )
        # Se for um valor pré-existente ainda não criptografado ou chave incompatível, retorna como está (fallback suave)
        return cipher_text

def mask_secret(secret_text: str, visible_chars: int = 4) -> str:
    """Retorna uma versão mascarada do segredo para preview em UI segura."""
    if not secret_text:
        return ""
    if len(secret_text) <= visible_chars * 2:
        return "•" * len(secret_text)
    return secret_text[:visible_chars] + "•" * 8 + secret_text[-visible_chars:]
