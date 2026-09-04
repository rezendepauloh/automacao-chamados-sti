import os
from pathlib import Path
from cryptography.fernet import Fernet

# Arquivo para armazenar a chave mestra caso não exista no ambiente
_KEY_FILE = Path(__file__).parent.parent / ".secret.key"

def _get_or_create_key() -> bytes:
    """
    Obtém a chave Fernet a partir da variável de ambiente APP_SECRET_KEY
    ou a partir do arquivo local .secret.key. Se não existir, gera e salva uma nova.
    """
    env_key = os.getenv("APP_SECRET_KEY")
    if env_key and env_key.strip():
        return env_key.strip().encode()

    if _KEY_FILE.exists():
        try:
            key_bytes = _KEY_FILE.read_bytes().strip()
            if key_bytes:
                return key_bytes
        except Exception:
            pass

    # Gera uma nova chave Fernet válida
    new_key = Fernet.generate_key()
    try:
        _KEY_FILE.write_bytes(new_key)
        # Tenta restringir permissões em sistemas POSIX
        if os.name == 'posix':
            os.chmod(_KEY_FILE, 0o600)
    except Exception:
        pass

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

def decrypt_value(cipher_text: str) -> str:
    """Decriptografa uma string cifrada via Fernet. Se não estiver cifrada ou falhar, retorna o texto original como fallback."""
    if not cipher_text:
        return ""
    try:
        key = _get_or_create_key()
        f = Fernet(key)
        return f.decrypt(cipher_text.encode('utf-8')).decode('utf-8')
    except Exception:
        # Se for um valor pré-existente ainda não criptografado, retorna como está (fallback suave)
        return cipher_text

def mask_secret(secret_text: str, visible_chars: int = 4) -> str:
    """Retorna uma versão mascarada do segredo para preview em UI segura."""
    if not secret_text:
        return ""
    if len(secret_text) <= visible_chars * 2:
        return "•" * len(secret_text)
    return secret_text[:visible_chars] + "•" * 8 + secret_text[-visible_chars:]
