import os
import re
import requests
import base64
from typing import Dict, Any, Optional
from src.config import _cfg

def _get_api_config() -> tuple[str, str, str]:
    """Retorna a URL base, chave de API e nome da instância da Evolution API (Fail-Closed)."""
    url = (_cfg("EVOLUTION_API_URL") or os.getenv("EVOLUTION_API_URL", "http://evolution-api:8080")).strip().rstrip("/")
    key = (_cfg("EVOLUTION_API_KEY") or os.getenv("EVOLUTION_API_KEY", "")).strip()
    instance = (_cfg("EVOLUTION_INSTANCE_NAME") or os.getenv("EVOLUTION_INSTANCE_NAME", "bancada_sti")).strip()
    return url, key, instance

def get_connection_status() -> Dict[str, Any]:
    """
    Retorna o status de conexão da instância do WhatsApp na Evolution API.
    Possíveis estados: 'open' (conectado), 'connecting', 'close' (desconectado) ou 'offline' (API inacessível).
    """
    base_url, api_key, instance = _get_api_config()
    if not api_key:
        return {"online": False, "state": "unauthorized", "error": "Chave EVOLUTION_API_KEY ausente ou não configurada (Fail-Closed)."}

    headers = {"apikey": api_key}
    endpoint = f"{base_url}/instance/connectionState/{instance}"

    try:
        resp = requests.get(endpoint, headers=headers, timeout=5)
        if resp.status_code == 200:
            data = resp.json()
            state = data.get("instance", {}).get("state", "close")
            return {"online": True, "state": state, "data": data}
        elif resp.status_code == 404:
            return {"online": True, "state": "not_created", "data": None}
        elif resp.status_code in [401, 403]:
            return {"online": False, "state": "unauthorized", "error": "Chave de API inválida ou não autorizada."}
        else:
            return {"online": False, "state": "error", "error": f"HTTP {resp.status_code}"}
    except Exception as e:
        return {"online": False, "state": "offline", "error": str(e)}

def create_instance_if_needed() -> bool:
    """Cria a instância com integração WHATSAPP-BAILEYS se não existir (Fail-Closed)."""
    base_url, api_key, instance = _get_api_config()
    if not api_key:
        return False

    headers = {"apikey": api_key, "Content-Type": "application/json"}
    payload = {
        "instanceName": instance,
        "token": api_key,
        "qrcode": True,
        "integration": "WHATSAPP-BAILEYS"
    }
    try:
        resp = requests.post(f"{base_url}/instance/create", json=payload, headers=headers, timeout=10)
        return resp.status_code in [200, 201, 403, 409]
    except Exception:
        return False

def get_qr_code() -> Dict[str, Any]:
    """
    Solicita a conexão da instância e retorna o QR Code em base64 e código de pareamento (Fail-Closed).
    """
    base_url, api_key, instance = _get_api_config()
    if not api_key:
        return {"success": False, "error": "Chave EVOLUTION_API_KEY ausente ou não configurada (Fail-Closed)."}

    headers = {"apikey": api_key}

    # Garante que a instância existe
    create_instance_if_needed()

    endpoint = f"{base_url}/instance/connect/{instance}"
    try:
        resp = requests.get(endpoint, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            # Evolution API v2 retorna no formato base64 ou code
            b64 = data.get("base64") or data.get("qrcode", {}).get("base64")
            code = data.get("code") or data.get("qrcode", {}).get("code")
            pairing_code = data.get("pairingCode")
            count = data.get("count", 0)

            return {
                "success": True,
                "base64": b64,
                "code": code,
                "pairing_code": pairing_code,
                "count": count
            }
        else:
            return {"success": False, "error": f"Status {resp.status_code}: {resp.text}"}
    except Exception as e:
        return {"success": False, "error": str(e)}

def disconnect_instance() -> bool:
    """Desconecta / encerra a sessão ativa do WhatsApp (Fail-Closed)."""
    base_url, api_key, instance = _get_api_config()
    if not api_key:
        return False

    headers = {"apikey": api_key}
    endpoint = f"{base_url}/instance/logout/{instance}"
    try:
        resp = requests.delete(endpoint, headers=headers, timeout=8)
        return resp.status_code == 200
    except Exception:
        return False

def format_clean_phone(phone_raw: str) -> str:
    """Extrai somente dígitos e garante prefixo DDI 55."""
    if not phone_raw:
        return ""
    digits = re.sub(r"\D", "", str(phone_raw))
    if len(digits) in [10, 11] and not digits.startswith("55"):
        digits = f"55{digits}"
    return digits

def send_whatsapp_text(phone_or_group: str, text: str) -> Dict[str, Any]:
    """
    Envia uma mensagem de texto pelo WhatsApp utilizando a Evolution API (Fail-Closed).
    Suporta telefones individuais (ex: 5567991455446) ou JIDs de grupo.
    """
    base_url, api_key, instance = _get_api_config()
    if not api_key:
        return {"success": False, "error": "Chave EVOLUTION_API_KEY ausente ou não configurada (Fail-Closed)."}

    headers = {"apikey": api_key, "Content-Type": "application/json"}

    destinatario = str(phone_or_group).strip()
    if not "@" in destinatario:
        destinatario = format_clean_phone(destinatario)

    payload = {
        "number": destinatario,
        "text": text
    }

    endpoint = f"{base_url}/message/sendText/{instance}"
    try:
        resp = requests.post(endpoint, json=payload, headers=headers, timeout=12)
        if resp.status_code in [200, 201]:
            return {"success": True, "data": resp.json()}
        else:
            return {"success": False, "error": f"HTTP {resp.status_code}: {resp.text}"}
    except Exception as e:
        return {"success": False, "error": str(e)}
