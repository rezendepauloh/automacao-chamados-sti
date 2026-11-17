# -*- coding: utf-8 -*-
"""
Módulo de segurança da API REST Bancada STI:
- Autenticação Fail-Closed por API Key (X-API-Key ou Authorization: Bearer)
- Comparação em tempo constante (hmac.compare_digest)
- Rate Limiting por janela deslizante por IP cliente
"""

import hmac
import os
import threading
import time
from collections import defaultdict, deque
from typing import Deque, Dict, Optional, Tuple

MIN_API_KEY_LENGTH = 16


def get_configured_api_key() -> str:
    """Obtém a chave configurada em BANCADA_API_KEY (Fail-Closed: sem fallback hardcoded)."""
    return os.getenv("BANCADA_API_KEY", "").strip()


def extract_api_key_from_headers(headers) -> str:
    """Extrai o token de X-API-Key ou Authorization: Bearer <token>."""
    if not headers:
        return ""
    api_key = (headers.get("x-api-key") or headers.get("X-API-Key") or "").strip()
    if api_key:
        return api_key

    auth_header = (headers.get("authorization") or headers.get("Authorization") or "").strip()
    if auth_header.lower().startswith("bearer "):
        return auth_header[7:].strip()
    return ""


def verify_api_key_headers(headers) -> Tuple[bool, int, str]:
    """
    Valida a API Key da requisição em modo Fail-Closed.
    Retorna (valido, status_http, mensagem_erro).
    """
    server_key = get_configured_api_key()
    if not server_key or len(server_key) < MIN_API_KEY_LENGTH:
        return (
            False,
            503,
            "BANCADA_API_KEY não configurada ou muito curta no servidor (mínimo 16 caracteres - Fail-Closed).",
        )

    provided_key = extract_api_key_from_headers(headers)
    if not provided_key:
        return (
            False,
            401,
            "Credencial ausente. Envie o header 'X-API-Key' ou 'Authorization: Bearer <token>'.",
        )

    if not hmac.compare_digest(provided_key.encode("utf-8"), server_key.encode("utf-8")):
        return False, 403, "API Key inválida."

    return True, 200, ""


class SlidingWindowRateLimiter:
    """Limitador de taxa em memória por IP usando janela deslizante de 60 segundos."""

    def __init__(self, max_requests_per_minute: Optional[int] = None, window_seconds: int = 60):
        self._default_limit = max_requests_per_minute
        self.window_seconds = max(1, int(window_seconds))
        self._hits: Dict[str, Deque[float]] = defaultdict(deque)
        self._lock = threading.Lock()

    def get_limit(self) -> int:
        if self._default_limit is not None:
            return max(1, int(self._default_limit))
        try:
            return max(1, int(os.getenv("BANCADA_API_RATE_LIMIT", "180") or "180"))
        except ValueError:
            return 180

    def is_allowed(self, client_key: str) -> Tuple[bool, int, int]:
        """
        Verifica se o cliente pode executar nova requisição.
        Retorna (permitido, restantes, retry_after_segundos).
        """
        now = time.monotonic()
        limit = self.get_limit()
        cutoff = now - self.window_seconds

        with self._lock:
            dq = self._hits[client_key]
            while dq and dq[0] <= cutoff:
                dq.popleft()

            if len(dq) >= limit:
                retry_after = max(1, int(self.window_seconds - (now - dq[0])) + 1)
                return False, 0, retry_after

            dq.append(now)
            remaining = max(0, limit - len(dq))
            return True, remaining, 0

    def reset(self) -> None:
        """Limpa contadores em memória (útil em testes unitários)."""
        with self._lock:
            self._hits.clear()
