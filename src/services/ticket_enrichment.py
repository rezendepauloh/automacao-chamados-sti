# -*- coding: utf-8 -*-
"""
Serviço unificado de enriquecimento e formatação de chamados.

Centraliza:
- Formatação padronizada de datas no padrão brasileiro (dd/mm/aaaa hh:mm:ss);
- Resolução do sAMAccountName do solicitante (id_cliente, ad_cache_users, sccm_cache_users e fallback LDAP);
- Descoberta em cascata de IP de Origem e Hostname (banco, mineração de texto/comentários,
  cache SCCM por usuário/host/IP/patrimônio, cache AD de computadores, histórico de chamados,
  DNS/rDNS rápido e consulta SCCM ao vivo com fallback garantido para vazio/"N/A").
"""

from __future__ import annotations

import concurrent.futures
import logging
import os
import re
import socket
import time
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)

INVALID_SENTINELS = {
    "",
    "nan",
    "none",
    "null",
    "<na>",
    "n/a",
    "na",
    "timeout",
    "erro",
    "acesso negado",
    "não informado",
    "nao informado",
    "não encontrada",
    "nao encontrada",
    "não encontrado",
    "nao encontrado",
}

_STOP_HOST_WORDS = {
    "nao",
    "não",
    "sem",
    "com",
    "para",
    "esta",
    "está",
    "novo",
    "nova",
    "velho",
    "computador",
    "maquina",
    "máquina",
    "estacao",
    "estação",
    "notebook",
    "desktop",
    "monitor",
    "reiniciando",
    "travando",
    "desligando",
    "nenhum",
    "nenhuma",
    "local",
    "remoto",
    "usuario",
    "usuário",
}

# Caches curtos em memória para evitar chamadas repetidas de rede em reruns do Streamlit
_LDAP_USER_CACHE: Dict[str, tuple[float, Dict[str, str]]] = {}
_DNS_CACHE: Dict[str, tuple[float, str]] = {}
_NETWORK_TTL_SECONDS = 180.0


def sanitize_network_value(val: Any) -> str:
    """Limpa e valida valores de rede/identificadores, descartando sentinelas inválidos."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s.lower() in INVALID_SENTINELS:
        return ""
    return s


def is_valid_ipv4(ip: str) -> bool:
    """Valida um endereço IPv4 descartando endereços reservados/coringa inválidos para estações."""
    s = sanitize_network_value(ip)
    if not s or not re.match(r"^\d{1,3}(?:\.\d{1,3}){3}$", s):
        return False
    parts = [int(p) for p in s.split(".")]
    if any(p < 0 or p > 255 for p in parts):
        return False
    if s in {"0.0.0.0", "127.0.0.1", "255.255.255.255"}:
        return False
    return True


def is_valid_hostname(host: str) -> bool:
    """Valida se uma string representa um hostname válido de estação."""
    s = sanitize_network_value(host)
    if not s or len(s) < 3 or len(s) > 63:
        return False
    if s.lower() in _STOP_HOST_WORDS:
        return False
    if re.match(r"^\d+(?:\.\d+)*$", s):
        return False
    return bool(re.match(r"^[A-Za-z0-9][A-Za-z0-9.\-_]{1,61}[A-Za-z0-9]$", s))


def format_br_datetime(val: Any, fallback: str = "N/A") -> str:
    """
    Converte datas/timestamps variados (ISO, SQL, BR com ou sem segundos, datetime)
    para o formato brasileiro estrito 'dd/mm/aaaa hh:mm:ss'.
    """
    if val is None:
        return fallback
    try:
        if pd.isna(val):
            return fallback
    except Exception:
        pass

    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%d/%m/%Y %H:%M:%S")

    s = str(val).strip()
    if not s or s.lower() in {"nan", "none", "null", "nat", "<na>", "n/a", "-"}:
        return fallback

    # 1. Formato brasileiro explícito (DD/MM/AAAA...) deve ser parseado antes do ISO
    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}", s):
        for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, fmt).strftime("%d/%m/%Y %H:%M:%S")
            except ValueError:
                continue
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        return s

    # 2. Formato ISO / banco (AAAA-MM-DD...)
    clean_iso = re.sub(r"\.\d+(?=(Z|[+-]\d{2}:?\d{2}|$))", "", s)
    clean_iso = clean_iso.rstrip("Z")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(clean_iso, fmt).strftime("%d/%m/%Y %H:%M:%S")
        except ValueError:
            continue

    dt = pd.to_datetime(s, dayfirst=False, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    return s


def extract_network_info_from_text(text: str) -> Dict[str, str]:
    """
    Extrai IPv4, Hostname corporativo e número de Patrimônio a partir de texto livre
    (título, descrição, andamento ou comentários do chamado).
    """
    result = {"ip": "", "hostname": "", "patrimonio": ""}
    if not text:
        return result

    txt = str(text)

    # 1. IP explícito por rótulo
    labeled_ip = re.search(
        r"(?i)\b(?:ip(?:\s+(?:de\s+origem|da\s+m[aá]quina|do\s+computador|local|fixo))?|ipv4|endere[cç]o\s+ip)\s*[:=\-]\s*(\d{1,3}(?:\.\d{1,3}){3})\b",
        txt,
    )
    if labeled_ip and is_valid_ipv4(labeled_ip.group(1)):
        result["ip"] = labeled_ip.group(1).strip()

    # 2. IP privado corporativo solto no texto (10.x, 172.16-31.x, 192.168.x), ignorando versões
    if not result["ip"]:
        for m in re.finditer(
            r"(?<!\d\.)\b((?:10\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])|192\.168)\.\d{1,3}\.\d{1,3})\b(?!\.\d)",
            txt,
        ):
            prefix = txt[max(0, m.start() - 16):m.start()].lower()
            if re.search(r"(?:vers[aã]o|version|build|ver\.?|v\.?)\s*$", prefix) or prefix.endswith("v"):
                continue
            cand_ip = m.group(1).strip()
            if is_valid_ipv4(cand_ip):
                result["ip"] = cand_ip
                break

    # 3. Hostname corporativo padrão MPMS (ex: MPE-58063, PGJ-NT-0123, DMP-43188, PJ-12345)
    corp_host = re.search(
        r"\b((?:MPE|PGJ|PJ[A-Z]{0,4}|DMP|STI)-[A-Za-z0-9\-]*\d[A-Za-z0-9\-]*)\b",
        txt,
        re.IGNORECASE,
    )
    if corp_host and is_valid_hostname(corp_host.group(1)):
        result["hostname"] = corp_host.group(1).strip().upper()

    # 4. Hostname explícito por rótulo
    if not result["hostname"]:
        labeled_host = re.search(
            r"(?i)\b(?:hostname|nome\s+d[ao]\s+(?:m[aá]quina|computador|esta[cç][aã]o|equipamento)|esta[cç][aã]o\s+de\s+trabalho)\s*[:=\-]\s*([A-Za-z0-9][A-Za-z0-9._\-]{2,30})\b",
            txt,
        )
        if labeled_host:
            cand = labeled_host.group(1).strip()
            if is_valid_hostname(cand) and not is_valid_ipv4(cand) and (re.search(r"\d", cand) or "-" in cand):
                result["hostname"] = cand.upper() if "-" in cand else cand

    # 5. Patrimônio / Tombamento para cruzamento posterior no inventário
    pat_match = re.search(
        r"(?i)\b(?:patrim[oô]nio|tombamento|plaqueta|pat\.?)\s*[:=#\-]?\s*(?:n[ºo°]?\.?\s*)?(\d{4,6})\b",
        txt,
    )
    if pat_match:
        result["patrimonio"] = pat_match.group(1).strip()

    return result


def _clean_sam_token(raw: str) -> str:
    s = sanitize_network_value(raw)
    if not s:
        return ""
    if "\\" in s:
        s = s.split("\\")[-1].strip()
    if "@" in s:
        s = s.split("@")[0].strip()
    return s


def _lookup_ad_user_ldap_cached(full_name: str) -> Dict[str, str]:
    """Consulta rápida e com cache TTL no AD via LDAP caso caches locais não possuam o usuário."""
    key = full_name.strip().lower()
    now = time.monotonic()
    cached = _LDAP_USER_CACHE.get(key)
    if cached and (now - cached[0]) < _NETWORK_TTL_SECONDS:
        return cached[1]

    res: Dict[str, str] = {}
    try:
        from ldap3 import ALL, SIMPLE, SUBTREE, Connection, Server
        from ldap3.utils.conv import escape_filter_chars
        from src.config import DOMINIO, DOMINIO_MMC, PASSWORD, USERNAME

        if not DOMINIO or not USERNAME or not PASSWORD:
            return {}

        server = Server(DOMINIO, port=389, get_info=ALL, connect_timeout=2)
        upn = f"{USERNAME}@{DOMINIO}"
        conn = Connection(server, user=upn, password=PASSWORD, authentication=SIMPLE, auto_bind=False, receive_timeout=2)
        if not conn.bind():
            return {}

        try:
            safe_q = escape_filter_chars(full_name.strip())
            for filt in (
                f"(displayName={safe_q})",
                f"(cn={safe_q})",
                f"(name={safe_q})",
                f"(displayName=*{safe_q}*)",
            ):
                conn.search(
                    search_base=f"{DOMINIO_MMC}",
                    search_filter=filt,
                    search_scope=SUBTREE,
                    attributes=["sAMAccountName", "displayName", "distinguishedName"],
                )
                if conn.entries:
                    attrs = conn.entries[0].entry_attributes_as_dict
                    sam_vals = attrs.get("sAMAccountName") or []
                    disp_vals = attrs.get("displayName") or []
                    dn_vals = attrs.get("distinguishedName") or []
                    if sam_vals:
                        res = {
                            "sam_account_name": str(sam_vals[0]).strip(),
                            "display_name": str(disp_vals[0]).strip() if disp_vals else "",
                            "user_dn": str(dn_vals[0]).strip() if dn_vals else "",
                        }
                        break
        finally:
            conn.unbind()
    except Exception as exc:
        logger.debug(f"Lookup LDAP de sAMAccountName indisponível para '{full_name}': {exc}")

    _LDAP_USER_CACHE[key] = (now, res)
    return res


def resolve_user_sam_account_name(
    usuario: str = "",
    id_cliente: str = "",
    allow_ldap: bool = True,
) -> Dict[str, str]:
    """
    Resolve o sAMAccountName (login de rede) do usuário do chamado através de:
    1. id_cliente já registrado (normalizado);
    2. Cache do Active Directory (ad_cache_users);
    3. Cache de usuários do SCCM (sccm_cache_users);
    4. Consulta LDAP ao vivo no AD (opcional, com timeout curto e cache).
    """
    from src.database.ad_db import find_ad_user_by_name, find_ad_user_by_sam
    from src.database.sccm_db import find_sccm_username_by_fullname

    raw_id = sanitize_network_value(id_cliente)
    raw_user = sanitize_network_value(usuario)

    # 1. Se id_cliente já veio preenchido e parece um login (sem espaços)
    if raw_id and " " not in raw_id:
        clean_id = _clean_sam_token(raw_id)
        if clean_id:
            ad_info = find_ad_user_by_sam(clean_id)
            if ad_info:
                return {
                    "sam_account_name": ad_info.get("sam_account_name") or clean_id,
                    "display_name": ad_info.get("display_name") or raw_user,
                    "user_dn": ad_info.get("dn", ""),
                }
            return {
                "sam_account_name": clean_id,
                "display_name": raw_user,
                "user_dn": "",
            }

    query_name = raw_user or raw_id
    if not query_name:
        return {"sam_account_name": "", "display_name": "", "user_dn": ""}

    # 2. Cache local do Active Directory
    ad_match = find_ad_user_by_name(query_name)
    if ad_match and ad_match.get("sam_account_name"):
        return {
            "sam_account_name": ad_match["sam_account_name"],
            "display_name": ad_match.get("display_name") or raw_user,
            "user_dn": ad_match.get("dn", ""),
        }

    # 3. Cache local de usuários do SCCM
    sccm_sam = find_sccm_username_by_fullname(query_name)
    if sccm_sam:
        return {
            "sam_account_name": sccm_sam,
            "display_name": raw_user,
            "user_dn": "",
        }

    # 4. Fallback LDAP em tempo real (somente quando habilitado)
    if allow_ldap:
        ldap_match = _lookup_ad_user_ldap_cached(query_name)
        if ldap_match and ldap_match.get("sam_account_name"):
            return ldap_match

    return {"sam_account_name": "", "display_name": raw_user, "user_dn": ""}


def _resolve_dns_cached(kind: str, target: str) -> str:
    """Executa DNS (hostname -> IP) ou rDNS (IP -> hostname) com timeout curto (1.0s) e cache TTL."""
    cache_key = f"{kind}:{target.lower()}"
    now = time.monotonic()
    cached = _DNS_CACHE.get(cache_key)
    if cached and (now - cached[0]) < _NETWORK_TTL_SECONDS:
        return cached[1]

    def _query() -> str:
        if kind == "host2ip":
            candidates = [target]
            try:
                from src.config import DOMINIO
                if DOMINIO and "." not in target:
                    candidates.append(f"{target}.{DOMINIO}")
            except Exception:
                pass
            for cand in candidates:
                try:
                    ip = socket.gethostbyname(cand)
                    if is_valid_ipv4(ip):
                        return ip
                except Exception:
                    continue
            return ""
        if kind == "ip2host":
            try:
                host, _, _ = socket.gethostbyaddr(target)
                short_host = str(host or "").split(".")[0].strip()
                if is_valid_hostname(short_host):
                    return short_host.upper()
            except Exception:
                return ""
        return ""

    resolved = ""
    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            fut = executor.submit(_query)
            resolved = fut.result(timeout=1.0)
    except Exception:
        resolved = ""

    _DNS_CACHE[cache_key] = (now, resolved)
    return resolved


def _is_sccm_server_reachable() -> bool:
    """Verifica rapidamente (<=0.5s) se o servidor SCCM está acessível antes de invocar PowerShell."""
    try:
        from src.config import _cfg
        site_server = _cfg("SCCM_SERVER", os.getenv("SCCM_SERVER", "")).strip()
        if not site_server:
            return False
        with socket.create_connection((site_server, 135), timeout=0.5):
            return True
    except Exception:
        return False


def resolve_ticket_device_and_user(
    row: Dict[str, Any],
    comments: Optional[List[Dict[str, Any]]] = None,
    allow_network_fallbacks: bool = True,
) -> Dict[str, str]:
    """
    Executa a cascata completa para descobrir:
    - sam_account_name (login do usuário no AD/SCCM)
    - display_name
    - ip_origem
    - hostname

    Retorna strings vazias caso não encontre, permitindo que a UI aplique 'N/A' como último fallback.
    """
    from src.database.ad_db import find_ad_computer_for_user
    from src.database.sccm_db import (
        find_sccm_device_by_hostname,
        find_sccm_device_by_ip,
        find_sccm_device_by_patrimonio,
        get_device_by_user,
    )
    from src.database.tickets_db import get_latest_device_info_by_user

    cid = sanitize_network_value(row.get("id") or row.get("Chamado#"))
    usuario = sanitize_network_value(row.get("usuario") or row.get("Nome do Usuário"))
    id_cliente = sanitize_network_value(row.get("id_cliente") or row.get("ID do Cliente"))

    user_info = resolve_user_sam_account_name(
        usuario=usuario,
        id_cliente=id_cliente,
        allow_ldap=allow_network_fallbacks,
    )
    sam = user_info.get("sam_account_name", "")
    user_dn = user_info.get("user_dn", "")

    curr_ip = sanitize_network_value(row.get("ip_origem") or row.get("IP_Origem"))
    if curr_ip and not is_valid_ipv4(curr_ip):
        curr_ip = ""

    curr_host = sanitize_network_value(row.get("hostname") or row.get("Hostname"))
    if curr_host and not is_valid_hostname(curr_host):
        curr_host = ""

    # Etapa 1: Mineração de texto (título, descrição, andamento e comentários)
    text_parts = [
        sanitize_network_value(row.get("titulo") or row.get("Título")),
        sanitize_network_value(row.get("descricao") or row.get("Descrição")),
        sanitize_network_value(row.get("andamento") or row.get("Andamento")),
    ]
    if comments:
        for c in comments:
            if isinstance(c, dict):
                text_parts.append(sanitize_network_value(c.get("texto")))
    combined_text = "\n".join(p for p in text_parts if p)
    extracted = extract_network_info_from_text(combined_text)

    if not curr_ip and extracted.get("ip"):
        curr_ip = extracted["ip"]
    if not curr_host and extracted.get("hostname"):
        curr_host = extracted["hostname"]
    patrimonio = extracted.get("patrimonio", "")

    # Etapa 2: Cruzamento no cache relacional do SCCM (sccm_cache_devices)
    if curr_host and not curr_ip:
        dev = find_sccm_device_by_hostname(curr_host)
        if dev and is_valid_ipv4(dev.get("ip", "")):
            curr_ip = dev["ip"]

    if curr_ip and not curr_host:
        dev = find_sccm_device_by_ip(curr_ip)
        if dev and is_valid_hostname(dev.get("hostname", "")):
            curr_host = dev["hostname"]

    if (not curr_ip or not curr_host) and sam:
        dev = get_device_by_user(sam)
        if dev:
            if not curr_ip and is_valid_ipv4(dev.get("ip", "")):
                curr_ip = dev["ip"]
            if not curr_host and is_valid_hostname(dev.get("hostname", "")):
                curr_host = dev["hostname"]

    if (not curr_ip or not curr_host) and patrimonio:
        dev = find_sccm_device_by_patrimonio(patrimonio)
        if dev:
            if not curr_ip and is_valid_ipv4(dev.get("ip", "")):
                curr_ip = dev["ip"]
            if not curr_host and is_valid_hostname(dev.get("hostname", "")):
                curr_host = dev["hostname"]

    # Etapa 3: Cache de computadores do AD (ad_cache_computers)
    if not curr_host:
        ad_comp = find_ad_computer_for_user(
            sam_account_name=sam,
            display_name=usuario,
            user_dn=user_dn,
            patrimonio=patrimonio,
        )
        if ad_comp and is_valid_hostname(ad_comp.get("hostname", "")):
            curr_host = ad_comp["hostname"]
            # Se descobriu hostname pelo AD, tenta buscar o IP dele no SCCM cache
            if not curr_ip:
                dev = find_sccm_device_by_hostname(curr_host)
                if dev and is_valid_ipv4(dev.get("ip", "")):
                    curr_ip = dev["ip"]

    # Etapa 4: Histórico de outros chamados do mesmo usuário no banco
    if (not curr_ip or not curr_host) and (sam or usuario):
        hist = get_latest_device_info_by_user(
            usuario=usuario,
            id_cliente=sam or id_cliente,
            exclude_ticket_id=cid,
        )
        if hist:
            if not curr_ip and is_valid_ipv4(hist.get("ip_origem", "")):
                curr_ip = hist["ip_origem"]
            if not curr_host and is_valid_hostname(hist.get("hostname", "")):
                curr_host = hist["hostname"]

    # Etapa 5: Fallbacks rápidos de rede (DNS/rDNS + SCCM ao vivo se acessível)
    if allow_network_fallbacks:
        if curr_host and not curr_ip:
            dns_ip = _resolve_dns_cached("host2ip", curr_host)
            if is_valid_ipv4(dns_ip):
                curr_ip = dns_ip

        if curr_ip and not curr_host:
            rdns_host = _resolve_dns_cached("ip2host", curr_ip)
            if is_valid_hostname(rdns_host):
                curr_host = rdns_host

        if (not curr_ip or not curr_host) and sam and _is_sccm_server_reachable():
            try:
                from src.config import fetch_sccm_data
                sccm_live = fetch_sccm_data(sam)
                live_ip = sanitize_network_value(sccm_live.get("ip"))
                live_host = sanitize_network_value(sccm_live.get("hostname"))
                if not curr_ip and is_valid_ipv4(live_ip):
                    curr_ip = live_ip
                if not curr_host and is_valid_hostname(live_host):
                    curr_host = live_host
            except Exception as exc:
                logger.debug(f"Fallback SCCM ao vivo indisponível para '{sam}': {exc}")

    return {
        "sam_account_name": sam,
        "display_name": user_info.get("display_name") or usuario,
        "ip_origem": curr_ip,
        "hostname": curr_host,
    }
