import os
import json
import logging
from datetime import datetime
import pandas as pd
from typing import List, Dict, Any, Optional
from .connection import get_connection, DB_TYPE

logger = logging.getLogger(__name__)

def setup_sccm_tables():
    """
    Cria as tabelas de cache relacional do SCCM (Dispositivos, Usuários e Coleções).
    Compatível com SQLite e PostgreSQL.
    """
    conn = get_connection()
    cursor = conn.cursor()

    is_pg = DB_TYPE in ["postgres", "postgresql"]
    id_pk = "SERIAL PRIMARY KEY" if is_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"

    # 1. Tabela de Dispositivos (SMS_R_System)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sccm_cache_devices (
        resource_id TEXT PRIMARY KEY,
        name TEXT,
        last_logon_user TEXT,
        ip_addresses TEXT,
        mac_addresses TEXT,
        manufacturer TEXT,
        model TEXT,
        operating_system TEXT,
        os_version TEXT,
        client_version TEXT,
        client_active INTEGER DEFAULT 1,
        last_active_time TEXT,
        ad_site_name TEXT,
        distinguished_name TEXT,
        raw_json TEXT,
        updated_at TEXT
    )
    """)

    # 2. Tabela de Usuários (SMS_R_User)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sccm_cache_users (
        resource_id TEXT PRIMARY KEY,
        user_name TEXT,
        full_user_name TEXT,
        user_group_name TEXT,
        windows_nt_domain TEXT,
        distinguished_name TEXT,
        raw_json TEXT,
        updated_at TEXT
    )
    """)

    # 3. Tabela de Coleções de Dispositivos e Usuários (SMS_Collection)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sccm_cache_collections (
        collection_id TEXT PRIMARY KEY,
        name TEXT,
        collection_type TEXT,
        member_count INTEGER DEFAULT 0,
        comment TEXT,
        last_refresh_time TEXT,
        updated_at TEXT
    )
    """)

    conn.commit()
    cursor.close()
    conn.close()


def save_sccm_devices(devices_list: List[Dict[str, Any]]) -> int:
    """Salva ou atualiza a lista de dispositivos do SCCM em lote."""
    if not devices_list:
        return 0

    setup_sccm_tables()
    conn = get_connection()
    cursor = conn.cursor()
    is_pg = DB_TYPE in ["postgres", "postgresql"]

    now_iso = datetime.now().isoformat()
    count = 0

    for dev in devices_list:
        res_id = str(dev.get("ResourceID") or dev.get("resource_id") or dev.get("Name") or "").strip()
        if not res_id:
            continue

        name = str(dev.get("Name") or "").strip()
        user = str(dev.get("LastLogonUserName") or dev.get("last_logon_user") or "").strip()
        
        # Trata IPs (pode vir como lista ou string)
        ips = dev.get("IPAddresses") or dev.get("ip_addresses") or []
        if isinstance(ips, list):
            ip_str = ", ".join([str(i) for i in ips if i and not str(i).startswith("169.254")])
        else:
            ip_str = str(ips)

        macs = dev.get("MACAddresses") or dev.get("mac_addresses") or []
        if isinstance(macs, list):
            mac_str = ", ".join([str(m) for m in macs if m])
        else:
            mac_str = str(macs)

        manufacturer = str(dev.get("Manufacturer") or dev.get("manufacturer") or "").strip()
        model = str(dev.get("Model") or dev.get("model") or "").strip()
        os_name = str(dev.get("OperatingSystemNameandVersion") or dev.get("operating_system") or "").strip()
        os_ver = str(dev.get("Build") or dev.get("os_version") or "").strip()
        client_ver = str(dev.get("ClientVersion") or dev.get("client_version") or "").strip()
        
        raw_act = dev.get("ClientActiveStatus")
        if raw_act is None:
            raw_act = dev.get("Active", 1)
        client_active = 1 if raw_act in [1, True, "1", "True"] else 0
        last_active = str(dev.get("LastActiveTime") or dev.get("last_active_time") or "").strip()
        ad_site = str(dev.get("ADSiteName") or dev.get("ad_site_name") or "").strip()
        dn = str(dev.get("DistinguishedName") or dev.get("distinguished_name") or "").strip()
        raw_json_str = json.dumps(dev, ensure_ascii=False)

        if is_pg:
            sql = """
            INSERT INTO sccm_cache_devices (
                resource_id, name, last_logon_user, ip_addresses, mac_addresses,
                manufacturer, model, operating_system, os_version, client_version,
                client_active, last_active_time, ad_site_name, distinguished_name,
                raw_json, updated_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (resource_id) DO UPDATE SET
                name = EXCLUDED.name,
                last_logon_user = EXCLUDED.last_logon_user,
                ip_addresses = EXCLUDED.ip_addresses,
                mac_addresses = EXCLUDED.mac_addresses,
                manufacturer = EXCLUDED.manufacturer,
                model = EXCLUDED.model,
                operating_system = EXCLUDED.operating_system,
                os_version = EXCLUDED.os_version,
                client_version = EXCLUDED.client_version,
                client_active = EXCLUDED.client_active,
                last_active_time = EXCLUDED.last_active_time,
                ad_site_name = EXCLUDED.ad_site_name,
                distinguished_name = EXCLUDED.distinguished_name,
                raw_json = EXCLUDED.raw_json,
                updated_at = EXCLUDED.updated_at;
            """
        else:
            sql = """
            INSERT OR REPLACE INTO sccm_cache_devices (
                resource_id, name, last_logon_user, ip_addresses, mac_addresses,
                manufacturer, model, operating_system, os_version, client_version,
                client_active, last_active_time, ad_site_name, distinguished_name,
                raw_json, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """

        params = (
            res_id, name, user, ip_str, mac_str,
            manufacturer, model, os_name, os_ver, client_ver,
            client_active, last_active, ad_site, dn,
            raw_json_str, now_iso
        )
        cursor.execute(sql, params)
        count += 1

    conn.commit()
    cursor.close()
    conn.close()
    return count


def save_sccm_users(users_list: List[Dict[str, Any]]) -> int:
    """Salva ou atualiza a lista de usuários do SCCM."""
    if not users_list:
        return 0

    setup_sccm_tables()
    conn = get_connection()
    cursor = conn.cursor()
    is_pg = DB_TYPE in ["postgres", "postgresql"]
    now_iso = datetime.now().isoformat()
    count = 0

    for u in users_list:
        res_id = str(u.get("ResourceID") or u.get("resource_id") or u.get("UserName") or "").strip()
        if not res_id:
            continue

        user_name = str(u.get("UserName") or "").strip()
        full_name = str(u.get("FullUserName") or u.get("full_user_name") or "").strip()
        user_grp = str(u.get("UserGroupName") or "").strip()
        domain = str(u.get("WindowsNTDomain") or "").strip()
        dn = str(u.get("DistinguishedName") or "").strip()
        raw_json_str = json.dumps(u, ensure_ascii=False)

        if is_pg:
            sql = """
            INSERT INTO sccm_cache_users (
                resource_id, user_name, full_user_name, user_group_name,
                windows_nt_domain, distinguished_name, raw_json, updated_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (resource_id) DO UPDATE SET
                user_name = EXCLUDED.user_name,
                full_user_name = EXCLUDED.full_user_name,
                user_group_name = EXCLUDED.user_group_name,
                windows_nt_domain = EXCLUDED.windows_nt_domain,
                distinguished_name = EXCLUDED.distinguished_name,
                raw_json = EXCLUDED.raw_json,
                updated_at = EXCLUDED.updated_at;
            """
        else:
            sql = """
            INSERT OR REPLACE INTO sccm_cache_users (
                resource_id, user_name, full_user_name, user_group_name,
                windows_nt_domain, distinguished_name, raw_json, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
            """

        params = (res_id, user_name, full_name, user_grp, domain, dn, raw_json_str, now_iso)
        cursor.execute(sql, params)
        count += 1

    conn.commit()
    cursor.close()
    conn.close()
    return count


def save_sccm_collections(collections_list: List[Dict[str, Any]]) -> int:
    """Salva ou atualiza a lista de coleções do SCCM."""
    if not collections_list:
        return 0

    setup_sccm_tables()
    conn = get_connection()
    cursor = conn.cursor()
    is_pg = DB_TYPE in ["postgres", "postgresql"]
    now_iso = datetime.now().isoformat()
    count = 0

    for c in collections_list:
        col_id = str(c.get("CollectionID") or c.get("collection_id") or "").strip()
        if not col_id:
            continue

        name = str(c.get("Name") or "").strip()
        col_type = "Dispositivos" if str(c.get("CollectionType", "2")) == "2" else "Usuários"
        member_count = int(c.get("MemberCount") or 0)
        comment = str(c.get("Comment") or "").strip()
        refresh = str(c.get("LastRefreshTime") or "").strip()

        if is_pg:
            sql = """
            INSERT INTO sccm_cache_collections (
                collection_id, name, collection_type, member_count,
                comment, last_refresh_time, updated_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (collection_id) DO UPDATE SET
                name = EXCLUDED.name,
                collection_type = EXCLUDED.collection_type,
                member_count = EXCLUDED.member_count,
                comment = EXCLUDED.comment,
                last_refresh_time = EXCLUDED.last_refresh_time,
                updated_at = EXCLUDED.updated_at;
            """
        else:
            sql = """
            INSERT OR REPLACE INTO sccm_cache_collections (
                collection_id, name, collection_type, member_count,
                comment, last_refresh_time, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?);
            """

        params = (col_id, name, col_type, member_count, comment, refresh, now_iso)
        cursor.execute(sql, params)
        count += 1

    conn.commit()
    cursor.close()
    conn.close()
    return count


def get_sccm_devices_df(search_text: str = "", active_only: bool = False) -> pd.DataFrame:
    """Recupera os computadores/dispositivos do SCCM cacheados localmente."""
    setup_sccm_tables()
    conn = get_connection()
    sql = "SELECT * FROM sccm_cache_devices WHERE 1=1"
    params = []

    if active_only:
        sql += " AND client_active = 1"

    if search_text:
        s = f"%{search_text.strip()}%"
        sql += " AND (name LIKE ? OR last_logon_user LIKE ? OR ip_addresses LIKE ? OR model LIKE ?)"
        params.extend([s, s, s, s])

    sql += " ORDER BY name ASC"

    try:
        if DB_TYPE in ["postgres", "postgresql"]:
            sql = sql.replace("?", "%s")
            df = pd.read_sql_query(sql, conn, params=params)
        else:
            df = pd.read_sql_query(sql, conn, params=params)
        return df
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()


def get_sccm_users_df(search_text: str = "") -> pd.DataFrame:
    """Recupera os usuários do SCCM cacheados localmente."""
    setup_sccm_tables()
    conn = get_connection()
    sql = "SELECT * FROM sccm_cache_users WHERE 1=1"
    params = []

    if search_text:
        s = f"%{search_text.strip()}%"
        sql += " AND (user_name LIKE ? OR full_user_name LIKE ? OR distinguished_name LIKE ?)"
        params.extend([s, s, s])

    sql += " ORDER BY user_name ASC"

    try:
        if DB_TYPE in ["postgres", "postgresql"]:
            sql = sql.replace("?", "%s")
            df = pd.read_sql_query(sql, conn, params=params)
        else:
            df = pd.read_sql_query(sql, conn, params=params)
        return df
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()


def get_sccm_collections_df(col_type: str = "") -> pd.DataFrame:
    """Recupera as coleções do SCCM cacheadas localmente."""
    setup_sccm_tables()
    conn = get_connection()
    sql = "SELECT * FROM sccm_cache_collections WHERE 1=1"
    params = []

    if col_type:
        sql += " AND collection_type = ?"
        params.append(col_type)

    sql += " ORDER BY name ASC"

    try:
        if DB_TYPE in ["postgres", "postgresql"]:
            sql = sql.replace("?", "%s")
            df = pd.read_sql_query(sql, conn, params=params)
        else:
            df = pd.read_sql_query(sql, conn, params=params)
        return df
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()


def get_device_by_user(username: str) -> Optional[Dict[str, str]]:
    """
    Busca no cache relacional do SCCM (sccm_cache_devices) o dispositivo associado ao usuário.
    Tenta correspondência exata e case-insensitive por last_logon_user.
    Retorna um dicionário com {'ip': ip, 'hostname': name} ou None se não encontrado.
    """
    if not username:
        return None
        
    u_clean = str(username).strip()
    if not u_clean or u_clean.lower() in ["none", "nan", "null", ""]:
        return None

    # Normaliza se vier com domínio (ex: MPE\usuario ou usuario@mpms.mp.br)
    if "\\" in u_clean:
        u_clean = u_clean.split("\\")[-1].strip()
    elif "@" in u_clean:
        u_clean = u_clean.split("@")[0].strip()

    setup_sccm_tables()
    conn = get_connection()
    cursor = conn.cursor()
    
    is_pg = DB_TYPE in ["postgres", "postgresql"]
    placeholder = "%s" if is_pg else "?"
    
    # Prioriza dispositivos ativos primeiro, ordenados por updated_at / last_active_time desc
    sql = f"""
    SELECT name, ip_addresses 
    FROM sccm_cache_devices 
    WHERE LOWER(TRIM(last_logon_user)) = LOWER(TRIM({placeholder}))
    ORDER BY client_active DESC, updated_at DESC
    LIMIT 1
    """
    
    device_info = None
    try:
        cursor.execute(sql, (u_clean,))
        row = cursor.fetchone()
        
        # Se não encontrou por igualdade estrita, tenta match com LIKE
        if not row:
            sql_like = f"""
            SELECT name, ip_addresses 
            FROM sccm_cache_devices 
            WHERE LOWER(last_logon_user) LIKE LOWER({placeholder})
            ORDER BY client_active DESC, updated_at DESC
            LIMIT 1
            """
            cursor.execute(sql_like, (f"%{u_clean}%",))
            row = cursor.fetchone()
            
        if row:
            name, ip_addresses = row[0], row[1]
            chosen_ip = ""
            if ip_addresses:
                # O campo pode conter IPs separados por vírgula ou JSON
                raw_ips = [ip.strip() for ip in str(ip_addresses).replace("[", "").replace("]", "").replace('"', '').replace("'", "").split(",") if ip.strip()]
                # Prioriza IP da rede interna (10.x)
                ip_10 = next((ip for ip in raw_ips if ip.startswith("10.")), None)
                if ip_10:
                    chosen_ip = ip_10
                elif raw_ips:
                    chosen_ip = raw_ips[0]
                    
            device_info = {
                "ip": chosen_ip or "",
                "hostname": str(name).strip() if name else ""
            }
    except Exception as e:
        logger = logging.getLogger(__name__) if "logging" in globals() else None
        if logger:
            logger.warning(f"Erro ao buscar dispositivo por usuário '{username}' no cache SCCM: {e}")
    finally:
        cursor.close()
        conn.close()
        
    return device_info

