import os
import re
import json
import logging
from datetime import datetime
import pandas as pd
from typing import List, Dict, Any, Optional
from .connection import get_connection, DB_TYPE

logger = logging.getLogger(__name__)

# ==============================================================================
# DICIONÁRIO DE ALIASES DE MODELOS DE HARDWARE
# ==============================================================================
# Mapeia códigos de fábrica (Machine Types / MTM da Lenovo, Dell, HP, etc.)
# para os nomes comerciais legíveis aos analistas.
# Se surgirem novos modelos, basta adicionar novas entradas nesta lista.
HARDWARE_MODEL_ALIASES: Dict[str, str] = {
    # Lenovo ThinkCentre
    "11DUSD3R00": "Lenovo ThinkCentre M70q Gen 2",
    "11DU": "Lenovo ThinkCentre M70q Gen 2",
    "M11DU": "Lenovo ThinkCentre M70q Gen 2",
    "12TES8R800": "Lenovo ThinkCentre M70q Gen 5",
    "12TE": "Lenovo ThinkCentre M70q Gen 5",
    "M12TES8R80": "Lenovo ThinkCentre M70q Gen 5",
    "M12TE": "Lenovo ThinkCentre M70q Gen 5",
    "10MUS09A00": "Lenovo ThinkCentre M710q",
    "10T7": "Lenovo ThinkCentre M720q",
    "11DT": "Lenovo ThinkCentre M70q Gen 1",
    "11T3": "Lenovo ThinkCentre M70q Gen 3",
    "12E3": "Lenovo ThinkCentre M70q Gen 4",
    "13E0S00400": "ThinkCentre M75q Gen 2 Tiny",
    # Lenovo ThinkPad
    "2349K9P": "Lenovo ThinkPad T430",
    "20AWA213BR": "Lenovo ThinkPad T440p",
    "20U9": "Lenovo ThinkPad X13 Gen 1",
    "20W0": "Lenovo ThinkPad T14 Gen 2",
    "21AH": "Lenovo ThinkPad T14 Gen 3",
    "20W1S6CB00": "Lenovo ThinkPad T14 Gen 1",
    "20W1S76U00": "Lenovo ThinkPad T14 Gen 1",
    "21M4000MBO": "Lenovo ThinkPad E14 Gen 1",
    "21JSS3FB00": "Lenovo ThinkPad E14 Gen 1",
    # Dell OptiPlex & Latitude (exemplos comuns)
    "0P6G5X": "Dell OptiPlex 7090",
}


def normalize_hardware_model(raw_model: str) -> str:
    """
    Retorna o nome comercial amigável correspondente ao modelo/código MTM de hardware.
    Caso não haja alias cadastrado, retorna a string original limpa.
    """
    if not raw_model:
        return ""
    cleaned = str(raw_model).strip()
    if not cleaned or cleaned.lower() in ["unknown", "none", "nan", "null"]:
        return ""

    # 1. Match exato
    if cleaned in HARDWARE_MODEL_ALIASES:
        return HARDWARE_MODEL_ALIASES[cleaned]

    # 2. Match por prefixo (ex: MTM de 4 caracteres iniciais da Lenovo como 11DU, 12TE)
    cleaned_upper = cleaned.upper()
    for code, alias in HARDWARE_MODEL_ALIASES.items():
        if cleaned_upper.startswith(code.upper()):
            return alias

    return cleaned


def extract_windows_build(os_version: str) -> Optional[int]:
    """
    Extrai o número inteiro da Build do Windows a partir de formatos como:
    '10.0.26100', '10.0.22631.3880', '22631', '19045', etc.
    """
    if not os_version:
        return None
    raw = str(os_version).strip()
    if not raw or raw.lower() in ["none", "nan", "null"]:
        return None

    # Tenta padrão semântico padrão Windows (Major.Minor.Build[.Revision])
    parts = raw.split(".")
    if len(parts) >= 3:
        try:
            return int(parts[2])
        except ValueError:
            pass

    # Tenta parte a parte do final para o início
    for p in reversed(parts):
        try:
            val = int(p)
            if val >= 1000:
                return val
        except ValueError:
            pass

    # Regex para capturar sequências de 4 a 5 dígitos
    matches = re.findall(r"\b(\d{4,5})\b", raw)
    if matches:
        try:
            return int(matches[0])
        except ValueError:
            pass

    return None


def normalize_os_name(os_name: str, os_version: str = "") -> str:
    """
    Normaliza o nome do Sistema Operacional reportado pelo SCCM (SMS_R_System).
    No SCCM, Windows 10 e Windows 11 frequentemente reportam 'Microsoft Windows NT Workstation 10.0'
    devido ao kernel compartilhado NT 10.0.
    
    Diferenciação técnica por número de Build (os_version):
      - Build >= 22000 (ex: 22000, 22621, 22631, 26100, 26200) -> 'Windows 11'
      - Build < 22000 (ex: 19045, 19044, 18363) e > 0 -> 'Windows 10'
      - Servidores ('Server' no nome) -> 'Windows Server 2022', 'Windows Server 2019', etc.
    """
    raw_name = str(os_name or "").strip()
    raw_ver = str(os_version or "").strip()

    if not raw_name and not raw_ver:
        return "Não Identificado"

    build = extract_windows_build(raw_ver)
    lower_name = raw_name.lower()

    # 1. Servidores
    if "server" in lower_name:
        if build:
            if build >= 26100:
                return "Windows Server 2025"
            elif build >= 20348:
                return "Windows Server 2022"
            elif build >= 17763:
                return "Windows Server 2019"
            elif build >= 14393:
                return "Windows Server 2016"
        return "Windows Server"

    # 2. Se já vier explicitamente como Windows 11
    if "windows 11" in lower_name:
        return raw_name.replace("Microsoft ", "").strip()

    # 3. Workstation NT 10.0 ou genérico Windows 10
    if "windows nt" in lower_name or "windows 10" in lower_name or "workstation" in lower_name:
        if build is not None:
            if build >= 22000:
                return "Windows 11"
            elif build > 0:
                return "Windows 10"
        # Sem build conhecida
        if "10.0" in raw_name or "windows 10" in lower_name:
            return "Windows 10"
        return "Windows"

    # 4. Caso tenha vindo sem nome mas com versão de build
    if not raw_name and build:
        if build >= 22000:
            return "Windows 11"
        elif build > 0:
            return "Windows 10"

    if lower_name in ["unknown unknown", "unknown", "nan", "null"]:
        return "Não Identificado"

    return raw_name


def migrate_sccm_os_normalization() -> int:
    """
    Migra e normaliza os registros existentes em sccm_cache_devices para refletir
    corretamente Windows 11, Windows 10 e Windows Server a partir da Build.
    Executa de forma rápida e idempotente.
    """
    conn = get_connection()
    cursor = conn.cursor()
    is_pg = DB_TYPE in ["postgres", "postgresql"]

    updated_count = 0
    try:
        # Checagem ultrarrápida: se não há registros com 'Windows NT' ou 'unknown', não precisa migrar
        cursor.execute("SELECT 1 FROM sccm_cache_devices WHERE operating_system LIKE '%Windows NT%' OR operating_system LIKE '%unknown%' LIMIT 1")
        if not cursor.fetchone():
            return 0

        cursor.execute("SELECT resource_id, operating_system, os_version FROM sccm_cache_devices")
        rows = cursor.fetchall()
        updates = []
        for res_id, raw_os, raw_ver in rows:
            norm_os = normalize_os_name(raw_os, raw_ver)
            if norm_os != (raw_os or ""):
                updates.append((norm_os, res_id))

        if updates:
            sql_up = "UPDATE sccm_cache_devices SET operating_system = %s WHERE resource_id = %s" if is_pg else "UPDATE sccm_cache_devices SET operating_system = ? WHERE resource_id = ?"
            cursor.executemany(sql_up, updates)
            conn.commit()
            updated_count = len(updates)
            logger.info(f"✨ [SCCM NORMALIZAÇÃO] {updated_count} dispositivos atualizados no cache relacional.")
    except Exception as e:
        logger.warning(f"Aviso na rotina de migração de SO do SCCM: {e}")
    finally:
        cursor.close()
        conn.close()

    return updated_count


def migrate_sccm_model_aliases() -> int:
    """
    Normaliza modelos que possuem código de fábrica (MTM/Product ID) para seus aliases
    comerciais conhecidos (ex: 11DUSD3R00 -> Lenovo ThinkCentre M70q Gen 2).
    Executa de forma rápida e idempotente.
    """
    conn = get_connection()
    cursor = conn.cursor()
    is_pg = DB_TYPE in ["postgres", "postgresql"]

    updated_count = 0
    try:
        cursor.execute("SELECT resource_id, model FROM sccm_cache_devices WHERE model IS NOT NULL AND model != ''")
        rows = cursor.fetchall()
        updates = []
        for res_id, raw_model in rows:
            norm_mod = normalize_hardware_model(raw_model)
            if norm_mod and norm_mod != raw_model:
                updates.append((norm_mod, res_id))

        if updates:
            sql_up = "UPDATE sccm_cache_devices SET model = %s WHERE resource_id = %s" if is_pg else "UPDATE sccm_cache_devices SET model = ? WHERE resource_id = ?"
            cursor.executemany(sql_up, updates)
            conn.commit()
            updated_count = len(updates)
            logger.info(f"✨ [SCCM ALIASES] {updated_count} modelos de computadores atualizados para nomes comerciais amigáveis.")
    except Exception as e:
        logger.warning(f"Aviso na rotina de migração de modelos do SCCM: {e}")
    finally:
        cursor.close()
        conn.close()

    return updated_count


def setup_sccm_tables():
    """
    Cria as tabelas de cache relacional do SCCM (Dispositivos, Usuários e Coleções).
    Compatível com SQLite e PostgreSQL.
    """
    conn = get_connection()
    cursor = conn.cursor()

    is_pg = DB_TYPE in ["postgres", "postgresql"]
    id_pk = "SERIAL PRIMARY KEY" if is_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"

    # 1. Tabela de Dispositivos (SMS_R_System + Hardware)
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
        processor TEXT,
        memory_ram TEXT,
        disk_drives TEXT,
        raw_json TEXT,
        updated_at TEXT
    )
    """)

    # Garante a existência das novas colunas de hardware em bancos preexistentes
    if is_pg:
        for col in ["processor", "memory_ram", "disk_drives"]:
            try:
                cursor.execute(f"ALTER TABLE sccm_cache_devices ADD COLUMN IF NOT EXISTS {col} TEXT;")
            except Exception:
                pass
    else:
        try:
            cursor.execute("PRAGMA table_info(sccm_cache_devices)")
            existing_cols = [c[1] for c in cursor.fetchall()]
            for col in ["processor", "memory_ram", "disk_drives"]:
                if col not in existing_cols:
                    cursor.execute(f"ALTER TABLE sccm_cache_devices ADD COLUMN {col} TEXT;")
        except Exception:
            pass

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

    # Normaliza dispositivos já persistidos se necessário
    migrate_sccm_os_normalization()
    migrate_sccm_model_aliases()


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
        raw_model = str(dev.get("Model") or dev.get("model") or "").strip()
        model = normalize_hardware_model(raw_model)
        raw_os_name = str(dev.get("OperatingSystemNameandVersion") or dev.get("operating_system") or "").strip()
        os_ver = str(dev.get("Build") or dev.get("os_version") or "").strip()
        os_name = normalize_os_name(raw_os_name, os_ver)
        client_ver = str(dev.get("ClientVersion") or dev.get("client_version") or "").strip()
        
        # Propriedades de Hardware adicionais
        processor = str(dev.get("Processor") or dev.get("processor") or dev.get("CPU") or "").strip()
        
        raw_mem = dev.get("MemoryRAM") or dev.get("memory_ram") or dev.get("TotalPhysicalMemory") or ""
        if isinstance(raw_mem, (int, float)) and raw_mem > 0:
            if raw_mem > 1024 * 1024 * 1024:
                memory_ram = f"{round(raw_mem / (1024**3))} GB"
            elif raw_mem > 1024 * 1024:
                memory_ram = f"{round(raw_mem / (1024**2))} GB"
            else:
                memory_ram = f"{raw_mem} MB"
        else:
            memory_ram = str(raw_mem).strip()

        disk_drives = str(dev.get("DiskDrives") or dev.get("disk_drives") or dev.get("Disks") or "").strip()

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
                processor, memory_ram, disk_drives, raw_json, updated_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
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
                processor = EXCLUDED.processor,
                memory_ram = EXCLUDED.memory_ram,
                disk_drives = EXCLUDED.disk_drives,
                raw_json = EXCLUDED.raw_json,
                updated_at = EXCLUDED.updated_at;
            """
        else:
            sql = """
            INSERT OR REPLACE INTO sccm_cache_devices (
                resource_id, name, last_logon_user, ip_addresses, mac_addresses,
                manufacturer, model, operating_system, os_version, client_version,
                client_active, last_active_time, ad_site_name, distinguished_name,
                processor, memory_ram, disk_drives, raw_json, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """

        params = (
            res_id, name, user, ip_str, mac_str,
            manufacturer, model, os_name, os_ver, client_ver,
            client_active, last_active, ad_site, dn,
            processor, memory_ram, disk_drives,
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


def get_sccm_devices_df(
    search_text: str = "",
    active_only: bool = False,
    os_filter: str = "",
    model_filter: str = ""
) -> pd.DataFrame:
    """Recupera os computadores/dispositivos do SCCM cacheados localmente."""
    setup_sccm_tables()
    conn = get_connection()
    sql = "SELECT * FROM sccm_cache_devices WHERE 1=1"
    params = []

    if active_only:
        sql += " AND client_active = 1"

    if search_text:
        s = f"%{search_text.strip()}%"
        sql += " AND (name LIKE ? OR last_logon_user LIKE ? OR ip_addresses LIKE ? OR model LIKE ? OR manufacturer LIKE ? OR raw_json LIKE ?)"
        params.extend([s, s, s, s, s, s])

    if os_filter and os_filter != "Todos":
        sql += " AND operating_system LIKE ?"
        params.append(f"%{os_filter.strip()}%")

    if model_filter and model_filter != "Todos":
        sql += " AND model LIKE ?"
        params.append(f"%{model_filter.strip()}%")

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

