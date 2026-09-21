import os
import sys
import json
import time
import subprocess
from typing import List, Dict, Any, Optional

from src.config import _cfg, DOMINIO_CURTO, setup_logging, DEBUG_DIR_SYNC
from src.database.sccm_db import (
    save_sccm_devices,
    save_sccm_users,
    save_sccm_collections
)

logger = setup_logging(DEBUG_DIR_SYNC / "sccm_service.log", "sccm_service")

def _get_pwsh_executable() -> Optional[str]:
    """Localiza o executável do PowerShell Core (pwsh) no ambiente Linux/Docker ou powershell.exe."""
    for exe in ["pwsh", "/usr/bin/pwsh", "/opt/microsoft/powershell/7/pwsh", "powershell.exe"]:
        try:
            res = subprocess.run([exe, "-v"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=3)
            if res.returncode == 0:
                return exe
        except Exception:
            pass
    return None


def execute_sccm_cim_query(query: str, class_target: str = "") -> List[Dict[str, Any]]:
    """
    Executa uma consulta CIM/WMI no namespace root\\sms\\site_<SITE_CODE> do servidor SCCM
    via PowerShell 7 (pwsh) utilizando credenciais administrativas do cofre/banco.
    """
    site_server = _cfg("SCCM_SERVER", os.getenv("SCCM_SERVER", ""))
    site_code = _cfg("SCCM_SITE_CODE", os.getenv("SCCM_SITE_CODE", "PGJ"))
    admin_user = _cfg("SCCM_ADMIN_USER", os.getenv("SCCM_ADMIN_USER", "paulo_admin"))
    admin_password = _cfg("SCCM_ADMIN_PASSWORD", "")

    if not site_server or not site_code:
        logger.warning("⚠️ 'SCCM_SERVER' ou 'SCCM_SITE_CODE' não configurados.")
        return []

    pwsh_exe = _get_pwsh_executable()
    if not pwsh_exe:
        logger.error("❌ Executável do PowerShell 7 ('pwsh') não encontrado no contêiner.")
        return []

    domain_user = admin_user
    if "\\" not in domain_user and "@" not in domain_user:
        short_domain = DOMINIO_CURTO or "MPE"
        domain_user = f"{short_domain}\\{admin_user}"

    escaped_password = admin_password.replace('"', '`"').replace('$', '`$')
    escaped_query = query.replace('"', '`"')

    # Script PowerShell para execução via Get-CimInstance com conversão para JSON
    ps_script = f"""
    $secpasswd = ConvertTo-SecureString "{escaped_password}" -AsPlainText -Force;
    $mycreds = New-Object System.Management.Automation.PSCredential ("{domain_user}", $secpasswd);
    $opt = New-CimSessionOption -Protocol Dcom;
    $sess = New-CimSession -ComputerName "{site_server}" -Credential $mycreds -SessionOption $opt -ErrorAction SilentlyContinue;
    if (-not $sess) {{
        $sess = New-CimSession -ComputerName "{site_server}" -Credential $mycreds;
    }}
    $res = Get-CimInstance -CimSession $sess -Namespace "root\\sms\\site_{site_code}" -Query "{escaped_query}" -ErrorAction Stop;
    $res | ConvertTo-Json -Depth 3 -Compress
    """

    logger.info(f"🔍 Executando consulta SCCM CIM no servidor '{site_server}' (Namespace: root\\sms\\site_{site_code})...")
    t0 = time.time()
    try:
        proc = subprocess.run(
            [pwsh_exe, "-NoProfile", "-NonInteractive", "-Command", ps_script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120
        )
        dur = time.time() - t0

        if proc.returncode != 0:
            err = proc.stderr.strip()
            logger.error(f"❌ Erro na consulta SCCM CIM ({dur:.2f}s): {err}")
            return []

        stdout = proc.stdout.strip()
        if not stdout:
            logger.info("ℹ️ Nenhum registro retornado pelo SCCM.")
            return []

        data = json.loads(stdout)
        if isinstance(data, dict):
            return [data]
        elif isinstance(data, list):
            return data
        return []
    except json.JSONDecodeError as je:
        logger.error(f"❌ Falha ao decodificar JSON do SCCM: {je}")
        return []
    except subprocess.TimeoutExpired:
        logger.error("⏰ Timeout na consulta CIM do SCCM (120s atingido).")
        return []
    except Exception as e:
        logger.error(f"❌ Exceção ao consultar SCCM: {e}", exc_info=True)
        return []


def sync_sccm_devices() -> int:
    """Sincroniza todos os computadores/dispositivos do SCCM (SMS_R_System) com o banco local."""
    logger.info("🔄 Sincronizando Dispositivos (SMS_R_System)...")
    # Consulta otimizada com atributos essenciais de hardware, rede e sistema operacional
    query = "SELECT ResourceID, Name, LastLogonUserName, IPAddresses, MACAddresses, Manufacturer, Model, OperatingSystemNameandVersion, Build, ClientVersion, ClientActiveStatus, LastActiveTime, ADSiteName, DistinguishedName FROM SMS_R_System WHERE Obsolete = 0 AND Decommissioned = 0"
    results = execute_sccm_cim_query(query)
    count = save_sccm_devices(results)
    logger.info(f"✅ Sincronização de dispositivos concluída: {count} estações salvas no cache.")
    return count


def sync_sccm_users() -> int:
    """Sincroniza os usuários catalogados no SCCM (SMS_R_User)."""
    logger.info("🔄 Sincronizando Usuários (SMS_R_User)...")
    query = "SELECT ResourceID, UserName, FullUserName, UserGroupName, WindowsNTDomain, DistinguishedName FROM SMS_R_User"
    results = execute_sccm_cim_query(query)
    count = save_sccm_users(results)
    logger.info(f"✅ Sincronização de usuários SCCM concluída: {count} usuários salvos no cache.")
    return count


def sync_sccm_collections() -> int:
    """Sincroniza as coleções de dispositivos e usuários (SMS_Collection)."""
    logger.info("🔄 Sincronizando Coleções (SMS_Collection)...")
    query = "SELECT CollectionID, Name, CollectionType, MemberCount, Comment, LastRefreshTime FROM SMS_Collection"
    results = execute_sccm_cim_query(query)
    count = save_sccm_collections(results)
    logger.info(f"✅ Sincronização de coleções concluída: {count} coleções salvas no cache.")
    return count


def import_sccm_inventory_json(file_path: Optional[str] = None) -> Dict[str, int]:
    """Importa o arquivo JSON de inventário gerado pelo bancada-launcher.ps1."""
    candidates = []
    if file_path:
        candidates.append(file_path)

    # Locais padrão onde o script do Windows salva ou copia o arquivo
    candidates.extend([
        "/app/01 - Dados Brutos/sccm_inventory.json",
        "/home/paulo/PythonProjects/automacao-chamados-sti/01 - Dados Brutos/sccm_inventory.json",
        "/mnt/c/Users/paulogoncalves/sccm_inventory.json",
        os.path.expanduser("~/sccm_inventory.json")
    ])

    chosen_file = None
    for c in candidates:
        if c and os.path.exists(c):
            chosen_file = c
            break

    if not chosen_file:
        logger.warning("⚠️ Nenhum arquivo de inventário SCCM encontrado.")
        return {"devices": 0, "users": 0, "collections": 0}

    logger.info(f"📂 Importando inventário SCCM de '{chosen_file}'...")
    try:
        with open(chosen_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        devices = data.get("devices", [])
        users = data.get("users", [])
        collections = data.get("collections", [])

        c_dev = save_sccm_devices(devices)
        c_usr = save_sccm_users(users)
        c_col = save_sccm_collections(collections)

        logger.info(f"✅ Importação finalizada: {c_dev} estações, {c_usr} usuários, {c_col} coleções.")
        return {"devices": c_dev, "users": c_usr, "collections": c_col}
    except Exception as e:
        logger.error(f"❌ Falha ao importar JSON de inventário SCCM: {e}")
        return {"devices": 0, "users": 0, "collections": 0}


def sync_all_sccm() -> Dict[str, int]:
    """Executa a sincronização completa de todos os módulos do SCCM."""
    # Tenta primeiro a importação de arquivo local gerado pelo bancada:// se existir
    json_res = import_sccm_inventory_json()
    if json_res["devices"] > 0 or json_res["collections"] > 0:
        return json_res

    dev_count = sync_sccm_devices()
    user_count = sync_sccm_users()
    col_count = sync_sccm_collections()
    return {
        "devices": dev_count,
        "users": user_count,
        "collections": col_count
    }
