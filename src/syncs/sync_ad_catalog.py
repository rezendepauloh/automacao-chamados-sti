# -*- coding: utf-8 -*-
"""
Worker assíncrono para sincronização do catálogo do Active Directory (OUs, Usuários, Computadores e Grupos)
via protocolo LDAP corporativo com persistência no banco SQLite local.
"""
import os
import sys
import tempfile
import time
from pathlib import Path

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from src.config import setup_logging, DEBUG_DIR_AD
from src.components.status_banner import check_process_running, read_log_lines
from src.terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_AD / "sync_ad.log", "sync_ad")

LOCK_FILE = Path(tempfile.gettempdir()) / "ad_sync.lock"
LOG_FILE = DEBUG_DIR_AD / "sync_ad.log"


def check_ad_sync_running() -> bool:
    """Verifica se o worker de sincronização do Active Directory está em execução."""
    return check_process_running(LOCK_FILE)


def read_ad_last_log_lines(n: int = 15) -> str:
    """Retorna as últimas N linhas do log de sincronização do AD."""
    return read_log_lines(LOG_FILE, n)


def run_ad_sync() -> dict:
    """Executa a sincronização completa do Active Directory via LDAP com banner no terminal e log."""
    print_header("WORKER - SINCRONIZAÇÃO DO ACTIVE DIRECTORY (LDAP)", color=CYAN)
    logger.info("Iniciando consulta completa ao Domain Controller (OUs, Usuários, Computadores e Grupos)...")

    from src.services.ad_ldap_service import sync_active_directory_cache
    start_time = time.time()
    
    def log_progress(pct: int, msg: str):
        logger.info(f"[{pct}%] {msg}")

    res = sync_active_directory_cache(page_size=500, progress_callback=log_progress)
    elapsed = time.time() - start_time

    if res.get("success"):
        logger.info(
            f"✅ Sincronização do Active Directory concluída com sucesso em {elapsed:.2f}s: "
            f"{res.get('total_ous', 0)} OUs, {res.get('total_users', 0)} Usuários, "
            f"{res.get('total_computers', 0)} Computadores, {res.get('total_groups', 0)} Grupos atualizados no cache."
        )
    else:
        logger.error(f"❌ Falha na sincronização do Active Directory após {elapsed:.2f}s: {res.get('error')}")

    return res


if __name__ == "__main__":
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_ad_sync()
    finally:
        if LOCK_FILE.exists():
            try:
                LOCK_FILE.unlink()
            except Exception:
                pass
