# -*- coding: utf-8 -*-
"""
Worker assíncrono para sincronização do catálogo de FAQs, Vídeos e Imagens
diretamente da API REST do SharePoint Online com persistência no SQLite.
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

from src.config import setup_logging, DEBUG_DIR_FAQ
from src.components.status_banner import check_process_running, read_log_lines
from src.terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_FAQ / "sync_sharepoint_catalog.log", "sync_sharepoint_catalog")

LOCK_FILE = Path(tempfile.gettempdir()) / "sharepoint_catalog_sync.lock"
LOG_FILE = DEBUG_DIR_FAQ / "sync_sharepoint_catalog.log"


def check_sharepoint_catalog_sync_running() -> bool:
    """Verifica se o worker de sincronização do SharePoint está em execução."""
    return check_process_running(LOCK_FILE)


def read_sharepoint_catalog_last_log_lines(n: int = 15) -> str:
    """Retorna as últimas N linhas do log de sincronização do catálogo SharePoint."""
    return read_log_lines(LOG_FILE, n)


def run_sharepoint_catalog_sync() -> dict:
    """Executa a sincronização do catálogo via API REST com cabeçalho estilizado no terminal e log."""
    print_header("WORKER - SINCRONIZAÇÃO DE CATÁLOGO SHAREPOINT", color=CYAN)
    logger.info("Iniciando verificação e sincronização de Artigos, Vídeos e Imagens do SharePoint...")

    from src.tabs.links_faqs import sync_sharepoint_catalog_via_api
    start_time = time.time()
    res = sync_sharepoint_catalog_via_api()
    elapsed = time.time() - start_time

    if res.get("success"):
        st_stats = res.get("stats", {})
        logger.info(
            f"✅ Sincronização do SharePoint concluída com sucesso em {elapsed:.2f}s: "
            f"{st_stats.get('faqs', 0)} artigos, {st_stats.get('videos', 0)} vídeos, {st_stats.get('imagens', 0)} imagens atualizados."
        )
    else:
        logger.error(f"❌ Falha na sincronização do SharePoint após {elapsed:.2f}s: {res.get('error')}")

    return res


if __name__ == "__main__":
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_sharepoint_catalog_sync()
    finally:
        if LOCK_FILE.exists():
            try:
                LOCK_FILE.unlink()
            except Exception:
                pass
