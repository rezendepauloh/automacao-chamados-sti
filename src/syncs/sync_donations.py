import sys
import tempfile
import os
from pathlib import Path

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from dotenv import load_dotenv
load_dotenv(override=True)

import time
import requests
from src.components.status_banner import check_process_running, read_log_lines
from src.config import (
    DONATIONS_FILE_PATH, setup_logging, DEBUG_DIR_DONATIONS,
    CITSMART_EMAIL, PASSWORD, HEADLESS, EXPLICIT_WAIT, get_chrome_driver,
    get_spreadsheet_setting, is_sharepoint_url, resolve_spreadsheet_path
)
from src.database import sync_donations_from_excel
from terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_DONATIONS / "donations.log", "sync_donations")

def check_donations_sync_running() -> bool:
    """Verifica se o processo de sincronização de doações está em execução."""
    lock_file = Path(tempfile.gettempdir()) / "donations_sync.lock"
    return check_process_running(lock_file)

def read_donations_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log de doações."""
    log_path = DEBUG_DIR_DONATIONS / "donations.log"
    return read_log_lines(log_path, n)

def download_sharepoint_donations_file(url: str) -> Path | None:
    """Faz o download da planilha de doações a partir da URL do SharePoint via cache rápido/Playwright."""
    from src.scrapers.sharepoint_downloader import download_sharepoint_file
    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "doacoes_redistribuicao.xlsx"
    return download_sharepoint_file(url, file_path)

def run_donations_sync():
    """Executa a sincronização da planilha de doações."""
    print_header("WORKER - SINCRONIZAÇÃO DE DOAÇÕES", color=CYAN)
    logger.info("Iniciando sincronização de doações em segundo plano...")
    excel_path_env = get_spreadsheet_setting("DONATIONS_EXCEL_RELATIVE_PATH", "").strip()
    target_file = None

    if is_sharepoint_url(excel_path_env):
        target_file = download_sharepoint_donations_file(excel_path_env)
    else:
        resolved = resolve_spreadsheet_path("DONATIONS_EXCEL_RELATIVE_PATH", "doacoes_redistribuicao.xlsx")
        if resolved.exists():
            target_file = resolved
        elif DONATIONS_FILE_PATH.exists():
            target_file = DONATIONS_FILE_PATH

    if not target_file or not Path(target_file).exists():
        logger.error(f"Planilha de Doações não localizada nem baixada (Configuração: {excel_path_env})")
        return

    try:
        sync_donations_from_excel(str(target_file))
        logger.info("Sincronização de doações finalizada com sucesso!")
    except Exception as e:
        logger.error(f"Erro durante a sincronização de doações: {e}")
        raise e

if __name__ == "__main__":
    lock_path = Path(tempfile.gettempdir()) / "donations_sync.lock"
    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_donations_sync()
    finally:
        if lock_path.exists():
            try:
                lock_path.unlink()
            except Exception:
                pass
