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
    VIAGENS_FILE_PATH, setup_logging, DEBUG_DIR_VIAGENS,
    CITSMART_EMAIL, PASSWORD, HEADLESS, EXPLICIT_WAIT, get_chrome_driver
)
from src.database import sync_viagens_from_excel
from terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_VIAGENS / "viagens_sync.log", "sync_viagens")

def check_viagens_sync_running() -> bool:
    """Verifica se o processo de sincronização de viagens está em execução."""
    lock_file = Path(tempfile.gettempdir()) / "viagens_sync.lock"
    return check_process_running(lock_file)

def read_viagens_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log de viagens."""
    log_path = DEBUG_DIR_VIAGENS / "viagens_sync.log"
    return read_log_lines(log_path, n)

def download_sharepoint_viagens_file(url: str) -> Path | None:
    """Faz o download da planilha de viagens a partir da URL do SharePoint via cache rápido/Playwright."""
    from src.scrapers.sharepoint_downloader import download_sharepoint_file
    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "planilha_viagens.xlsx"
    return download_sharepoint_file(url, file_path)

def run_viagens_sync():
    """Executa a sincronização da planilha de viagens."""
    print_header("WORKER - SINCRONIZAÇÃO DE VIAGENS DA BANCADA", color=CYAN)
    logger.info("Iniciando sincronização de viagens...")
    from src.config import get_spreadsheet_setting, is_sharepoint_url, resolve_spreadsheet_path
    excel_path_env = get_spreadsheet_setting("VIAGENS_EXCEL_RELATIVE_PATH", "https://ministeriopublicoms.sharepoint.com/:x:/s/dit-manutencao/IQBQonhEup-USL9RTQTk8W6vAS9SWFl5jhGXAOR_bezve5E?e=PYfupT").strip()
    target_file = None

    if is_sharepoint_url(excel_path_env):
        target_file = download_sharepoint_viagens_file(excel_path_env)
    elif excel_path_env:
        resolved = resolve_spreadsheet_path("VIAGENS_EXCEL_RELATIVE_PATH", "viagens_bancada.xlsx")
        if resolved.is_file():
            target_file = resolved
        elif VIAGENS_FILE_PATH.is_file():
            target_file = VIAGENS_FILE_PATH

    # Fallback para uploads/temp_viagens.xlsx se disponível
    if not target_file or not Path(target_file).is_file():
        temp_up = Path("uploads/temp_viagens.xlsx")
        if temp_up.is_file():
            target_file = temp_up

    # Fallback para caminho do OneDrive do usuário no Windows se no WSL
    if not target_file or not Path(target_file).is_file():
        onedrive_win = Path("/mnt/c/Users/paulogoncalves/OneDrive - Ministerio Público do Estado de Mato Grosso do Sul/Documentos SharePoint DIT-Manutenção/Viagens/Planilha de Viagens.xlsx")
        if onedrive_win.is_file():
            target_file = onedrive_win

    if not target_file or not Path(target_file).is_file():
        logger.error(f"Planilha de Viagens não localizada nem baixada (Configuração: {excel_path_env})")
        return False

    try:
        success = sync_viagens_from_excel(str(target_file))
        if success:
            logger.info("Sincronização de viagens finalizada com sucesso!")
            return True
        return False
    except Exception as e:
        logger.error(f"Erro durante a sincronização de viagens: {e}")
        raise e

if __name__ == "__main__":
    lock_path = Path(tempfile.gettempdir()) / "viagens_sync.lock"
    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_viagens_sync()
    finally:
        if lock_path.exists():
            try:
                lock_path.unlink()
            except Exception:
                pass
