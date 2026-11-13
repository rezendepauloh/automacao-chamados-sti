# -*- coding: utf-8 -*-
"""
Worker assíncrono para sincronização da planilha de Previsão de Férias da Bancada
a partir do SharePoint Online ou arquivo local com persistência no SQLite.
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

from dotenv import load_dotenv
load_dotenv(override=True)

import requests
from src.components.status_banner import check_process_running, read_log_lines
from src.config import (
    setup_logging, DEBUG_DIR_FERIAS, FERIAS_EXCEL_RELATIVE_PATH,
    CITSMART_EMAIL, PASSWORD, HEADLESS, EXPLICIT_WAIT, get_chrome_driver
)
from src.database import sync_ferias_from_excel
from terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_FERIAS / "sync_ferias.log", "sync_ferias")

LOCK_FILE = Path(tempfile.gettempdir()) / "ferias_sync.lock"
LOG_FILE = DEBUG_DIR_FERIAS / "sync_ferias.log"


def check_ferias_sync_running() -> bool:
    """Verifica se o processo de sincronização de férias está em execução."""
    return check_process_running(LOCK_FILE)


def read_ferias_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log de férias."""
    return read_log_lines(LOG_FILE, n)


def download_sharepoint_ferias_file(url: str) -> Path | None:
    """Faz o download da planilha de férias a partir da URL do SharePoint via cache rápido/Playwright."""
    from src.scrapers.sharepoint_downloader import download_sharepoint_file
    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "Previsao_de_Ferias-Manutencao.xlsx"
    return download_sharepoint_file(url, file_path)


def run_ferias_sync() -> bool:
    """Executa a sincronização da planilha de férias."""
    print_header("WORKER - SINCRONIZAÇÃO DE FÉRIAS DA BANCADA", color=CYAN)
    logger.info("Iniciando rotina de sincronização de férias...")
    from src.config import get_spreadsheet_setting, is_sharepoint_url, resolve_spreadsheet_path
    excel_path_env = get_spreadsheet_setting("FERIAS_EXCEL_RELATIVE_PATH", "https://ministeriopublicoms.sharepoint.com/:x:/r/sites/dit-manutencao/_layouts/15/Doc.aspx?sourcedoc=%7BE197F2AD-7143-4E92-A56B-B049D930E4C5%7D&file=Previs%C3%A3o%20de%20F%C3%A9rias-Manutencao.xlsx&action=default&mobileredirect=true&wdwpf=doclib-t").strip()
    target_file = None

    if is_sharepoint_url(excel_path_env):
        target_file = download_sharepoint_ferias_file(excel_path_env)
    elif excel_path_env:
        resolved = resolve_spreadsheet_path("FERIAS_EXCEL_RELATIVE_PATH", "Previsao_de_Ferias-Manutencao.xlsx")
        if resolved.is_file():
            target_file = resolved

    # Fallbacks locais
    if not target_file or not Path(target_file).is_file():
        default_up = Path("uploads/Previsao_de_Ferias-Manutencao.xlsx")
        if default_up.is_file():
            target_file = default_up

    if not target_file or not Path(target_file).is_file():
        logger.error(f"Planilha de Férias não localizada nem baixada (Configuração: {excel_path_env})")
        return False

    try:
        success = sync_ferias_from_excel(str(target_file))
        if success:
            logger.info("✅ Sincronização de férias finalizada com sucesso!")
            return True
        return False
    except Exception as e:
        logger.error(f"Erro durante a sincronização de férias: {e}")
        raise e


if __name__ == "__main__":
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_ferias_sync()
    finally:
        if LOCK_FILE.exists():
            try:
                LOCK_FILE.unlink()
            except Exception:
                pass
