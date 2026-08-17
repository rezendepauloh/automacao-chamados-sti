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

from src.components.status_banner import check_process_running, read_log_lines
from src.config import setup_logging, DEBUG_DIR_FISCALIZACAO
from src.database import sync_fiscalizacao_from_excel
from terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_FISCALIZACAO / "sync.log", "fiscalizacao_sync")

def check_fiscalizacao_sync_running() -> bool:
    """Verifica se o processo de sincronização de fiscalização está em execução."""
    lock_file = Path(tempfile.gettempdir()) / "fiscalizacao_sync.lock"
    return check_process_running(lock_file)

def read_fiscalizacao_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log da fiscalização."""
    log_path = DEBUG_DIR_FISCALIZACAO / "sync.log"
    return read_log_lines(log_path, n)

def run_fiscalizacao_sync():
    """Executa a leitura da planilha do OneDrive/SharePoint e grava no SQLite."""
    print_header("WORKER - SINCRONIZAÇÃO DE FISCALIZAÇÃO", color=CYAN)
    logger.info("Iniciando sincronização de fiscalização em segundo plano...")
    relative_path = os.getenv("FISCAL_EXCEL_RELATIVE_PATH", "")
    excel_file = Path.home() / relative_path if relative_path else None

    if not excel_file or not excel_file.exists():
        logger.error(f"Planilha de Fiscais não localizada no caminho: {excel_file}")
        return

    try:
        sync_fiscalizacao_from_excel(str(excel_file))
        logger.info("Sincronização de fiscalização finalizada com sucesso!")
    except Exception as e:
        logger.error(f"Erro durante a sincronização de fiscalização: {e}")
        raise e

if __name__ == "__main__":
    lock_path = Path(tempfile.gettempdir()) / "fiscalizacao_sync.lock"
    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_fiscalizacao_sync()
    finally:
        if lock_path.exists():
            try:
                lock_path.unlink()
            except Exception:
                pass
