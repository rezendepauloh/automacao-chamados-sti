import sys
import tempfile
import os
from pathlib import Path

# Garante importações dos módulos do projeto
root_dir = Path(__file__).parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(root_dir / "src") not in sys.path:
    sys.path.insert(0, str(root_dir / "src"))

from src.components.status_banner import check_process_running, read_log_lines
from src.config import DONATIONS_FILE_PATH, setup_logging, DEBUG_DIR_DONATIONS
from src.database import sync_donations_from_excel

logger = setup_logging(DEBUG_DIR_DONATIONS / "donations.log", "sync_donations")


def check_donations_sync_running() -> bool:
    """Verifica se o processo de sincronização de doações está em execução."""
    lock_file = Path(tempfile.gettempdir()) / "donations_sync.lock"
    return check_process_running(lock_file)


def read_donations_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log de doações."""
    log_path = DEBUG_DIR_DONATIONS / "donations.log"
    return read_log_lines(log_path, n)


def run_donations_sync():
    """Executa a sincronização da planilha de doações."""
    logger.info("Iniciando sincronização de doações em segundo plano...")
    try:
        sync_donations_from_excel(str(DONATIONS_FILE_PATH))
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
