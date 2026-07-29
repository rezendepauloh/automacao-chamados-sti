import tempfile
import ctypes
from pathlib import Path

def check_orquestrador_running() -> bool:
    """Verifica se o orquestrador está rodando de forma ativa analisando o arquivo de lock no Windows."""
    lock_file = Path(tempfile.gettempdir()) / "automated_otrs_citsmart.lock"
    if not lock_file.exists():
        return False
        
    try:
        with open(lock_file, "r") as f:
            pid = int(f.read().strip())
        
        # Verifica se o processo com esse PID está ativo
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            exit_code = ctypes.c_ulong()
            if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                kernel32.CloseHandle(handle)
                return exit_code.value == 259  # 259 significa STILL_ACTIVE
            kernel32.CloseHandle(handle)
    except:
        pass
    return False

def read_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log do orquestrador."""
    log_path = Path("debug_logs") / "orquestrador" / "orquestrador.log"
    if not log_path.exists():
        return "Nenhum log gerado ainda. Aguardando início..."
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
            return "".join(lines[-n:])
    except Exception as e:
        return f"Erro ao ler arquivo de log: {e}"
