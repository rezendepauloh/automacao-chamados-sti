import tempfile
import ctypes
from pathlib import Path
import streamlit as st

def check_process_running(lock_file: Path) -> bool:
    """Verifica se um processo está rodando de forma ativa analisando o arquivo de lock no Windows."""
    if not lock_file.exists():
        return False
        
    try:
        with open(lock_file, "r") as f:
            pid = int(f.read().strip())
        
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


def read_log_lines(log_path: Path, n: int = 15) -> str:
    """Lê as últimas N linhas de um arquivo de log arbitrário."""
    if not log_path.exists():
        return "Nenhum log gerado ainda. Aguardando início..."
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
            return "".join(lines[-n:])
    except Exception as e:
        return f"Erro ao ler arquivo de log: {e}"


def check_orquestrador_running() -> bool:
    """Retrocompatibilidade: Verifica se o orquestrador principal está rodando."""
    lock_file = Path(tempfile.gettempdir()) / "automated_otrs_citsmart.lock"
    return check_process_running(lock_file)


def read_last_log_lines(n: int = 15) -> str:
    """Retrocompatibilidade: Lê as últimas N linhas do log do orquestrador."""
    log_path = Path("debug_logs") / "orquestrador" / "orquestrador.log"
    return read_log_lines(log_path, n)


def render_log_expander(title: str, is_running: bool, read_log_func, check_func, info_text: str):
    """Renderiza um accordion de log que se atualiza sozinho a cada 3 segundos e auto-encerra quando a checagem retorna False."""
    if not is_running:
        return

    with st.expander(title, expanded=False):
        st.info(info_text)

        @st.fragment(run_every="3s")
        def show_logs():
            if check_func and not check_func():
                st.rerun()

            logs = read_log_func(15)
            st.code(logs, language="text")
            st.button("🔄 Atualizar Progresso Manualmente", key=f"btn_refresh_{title}")

        show_logs()
