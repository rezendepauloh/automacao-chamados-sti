import subprocess
import os
import sys
import time
import tempfile
import ctypes
from pathlib import Path
from datetime import datetime
from config import setup_logging, LOG_FILE_ORQUESTRADOR

# CREATE_NO_WINDOW = 0x08000000 é um comando da API do Windows que proíbe a criação de janelas de terminal para subprocessos
CREATE_NO_WINDOW = 0x08000000

python_exe = os.path.join("venv", "Scripts", "pythonw.exe")

# Inicializa o logging usando a biblioteca central unificada com proteção de encoding
logger = setup_logging(LOG_FILE_ORQUESTRADOR, "ORQUESTRADOR")

LOCK_FILE = Path(tempfile.gettempdir()) / "automated_otrs_citsmart.lock"

def is_pid_running(pid: int) -> bool:
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if handle:
        exit_code = ctypes.c_ulong()
        if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
            kernel32.CloseHandle(handle)
            return exit_code.value == 259  # 259 representa STILL_ACTIVE no Windows
        kernel32.CloseHandle(handle)
    return False

def acquire_lock():
    if LOCK_FILE.exists():
        try:
            with open(LOCK_FILE, "r") as f:
                old_pid = int(f.read().strip())
            if is_pid_running(old_pid):
                return False, old_pid
        except Exception:
            pass  # Se houver erro ao ler, trata como lock orfão ou corrompido e sobrescreve
            
    try:
        LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
        LOCK_FILE.write_text(str(os.getpid()))
        return True, os.getpid()
    except Exception:
        return False, None

def release_lock():
    try:
        if LOCK_FILE.exists():
            LOCK_FILE.unlink()
    except Exception:
        pass

def format_duration(seconds: float) -> str:
    mins, secs = divmod(int(seconds), 60)
    hours, mins = divmod(mins, 60)
    return f"{hours:02d}:{mins:02d}:{secs:02d}"

def run_script(script_name: str) -> subprocess.CompletedProcess:
    # Executa o sub-processo de forma oculta capturando a saída
    return subprocess.run(
        [python_exe, script_name],
        creationflags=CREATE_NO_WINDOW,
        capture_output=True,
        text=True
    )

def main():
    logger.info("=============================================================")
    logger.info("INICIANDO ORQUESTRAÇÃO DOS ROBÔS")
    logger.info("=============================================================")

    # 1. Tenta obter o Lock para evitar colisões de agendamento do Windows Task Scheduler
    acquired, active_pid = acquire_lock()
    if not acquired:
        logger.info(f"⚠️ ABORTOU: Já existe uma instância do robô ativa (PID: {active_pid}).")
        logger.info("Aguardando finalização do processo anterior para evitar colisões de arquivos e travas de planilhas.")
        sys.exit(1)

    scripts = [
        ("otrs_scraper.py", "Coleta OTRS"),
        ("citsmart_scraper.py", "Coleta CitSmart"),
        ("preprocess_chamados.py", "Pré-processamento"),
        ("tag_classifier.py", "Classificação de TAGs por IA"),
        ("sync_master.py", "Sincronização Planilha Master")
    ]

    reports = []
    total_start = time.time()
    all_ok = True

    try:
        for idx, (script, desc) in enumerate(scripts, 1):
            logger.info(f"[{idx}/{len(scripts)}] Executando {script} ({desc})...")
            start_time = time.time()
            
            # Roda o subprocesso de forma totalmente invisível
            res = run_script(script)
            
            duration = time.time() - start_time
            dur_str = format_duration(duration)
            
            status = "SUCESSO" if res.returncode == 0 else "FALHA"
            reports.append({
                "passo": f"{idx}. {script}",
                "status": status,
                "duracao": dur_str,
                "retorno": res.returncode
            })
            
            if res.returncode == 0:
                logger.info(f"   └─ Concluído com SUCESSO em {dur_str}.")
            else:
                logger.info(f"   └─ ❌ FALHA com código de retorno {res.returncode} em {dur_str}.")
                if res.stderr:
                    logger.info(f"      Erro capturado:\n{res.stderr.strip()}")
                
                # Fail-fast: interrompe imediatamente no primeiro erro crítico para não estragar as planilhas
                logger.info("🛑 Sequência abortada para proteger a integridade e segurança dos dados.")
                all_ok = False
                break

    except Exception as e:
        logger.info(f"💥 Erro fatal inesperado no orquestrador: {e}")
        all_ok = False

    finally:
        total_duration = format_duration(time.time() - total_start)
        release_lock()

        # Monta a tabela ASCII final de relatório
        logger.info("=============================================================")
        logger.info("                      PAINEL DE EXECUÇÃO                      ")
        logger.info("=============================================================")
        logger.info(f"{'Passo / Script':<30} | {'Status':<8} | {'Duração':<8} | {'Retorno':<7}")
        logger.info("-" * 61)
        for rep in reports:
            logger.info(f"{rep['passo']:<30} | {rep['status']:<8} | {rep['duracao']:<8} | {rep['retorno']:<7}")
        logger.info("-" * 61)
        status_final = "CONCLUÍDO COM SUCESSO!" if all_ok else "CONCLUÍDO COM ERROS!"
        logger.info(f"Tempo Total: {total_duration:<18} | Status: {status_final}")
        logger.info("=============================================================\n")

        if not all_ok:
            sys.exit(1)
        else:
            sys.exit(0)

if __name__ == "__main__":
    main()