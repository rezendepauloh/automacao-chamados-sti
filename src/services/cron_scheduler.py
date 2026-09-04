import os
import sys
import time
import threading
from pathlib import Path
from datetime import datetime, timedelta
import traceback

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from src.config import setup_logging, DEBUG_DIR_SYNC
from src.database import (
    setup_cron_tables,
    get_cron_schedules,
    get_cron_schedule_by_id,
    log_cron_execution_start,
    log_cron_execution_end
)

logger = setup_logging(DEBUG_DIR_SYNC / "cron_scheduler.log", "cron_scheduler")

# Mapeamento de executores das rotinas
def execute_task_by_id(task_id: str) -> str:
    """Invoca o worker correspondente à tarefa configurada."""
    import tempfile
    logger.info(f"🚀 Iniciando execução da tarefa: {task_id}")

    # Helper para criar e remover lock file de compatibilidade com a UI (apenas para rotinas que não gerenciam lock próprio)
    lock_name_map = {
        "sync_portarias": "portarias_sync.lock",
        "sync_viagens": "viagens_sync.lock",
        "sync_fiscalizacao": "fiscalizacao_sync.lock"
    }

    lock_file = None
    if task_id in lock_name_map:
        lock_file = Path(tempfile.gettempdir()) / lock_name_map[task_id]
        try:
            with open(lock_file, "w") as f:
                f.write(str(os.getpid()))
        except Exception:
            lock_file = None

    try:
        if task_id == "whatsapp_d1":
            from src.syncs.sync_whatsapp_scheduler import run_whatsapp_scheduler
            res = run_whatsapp_scheduler(dry_run=False, force=False)
            return f"Finalizado com sucesso: {res}"

        elif task_id == "sync_portarias":
            from src.syncs.sync_portarias import sync_portarias_and_generate_alerts
            sync_portarias_and_generate_alerts()
            return "Sincronização de portarias concluída com sucesso."

        elif task_id == "sync_viagens":
            from src.syncs.sync_viagens import run_viagens_sync
            ok = run_viagens_sync()
            return "Sincronização de viagens finalizada com sucesso." if ok else "Falha ou planilha não localizada."

        elif task_id == "sync_plantoes":
            from src.syncs.sync_plantoes_alerts import check_and_generate_plantao_alerts
            check_and_generate_plantao_alerts()
            return "Verificação de escalas de plantão concluída com sucesso."

        elif task_id == "sync_fiscalizacao":
            from src.syncs.sync_fiscalizacao import run_fiscalizacao_sync
            from src.syncs.sync_garantia import run_garantia_sync
            try:
                run_fiscalizacao_sync()
            except Exception as e1:
                logger.warning(f"Aviso na sincronização de fiscalização: {e1}")
            try:
                run_garantia_sync()
            except Exception as e2:
                logger.warning(f"Aviso na sincronização de garantia: {e2}")
            return "Sincronização de fiscalização e garantia concluída."

        elif task_id == "orquestrador_chamados":
            import subprocess
            proc = subprocess.run([sys.executable, "orquestrador.py"], capture_output=True, text=True)
            if proc.returncode != 0:
                raise RuntimeError(f"Orquestrador finalizou com código {proc.returncode}: {proc.stderr[:300] if proc.stderr else proc.stdout[:300]}")
            return "Orquestrador de chamados executado com sucesso."

        else:
            raise ValueError(f"Rotina desconhecida: {task_id}")
    finally:
        if lock_file and lock_file.exists():
            try:
                lock_file.unlink()
            except Exception:
                pass

class BancadaCronDaemon:
    """Motor de execução em segundo plano para tarefas periódicas da Bancada STI."""
    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(BancadaCronDaemon, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        self._running = False
        self._thread = None
        self._executing_tasks = set()
        logger.info("BancadaCronDaemon instanciado.")

    def start(self):
        """Inicia a thread do agendador em segundo plano."""
        with self._lock:
            if self._running:
                logger.info("BancadaCronDaemon já está em execução.")
                return
            self._running = True
            self._thread = threading.Thread(target=self._loop, name="BancadaCronDaemonThread", daemon=True)
            self._thread.start()
            logger.info("🟢 Thread do BancadaCronDaemon iniciada em background.")

    def stop(self):
        """Sinaliza parada do agendador."""
        with self._lock:
            self._running = False
            logger.info("🔴 Sinal de parada enviado ao BancadaCronDaemon.")

    def is_alive(self) -> bool:
        return self._running and self._thread is not None and self._thread.is_alive()

    def trigger_task_now(self, task_id: str) -> bool:
        """Dispara uma tarefa imediatamente em uma thread separada."""
        if task_id in self._executing_tasks:
            logger.warning(f"Tarefa {task_id} já está em execução.")
            return False
            
        t = threading.Thread(target=self._run_single_task, args=(task_id,), daemon=True)
        t.start()
        return True

    def _should_run(self, row: dict, now: datetime) -> bool:
        """Determina se uma rotina deve ser executada no instante atual."""
        if not bool(row.get("ativo")):
            return False

        task_id = str(row.get("task_id", ""))
        if task_id in self._executing_tasks:
            return False

        tipo = str(row.get("tipo_agendamento", "intervalo")).lower()
        apenas_dias_uteis = bool(row.get("apenas_dias_uteis", 0))

        # Checa dia útil (segunda=0 a sexta=4)
        if apenas_dias_uteis and now.weekday() > 4:
            return False

        ult_exec_str = row.get("ultima_execucao")
        ult_exec = None
        if ult_exec_str:
            try:
                ult_exec = datetime.strptime(str(ult_exec_str).split(".")[0], "%Y-%m-%d %H:%M:%S")
            except Exception:
                pass

        if tipo == "horario_fixo":
            target_time_str = str(row.get("horario_fixo", "12:00")).strip()
            try:
                target_hour, target_minute = map(int, target_time_str.split(":"))
            except Exception:
                target_hour, target_minute = 12, 0

            # Dispara se estamos no minuto exato ou até 3 minutos após
            if now.hour == target_hour and 0 <= (now.minute - target_minute) <= 3:
                # Se já executou hoje no mesmo horário, não roda de novo
                if ult_exec and ult_exec.date() == now.date() and ult_exec.hour == target_hour:
                    return False
                return True
            return False

        elif tipo == "intervalo":
            valor = int(row.get("intervalo_valor", 2) or 2)
            unidade = str(row.get("intervalo_unidade", "horas")).lower()

            if unidade == "minutos":
                delta = timedelta(minutes=valor)
            elif unidade == "dias":
                delta = timedelta(days=valor)
            else: # horas
                delta = timedelta(hours=valor)

            if not ult_exec:
                return True

            return (now - ult_exec) >= delta

        return False

    def _run_single_task(self, task_id: str):
        """Executa a tarefa, calcula a duração e registra logs e status."""
        self._executing_tasks.add(task_id)
        log_id = log_cron_execution_start(task_id)
        start_time = time.time()

        try:
            msg = execute_task_by_id(task_id)
            duracao = round(time.time() - start_time, 2)
            log_cron_execution_end(log_id, task_id, status="sucesso", mensagem=msg, duracao_seg=duracao)
            logger.info(f"✅ Tarefa {task_id} finalizada com sucesso em {duracao}s: {msg}")
        except BaseException as e:
            duracao = round(time.time() - start_time, 2)
            err_msg = f"Erro: {str(e)}\n{traceback.format_exc()}"
            log_cron_execution_end(log_id, task_id, status="erro", mensagem=err_msg[:500], duracao_seg=duracao)
            logger.error(f"❌ Falha na execução da tarefa {task_id}: {err_msg}")
        finally:
            self._executing_tasks.discard(task_id)

    def _loop(self):
        """Loop contínuo de verificação a cada 30 segundos."""
        logger.info("Loop do BancadaCronDaemon ativo.")
        setup_cron_tables()

        while self._running:
            try:
                now = datetime.now()
                df = get_cron_schedules()
                if not df.empty:
                    for _, row in df.iterrows():
                        task_dict = row.to_dict()
                        task_id = str(task_dict.get("task_id", ""))
                        if self._should_run(task_dict, now):
                            logger.info(f"⏰ Disparando tarefa agendada: {task_id} ({task_dict.get('nome')})")
                            threading.Thread(target=self._run_single_task, args=(task_id,), daemon=True).start()
            except Exception as e:
                logger.error(f"Exceção no loop do Cron Daemon: {e}")

            # Dorme 30 segundos verificando flag de parada
            for _ in range(30):
                if not self._running:
                    break
                time.sleep(1)

        logger.info("Loop do BancadaCronDaemon encerrado.")

# Função utilitária para obter instância global
def get_cron_daemon() -> BancadaCronDaemon:
    daemon = BancadaCronDaemon()
    if not daemon.is_alive():
        daemon.start()
    return daemon
