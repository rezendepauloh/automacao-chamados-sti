import sys
from pathlib import Path

# Adiciona a raiz do projeto e a pasta src ao sys.path
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

import tempfile
import os
from pathlib import Path
from src.database import add_notification
from src.config import setup_logging, DEBUG_DIR_FAQ
from src.components.status_banner import check_process_running, read_log_lines

logger = setup_logging(DEBUG_DIR_FAQ / "sync_portarias.log", "sync_portarias")


def check_portarias_sync_running() -> bool:
    lock_file = Path(tempfile.gettempdir()) / "portarias_sync.lock"
    return check_process_running(lock_file)


def read_portarias_last_log_lines(n: int = 15) -> str:
    log_path = Path("debug_logs") / "faq" / "sync_portarias.log"
    return read_log_lines(log_path, n)


def sync_portarias_and_generate_alerts():
    """Busca as portarias da API do MPMS e gera notificações para novas ocorrências."""
    from src.tabs.portarias import fetch_portarias_bancada
    logger.info("Iniciando verificação de novas Portarias da Bancada...")
    
    portarias = fetch_portarias_bancada()
    novas_count = 0

    for item in portarias:
        ato_id = item.get("id")
        numero = item.get("numero", "S/N")
        data_emissao = item.get("data_emissao", "")
        membros_str = ", ".join(item.get("membros", []))
        titulo = item.get("titulo", "")
        titulo_resumo = titulo[:120] + "..." if len(titulo) > 120 else titulo

        titulo_notif = f"Nova Portaria nº {numero}"
        msg_notif = f"Servidores: {membros_str}. Ementa: {titulo_resumo}"

        inserted = add_notification(
            tipo="Portaria",
            titulo=titulo_notif,
            mensagem=msg_notif,
            data_evento=data_emissao,
            link_pagina="📜 Portarias da Bancada"
        )

        if inserted:
            novas_count += 1
            logger.info(f"🔔 Nova notificação gerada para Portaria nº {numero}")

    logger.info(f"Sincronização de portarias concluída. Total de novos alertas gerados: {novas_count}")


if __name__ == "__main__":
    lock_path = Path(tempfile.gettempdir()) / "portarias_sync.lock"
    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))

    try:
        sync_portarias_and_generate_alerts()
    finally:
        if lock_path.exists():
            try:
                lock_path.unlink()
            except Exception:
                pass
