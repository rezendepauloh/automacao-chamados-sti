import sys
from pathlib import Path
from datetime import datetime, timedelta

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from src.database import get_plantoes_matutino, get_plantoes_semanal, add_notification
from src.tabs.plantoes import is_bancada_member
from src.config import setup_logging, DEBUG_DIR_PLANTOES

logger = setup_logging(DEBUG_DIR_PLANTOES / "sync_alerts.log", "sync_plantoes_alerts")

def check_and_generate_plantao_alerts():
    """Gera notificações automáticas para os plantões matutino e semanal seguindo as regras de antecedência."""
    logger.info("Iniciando verificação de alertas de plantão...")
    today = datetime.now().date()

    df_mat = get_plantoes_matutino(today.year)
    if not df_mat.empty:
        for _, row in df_mat.iterrows():
            servidor = str(row.get("servidor", "")).strip()
            if not is_bancada_member(servidor):
                continue

            dt_iso_str = str(row.get("data_iso", "")).strip()
            if not dt_iso_str:
                continue

            try:
                dt_duty = datetime.strptime(dt_iso_str, "%Y-%m-%d").date()
            except Exception:
                continue

            if dt_duty.weekday() == 0:
                notify_date = dt_duty - timedelta(days=3)
            else:
                notify_date = dt_duty - timedelta(days=1)

            if notify_date <= today <= dt_duty:
                dt_fmt = dt_duty.strftime("%d/%m/%Y")
                dia_semana_str = row.get("dia_semana", "")
                
                titulo = f"Plantão Matutino Amanhã ({servidor.split()[0]})" if today == dt_duty - timedelta(days=1) else f"Lembrete: Plantão Matutino ({servidor.split()[0]})"
                msg = f"Servidor {servidor} escalado para o Plantão Matutino dia {dt_fmt} ({dia_semana_str}) das 08h às 15h."

                inserted = add_notification(
                    tipo="Plantão Matutino",
                    titulo=titulo,
                    mensagem=msg,
                    data_evento=dt_iso_str,
                    link_pagina="📅 Plantões da Bancada"
                )
                if inserted:
                    logger.info(f"🔔 Notificação criada para Plantão Matutino de {servidor} em {dt_fmt}")

    df_sem = get_plantoes_semanal(today.year)
    if not df_sem.empty:
        for _, row in df_sem.iterrows():
            dt_ini_str = str(row.get("data_inicio", "")).strip()
            if not dt_ini_str:
                continue

            try:
                dt_ini_date = datetime.strptime(dt_ini_str.split()[0], "%Y-%m-%d").date()
            except Exception:
                continue

            manut = str(row.get("manutencao", "")).strip()
            sdesk = str(row.get("service_desk", "")).strip()
            infra = str(row.get("infraestrutura", "")).strip()
            dev = str(row.get("desenvolvimento", "")).strip()

            bancada_escalados = [s for s in [manut, sdesk, infra, dev] if is_bancada_member(s)]
            if not bancada_escalados:
                continue

            monday_of_week = dt_ini_date - timedelta(days=dt_ini_date.weekday())

            if monday_of_week <= today <= dt_ini_date + timedelta(days=7):
                dt_fmt = dt_ini_date.strftime("%d/%m/%Y")
                membros_nomes = ", ".join([s.split()[0] for s in bancada_escalados])
                
                titulo = f"Lembrete: Plantão Semanal SIMP ({membros_nomes})"
                msg = f"Escala de Plantão Semanal com início em {dt_fmt}. Integrante(s): {membros_nomes}."

                inserted = add_notification(
                    tipo="Plantão Semanal",
                    titulo=titulo,
                    mensagem=msg,
                    data_evento=dt_ini_str.split()[0],
                    link_pagina="📅 Plantões da Bancada"
                )
                if inserted:
                    logger.info(f"🔔 Notificação criada para Plantão Semanal em {dt_fmt}")

    logger.info("Verificação de alertas de plantão concluída.")

if __name__ == "__main__":
    check_and_generate_plantao_alerts()
