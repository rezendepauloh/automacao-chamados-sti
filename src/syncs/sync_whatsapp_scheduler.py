import sys
import argparse
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from src.config import setup_logging, DEBUG_DIR_PLANTOES
from src.database import (
    get_plantoes_matutino,
    get_plantoes_semanal,
    get_viagens_df,
    has_whatsapp_been_sent,
    log_whatsapp_dispatch,
    get_whatsapp_destinatarios
)
from src.services.member_matcher import (
    BANCADA_MEMBROS,
    resolve_bancada_member,
    extract_all_bancada_members
)
from src.services.evolution_client import send_whatsapp_text, get_connection_status

logger = setup_logging(DEBUG_DIR_PLANTOES / "whatsapp_scheduler.log", "sync_whatsapp_scheduler")

def get_target_dates_for_today(reference_date: datetime.date = None) -> list[datetime.date]:
    """
    Retorna as datas de eventos que devem ser avisadas hoje considerando regra D-1 em dias úteis:
    - Segunda a Quinta: avisa o dia seguinte (D+1).
    - Sexta-feira: avisa Sábado (D+1), Domingo (D+2) e Segunda-feira (D+3).
    - Fim de semana (Sábado/Domingo): não envia disparos rotineiros de D-1 (avisados na sexta).
    """
    if reference_date is None:
        reference_date = datetime.now().date()
        
    weekday = reference_date.weekday() # 0 = Segunda, 4 = Sexta, 5 = Sábado, 6 = Domingo
    if weekday == 4: # Sexta-feira
        return [
            reference_date + timedelta(days=1), # Sábado
            reference_date + timedelta(days=2), # Domingo
            reference_date + timedelta(days=3)  # Segunda-feira
        ]
    elif weekday < 4: # Segunda a Quinta
        return [reference_date + timedelta(days=1)]
    else:
        return []

def run_whatsapp_scheduler(dry_run: bool = False, force: bool = False) -> dict:
    """
    Executa a verificação de plantões, viagens e portarias para disparos no WhatsApp.
    Retorna sumário da execução.
    """
    logger.info(f"Iniciando rotina do WhatsApp Scheduler (dry_run={dry_run}, force={force})...")
    
    # Verifica status da conexão da Evolution API
    status = get_connection_status()
    if not status.get("online") or status.get("state") != "open":
        msg = f"Evolution API não está pronta/conectada (Estado: {status.get('state')}). Cancelando disparos."
        logger.warning(msg)
        return {"status": "aborted", "reason": msg, "sent_count": 0}

    today = datetime.now().date()
    target_dates = get_target_dates_for_today(today)
    target_iso_set = {d.strftime("%Y-%m-%d") for d in target_dates}
    
    logger.info(f"Data de referência: {today} | Datas alvo D-1: {list(target_iso_set)}")
    
    sent_count = 0
    skipped_count = 0
    errors_count = 0

    # Carrega destinatários ativos da Bancada
    dest_df = get_whatsapp_destinatarios(only_active=True)
    active_phones = set(dest_df["telefone"].astype(str).tolist()) if not dest_df.empty else set()

    # =========================================================================
    # 1. PLANTÕES MATUTINOS (D-1 para o membro escalado)
    # =========================================================================
    try:
        df_mat = get_plantoes_matutino(today.year)
        if not df_mat.empty:
            for _, row in df_mat.iterrows():
                servidor_str = str(row.get("servidor", "")).strip()
                membro = resolve_bancada_member(servidor_str)
                if not membro:
                    continue
                    
                data_iso = str(row.get("data_iso", "")).strip()
                if not data_iso:
                    continue
                    
                # Checa se é uma das datas alvo
                if not force and data_iso not in target_iso_set:
                    continue

                phone = membro["telefone"]
                if active_phones and phone not in active_phones:
                    continue

                evento_id = f"matutino_{data_iso}_{phone}"
                tipo_evento = "Plantão Matutino"

                if not force and has_whatsapp_been_sent(tipo_evento, evento_id, phone):
                    skipped_count += 1
                    continue

                # Formata mensagem
                try:
                    dt_obj = datetime.strptime(data_iso, "%Y-%m-%d").date()
                    data_fmt = dt_obj.strftime("%d/%m/%Y")
                except Exception:
                    data_fmt = data_iso
                    
                dia_sem = str(row.get("dia_semana", "")).strip()
                primeiro_nome = membro["primeiro_nome"]
                
                texto = (
                    f"☀️ *Lembrete de Plantão Matutino - Bancada STI*\n\n"
                    f"Olá, *{primeiro_nome}*! 👋\n\n"
                    f"Você está escalado(a) para o *Plantão Matutino*:\n"
                    f"📅 *Data:* {data_fmt} ({dia_sem})\n"
                    f"⏰ *Horário:* 08:00 às 15:00\n"
                    f"🏢 *Setor:* Bancada de Atendimento STI / DIT\n\n"
                    f"_Mensagem automática enviada pelo Sistema de Gestão da Bancada._"
                )

                if dry_run:
                    logger.info(f"[DRY-RUN] Enviaria Plantão Matutino para {membro['nome']} ({phone})")
                    sent_count += 1
                else:
                    res = send_whatsapp_text(phone, texto)
                    if res.get("success"):
                        log_whatsapp_dispatch(tipo_evento, evento_id, data_iso, phone, texto, "enviado", str(res))
                        logger.info(f"✅ Alerta Plantão Matutino enviado para {membro['nome']} ({phone})")
                        sent_count += 1
                    else:
                        log_whatsapp_dispatch(tipo_evento, evento_id, data_iso, phone, texto, "erro", str(res))
                        logger.error(f"❌ Falha ao enviar para {phone}: {res.get('error')}")
                        errors_count += 1
    except Exception as e:
        logger.error(f"Erro ao processar plantões matutinos: {e}")

    # =========================================================================
    # 2. PLANTÕES SEMANAIS (D-1 na sexta-feira ou dia anterior à escala)
    # =========================================================================
    try:
        df_sem = get_plantoes_semanal(today.year)
        if not df_sem.empty:
            for _, row in df_sem.iterrows():
                dt_ini_str = str(row.get("data_inicio", "")).strip().split()[0]
                if not dt_ini_str:
                    continue

                if not force and dt_ini_str not in target_iso_set:
                    continue

                manut = str(row.get("manutencao", "")).strip()
                sdesk = str(row.get("service_desk", "")).strip()
                infra = str(row.get("infraestrutura", "")).strip()
                dev = str(row.get("desenvolvimento", "")).strip()

                for setor, serv in [("Manutenção", manut), ("Service Desk", sdesk), ("Infraestrutura", infra), ("Desenvolvimento", dev)]:
                    membro = resolve_bancada_member(serv)
                    if not membro:
                        continue

                    phone = membro["telefone"]
                    if active_phones and phone not in active_phones:
                        continue

                    evento_id = f"semanal_{dt_ini_str}_{phone}"
                    tipo_evento = "Plantão Semanal"

                    if not force and has_whatsapp_been_sent(tipo_evento, evento_id, phone):
                        skipped_count += 1
                        continue

                    try:
                        dt_obj = datetime.strptime(dt_ini_str, "%Y-%m-%d").date()
                        data_fmt = dt_obj.strftime("%d/%m/%Y")
                    except Exception:
                        data_fmt = dt_ini_str

                    primeiro_nome = membro["primeiro_nome"]
                    texto = (
                        f"🌙 *Lembrete de Plantão Semanal SIMP - Bancada STI*\n\n"
                        f"Olá, *{primeiro_nome}*! 👋\n\n"
                        f"A sua semana de *Plantão Semanal* tem início em breve:\n"
                        f"📅 *Início da Escala:* {data_fmt}\n"
                        f"🛠️ *Atribuição:* {setor}\n\n"
                        f"Por favor, mantenha seu telefone funcional a postos.\n\n"
                        f"_Mensagem automática enviada pelo Sistema de Gestão da Bancada._"
                    )

                    if dry_run:
                        logger.info(f"[DRY-RUN] Enviaria Plantão Semanal para {membro['nome']} ({phone})")
                        sent_count += 1
                    else:
                        res = send_whatsapp_text(phone, texto)
                        if res.get("success"):
                            log_whatsapp_dispatch(tipo_evento, evento_id, dt_ini_str, phone, texto, "enviado", str(res))
                            logger.info(f"✅ Alerta Plantão Semanal enviado para {membro['nome']} ({phone})")
                            sent_count += 1
                        else:
                            log_whatsapp_dispatch(tipo_evento, evento_id, dt_ini_str, phone, texto, "erro", str(res))
                            logger.error(f"❌ Falha ao enviar semanal para {phone}: {res.get('error')}")
                            errors_count += 1
    except Exception as e:
        logger.error(f"Erro ao processar plantões semanais: {e}")

    # =========================================================================
    # 3. VIAGENS DA BANCADA (D-1 para todos os integrantes que vão viajar)
    # =========================================================================
    try:
        df_viagens = get_viagens_df()
        if not df_viagens.empty:
            for _, row in df_viagens.iterrows():
                saida_iso = str(row.get("saida_iso", "")).strip()
                if not saida_iso:
                    continue

                if not force and saida_iso not in target_iso_set:
                    continue

                quem_foi = str(row.get("quem_foi", "")).strip()
                membros_viagem = extract_all_bancada_members(quem_foi)
                if not membros_viagem:
                    continue

                localidade = str(row.get("localidade", "")).strip() or "Destino a definir"
                chamado = str(row.get("chamado", "")).strip()
                retorno_br = str(row.get("retorno_br", "")).strip()
                saida_br = str(row.get("saida_br", "")).strip() or saida_iso

                for membro in membros_viagem:
                    phone = membro["telefone"]
                    if active_phones and phone not in active_phones:
                        continue

                    v_id = str(row.get("id", ""))
                    evento_id = f"viagem_{v_id}_{saida_iso}_{phone}"
                    tipo_evento = "Viagem da Bancada"

                    if not force and has_whatsapp_been_sent(tipo_evento, evento_id, phone):
                        skipped_count += 1
                        continue

                    primeiro_nome = membro["primeiro_nome"]
                    texto = (
                        f"🚗 *Lembrete de Viagem da Bancada STI*\n\n"
                        f"Olá, *{primeiro_nome}*! 👋\n\n"
                        f"Você tem uma viagem agendada com saída prevista para breve:\n"
                        f"📍 *Destino:* {localidade}\n"
                        f"🗓️ *Data de Saída:* {saida_br}\n"
                        f"🗓️ *Data de Retorno:* {retorno_br if retorno_br else 'A definir'}\n"
                    )
                    if chamado and chamado.lower() not in ["none", "nan", ""]:
                        texto += f"🎫 *Chamado Relacionado:* {chamado}\n"
                    texto += f"👥 *Integrantes:* {quem_foi}\n\n_Boa viagem e bom trabalho!_\n_Sistema Bancada STI_"

                    if dry_run:
                        logger.info(f"[DRY-RUN] Enviaria Viagem para {membro['nome']} ({phone})")
                        sent_count += 1
                    else:
                        res = send_whatsapp_text(phone, texto)
                        if res.get("success"):
                            log_whatsapp_dispatch(tipo_evento, evento_id, saida_iso, phone, texto, "enviado", str(res))
                            logger.info(f"✅ Alerta de Viagem enviado para {membro['nome']} ({phone})")
                            sent_count += 1
                        else:
                            log_whatsapp_dispatch(tipo_evento, evento_id, saida_iso, phone, texto, "erro", str(res))
                            logger.error(f"❌ Falha ao enviar viagem para {phone}: {res.get('error')}")
                            errors_count += 1
    except Exception as e:
        logger.error(f"Erro ao processar viagens: {e}")

    # =========================================================================
    # 4. NOVAS PORTARIAS (Avisa os 3 membros quando uma portaria nova é detectada)
    # =========================================================================
    try:
        from src.tabs.portarias import fetch_portarias_bancada
        portarias = fetch_portarias_bancada()
        for p in portarias:
            ato_id = str(p.get("id", ""))
            numero = str(p.get("numero", "S/N"))
            membros_portaria = p.get("membros", [])
            titulo_ementa = str(p.get("titulo", ""))
            data_emissao = str(p.get("data_emissao", "")).strip()
            data_publicacao = str(p.get("data_publicacao", "")).strip()

            # Converte data para checagem do mesmo dia (ou dias recentes se ainda não enviada)
            # Portarias publicadas hoje às 12h ou emitidas hoje
            today_br = today.strftime("%d/%m/%Y")
            if not force:
                # Dispara se a data de publicação ou de emissão for hoje (ou recente até 2 dias caso publicada no fim de semana)
                is_recent = False
                for d_str in [data_publicacao, data_emissao]:
                    if d_str == today_br:
                        is_recent = True
                        break
                    try:
                        dt_p = datetime.strptime(d_str, "%d/%m/%Y").date()
                        # Se foi publicada nas últimas 48h e ainda não enviada
                        if 0 <= (today - dt_p).days <= 2:
                            is_recent = True
                            break
                    except Exception:
                        continue
                if not is_recent:
                    continue

            # Verifica quais membros autorizados estão citados
            for membro in BANCADA_MEMBROS:
                if any(membro["nome"].lower() in m.lower() for m in membros_portaria):
                    phone = membro["telefone"]
                    if active_phones and phone not in active_phones:
                        continue

                    evento_id = f"portaria_{ato_id}_{phone}"
                    tipo_evento = "Portaria Bancada"

                    if not force and has_whatsapp_been_sent(tipo_evento, evento_id, phone):
                        skipped_count += 1
                        continue

                    primeiro_nome = membro["primeiro_nome"]
                    ementa_resumida = titulo_ementa[:150] + "..." if len(titulo_ementa) > 150 else titulo_ementa
                    data_ref = data_publicacao or data_emissao
                    texto = (
                        f"📜 *Nova Portaria Publicada - Bancada STI*\n\n"
                        f"Olá, *{primeiro_nome}*!\n\n"
                        f"Foi identificada uma portaria publicada envolvendo a bancada:\n"
                        f"📄 *Portaria nº:* {numero}\n"
                        f"📅 *Data:* {data_ref}\n"
                        f"📝 *Ementa:* {ementa_resumida}\n\n"
                        f"Acesse o sistema da Bancada para visualizar os detalhes completos do documento."
                    )

                    if dry_run:
                        logger.info(f"[DRY-RUN] Enviaria Portaria para {membro['nome']} ({phone})")
                        sent_count += 1
                    else:
                        res = send_whatsapp_text(phone, texto)
                        if res.get("success"):
                            log_whatsapp_dispatch(tipo_evento, evento_id, data_emissao, phone, texto, "enviado", str(res))
                            logger.info(f"✅ Alerta de Portaria enviado para {membro['nome']} ({phone})")
                            sent_count += 1
                        else:
                            log_whatsapp_dispatch(tipo_evento, evento_id, data_emissao, phone, texto, "erro", str(res))
                            logger.error(f"❌ Falha ao enviar portaria para {phone}: {res.get('error')}")
                            errors_count += 1
    except Exception as e:
        logger.error(f"Erro ao processar portarias: {e}")

    summary = {
        "status": "completed",
        "sent_count": sent_count,
        "skipped_count": skipped_count,
        "errors_count": errors_count,
        "target_dates": list(target_iso_set)
    }
    logger.info(f"Finalizado WhatsApp Scheduler. Sumário: {summary}")
    return summary

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scheduler de notificações WhatsApp para Bancada STI")
    parser.add_argument("--dry-run", action="store_true", help="Apenas simula os disparos sem enviar")
    parser.add_argument("--force", action="store_true", help="Ignora verificação de D-1 e envia mesmo que já tenha sido disparado")
    args = parser.parse_args()

    res = run_whatsapp_scheduler(dry_run=args.dry_run, force=args.force)
    print(res)
