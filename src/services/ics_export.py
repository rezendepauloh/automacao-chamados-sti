import re
from datetime import datetime, date, timedelta
from typing import List, Dict, Any


def _clean_ics_text(text: str) -> str:
    """Escapa caracteres especiais conforme o padrão iCalendar RFC 5545."""
    if not text:
        return ""
    # Remove ou limpa tags HTML caso existam
    clean = re.sub(r'<[^>]+>', ' ', str(text))
    # Normaliza quebras de linha
    clean = clean.replace('\r\n', '\n').replace('\r', '\n')
    # Escapa contra-barra, ponto e vírgula e vírgula
    clean = clean.replace('\\', '\\\\')
    clean = clean.replace(';', '\\;')
    clean = clean.replace(',', '\\,')
    clean = clean.replace('\n', '\\n')
    return clean.strip()


def _format_datetime_or_date(dt_str: str) -> tuple[str, bool]:
    """
    Identifica se a string é data/hora (YYYY-MM-DDTHH:MM:SS) ou data pura (YYYY-MM-DD).
    Retorna a string formatada para o formato ICS e uma flag indicando se é dia inteiro (VALUE=DATE).
    """
    if not dt_str:
        return "", False

    dt_str = str(dt_str).strip().replace(" ", "T")

    # Caso 1: Data completa com horário (ex: 2026-09-21T08:00:00 ou 2026-09-21T08:00)
    if "T" in dt_str:
        try:
            # Tenta com segundos
            dt = datetime.strptime(dt_str[:19], "%Y-%m-%dT%H:%M:%S")
            return dt.strftime("%Y%m%dT%H%M%S"), False
        except Exception:
            pass
        try:
            # Tenta sem segundos
            dt = datetime.strptime(dt_str[:16], "%Y-%m-%dT%H:%M")
            return dt.strftime("%Y%m%dT%H%M00"), False
        except Exception:
            pass

    # Caso 2: Data pura (YYYY-MM-DD)
    if re.match(r'^\d{4}-\d{2}-\d{2}$', dt_str[:10]):
        try:
            dt = datetime.strptime(dt_str[:10], "%Y-%m-%d")
            return dt.strftime("%Y%m%d"), True
        except Exception:
            pass

    return "", False


def generate_ics_calendar(events: List[Dict[str, Any]], calendar_name: str = "Bancada STI - Calendário Geral") -> str:
    """
    Gera o conteúdo de um arquivo .ics (RFC 5545) a partir da lista unificada de eventos do FullCalendar.
    Compatível com Microsoft Outlook, Office 365, Google Agenda e Apple Calendar.
    """
    now_utc = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Bancada STI//Calendario Geral v1.0//PT-BR",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{_clean_ics_text(calendar_name)}",
        "X-WR-TIMEZONE:America/Campo_Grande",
    ]

    for idx, ev in enumerate(events):
        start_raw = ev.get("start")
        end_raw = ev.get("end") or start_raw
        title = ev.get("title", "Evento STI")
        ev_id = ev.get("id") or f"sti_event_{idx}_{start_raw}"
        props = ev.get("extendedProps", {})

        dt_start_ics, is_all_day_start = _format_datetime_or_date(start_raw)
        dt_end_ics, is_all_day_end = _format_datetime_or_date(end_raw)

        if not dt_start_ics:
            continue

        # Se for evento de dia inteiro e a data final for igual à inicial, no padrão RFC 5545
        # DTEND para dia inteiro deve ser o dia seguinte (exclusivo).
        if is_all_day_start:
            if not dt_end_ics or dt_end_ics == dt_start_ics:
                try:
                    d_obj = datetime.strptime(dt_start_ics, "%Y%m%d") + timedelta(days=1)
                    dt_end_ics = d_obj.strftime("%Y%m%d")
                except Exception:
                    dt_end_ics = dt_start_ics

        # Monta a descrição detalhada para o Outlook
        desc_parts = []
        tipo = props.get("tipo")
        if tipo:
            desc_parts.append(f"Tipo: {tipo}")

        cat = props.get("categoria_evento", "")
        location_val = ""

        if cat == "chamado":
            cid = props.get("id", "")
            base = props.get("base", "")
            status = props.get("status", "")
            tag = props.get("tag", "")
            usuario = props.get("usuario", "")
            localidade = props.get("localidade", "")
            unidade = props.get("unidade", "")
            descricao = props.get("descricao", "")

            if cid:
                desc_parts.append(f"Chamado: #{cid} ({base})")
            if status:
                desc_parts.append(f"Status: {status}")
            if tag:
                desc_parts.append(f"TAG: {tag}")
            if usuario:
                desc_parts.append(f"Solicitante: {usuario}")
            if unidade or localidade:
                desc_parts.append(f"Lotação: {unidade} ({localidade})")
                location_val = f"{unidade} - {localidade}"
            if descricao:
                desc_parts.append(f"\nDescrição:\n{descricao}")

        elif cat == "plantao":
            servidor = props.get("servidor", "")
            telefone = props.get("telefone", "")
            if servidor:
                desc_parts.append(f"Servidor: {servidor}")
            if telefone:
                desc_parts.append(f"Contato: {telefone}")
            location_val = "PGJ / STI"

        elif cat == "viagem":
            localidade = props.get("localidade", "")
            quem_foi = props.get("quem_foi", "")
            chamado = props.get("chamado", "")
            saida = props.get("saida_br", "")
            retorno = props.get("retorno_br", "")
            if localidade:
                desc_parts.append(f"Destino: {localidade}")
                location_val = localidade
            if quem_foi:
                desc_parts.append(f"Técnico(s): {quem_foi}")
            if chamado:
                desc_parts.append(f"Chamado/Demanda: {chamado}")
            if saida or retorno:
                desc_parts.append(f"Período: {saida} até {retorno}")

        elif cat == "garantia":
            contrato = props.get("contrato", "")
            fornecedor = props.get("fornecedor", "")
            item = props.get("item", "")
            status_g = props.get("status_garantia", "")
            if contrato:
                desc_parts.append(f"Contrato: {contrato}")
            if fornecedor:
                desc_parts.append(f"Fornecedor: {fornecedor}")
            if item:
                desc_parts.append(f"Equipamento/Item: {item}")
            if status_g:
                desc_parts.append(f"Status: {status_g}")

        elif cat == "portaria":
            membros = props.get("membros", "")
            ementa = props.get("ementa", "")
            pdf_url = props.get("pdf_url", "")
            if membros:
                desc_parts.append(f"Servidores/Membros: {membros}")
            if ementa:
                desc_parts.append(f"Ementa: {ementa}")
            if pdf_url:
                desc_parts.append(f"Documento/Link: {pdf_url}")

        elif cat == "manual":
            autor = props.get("autor", "")
            descricao = props.get("descricao", "")
            if autor:
                desc_parts.append(f"Autor: {autor}")
            if descricao:
                desc_parts.append(f"Detalhes: {descricao}")

        full_description = "\n".join(desc_parts)

        # Inicia o bloco do evento
        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{ev_id}@bancada-sti.mpms.mp.br")
        lines.append(f"DTSTAMP:{now_utc}")

        if is_all_day_start:
            lines.append(f"DTSTART;VALUE=DATE:{dt_start_ics}")
            if dt_end_ics:
                lines.append(f"DTEND;VALUE=DATE:{dt_end_ics}")
        else:
            lines.append(f"DTSTART:{dt_start_ics}")
            if dt_end_ics:
                lines.append(f"DTEND:{dt_end_ics}")

        lines.append(f"SUMMARY:{_clean_ics_text(title)}")
        if full_description:
            lines.append(f"DESCRIPTION:{_clean_ics_text(full_description)}")
        if location_val:
            lines.append(f"LOCATION:{_clean_ics_text(location_val)}")

        # Categoria para cores e organização no Outlook
        cat_label = tipo or cat.capitalize()
        lines.append(f"CATEGORIES:{_clean_ics_text(cat_label)}")
        lines.append("STATUS:CONFIRMED")
        lines.append("TRANSP:OPAQUE")
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def fetch_all_unified_calendar_events(bancada_only: bool = True, ano: int = None) -> List[Dict[str, Any]]:
    """
    Carrega todos os eventos consolidados de todas as fontes da Bancada STI:
    Eventos Manuais, Plantões Matutinos, Plantões Semanais, Garantias, Viagens, Chamados e Portarias.
    """
    if ano is None:
        ano = datetime.now().year

    from src.database import (
        get_plantoes_matutino,
        get_plantoes_semanal,
        get_garantia_contratos_df,
        get_viagens_df,
        load_data,
        get_eventos_manuais
    )
    from src.tabs.plantoes import format_phone_number, is_bancada_member
    from src.tabs.garantia import parse_date_to_iso_and_br
    from src.tabs.portarias import fetch_portarias_bancada

    events = []

    # 1. Eventos Manuais
    try:
        df_manuais = get_eventos_manuais()
        if not df_manuais.empty:
            for idx, row in df_manuais.iterrows():
                titulo = str(row.get('titulo', '')).strip()
                dt_ini = str(row.get('data_inicio', '')).strip()
                dt_fim = str(row.get('data_fim', '')).strip()
                if dt_ini:
                    events.append({
                        "id": f"manual_{row.get('id', idx)}",
                        "title": f"📝 {titulo}",
                        "start": dt_ini,
                        "end": dt_fim if dt_fim else dt_ini,
                        "extendedProps": {
                            "categoria_evento": "manual",
                            "tipo": "Registro Manual",
                            "autor": str(row.get('autor', 'Bancada STI')).strip(),
                            "descricao": str(row.get('descricao', '')).strip()
                        }
                    })
    except Exception:
        pass

    # 2. Plantão Matutino
    try:
        df_mat = get_plantoes_matutino(ano)
        if not df_mat.empty:
            for _, row in df_mat.iterrows():
                servidor = str(row.get('servidor', '')).strip()
                if bancada_only and not is_bancada_member(servidor):
                    continue
                dt_iso = str(row.get('data_iso', '')).strip()
                if dt_iso:
                    events.append({
                        "id": f"plantao_mat_{dt_iso}_{servidor.split()[0]}",
                        "title": f"☀️ Matutino: {servidor.split()[0]} ({servidor.split()[-1]})",
                        "start": f"{dt_iso}T08:00:00",
                        "end": f"{dt_iso}T15:00:00",
                        "extendedProps": {
                            "categoria_evento": "plantao",
                            "servidor": servidor,
                            "telefone": format_phone_number(str(row.get('telefone', ''))),
                            "tipo": "Plantão Matutino PGJ (08h-15h)"
                        }
                    })
    except Exception:
        pass

    # 3. Plantão Semanal
    try:
        df_sem = get_plantoes_semanal(ano)
        if not df_sem.empty:
            for _, row in df_sem.iterrows():
                manut = str(row.get('manutencao', '')).strip()
                sdesk = str(row.get('service_desk', '')).strip()
                infra = str(row.get('infraestrutura', '')).strip()
                dev = str(row.get('desenvolvimento', '')).strip()
                dt_ini_raw = str(row.get('data_inicio', '')).strip()
                dt_fim_raw = str(row.get('data_fim', '')).strip()

                if bancada_only:
                    bancada_na_escala = [s for s in [manut, sdesk, infra, dev] if is_bancada_member(s)]
                    if not bancada_na_escala:
                        continue
                    display_name = ", ".join([s.split()[0] for s in bancada_na_escala])
                else:
                    display_name = manut.split()[0] if manut else "STI"

                if dt_ini_raw:
                    events.append({
                        "id": f"plantao_sem_{dt_ini_raw[:10]}",
                        "title": f"🌙 Plantão Semanal: {display_name}",
                        "start": dt_ini_raw.replace(" ", "T"),
                        "end": dt_fim_raw.replace(" ", "T") if dt_fim_raw else dt_ini_raw.replace(" ", "T"),
                        "extendedProps": {
                            "categoria_evento": "plantao",
                            "servidor": f"Manutenção: {manut} | Service Desk: {sdesk} | Infra: {infra} | Dev: {dev}",
                            "tipo": "Plantão Semanal SIMP"
                        }
                    })
    except Exception:
        pass

    # 4. Viagens
    try:
        df_viagens = get_viagens_df()
        if not df_viagens.empty:
            for idx, row in df_viagens.iterrows():
                saida_iso = row.get("saida_iso", "")
                retorno_iso = row.get("retorno_iso", "")
                localidade = row.get("localidade", "")
                quem_foi = row.get("quem_foi", "")
                if saida_iso:
                    cal_end = saida_iso
                    if retorno_iso:
                        try:
                            dt_ret = datetime.strptime(retorno_iso, "%Y-%m-%d")
                            cal_end = (dt_ret + timedelta(days=1)).strftime("%Y-%m-%d")
                        except Exception:
                            cal_end = retorno_iso
                    events.append({
                        "id": f"viagem_{row.get('id', idx)}",
                        "title": f"✈️ Viagem: {localidade}" + (f" ({quem_foi})" if quem_foi else ""),
                        "start": saida_iso,
                        "end": cal_end,
                        "extendedProps": {
                            "categoria_evento": "viagem",
                            "tipo": "Viagem da Bancada",
                            "localidade": localidade,
                            "quem_foi": quem_foi,
                            "chamado": str(row.get("chamado", "")),
                            "saida_br": str(row.get("saida_br", "")),
                            "retorno_br": str(row.get("retorno_br", ""))
                        }
                    })
    except Exception:
        pass

    # 5. Garantias
    try:
        df_garantia = get_garantia_contratos_df()
        if not df_garantia.empty:
            for idx, row in df_garantia.iterrows():
                item = str(row.get('item', '')).strip()
                forn = str(row.get('fornecedor', '')).strip()
                iso_fim, br_fim = parse_date_to_iso_and_br(row.get('garantia_fim'))
                if iso_fim:
                    events.append({
                        "id": f"garantia_fim_{idx}",
                        "title": f"🔴 Fim Garantia: {item} ({forn})",
                        "start": iso_fim,
                        "extendedProps": {
                            "categoria_evento": "garantia",
                            "tipo": "Fim de Garantia",
                            "contrato": str(row.get('contrato', '')),
                            "item": item,
                            "fornecedor": forn,
                            "status_garantia": str(row.get('status_garantia', ''))
                        }
                    })
    except Exception:
        pass

    # 6. Portarias
    try:
        portarias = fetch_portarias_bancada()
        for p in portarias:
            dt_emissao = p.get("data_emissao", "")
            if not dt_emissao:
                continue
            try:
                dt_obj = datetime.strptime(dt_emissao, "%d/%m/%Y")
                iso_dt = dt_obj.strftime("%Y-%m-%d")
            except Exception:
                continue
            membros = ", ".join(p.get("membros", []))
            num = p.get("numero", p.get("id"))
            events.append({
                "id": f"portaria_{p.get('id')}",
                "title": f"📜 Portaria #{num}: {membros.split(',')[0]}",
                "start": iso_dt,
                "extendedProps": {
                    "categoria_evento": "portaria",
                    "tipo": "Portaria da Bancada",
                    "membros": membros,
                    "ementa": p.get("texto", "")[:300],
                    "pdf_url": p.get("pdf_url", "")
                }
            })
    except Exception:
        pass

    return events


def update_published_ics_file(target_path: str = None) -> str:
    """
    Gera o calendário consolidado da Bancada STI e salva em um arquivo estático para acesso web contínuo.
    Por padrão salva em 'src/js/calendario.ics' e na raiz pública.
    """
    from pathlib import Path
    root = Path(__file__).resolve().parent.parent.parent

    events = fetch_all_unified_calendar_events(bancada_only=True)
    ics_text = generate_ics_calendar(events, calendar_name="Bancada STI - Calendário Unificado")

    # Caminhos onde o arquivo fica disponível
    destinos = [
        root / "src" / "js" / "calendario.ics",
        root / "calendario.ics"
    ]
    if target_path:
        destinos.append(Path(target_path))

    for dest in destinos:
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(ics_text, encoding="utf-8")
        except Exception:
            pass

    return destinos[0].as_posix()
