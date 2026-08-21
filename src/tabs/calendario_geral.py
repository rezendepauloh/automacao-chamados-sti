import json
import re
from datetime import datetime, date, time
import pandas as pd
import streamlit as st
from src.components.calendar import render_master_calendar
from src.database import (
    get_plantoes_matutino,
    get_plantoes_semanal,
    get_garantia_contratos_df,
    load_data,
    save_evento_manual,
    get_eventos_manuais
)
from src.tabs.plantoes import format_phone_number, is_bancada_member
from src.tabs.garantia import parse_date_to_iso_and_br
from src.tabs.portarias import fetch_portarias_bancada



def parse_ticket_date_iso_and_br(date_val):
    """Auxiliar para converter datas de chamados (ISO ou BR string) em (ISO_str, BR_str)."""
    if pd.isna(date_val) or not date_val or str(date_val).strip() == "":
        return None, None
    s = str(date_val).strip()
    try:
        # Se a data já for ISO (YYYY-MM-DD...)
        if re.match(r'^\d{4}-\d{2}-\d{2}', s):
            dt = pd.to_datetime(s, errors='coerce')
        else:
            # Se for formato brasileiro (DD/MM/YYYY...)
            dt = pd.to_datetime(s, dayfirst=True, errors='coerce')

        if pd.isna(dt):
            return None, None

        if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
            iso_str = dt.strftime('%Y-%m-%d')
        else:
            iso_str = dt.strftime('%Y-%m-%dT%H:%M:%S')

        br_str = dt.strftime('%d/%m/%Y %H:%M:%S')
        return iso_str, br_str
    except Exception:
        return None, None


def extract_vacation_dates(text: str, fallback_date_str: str):
    """
    Procura por padrões de período de férias em formato de texto usando Regex.
    Exemplos: 'de 8 a 17.9.2026', '01/10 a 15/10', 'de 05 a 19/08/2026', '10/05/2026 a 25/05/2026'.
    Retorna uma tupla (start_iso, end_iso, start_br, end_br). Se falhar, faz fallback para a data de emissão.
    """
    if not text:
        return None, None, None, None

    # Tenta extrair ano de referência do texto ou fallback_date_str
    current_year = datetime.now().year
    if fallback_date_str:
        try:
            current_year = datetime.strptime(fallback_date_str, "%d/%m/%Y").year
        except Exception:
            pass

    # Padrão 1: "de DD a DD/MM/YYYY" ou "de DD.MM a DD.MM.YYYY" ou "DD/MM a DD/MM/YYYY"
    p1 = r'(?:no\s+período\s+)?(?:de\s+)?(\d{1,2})[\/\.]?(\d{1,2})?\s*(?:a|até|-)\s*(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{4})'
    m1 = re.search(p1, text, re.IGNORECASE)
    if m1:
        d1, m1_m, d2, m2, y2 = m1.groups()
        m1_val = int(m1_m) if m1_m else int(m2)
        try:
            dt1 = datetime(int(y2), m1_val, int(d1))
            dt2 = datetime(int(y2), int(m2), int(d2))
            return dt1.strftime("%Y-%m-%d"), dt2.strftime("%Y-%m-%d"), dt1.strftime("%d/%m/%Y"), dt2.strftime("%d/%m/%Y")
        except Exception:
            pass

    # Padrão 2: "de DD/MM a DD/MM" (sem ano explícito)
    p2 = r'(?:no\s+período\s+)?(?:de\s+)?(\d{1,2})[\/\.](\d{1,2})\s*(?:a|até|-)\s*(\d{1,2})[\/\.](\d{1,2})'
    m2 = re.search(p2, text, re.IGNORECASE)
    if m2:
        d1, m1, d2, m2_m = m2.groups()
        try:
            dt1 = datetime(current_year, int(m1), int(d1))
            dt2 = datetime(current_year, int(m2_m), int(d2))
            return dt1.strftime("%Y-%m-%d"), dt2.strftime("%Y-%m-%d"), dt1.strftime("%d/%m/%Y"), dt2.strftime("%d/%m/%Y")
        except Exception:
            pass

    # Padrão 3: "de DD a DD.MM.YYYY" (ex: "de 8 a 17.9.2026")
    p3 = r'(?:de\s+)?(\d{1,2})\s*a\s*(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{4})'
    m3 = re.search(p3, text, re.IGNORECASE)
    if m3:
        d1, d2, m, y = m3.groups()
        try:
            dt1 = datetime(int(y), int(m), int(d1))
            dt2 = datetime(int(y), int(m), int(d2))
            return dt1.strftime("%Y-%m-%d"), dt2.strftime("%Y-%m-%d"), dt1.strftime("%d/%m/%Y"), dt2.strftime("%d/%m/%Y")
        except Exception:
            pass

    return None, None, None, None


@st.dialog("➕ Novo Evento Manual")
def modal_novo_evento_manual():
    """Modal nativo do Streamlit (@st.dialog) para cadastro de eventos manuais."""
    st.write("Preencha as informações do evento para adicionar ao Calendário Geral:")

    titulo = st.text_input("📌 Título do Evento *", placeholder="Ex: Manutenção Preventiva no Servidor X")
    
    col_d1, col_t1 = st.columns(2)
    with col_d1:
        d_ini = st.date_input("📅 Data de Início *", value=date.today(), format="DD/MM/YYYY")
    with col_t1:
        t_ini = st.time_input("⏰ Hora de Início", value=time(8, 0))

    has_end_date = st.checkbox("Definir Data/Hora de Término")
    d_fim, t_fim = None, None
    if has_end_date:
        col_d2, col_t2 = st.columns(2)
        with col_d2:
            d_fim = st.date_input("📅 Data de Término", value=date.today(), format="DD/MM/YYYY")
        with col_t2:
            t_fim = st.time_input("⏰ Hora de Término", value=time(17, 0))

    autor = st.text_input("👤 Autor / Responsável", value="Bancada STI")
    descricao = st.text_area("📝 Descrição / Detalhes", placeholder="Descreva os detalhes e orientações do evento...")

    if st.button("💾 Salvar Evento", type="primary", width='stretch'):
        if not titulo.strip():
            st.error("Por favor, preencha o título do evento.")
            return

        dt_inicio_iso = f"{d_ini.strftime('%Y-%m-%d')}T{t_ini.strftime('%H:%M:%S')}"
        dt_fim_iso = ""
        if has_end_date and d_fim:
            dt_fim_iso = f"{d_fim.strftime('%Y-%m-%d')}T{t_fim.strftime('%H:%M:%S')}"

        save_evento_manual(
            titulo=titulo.strip(),
            data_inicio=dt_inicio_iso,
            data_fim=dt_fim_iso,
            descricao=descricao.strip(),
            autor=autor.strip()
        )
        st.toast("✅ Evento manual salvo com sucesso!", icon="🎉")
        st.rerun()





def render_calendario_geral_page():
    """Renderiza a página principal do Calendário Geral Unificado com Filtros Laterais e Botão de Novo Evento."""
    st.cache_data.clear()

    st.title("📅 Calendário Geral Unificado")
    st.caption("Visão centralizada de registros manuais, plantões da bancada, vigências de contratos de garantia, portarias e chamados técnicos.")

    # --- BOTÃO DE DESTAQUE NO TOPO DA SIDEBAR ---
    if st.sidebar.button("➕ Novo Evento Manual", type="primary", width='stretch'):
        modal_novo_evento_manual()

    st.sidebar.markdown("---")

    # --- FILTROS SIDEBAR (Categorias Desagrupadas) ---
    st.sidebar.markdown("## 📅 Agendas Visíveis")
    
    chk_manuais = st.sidebar.checkbox("Registros Manuais", value=True)
    chk_matutino = st.sidebar.checkbox("Plantão Matutino", value=True)
    chk_semanal = st.sidebar.checkbox("Plantão Semanal", value=True)
    chk_garantias = st.sidebar.checkbox("Garantias", value=True)
    chk_otrs = st.sidebar.checkbox("Chamados OTRS", value=True)
    chk_citsmart = st.sidebar.checkbox("Chamados CitSmart", value=True)
    chk_portarias = st.sidebar.checkbox("Portarias (Geral)", value=True)
    chk_ferias = st.sidebar.checkbox("Portarias (Férias)", value=True)
    chk_portarias_fiscais = st.sidebar.checkbox("Portarias (Fiscais de Contrato)", value=True)

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🔍 Opções Adicionais")

    opcoes_servidores = [
        "🟢 Apenas Bancada (Paulo, Reginaldo, Luiz, Murillo)",
        "🌐 Todos os Servidores da STI"
    ]
    selected_servidor_mode = st.sidebar.radio("👥 Servidores Exibidos (Plantões):", opcoes_servidores)
    bancada_only = "Apenas Bancada" in selected_servidor_mode

    search_query = st.sidebar.text_input("🔎 Pesquisar Evento:", placeholder="Ex: Paulo, impressora, #4645...").strip().lower()

    anos_disponiveis = [2026, 2025, 2024]
    selected_ano = st.sidebar.selectbox("📅 Selecionar Ano (Plantões):", anos_disponiveis)

    # --- CONSTRUÇÃO DA LISTA UNIFICADA DE EVENTOS ---
    events = []

    # 1. EVENTOS MANUAIS
    if chk_manuais:
        df_manuais = get_eventos_manuais()
        if not df_manuais.empty:
            for idx, row in df_manuais.iterrows():
                titulo = str(row.get('titulo', '')).strip()
                dt_ini = str(row.get('data_inicio', '')).strip()
                dt_fim = str(row.get('data_fim', '')).strip()
                autor = str(row.get('autor', 'Bancada STI')).strip()
                descricao = str(row.get('descricao', '')).strip()

                if not dt_ini:
                    continue

                events.append({
                    "id": f"manual_{row.get('id', idx)}",
                    "title": f"📝 {titulo}",
                    "start": dt_ini,
                    "end": dt_fim if dt_fim else dt_ini,
                    "backgroundColor": "#10b981",
                    "borderColor": "#059669",
                    "extendedProps": {
                        "categoria_evento": "manual",
                        "tipo": "Registro Manual",
                        "autor": autor,
                        "descricao": descricao,
                        "raw_data_inicio": dt_ini,
                        "raw_data_fim": dt_fim
                    }
                })

    # 2. PLANTÃO MATUTINO
    if chk_matutino:
        df_matutino = get_plantoes_matutino(selected_ano)
        if not df_matutino.empty:
            for _, row in df_matutino.iterrows():
                servidor = str(row.get('servidor', '')).strip()
                if bancada_only and not is_bancada_member(servidor):
                    continue

                dt_iso = str(row.get('data_iso', '')).strip()
                if not dt_iso:
                    continue
                tel_formatted = format_phone_number(str(row.get('telefone', '')))
                
                dt_ini_str = f"{dt_iso}T08:00:00"
                dt_fim_str = f"{dt_iso}T15:00:00"
                
                raw_ini_br, raw_fim_br = "", ""
                try:
                    d_obj = datetime.strptime(dt_iso, "%Y-%m-%d")
                    raw_ini_br = d_obj.strftime("%d/%m/%Y 08:00:00")
                    raw_fim_br = d_obj.strftime("%d/%m/%Y 15:00:00")
                except Exception:
                    pass

                events.append({
                    "title": f"☀️ Matutino: {servidor.split()[0]} ({servidor.split()[-1]})",
                    "start": dt_ini_str,
                    "end": dt_fim_str,
                    "backgroundColor": "#e67e22",
                    "borderColor": "#d35400",
                    "extendedProps": {
                        "categoria_evento": "plantao",
                        "servidor": servidor,
                        "telefone": tel_formatted,
                        "tipo": "Plantão Matutino PGJ (08h-15h)",
                        "raw_data_inicio": raw_ini_br,
                        "raw_data_fim": raw_fim_br
                    }
                })

    # 3. PLANTÃO SEMANAL
    if chk_semanal:
        df_semanal = get_plantoes_semanal(selected_ano)
        if not df_semanal.empty:
            for _, row in df_semanal.iterrows():
                manut = str(row.get('manutencao', '')).strip()
                sdesk = str(row.get('service_desk', '')).strip()
                infra = str(row.get('infraestrutura', '')).strip()
                dev = str(row.get('desenvolvimento', '')).strip()
                
                dt_ini_raw = str(row.get('data_inicio', '')).strip()
                dt_fim_raw = str(row.get('data_fim', '')).strip()
                
                dt_ini = dt_ini_raw.replace(" ", "T")
                dt_fim = dt_fim_raw.replace(" ", "T")
                
                raw_ini_br, raw_fim_br = "", ""
                try:
                    if dt_ini_raw:
                        dt_o = datetime.strptime(dt_ini_raw, "%Y-%m-%d %H:%M:%S")
                        raw_ini_br = dt_o.strftime("%d/%m/%Y %H:%M:%S")
                    if dt_fim_raw:
                        dt_f = datetime.strptime(dt_fim_raw, "%Y-%m-%d %H:%M:%S")
                        raw_fim_br = dt_f.strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    pass

                if bancada_only:
                    bancada_na_escala = [s for s in [manut, sdesk, infra, dev] if is_bancada_member(s)]
                    if not bancada_na_escala:
                        continue
                    display_name = ", ".join([s.split()[0] for s in bancada_na_escala])
                else:
                    display_name = manut.split()[0] if manut else "STI"

                details_html = (
                    f"<b>Manutenção:</b> {manut if manut else 'Não informado'}<br>"
                    f"<b>Service Desk:</b> {sdesk if sdesk else 'Não informado'}<br>"
                    f"<b>Infraestrutura:</b> {infra if infra else 'Não informado'}<br>"
                    f"<b>Desenvolvimento:</b> {dev if dev else 'Não informado'}"
                )
                
                if dt_ini:
                    events.append({
                        "title": f"🌙 Plantão Semanal: {display_name}",
                        "start": dt_ini,
                        "end": dt_fim if dt_fim else dt_ini,
                        "backgroundColor": "#8e44ad",
                        "borderColor": "#9b59b6",
                        "extendedProps": {
                            "categoria_evento": "plantao",
                            "servidor": f"Manutenção: {manut} | Service Desk: {sdesk} | Infra: {infra} | Dev: {dev}",
                            "detailsHtml": details_html,
                            "tipo": "Plantão Semanal SIMP",
                            "raw_data_inicio": raw_ini_br,
                            "raw_data_fim": raw_fim_br
                        }
                    })

    # 4. CONTRATOS DE GARANTIA
    if chk_garantias:
        df_garantia = get_garantia_contratos_df()
        if not df_garantia.empty:
            for idx, row in df_garantia.iterrows():
                contrato = str(row.get('contrato', '')).strip()
                pu_saj = str(row.get('pu_saj', '')).strip()
                item = str(row.get('item', '')).strip()
                fornecedor = str(row.get('fornecedor', '')).strip()
                nota_fiscal = str(row.get('nota_fiscal', '')).strip()
                status_garantia = str(row.get('status_garantia', '')).strip()
                link_suporte = str(row.get('link_suporte', '')).strip()

                iso_ini, br_ini = parse_date_to_iso_and_br(row.get('garantia_inicio'))
                if iso_ini:
                    events.append({
                        "id": f"garantia_ini_{idx}",
                        "title": f"🟢 Início Garantia: {item} ({fornecedor})",
                        "start": iso_ini,
                        "backgroundColor": "#3b82f6",
                        "borderColor": "#1d4ed8",
                        "extendedProps": {
                            "categoria_evento": "garantia",
                            "tipo": "🟢 Início da Garantia",
                            "contrato": contrato,
                            "pu_saj": pu_saj,
                            "item": item,
                            "fornecedor": fornecedor,
                            "nota_fiscal": nota_fiscal,
                            "status_garantia": status_garantia,
                            "data_formatada": br_ini,
                            "link_suporte": link_suporte
                        }
                    })

                iso_fim, br_fim = parse_date_to_iso_and_br(row.get('garantia_fim'))
                if iso_fim:
                    bg_col = "#ef4444" if status_garantia == "Garantia Vencida" else ("#f59e0b" if "30" in status_garantia else "#3b82f6")
                    events.append({
                        "id": f"garantia_fim_{idx}",
                        "title": f"🔴 Fim Garantia: {item} ({fornecedor})",
                        "start": iso_fim,
                        "backgroundColor": bg_col,
                        "borderColor": bg_col,
                        "extendedProps": {
                            "categoria_evento": "garantia",
                            "tipo": "🔴 Fim / Vencimento da Garantia",
                            "contrato": contrato,
                            "pu_saj": pu_saj,
                            "item": item,
                            "fornecedor": fornecedor,
                            "nota_fiscal": nota_fiscal,
                            "status_garantia": status_garantia,
                            "data_formatada": br_fim,
                            "link_suporte": link_suporte
                        }
                    })

    # 5. CHAMADOS (OTRS e CITSMART)
    if chk_otrs or chk_citsmart:
        df_chamados = load_data()
        if not df_chamados.empty:
            from src.database import get_comments_by_ticket
            for _, row in df_chamados.iterrows():
                base_raw = str(row.get('base', 'OTRS')).strip()
                base_upper = base_raw.upper()

                is_otrs = "OTRS" in base_upper
                is_citsmart = "CITSMART" in base_upper or "CIT" in base_upper

                if is_otrs and not chk_otrs:
                    continue
                if is_citsmart and not chk_citsmart:
                    continue
                if not is_otrs and not is_citsmart and not (chk_otrs or chk_citsmart):
                    continue

                base = "OTRS" if is_otrs else ("CitSmart" if is_citsmart else base_raw)

                cid = str(row.get('id', '')).strip()
                titulo = str(row.get('titulo', '')).strip()
                if not titulo or titulo.lower() in ["none", "nan", "null"]:
                    titulo = "Sem Título"

                status = str(row.get('status', 'Aberto')).strip()
                tag = str(row.get('tag', '')).strip()
                usuario = str(row.get('usuario', '')).strip()
                localidade = str(row.get('localidade_fisica', '')).strip()
                unidade = str(row.get('unidade', '')).strip()
                descricao = str(row.get('descricao', '')).strip()
                dt_criacao_raw = row.get('data_criacao')

                iso_dt, br_dt = parse_ticket_date_iso_and_br(dt_criacao_raw)
                if not iso_dt:
                    continue

                bg_col = "#0ea5e9" if is_otrs else "#f59e0b"
                border_col = "#0284c7" if is_otrs else "#d97706"

                desc_resumo = (descricao[:350] + "...") if len(descricao) > 350 else descricao
                comments_list = get_comments_by_ticket(cid)

                events.append({
                    "id": f"chamado_{cid}",
                    "title": f"📋 #{cid} ({base}): {titulo[:35]}",
                    "start": iso_dt,
                    "backgroundColor": bg_col,
                    "borderColor": border_col,
                    "extendedProps": {
                        "categoria_evento": "chamado",
                        "tipo": f"Chamado {base}",
                        "id": cid,
                        "base": base,
                        "status": status,
                        "tag": tag if tag else "Sem TAG",
                        "titulo_completo": titulo,
                        "usuario": usuario,
                        "localidade": localidade,
                        "unidade": unidade,
                        "data_criacao": br_dt if br_dt else str(dt_criacao_raw),
                        "descricao": desc_resumo if desc_resumo else "Sem descrição.",
                        "comentarios": comments_list if comments_list else []
                    }
                })

    # 6. PORTARIAS DA BANCADA (3 NÍVEIS: FÉRIAS, FISCAIS E GERAL)
    if chk_portarias or chk_ferias or chk_portarias_fiscais:
        try:
            portarias_list = fetch_portarias_bancada()
            for p_item in portarias_list:
                p_id = p_item.get("id")
                titulo = p_item.get("titulo", "")
                texto = p_item.get("texto", "")
                dt_emissao = p_item.get("data_emissao", "")
                dt_pub = p_item.get("data_publicacao", "")
                membros = ", ".join(p_item.get("membros", []))
                pdf_url = p_item.get("pdf_url", "")

                full_text = f"{titulo} {texto}".lower()

                # CLASSIFICAÇÃO EM 3 NÍVEIS
                is_ferias = "férias" in full_text or "ferias" in full_text
                
                is_fiscal = False
                if not is_ferias:
                    keywords_fiscal = ["fiscalização", "fiscalizacao", "fiscal", "fiscais", "nota de empenho", "gestão", "gestao", "contrato"]
                    is_fiscal = any(kw in full_text for kw in keywords_fiscal)

                # FILTRAGEM POR CHECKBOX
                if is_ferias and not chk_ferias:
                    continue
                if is_fiscal and not chk_portarias_fiscais:
                    continue
                if not is_ferias and not is_fiscal and not chk_portarias:
                    continue

                # Converte data de emissão para ISO
                iso_date_emissao, br_date_emissao = None, None
                if dt_emissao:
                    try:
                        dt_obj = datetime.strptime(dt_emissao, "%d/%m/%Y")
                        iso_date_emissao = dt_obj.strftime("%Y-%m-%d")
                        br_date_emissao = dt_emissao
                    except Exception:
                        pass

                if not iso_date_emissao:
                    continue

                start_dt, end_dt = iso_date_emissao, iso_date_emissao
                start_br, end_br = br_date_emissao, br_date_emissao

                # Se for Férias, tenta extrair o período via Regex (com fallback)
                if is_ferias:
                    r_start_iso, r_end_iso, r_start_br, r_end_br = extract_vacation_dates(full_text, br_date_emissao)
                    if r_start_iso and r_end_iso:
                        start_dt, end_dt = r_start_iso, r_end_iso
                        start_br, end_br = r_start_br, r_end_br

                # CORES E LABELS POR NÍVEL
                if is_ferias:
                    bg_col, border_col = "#0d9488", "#0f766e"
                    event_type_label = "🏖️ Portaria de Férias"
                    icon = "🏖️"
                elif is_fiscal:
                    bg_col, border_col = "#8b5cf6", "#7c3aed"
                    event_type_label = "📜 Fiscalização de Contrato"
                    icon = "📜"
                else:
                    bg_col, border_col = "#475569", "#334155"
                    event_type_label = "📜 Portaria Geral"
                    icon = "📜"

                ementa_snippet = (texto[:300] + "...") if len(texto) > 300 else (texto if texto else titulo)

                events.append({
                    "id": f"portaria_{p_id}",
                    "title": f"{icon} Portaria #{p_item.get('numero', p_id)}: {membros.split(',')[0]}",
                    "start": start_dt,
                    "end": end_dt,
                    "backgroundColor": bg_col,
                    "borderColor": border_col,
                    "extendedProps": {
                        "categoria_evento": "portaria",
                        "tipo": event_type_label,
                        "is_ferias": is_ferias,
                        "is_fiscal": is_fiscal,
                        "titulo": titulo,
                        "membros": membros,
                        "data_emissao": start_br,
                        "data_publicacao": dt_pub if dt_pub else start_br,
                        "ementa": ementa_snippet,
                        "pdf_url": pdf_url,
                        "raw_data_inicio": start_br,
                        "raw_data_fim": end_br
                    }
                })
        except Exception as e:
            pass

    # --- APLICA PESQUISA GLOBAL (FILTRO DE EVENTOS) ---
    filtered_events = events
    if search_query:
        matching_events = []
        for ev in events:
            # Varre o título e todas as extendedProps em busca da palavra-chave
            ev_title = str(ev.get("title", "")).lower()
            props = ev.get("extendedProps", {})
            props_str_list = []
            for v in props.values():
                if isinstance(v, list):
                    props_str_list.append(" ".join([str(item).lower() for item in v if item]))
                elif v:
                    props_str_list.append(str(v).lower())
            props_str = " ".join(props_str_list)
            
            combined_search_text = f"{ev_title} {props_str}"
            if search_query in combined_search_text:
                matching_events.append(ev)
        filtered_events = matching_events

    render_master_calendar(filtered_events)

    # --- TABELA DE RESULTADOS DA PESQUISA (EXIBIDA SE HOUVER BUSCA) ---
    if search_query:
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader(f"🔎 Resultados da Pesquisa para '{search_query}' ({len(filtered_events)} eventos encontrados)")
        
        if not filtered_events:
            st.info("Nenhum evento encontrado para a palavra-chave informada.")
        else:
            table_rows = []
            for ev in filtered_events:
                props = ev.get("extendedProps", {})
                
                # Monta detalhes amigáveis por categoria
                cat = props.get("categoria_evento", "")
                detalhes = ""
                if cat == "chamado":
                    detalhes = f"Solicitante: {props.get('usuario', 'N/A')} | Local: {props.get('localidade', 'N/A')}"
                elif cat == "plantao":
                    detalhes = props.get("servidor", "")
                elif cat == "garantia":
                    detalhes = f"Contrato: {props.get('contrato', 'N/A')} | Fornecedor: {props.get('fornecedor', 'N/A')}"
                elif cat == "portaria":
                    detalhes = f"Membros: {props.get('membros', 'N/A')} | Ementa: {props.get('ementa', '')[:100]}..."
                elif cat == "manual":
                    detalhes = f"Autor: {props.get('autor', 'N/A')} | {props.get('descricao', '')}"

                dt_exibicao = props.get("raw_data_inicio") or str(ev.get("start", ""))

                table_rows.append({
                    "Título": ev.get("title", ""),
                    "Categoria / Tipo": props.get("tipo", "Evento"),
                    "Data / Período": dt_exibicao,
                    "Detalhes / Resumo": detalhes
                })

            df_search_res = pd.DataFrame(table_rows)
            st.dataframe(
                df_search_res,
                column_config={
                    "Título": st.column_config.TextColumn("Título do Evento"),
                    "Categoria / Tipo": st.column_config.TextColumn("Categoria / Tipo"),
                    "Data / Período": st.column_config.TextColumn("Data / Período"),
                    "Detalhes / Resumo": st.column_config.TextColumn("Detalhes / Resumo"),
                },
                hide_index=True,
                width='stretch'
            )
