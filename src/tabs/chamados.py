import sys
import time
from pathlib import Path

root_dir = Path(__file__).parent.parent.parent

import asyncio
import os
import re
from datetime import datetime
import pandas as pd
import sqlite3
import logging
import streamlit as st
import streamlit.components.v1 as components

logger = logging.getLogger(__name__)

from src.components.status_banner import check_orquestrador_running, read_last_log_lines, render_log_expander
from src.database import (
    load_data,
    get_comments_by_ticket,
    update_ticket_location_details,
    update_ticket_tag,
    update_ticket_andamento,
    update_ticket_title,
    update_ticket_device_info,
    update_ticket_client_id,
    get_ticket_by_id
)
from src.services.ticket_enrichment import (
    format_br_datetime,
    resolve_ticket_device_and_user,
    sanitize_network_value,
)
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.metric_cards import render_metric_cards

from src.config import CITSMART_URL, OTRS_URL
from src.database.connection import DB_PATH

TAG_COLORS = {
    "BACKUP": "#dd5358",
    "EVENTO": "#ce66ce",
    "FORMATAÇÃO": "#d38a62",
    "GARANTIA": "#518bbb",
    "IMPRESSORA": "#C6EFCE",
    "INSTALAÇÃO HARDWARE": "#FCE4D6",
    "INSTALAÇÃO SOFTWARE": "#86BEEE",
    "MANUTENÇÃO": "#E9CF69",
    "MONITOR": "#cbdd6f",
    "MUDANÇA": "#21ffe0",
    "PREPARAÇÃO COMPUTADORES": "#f09c72",
    "REDE": "#B7F391",
    "SOLICITAÇÃO SSD": "#f5a89b",
    "SUPORTE": "#FFE699",
    "TELEFONIA FIXA": "#e273a1",
    "VIAGEM": "#61e7c6",
    "VISTORIA CPDS": "#b2740e",
}




def _on_modal_dismiss():
    """Callback disparado automaticamente quando o modal é fechado (pelo X, botão Fechar ou clicando fora)."""
    curr = st.session_state.get("active_ticket_modal")
    if curr:
        st.session_state["dismissed_ticket_id"] = curr
    st.session_state["active_ticket_modal"] = None
    if st.session_state.get("ticket_data_dirty"):
        st.session_state["ticket_data_dirty"] = False
        st.cache_data.clear()


@st.dialog("Detalhes do Chamado", width="large", on_dismiss=_on_modal_dismiss)
def show_ticket_details(row_or_id):
    """Exibe o modal com todos os detalhes e ações rápidas do chamado."""
    if isinstance(row_or_id, dict):
        ticket_id = str(row_or_id.get('id', '')).strip()
    else:
        ticket_id = str(row_or_id).strip()

    # Busca os dados mais recentes diretamente do banco
    fresh_row = get_ticket_by_id(ticket_id)
    if not fresh_row and isinstance(row_or_id, dict):
        fresh_row = row_or_id
    elif not fresh_row:
        st.error(f"Chamado #{ticket_id} não encontrado no banco de dados.")
        if st.button("Fechar", key="close_modal_err_btn", width="stretch"):
            _on_modal_dismiss()
            st.rerun()
        return

    row = fresh_row
    cid_str = str(row.get('id', '')).strip()
    title = str(row.get('titulo', '')).strip()
    if title and title.lower() not in ["none", "nan", "null", ""]:
        header_text = f"🎫 Chamado #{cid_str} – {title}"
    else:
        header_text = f"🎫 Chamado #{cid_str}"

    def _sanitize_val(v):
        if v is None or pd.isna(v):
            return ""
        s = str(v).strip()
        if s.lower() in ["nan", "none", "null", "<na>"]:
            return ""
        return s

    comments = get_comments_by_ticket(cid_str)
    resolved = resolve_ticket_device_and_user(
        row,
        comments=comments,
        allow_network_fallbacks=True,
    )

    sam_account = resolved.get("sam_account_name", "")
    curr_ip = resolved.get("ip_origem", "")
    curr_host = resolved.get("hostname", "")

    # Persiste silenciosamente novos dados descobertos sem alterar data_atualizacao
    orig_client_id = sanitize_network_value(row.get('id_cliente'))
    orig_ip = sanitize_network_value(row.get('ip_origem'))
    orig_host = sanitize_network_value(row.get('hostname'))

    if sam_account and not orig_client_id and cid_str:
        try:
            update_ticket_client_id(cid_str, sam_account, touch_updated_at=False)
        except Exception as cli_err:
            logger.debug(f"Não foi possível persistir id_cliente do chamado {cid_str}: {cli_err}")

    if ((curr_ip and curr_ip != orig_ip) or (curr_host and curr_host != orig_host)) and cid_str:
        try:
            update_ticket_device_info(cid_str, curr_ip, curr_host, touch_updated_at=False)
        except Exception as dev_err:
            logger.debug(f"Não foi possível atualizar device info do chamado {cid_str}: {dev_err}")

    criado_br = format_br_datetime(
        row.get('data_criacao') or row.get('Data Formatada') or row.get('datetime_obj'),
        fallback="N/A",
    )
    atualizado_br = format_br_datetime(row.get('data_atualizacao'), fallback="N/A")

    from src.auth import is_admin
    admin_mode = is_admin()

    with st.expander(header_text, expanded=True):
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### 👤 Informações do Usuário")
            raw_user = _sanitize_val(row.get('usuario')) or resolved.get('display_name') or 'N/A'
            user_display = raw_user
            if sam_account:
                if raw_user == 'N/A':
                    user_display = f"`{sam_account}`"
                elif raw_user.lower() != sam_account.lower() and f"({sam_account.lower()})" not in raw_user.lower():
                    user_display = f"{raw_user} (`{sam_account}`)"

            clean_loc = _sanitize_val(row.get('localidade_fisica'))
            if not clean_loc or clean_loc.lower() in ["nan", "none", "null", "não identificada"]:
                clean_loc = "Não identificada"
            elif "nan -" in clean_loc.lower() or "- nan" in clean_loc.lower():
                clean_loc = "Não identificada"

            st.markdown(f"**Usuário:** {user_display}")
            st.markdown(f"**Localidade:** {clean_loc}")
            st.markdown(f"**IP de Origem:** `{curr_ip or 'N/A'}`")
            st.markdown(f"**Hostname:** `{curr_host or 'N/A'}`")

            target_host = curr_host or curr_ip
            if target_host and target_host != "N/A":
                col_act1, col_act2 = st.columns(2)
                with col_act1:
                    cmrc_url = f"bancada://run?tool=cmrc&host={target_host}"
                    st.link_button("🎮 CmRcViewer (SCCM)", url=cmrc_url, width="stretch", help="Abre o Controle Remoto do SCCM.")
                with col_act2:
                    rdp_url = f"bancada://run?tool=rdp&host={target_host}"
                    st.link_button("🖥️ MSTSC (RDP)", url=rdp_url, width="stretch", help="Conecta via Área de Trabalho Remota.")

        with col2:
            st.markdown("### ⚙️ Classificação & Status")
            base_val = _sanitize_val(row.get('base')) or 'N/A'
            status_val = _sanitize_val(row.get('status')) or 'Aberto'
            st.markdown(f"**Base de Origem:** `{base_val}` &nbsp;|&nbsp; **Status:** `{status_val}`")
            st.markdown(f"**Criado em:** `{criado_br}`")
            st.markdown(f"**Atualizado em:** `{atualizado_br}`")

            tag_name = str(row.get('tag', '')).upper().strip()
            bg_color = TAG_COLORS.get(tag_name, "#262730")

            hex_color = bg_color.lstrip('#')
            try:
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                text_color = "#ffffff" if luminance < 0.6 else "#212529"
            except Exception:
                text_color = "#ffffff"

            tag_html = f'<span style="background-color: {bg_color}; color: {text_color}; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-family: inherit; font-size: 13px;">{row.get("tag", "")}</span>'
            st.markdown(f"**TAG Atual:** {tag_html}", unsafe_allow_html=True)

            if admin_mode:
                tag_options = sorted(list(TAG_COLORS.keys()))
                try:
                    default_idx = tag_options.index(tag_name)
                except ValueError:
                    default_idx = 0

                tag_key = f"select_tag_{cid_str}"
                new_tag = st.selectbox("🏷️ Alterar TAG Manualmente", options=tag_options, index=default_idx, key=tag_key)

                if new_tag != tag_name:
                    if st.button("💾 Salvar Nova TAG", key=f"save_tag_btn_{cid_str}"):
                        update_ticket_tag(cid_str, new_tag)
                        st.cache_data.clear()
                        st.success(f"✅ TAG alterada com sucesso para {new_tag}!")
                        st.toast(f"✅ TAG alterada para {new_tag}!", icon="🏷️")

            link_url = row.get('link')
            if link_url:
                link_url = str(link_url).strip()
                if not link_url.startswith("http://") and not link_url.startswith("https://"):
                    otrs_base = (OTRS_URL).rstrip('/')
                    link_url = f"{otrs_base}/{link_url.lstrip('/')}"
                st.link_button("🔗 Abrir Chamado Original", link_url, width="stretch")

    if admin_mode:
        edit_col1, edit_col2 = st.columns(2)
        with edit_col1:
            with st.expander("✏️ Editar Título do Chamado", expanded=False):
                curr_title = str(row.get('titulo', '')).strip()
                if curr_title.lower() in ["none", "nan", "null", "sem título"]:
                    curr_title = ""
                new_titulo = st.text_input("Título do Chamado:", value=curr_title, key=f"edit_titulo_{cid_str}")
                if st.button("💾 Salvar Título", key=f"save_title_btn_{cid_str}", width="stretch"):
                    update_ticket_title(cid_str, new_titulo)
                    st.cache_data.clear()
                    st.success("✅ Título atualizado com sucesso!")
                    st.toast("✅ Título salvo com sucesso!", icon="💾")

        with edit_col2:
            with st.expander("📍 Editar Localização Manual", expanded=False):
                curr_cp = _sanitize_val(row.get('cidade_predio'))
                curr_un = _sanitize_val(row.get('unidade'))
                if "não encontrad" in curr_un.lower() or "nao encontrad" in curr_un.lower():
                    curr_un = ""
                curr_loc = _sanitize_val(row.get('localidade_fisica'))
                if "nan" in curr_loc.lower() or "não encontrad" in curr_loc.lower():
                    curr_loc = ""

                new_cidade = st.text_input("Cidade - Prédio", value=curr_cp, key=f"edit_cidade_{cid_str}")
                new_unidade = st.text_input("Unidade", value=curr_un, key=f"edit_unidade_{cid_str}")
                new_localidade = st.text_input("Localidade Física", value=curr_loc, key=f"edit_localidade_{cid_str}")

                if st.button("💾 Salvar Localização", key=f"save_loc_btn_{cid_str}", width="stretch"):
                    update_ticket_location_details(cid_str, new_localidade, new_cidade, new_unidade)
                    st.cache_data.clear()
                    st.success("✅ Localização salva com sucesso!")
                    st.toast("✅ Localização salva com sucesso!", icon="💾")

    with st.expander("📝 Andamento / Nota de Atendimento", expanded=True):
        current_andamento = str(row.get('andamento', '')).strip()
        if current_andamento.lower() in ["none", "nan", "null", ""]:
            current_andamento = ""
        if admin_mode:
            andamento_key = f"andamento_modal_ta_{cid_str}"
            new_andamento = st.text_area(
                "Nota rápida sobre o andamento do chamado:",
                value=current_andamento,
                key=andamento_key
            )

            if st.button("💾 Salvar Nota de Andamento", key=f"save_andamento_modal_btn_{cid_str}"):
                update_ticket_andamento(cid_str, new_andamento)
                st.cache_data.clear()
                st.success("✅ Nota de andamento salva com sucesso!")
                st.toast("✅ Nota de andamento salva com sucesso!", icon="💾")
        else:
            st.write(current_andamento if current_andamento else "*(Nenhuma nota de andamento registrada)*")

    desc_header = f"📝 #1 – {criado_br} (Descrição)" if criado_br != "N/A" else "📝 #1 – (Descrição)"
    with st.expander(desc_header, expanded=True):
        st.text(row.get('descricao', ''))

    if comments:
        st.markdown("### 💬 Histórico de Notas e Acompanhamentos")
        for i, c in enumerate(comments, start=2):
            c_data_br = format_br_datetime(c.get('data', ''), fallback="")
            c_data_txt = c_data_br if c_data_br else str(c.get('data', '')).strip()
            header = f"🕒 #{i} – {c_data_txt} – por {c.get('autor', '')}"
            with st.expander(header):
                st.text(c.get('texto', ''))
        
    if st.button("Fechar", key="close_modal_btn", width="stretch"):
        _on_modal_dismiss()
        st.rerun()





def render_chamados_page():
    """Renderiza a página principal do Painel de Chamados."""
    col_title, col_btn = st.columns([3, 1])
    with col_title:
        st.title("📊 Painel de Chamados Centralizado")
        st.write("Visualize e interaja com os chamados do OTRS e CitSmart.")

    robo_ativo = check_orquestrador_running()

    with col_btn:
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        from src.auth import is_admin
        if is_admin():
            if robo_ativo:
                st.button("🤖 Robô em Execução...", width='stretch', disabled=True)
            else:
                run_orquestrador = st.button(
                    "🔄 Atualizar Chamados", 
                    width='stretch', 
                    help="Executa o orquestrador completo em segundo plano.",
                    type="primary"
                )
                if run_orquestrador:
                    import time
                    from src.services.cron_scheduler import get_cron_daemon
                    cron_daemon = get_cron_daemon()
                    if cron_daemon.trigger_task_now("orquestrador_chamados"):
                        time.sleep(0.8)
                        st.toast("🚀 Robô iniciado em segundo plano!", icon="🤖")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.warning("O robô já está em execução neste momento.")

    if "was_orquestrador_running" not in st.session_state:
        st.session_state["was_orquestrador_running"] = False

    if st.session_state["was_orquestrador_running"] and not robo_ativo:
        st.session_state["was_orquestrador_running"] = False
        try:
            import json as _json
            from src.config import DEBUG_DIR_SYNC
            summary_file = DEBUG_DIR_SYNC / "last_run_summary.json"
            new_cnt = 0
            if summary_file.exists():
                new_cnt = int(_json.loads(summary_file.read_text(encoding="utf-8")).get("new_tickets", 0))
            if new_cnt > 0:
                st.toast(f"🆕 {new_cnt} chamado(s) novo(s)!", icon="🎫")
            else:
                st.toast("✅ Sincronização concluída: nenhum chamado novo.", icon="ℹ️")
        except Exception:
            st.toast("✅ Sincronização de chamados concluída!", icon="🤖")

    if robo_ativo:
        st.session_state["was_orquestrador_running"] = True

    render_log_expander(
        "🤖 Robô Rodando em Segundo Plano – Acompanhar Progresso",
        robo_ativo,
        read_last_log_lines,
        check_orquestrador_running,
        "O robô está coletando novos chamados e classificando com IA neste momento. Você pode continuar usando o painel normalmente!"
    )

    df = load_data()

    if df.empty:
        st.warning("Nenhum dado encontrado no banco de dados. Execute o orquestrador primeiro!")
        return

    df['datetime_obj'] = pd.to_datetime(df['data_criacao'], format='mixed', errors='coerce')
    df['Data Formatada'] = df['datetime_obj'].dt.strftime('%d/%m/%Y %H:%M:%S').fillna(df['data_criacao'])

    # Ordenação padrão decrescente pela data correta
    df = df.sort_values(by='datetime_obj', ascending=False)

    min_date = df['datetime_obj'].dropna().min().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    max_date = df['datetime_obj'].dropna().max().date() if not df['datetime_obj'].dropna().empty else datetime.now().date()
    
    if "f_date_range" not in st.session_state:
        st.session_state["f_date_range"] = (min_date, max_date)
    if "f_status" not in st.session_state:
        st.session_state["f_status"] = []
    if "f_tags" not in st.session_state:
        st.session_state["f_tags"] = []
    if "custom_loc_selection" not in st.session_state:
        st.session_state["custom_loc_selection"] = []
    if "f_cities" not in st.session_state:
        st.session_state["f_cities"] = []
    if "f_units" not in st.session_state:
        st.session_state["f_units"] = []
    if "f_bases" not in st.session_state:
        st.session_state["f_bases"] = []
    if "f_user" not in st.session_state:
        st.session_state["f_user"] = ""
    if "f_mode" not in st.session_state:
        st.session_state["f_mode"] = "🟢 Manter Selecionados"
    if "f_ticket_ids" not in st.session_state:
        st.session_state["f_ticket_ids"] = []

    def get_filtered_options(col_name: str) -> list:
        temp_df = df.copy()
        is_exclude_mode = (st.session_state.get("f_mode") == "🔴 Ocultar Selecionados")
        
        dr = st.session_state.get("f_date_range", (min_date, max_date))
        if isinstance(dr, tuple) and len(dr) == 2:
            start_date, end_date = dr
            temp_df = temp_df[
                (temp_df['datetime_obj'].dt.date >= start_date) & 
                (temp_df['datetime_obj'].dt.date <= end_date)
            ]
            
        if col_name != 'status' and st.session_state.get("f_status"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['status'].isin(st.session_state["f_status"])]
            else:
                temp_df = temp_df[temp_df['status'].isin(st.session_state["f_status"])]
            
        if col_name != 'tag' and st.session_state.get("f_tags"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['tag'].isin(st.session_state["f_tags"])]
            else:
                temp_df = temp_df[temp_df['tag'].isin(st.session_state["f_tags"])]
            
        if col_name != 'localidade_fisica' and st.session_state.get("custom_loc_selection"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['localidade_fisica'].isin(st.session_state["custom_loc_selection"])]
            else:
                temp_df = temp_df[temp_df['localidade_fisica'].isin(st.session_state["custom_loc_selection"])]
            
        if col_name != 'cidade_predio' and st.session_state.get("f_cities"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['cidade_predio'].isin(st.session_state["f_cities"])]
            else:
                temp_df = temp_df[temp_df['cidade_predio'].isin(st.session_state["f_cities"])]
            
        if col_name != 'unidade' and st.session_state.get("f_units"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['unidade'].isin(st.session_state["f_units"])]
            else:
                temp_df = temp_df[temp_df['unidade'].isin(st.session_state["f_units"])]
            
        if col_name != 'base' and st.session_state.get("f_bases"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['base'].isin(st.session_state["f_bases"])]
            else:
                temp_df = temp_df[temp_df['base'].isin(st.session_state["f_bases"])]
            
        if col_name != 'usuario' and st.session_state.get("f_user"):
            if is_exclude_mode:
                temp_df = temp_df[~temp_df['usuario'].str.contains(st.session_state["f_user"], case=False, na=False)]
            else:
                temp_df = temp_df[temp_df['usuario'].str.contains(st.session_state["f_user"], case=False, na=False)]
            
        options = sorted(list(temp_df[col_name].dropna().unique()))
        
        key_map = {
            'status': 'f_status',
            'tag': 'f_tags',
            'localidade_fisica': 'custom_loc_selection',
            'cidade_predio': 'f_cities',
            'unidade': 'f_units',
            'base': 'f_bases',
            'usuario': 'f_user'
        }
        session_key = key_map.get(col_name)
        if session_key:
            current_selection = st.session_state.get(session_key, [])
            if current_selection:
                if isinstance(current_selection, list):
                    for val in current_selection:
                        if val not in options:
                            options.append(val)
                elif current_selection not in options:
                    options.append(current_selection)
                    
        return options

    st.sidebar.markdown("### 🔍 Filtros de Chamados")
    
    filter_mode = st.sidebar.radio(
        "Modo de Filtragem:",
        options=["🟢 Manter Selecionados", "🔴 Ocultar Selecionados"],
        key="f_mode",
        horizontal=True,
        help="🟢 Manter: Mostra apenas os itens escolhidos.\n🔴 Ocultar: Exibe tudo EXCETO os itens escolhidos."
    )
    
    status_options = get_filtered_options('status')
    tag_options = get_filtered_options('tag')
    city_options = get_filtered_options('cidade_predio')
    unit_options = get_filtered_options('unidade')
    base_options = get_filtered_options('base')

    col_btn_sel1, col_btn_sel2 = st.sidebar.columns(2)
    with col_btn_sel1:
        if st.button("☑️ Marcar Todos", width='stretch', help="Seleciona todas as opções"):
            st.session_state["f_status"] = list(status_options)
            st.session_state["f_tags"] = list(tag_options)
            st.session_state["f_cities"] = list(city_options)
            st.session_state["f_units"] = list(unit_options)
            st.session_state["f_bases"] = list(base_options)
            st.rerun()
    with col_btn_sel2:
        if st.button("🧹 Limpar Tudo", width='stretch', help="Limpa todos os filtros ativos"):
            st.session_state["f_date_range"] = (min_date, max_date)
            st.session_state["f_status"] = []
            st.session_state["f_tags"] = []
            st.session_state["custom_loc_selection"] = []
            st.session_state["f_cities"] = []
            st.session_state["f_units"] = []
            st.session_state["f_bases"] = []
            st.session_state["f_user"] = ""
            st.session_state["f_ticket_ids"] = []
            st.rerun()

    st.sidebar.markdown("---")

    date_range = st.sidebar.date_input(
        "Intervalo de Datas",
        value=st.session_state["f_date_range"],
        min_value=min_date,
        max_value=max_date,
        format="DD/MM/YYYY",
        key="f_date_range"
    )
    
    selected_status = st.sidebar.multiselect(
        "Status", 
        options=status_options, 
        key="f_status", 
        placeholder="Escolha as opções..."
    )
    
    selected_tags = st.sidebar.multiselect(
        "TAG (Categoria de IA)", 
        options=tag_options, 
        key="f_tags", 
        placeholder="Escolha as opções..."
    )
    
    loc_options = get_filtered_options('localidade_fisica')
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📍 Seleção de Localidades")
    
    selected_locs = st.sidebar.multiselect(
        "Localidade Física",
        options=loc_options,
        default=st.session_state.get("custom_loc_selection", []),
        key="custom_loc_selection",
        placeholder="Escolha as localidades..."
    )
    
    selected_cities = st.sidebar.multiselect(
        "Cidade - Prédio", 
        options=city_options, 
        key="f_cities", 
        placeholder="Escolha as opções..."
    )
    
    selected_units = st.sidebar.multiselect(
        "Unidade", 
        options=unit_options, 
        key="f_units", 
        placeholder="Escolha as opções..."
    )
    
    selected_bases = st.sidebar.multiselect(
        "Base de Origem", 
        options=base_options, 
        key="f_bases", 
        placeholder="Escolha as opções..."
    )
    
    user_search = st.sidebar.text_input(
        "Buscar por Usuário", 
        key="f_user", 
        placeholder="Digite o nome do usuário..."
    )

    filtered_df = df.copy()
    is_exclude_mode = (filter_mode == "🔴 Ocultar Selecionados")
    
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['datetime_obj'].dt.date >= start_date) & 
            (filtered_df['datetime_obj'].dt.date <= end_date)
        ]
        
    if selected_status:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['status'].isin(selected_status)]
        else:
            filtered_df = filtered_df[filtered_df['status'].isin(selected_status)]

    if selected_tags:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['tag'].isin(selected_tags)]
        else:
            filtered_df = filtered_df[filtered_df['tag'].isin(selected_tags)]

    if selected_locs:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['localidade_fisica'].isin(selected_locs)]
        else:
            filtered_df = filtered_df[filtered_df['localidade_fisica'].isin(selected_locs)]

    if selected_cities:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['cidade_predio'].isin(selected_cities)]
        else:
            filtered_df = filtered_df[filtered_df['cidade_predio'].isin(selected_cities)]

    if selected_units:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['unidade'].isin(selected_units)]
        else:
            filtered_df = filtered_df[filtered_df['unidade'].isin(selected_units)]

    if selected_bases:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['base'].isin(selected_bases)]
        else:
            filtered_df = filtered_df[filtered_df['base'].isin(selected_bases)]

    if user_search:
        if is_exclude_mode:
            filtered_df = filtered_df[~filtered_df['usuario'].str.contains(user_search, case=False, na=False)]
        else:
            filtered_df = filtered_df[filtered_df['usuario'].str.contains(user_search, case=False, na=False)]

    st.sidebar.markdown("---")
    
    source_df = filtered_df if not filtered_df.empty else df
    df_tickets_list = source_df[['id', 'titulo']].copy()
    ids_clean = df_tickets_list['id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    titles_clean = df_tickets_list['titulo'].fillna("Sem Título").astype(str).str.strip()
    titles_clean = titles_clean.replace(["nan", "NaN", "None", "null", ""], "Sem Título")
    labels = ids_clean + " - " + titles_clean
    ticket_options = sorted(labels.unique().tolist())
    
    current_sel_tickets = st.session_state.get("f_ticket_ids", [])
    for sel_t in current_sel_tickets:
        if sel_t not in ticket_options:
            ticket_options.append(sel_t)

    selected_tickets = st.sidebar.multiselect(
        "🎫 Chamados Específicos (por ID / Título)",
        options=ticket_options,
        key="f_ticket_ids",
        placeholder="Selecione chamados individuais..."
    )

    items_per_page = render_items_per_page_selector(
        key_prefix="chamados_v5",
        options=[10, 50, 100, "Todos"],
        default_index=2,
        label="📄 Chamados por página:"
    )

    st.sidebar.markdown("---")


    if selected_tickets:
        selected_ids = [t.split(" - ")[0].strip() for t in selected_tickets]
        clean_df_ids = filtered_df['id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        if is_exclude_mode:
            filtered_df = filtered_df[~clean_df_ids.isin(selected_ids)]
        else:
            filtered_df = filtered_df[clean_df_ids.isin(selected_ids)]
        
    total_ch = len(filtered_df)
    abertos_ch = len(filtered_df[filtered_df['status'] == 'Aberto']) if 'status' in filtered_df.columns else 0
    fechados_ch = len(filtered_df[filtered_df['status'] == 'Fechado']) if 'status' in filtered_df.columns else 0

    render_metric_cards([
        {
            "title": "TOTAL DE CHAMADOS",
            "value": total_ch,
            "border_color": "#3b82f6",
        },
        {
            "title": "ABERTOS",
            "value": abertos_ch,
            "border_color": "#f59e0b",
            "value_color": "#f59e0b",
        },
        {
            "title": "FECHADOS",
            "value": fechados_ch,
            "border_color": "#10b981",
            "value_color": "#10b981",
        },
    ])


    
    # ABAS INTERNAS DE VISUALIZAÇÃO COM QUERY PARAMETERS (?subtab=slug)
    CHAMADOS_SUBTAB_MAP = {
        "tabela": "📋 Tabela Geral de Chamados",
        "graficos": "📈 Gráficos & Estatísticas do Painel"
    }
    CHAMADOS_SUBTAB_REVERSE = {v: k for k, v in CHAMADOS_SUBTAB_MAP.items()}

    url_subtab = st.query_params.get("subtab", "tabela")
    default_title = CHAMADOS_SUBTAB_MAP.get(url_subtab, "📋 Tabela Geral de Chamados")
    options = list(CHAMADOS_SUBTAB_MAP.values())
    default_idx = options.index(default_title) if default_title in options else 0

    selected_subtab = st.radio(
        "Navegação do Painel:",
        options=options,
        index=default_idx,
        horizontal=True,
        label_visibility="collapsed",
        key="chamados_subtab_radio"
    )

    new_slug = CHAMADOS_SUBTAB_REVERSE.get(selected_subtab, "tabela")
    if st.query_params.get("subtab") != new_slug:
        st.query_params["subtab"] = new_slug

    st.markdown("<br>", unsafe_allow_html=True)

    if selected_subtab == "📈 Gráficos & Estatísticas do Painel":

        st.subheader("📊 Análise Estatística dos Chamados Filtrados")
        st.write("Visualização consolidada de abertura de chamados por Prédio, Unidade, Categorias (TAGs) e Usuários.")
        st.markdown("<br>", unsafe_allow_html=True)

        if not filtered_df.empty:
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("#### 🏢 Top Prédios / Cidades com Mais Chamados")
                city_counts = filtered_df['cidade_predio'].value_counts().head(10).reset_index()
                city_counts.columns = ['Prédio / Cidade', 'Quantidade']
                st.bar_chart(city_counts.set_index('Prédio / Cidade'), width='stretch')

            with g_col2:
                st.markdown("#### 🏛️ Top Unidades / Setores Mais Demandantes")
                unit_counts = filtered_df['unidade'].value_counts().head(10).reset_index()
                unit_counts.columns = ['Unidade / Setor', 'Quantidade']
                st.bar_chart(unit_counts.set_index('Unidade / Setor'), width='stretch')

            st.markdown("---")

            g_col3, g_col4 = st.columns(2)
            with g_col3:
                st.markdown("#### 🏷️ Distribuição por Categoria (TAG de IA)")
                tag_counts = filtered_df['tag'].value_counts().reset_index()
                tag_counts.columns = ['Categoria (TAG)', 'Quantidade']
                st.bar_chart(tag_counts.set_index('Categoria (TAG)'), width='stretch')

            with g_col4:
                st.markdown("#### 👤 Top Usuários que Mais Abrem Chamados")
                user_counts = filtered_df['usuario'].value_counts().head(10).reset_index()
                user_counts.columns = ['Usuário', 'Quantidade']
                st.bar_chart(user_counts.set_index('Usuário'), width='stretch')
                
            st.markdown("---")

            g_col5, g_col6 = st.columns(2)
            with g_col5:
                st.markdown("#### 🔄 Origem dos Chamados (Base)")
                base_counts = filtered_df['base'].value_counts().reset_index()
                base_counts.columns = ['Base de Origem', 'Quantidade']
                st.bar_chart(base_counts.set_index('Base de Origem'), width='stretch')

            with g_col6:
                st.markdown("#### 📍 Status por Prédio / Cidade (Abertos x Fechados)")
                status_city = filtered_df.groupby(['cidade_predio', 'status']).size().unstack(fill_value=0)
                st.bar_chart(status_city.head(10), width='stretch')
        else:
            st.info("Sem chamados no filtro selecionado para renderizar gráficos.")

    else:

        cols_to_show = [
            'id', 'status', 'tag', 'andamento', 'localidade_fisica', 
            'cidade_predio', 'unidade', 'usuario', 'datetime_obj', 'base'
        ]
            
        cols_to_generate = list(cols_to_show)
        if 'link' not in cols_to_generate and 'link' in filtered_df.columns:
            cols_to_generate.append('link')
        if 'ip_origem' not in cols_to_generate and 'ip_origem' in filtered_df.columns:
            cols_to_generate.append('ip_origem')
            
        df_display = filtered_df[cols_to_generate].copy()
        
        def format_id_link(row):
            cid = str(row['id']).strip()
            link = str(row.get('link', '')).strip() if 'link' in row else ''
            citsmart_base = (CITSMART_URL).rstrip('/')
            otrs_base = (OTRS_URL).rstrip('/')
            if not link or link.lower() in ["none", "nan", "null", ""]:
                if row.get('base') == 'CitSmart':
                    link = f"{citsmart_base}/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}"
                else:
                    link = f"{otrs_base}/otrs/index.pl"
            elif not link.startswith("http://") and not link.startswith("https://"):
                link = f"{otrs_base}/{link.lstrip('/')}"
            return f"{link}#id:{cid}"

        df_display['id'] = df_display.apply(format_id_link, axis=1)
        
        cols_for_dataframe = list(cols_to_show)
        if 'ip_origem' in df_display.columns:
            df_display['ip_origem'] = df_display['ip_origem'].fillna("").astype(str).replace(r'^(?i:nan|none|null|<na>)$', '', regex=True).str.strip()
            cols_for_dataframe.append('ip_origem')
            
        df_final_display = df_display[cols_for_dataframe]
        
        col_tbl_head, col_tbl_html, col_tbl_cap = st.columns([2.6, 1.2, 1.5])

        with col_tbl_head:
            st.subheader(f"📋 Lista de Chamados ({len(filtered_df)} registros)")
            st.write("Dica: Clique no **checkbox (caixinha de seleção)** no início de qualquer linha para abrir os Detalhes no Modal.")
        with col_tbl_html:
            st.write("")
            import hashlib
            _export_sig = hashlib.md5(
                "|".join(filtered_df["id"].astype(str).tolist()).encode("utf-8")
            ).hexdigest()[:10]
            html_report = build_chamados_html_report(filtered_df, _load_comments_batch_for_export(filtered_df))
            st.download_button(
                label="📄 Exportar Filtrados (HTML)",
                data=html_report.encode("utf-8"),
                file_name=f"chamados_relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                mime="text/html",
                key=f"btn_export_chamados_html_{len(filtered_df)}_{_export_sig}",
                width="stretch",
            )
        with col_tbl_cap:
            st.components.v1.html("""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <div style="text-align: right; padding-top: 5px;">
                <button id="btn-cap-tbl" onclick="captureTable()" style="
                    background: linear-gradient(135deg, #10b981 0%, #059669 100%);
                    color: white;
                    border: none;
                    padding: 8px 14px;
                    font-size: 0.85rem;
                    font-weight: bold;
                    border-radius: 6px;
                    cursor: pointer;
                    box-shadow: 0 2px 6px rgba(0,0,0,0.3);
                    transition: all 0.2s ease;
                ">
                    📸 Baixar Tabela em Alta Resolução (PNG)
                </button>
            </div>
            <script>
            function captureTable() {
                const btn = document.getElementById("btn-cap-tbl");
                btn.innerText = "⏳ Gerando PNG em alta resolução...";
                btn.disabled = true;

                const tableEl = window.parent.document.querySelector('div[data-testid="stDataFrame"]') || 
                                window.parent.document.querySelector('.stDataFrame');

                if (!tableEl) {
                    alert("Não foi possível encontrar a tabela na tela.");
                    btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                    btn.disabled = false;
                    return;
                }

                // O st.dataframe usa a biblioteca glide-data-grid baseada em HTML5 Canvas.
                // A glide-data-grid renderiza no fundo uma árvore de nós DOM invisíveis para leitores de tela / acessibilidade.
                // O html2canvas interpreta esses nós e os desenha como texto sobreposto ("alt"), poluindo a imagem.
                // Usamos ignoreElements e onclone para ocultar/eliminar esses nós de acessibilidade.
                html2canvas(tableEl, {
                    scale: 2.5,
                    useCORS: true,
                    backgroundColor: "#0e1117",
                    logging: false,
                    ignoreElements: function(element) {
                        if (!element) return false;
                        // Não ignora os elementos principais nem canvas
                        const tag = (element.tagName || "").toLowerCase();
                        if (tag === "canvas" || tag === "div") {
                            // Se for container explícito de acessibilidade do glide-data-grid
                            const cls = (element.className || "").toString().toLowerCase();
                            const role = (element.getAttribute("role") || "").toLowerCase();
                            if (role === "presentation" || role === "row" || role === "gridcell" || role === "rowgroup") {
                                return true;
                            }
                            if (cls.includes("clip") || cls.includes("sr-only") || cls.includes("accessibility") || cls.includes("hidden")) {
                                return true;
                            }
                        }
                        // Ignora tabelas ou textos ocultos criados para acessibilidade dentro do glide-data-grid
                        if (tag === "table" || tag === "tr" || tag === "td" || tag === "th" || tag === "tbody" || tag === "thead") {
                            return true;
                        }
                        // Ignora elementos com aria-hidden falso/texto invisível
                        const style = window.getComputedStyle ? window.getComputedStyle(element) : null;
                        if (style && (style.opacity === "0" || style.visibility === "hidden")) {
                            return true;
                        }
                        return false;
                    },
                    onclone: function(clonedDoc) {
                        // Varre o clone e esconde qualquer texto solto ou elemento de acessibilidade que não seja o próprio canvas
                        const clonedTable = clonedDoc.querySelector('div[data-testid="stDataFrame"]') || clonedDoc.querySelector('.stDataFrame');
                        if (clonedTable) {
                            // Remove nós de texto que não estejam dentro de elementos visuais legítimos
                            const elements = clonedTable.querySelectorAll('*');
                            elements.forEach(el => {
                                if (el.tagName.toLowerCase() !== 'canvas' && !el.querySelector('canvas')) {
                                    el.style.color = 'transparent';
                                    el.style.textShadow = 'none';
                                }
                            });
                        }
                    }
                }).then(canvas => {
                    const link = document.createElement("a");
                    const d = new Date();
                    const dateStr = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0') + "-" + String(d.getDate()).padStart(2,'0');
                    link.download = `tabela_chamados_sti_${dateStr}.png`;
                    link.href = canvas.toDataURL("image/png");
                    link.click();

                    btn.innerText = "✅ Imagem Salva!";
                    setTimeout(() => {
                        btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                        btn.disabled = false;
                    }, 3000);
                }).catch(err => {
                    console.error("Erro no html2canvas:", err);
                    alert("Erro ao capturar tabela: " + err);
                    btn.innerText = "📸 Baixar Tabela em Alta Resolução (PNG)";
                    btn.disabled = false;
                });
            }
            </script>
            """, height=65)
        

        def style_dataframe(row):
            tag = str(row.get('tag', '')).upper().strip()
            bg_color = TAG_COLORS.get(tag, "")
            if bg_color:
                hex_color = bg_color.lstrip('#')
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                text_color = "#ffffff" if luminance < 0.6 else "#212529"
                style = f"background-color: {bg_color}; color: {text_color};"
            else:
                style = ""
                
            return [style] * len(row)

        if "chamados_reset_counter" not in st.session_state:
            st.session_state["chamados_reset_counter"] = 0

        df_page, current_page, total_pages, total_items = paginate_items(
            df_final_display,
            page_key="chamados",
            items_per_page=items_per_page
        )

        table_height = "content" if items_per_page >= 999999 else 600

        # Chave estável por página para preservar a ordenação de colunas (asc/desc) feita pelo usuário
        grid_key = f"tabela_chamados_p{current_page}"

        selection_event = st.dataframe(
            df_page.style.apply(style_dataframe, axis=1),
            column_order=cols_to_show,
            column_config={
                "id": st.column_config.LinkColumn("Chamado #", display_text=r".*#id:(.*)"),
                "status": st.column_config.TextColumn("Status"),
                "tag": st.column_config.TextColumn("TAG"),
                "andamento": st.column_config.TextColumn("Andamento"),
                "localidade_fisica": st.column_config.TextColumn("Localidade Física"),
                "cidade_predio": st.column_config.TextColumn("Cidade - Prédio"),
                "unidade": st.column_config.TextColumn("Unidade"),
                "usuario": st.column_config.TextColumn("Usuário"),
                "datetime_obj": st.column_config.DatetimeColumn("Data Criação", format="DD/MM/YYYY HH:mm:ss"),
                "ip_origem": st.column_config.TextColumn("IP de Origem"),
                "base": st.column_config.TextColumn("Base"),
            },
            hide_index=True,
            width="stretch",
            height=table_height,
            on_select="rerun",
            selection_mode="single-row",
            key=grid_key
        )

        selected_rows = selection_event.selection.rows if hasattr(selection_event, "selection") else []
        
        # Identifica se houve uma nova seleção feita pelo usuário
        if selected_rows:
            selected_idx = selected_rows[0]
            if selected_idx < len(df_page):
                sel_row_display = df_page.iloc[selected_idx]
                raw_id_str = str(sel_row_display['id'])
                # Extrai o ID limpo a partir de #id: ou valor direto
                if "#id:" in raw_id_str:
                    clean_ticket_id = raw_id_str.split("#id:")[-1].strip()
                else:
                    clean_ticket_id = raw_id_str.strip()

                # Só reabre o modal se for um chamado diferente do último dispensado/fechado
                if st.session_state.get("dismissed_ticket_id") != clean_ticket_id:
                    st.session_state["active_ticket_modal"] = clean_ticket_id
                    st.session_state["dismissed_ticket_id"] = None
        else:
            # Se o usuário desmarcou a linha, limpa o registro de dispensado
            st.session_state["dismissed_ticket_id"] = None

        if st.session_state.get("active_ticket_modal"):
            show_ticket_details(st.session_state["active_ticket_modal"])

        render_pagination_controls(
            page_key="chamados",
            current_page=current_page,
            total_pages=total_pages,
            total_items=total_items,
            items_per_page=items_per_page
        )


def _load_comments_batch_for_export(df: pd.DataFrame) -> dict:
    """Carrega comentários em lote para todos os chamados presentes no DataFrame filtrado."""
    result: dict = {}
    if df is None or df.empty or "id" not in df.columns:
        return result

    raw_ids = [
        str(cid).strip()[:-2] if str(cid).strip().endswith(".0") else str(cid).strip()
        for cid in df["id"].tolist()
        if cid is not None and str(cid).strip()
    ]
    unique_ids = list(dict.fromkeys(raw_ids))
    if not unique_ids:
        return result

    try:
        from src.database.connection import get_connection
        from src.config import clean_otrs_comments

        conn = get_connection()
        cursor = conn.cursor()
        try:
            for start in range(0, len(unique_ids), 250):
                chunk = unique_ids[start:start + 250]
                placeholders = ", ".join(["?"] * len(chunk))
                cursor.execute(
                    f"""
                    SELECT COALESCE(chamado_id, ''), COALESCE(data, ''), COALESCE(autor, ''), COALESCE(texto, '')
                    FROM comentarios
                    WHERE chamado_id IN ({placeholders})
                    ORDER BY id ASC
                    """,
                    tuple(chunk),
                )
                for row in cursor.fetchall():
                    cid = str(row[0] or "").strip()
                    result.setdefault(cid, []).append(
                        {"data": str(row[1] or ""), "autor": str(row[2] or ""), "texto": str(row[3] or "")}
                    )
        finally:
            cursor.close()
            conn.close()

        for cid, raw_list in list(result.items()):
            result[cid] = clean_otrs_comments(raw_list)
    except Exception:
        pass
    return result


def build_chamados_html_report(df: pd.DataFrame, comments_by_ticket: dict | None = None) -> str:
    """Gera relatório HTML interativo dos chamados filtrados (com descrição, andamento e comentários)."""
    import html as html_lib

    comments_map = comments_by_ticket or {}
    total = len(df) if df is not None else 0
    now_br = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    def _safe(val) -> str:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        s = str(val).strip()
        return "" if s.lower() in ("nan", "none", "null", "<na>") else s

    def _fmt_br_dt(val) -> str:
        return format_br_datetime(val, fallback="-")

    cards_html: list[str] = []
    loc_options: set[str] = set()
    user_options: set[str] = set()
    base_options: set[str] = set()
    tag_options: set[str] = set()

    if df is not None and not df.empty:
        for _, row in df.iterrows():
            cid = _safe(row.get("id", ""))
            if cid.endswith(".0"):
                cid = cid[:-2]
            base = _safe(row.get("base", "")) or "N/D"
            status = _safe(row.get("status", "")) or "Aberto"
            tag = _safe(row.get("tag", "")) or "SEM TAG"
            titulo = _safe(row.get("titulo", "")) or f"Chamado #{cid}"
            cidade_predio = _safe(row.get("cidade_predio", ""))
            unidade = _safe(row.get("unidade", ""))
            localidade = _safe(row.get("localidade_fisica", ""))
            loc_parts = list(dict.fromkeys([p for p in (cidade_predio, unidade, localidade) if p]))
            loc_display = " • ".join(loc_parts) if loc_parts else "Localidade não informada"

            usuario = _safe(row.get("usuario", "")) or _safe(row.get("id_cliente", "")) or "Não informado"
            loc_options.add(loc_display)
            user_options.add(usuario)
            base_options.add(base.upper())
            tag_options.add(tag.upper())
            criado = _fmt_br_dt(row.get("data_criacao", ""))
            atualizado = _fmt_br_dt(row.get("data_atualizacao", ""))
            hostname = _safe(row.get("hostname", "")) or "-"
            ip_origem = _safe(row.get("ip_origem", "")) or "-"
            link = _safe(row.get("link", ""))
            descricao = _safe(row.get("descricao", ""))
            andamento = _safe(row.get("andamento", ""))

            base_cls = "badge-otrs" if "otrs" in base.lower() else ("badge-cit" if "citsmart" in base.lower() else "")
            status_cls = "badge-open" if status.lower() == "aberto" else ""
            search_blob = html_lib.escape(
                f"{cid} {titulo} {loc_display} {usuario} {tag} {hostname}".lower(),
                quote=True,
            )

            link_html = ""
            if link.lower().startswith(("http://", "https://")):
                safe_link = html_lib.escape(link, quote=True)
                link_html = f"<div class='info-item no-print'><b>🔗 Link:</b> <a href='{safe_link}' target='_blank' rel='noopener'>Abrir no sistema</a></div>"

            desc_html = f"<div class='sec'><div class='sec-title'>Descrição</div>{html_lib.escape(descricao)}</div>" if descricao else ""
            and_html = f"<div class='sec'><div class='sec-title'>Andamento / Anotações</div>{html_lib.escape(andamento)}</div>" if andamento else ""

            ticket_comments = comments_map.get(cid, [])
            comments_html = ""
            if ticket_comments:
                c_blocks = []
                for c in ticket_comments:
                    c_data = html_lib.escape(_safe(c.get("data", "")))
                    c_autor = html_lib.escape(_safe(c.get("autor", "")) or "Sistema")
                    c_texto = html_lib.escape(_safe(c.get("texto", "")))
                    c_blocks.append(
                        f"<div class='comment'><div class='comment-meta'>{c_autor} • {c_data}</div>{c_texto}</div>"
                    )
                comments_html = f"<div class='sec'><div class='sec-title'>Comentários ({len(ticket_comments)})</div>{''.join(c_blocks)}</div>"

            cards_html.append(
                f"""<details class='card' open data-id='{html_lib.escape(cid, quote=True)}' data-search='{search_blob}' data-loc='{html_lib.escape(loc_display, quote=True)}' data-user='{html_lib.escape(usuario, quote=True)}' data-base='{html_lib.escape(base.upper(), quote=True)}' data-tag='{html_lib.escape(tag.upper(), quote=True)}'>
<summary class='card-head'>
<div class='head-top'>
<div class='badges'><span class='badge {base_cls}'>{html_lib.escape(base.upper())}</span><span class='badge {status_cls}'>{html_lib.escape(status)}</span><span class='badge'>{html_lib.escape(tag.upper())}</span><span class='id-tag'>#{html_lib.escape(cid)}</span></div>
<label class='done-wrap no-print' onclick='event.stopPropagation()'><input type='checkbox' onchange='toggleDone("{html_lib.escape(cid, quote=True)}", this)' /> Concluído</label>
</div>
<div class='title'>{html_lib.escape(titulo)}</div>
<div class='loc-banner'>📍 Localidade: {html_lib.escape(loc_display)}</div>
</summary>
<div class='card-body'>
<div class='info-grid'>
<div class='info-item'><b>👤 Solicitante:</b> {html_lib.escape(usuario)}</div>
<div class='info-item'><b>🕒 Criado:</b> {html_lib.escape(criado)}</div>
<div class='info-item'><b>🔄 Atualizado:</b> {html_lib.escape(atualizado)}</div>
<div class='info-item'><b>💻 Hostname:</b> {html_lib.escape(hostname)}</div>
<div class='info-item'><b>🌐 IP:</b> {html_lib.escape(ip_origem)}</div>
{link_html}
</div>
{desc_html}
{and_html}
{comments_html}
</div></details>"""
            )

    def _build_select_options(values: set[str]) -> str:
        return "".join(
            f"<option value='{html_lib.escape(v, quote=True)}'>{html_lib.escape(v)}</option>"
            for v in sorted(values, key=lambda x: x.lower())
        )

    loc_opts_html = _build_select_options(loc_options)
    user_opts_html = _build_select_options(user_options)
    base_opts_html = _build_select_options(base_options)
    tag_opts_html = _build_select_options(tag_options)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Relatório de Chamados - Bancada STI</title>
<style>
  :root {{ --bg: #f1f5f9; --card: #ffffff; --border: #cbd5e1; --text: #0f172a; --muted: #475569; --accent: #0284c7; --loc-bg: #fef3c7; --loc-border: #f59e0b; --loc-text: #92400e; }}
  * {{ box-sizing: border-box; }}
  body {{ font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 14px; line-height: 1.45; }}
  .container {{ max-width: 1020px; margin: 0 auto; }}
  .header {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 16px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
  h1 {{ font-size: 1.25rem; margin: 0 0 4px 0; }}
  .meta {{ color: var(--muted); font-size: 0.85rem; }}
  .toolbar {{ display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }}
  .toolbar input {{ flex: 1 1 220px; padding: 8px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.9rem; }}
  .filters-row {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 8px; margin-top: 10px; }}
  .filters-row select {{ width: 100%; padding: 8px 10px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.85rem; background: #fff; color: var(--text); }}
  .btn {{ border: 1px solid var(--border); background: #fff; color: var(--text); padding: 8px 12px; border-radius: 8px; font-size: 0.85rem; font-weight: 600; cursor: pointer; }}
  .btn-primary {{ background: var(--accent); color: #fff; border-color: var(--accent); }}
  details.card {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; margin-bottom: 12px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.04); transition: opacity 0.2s ease; }}
  details.card.done {{ opacity: 0.58; border-left: 5px solid #10b981; }}
  summary.card-head {{ list-style: none; padding: 14px 16px; cursor: pointer; display: flex; flex-direction: column; gap: 8px; user-select: none; }}
  summary.card-head::-webkit-details-marker {{ display: none; }}
  .head-top {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; }}
  .badges {{ display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }}
  .badge {{ font-size: 0.72rem; font-weight: 700; padding: 3px 8px; border-radius: 6px; background: #e2e8f0; color: #1e293b; text-transform: uppercase; }}
  .badge-otrs {{ background: #dbeafe; color: #1e40af; }}
  .badge-cit {{ background: #ccfbf1; color: #0f766e; }}
  .badge-open {{ background: #dcfce7; color: #166534; }}
  .id-tag {{ font-weight: 800; color: var(--accent); font-size: 0.95rem; }}
  .done-wrap {{ display: inline-flex; align-items: center; gap: 5px; font-size: 0.8rem; font-weight: 600; color: #047857; background: #ecfdf5; border: 1px solid #a7f3d0; padding: 4px 8px; border-radius: 8px; white-space: nowrap; }}
  .title {{ font-size: 1.05rem; font-weight: 700; color: var(--text); word-break: break-word; }}
  .loc-banner {{ background: var(--loc-bg); border: 1px solid var(--loc-border); color: var(--loc-text); padding: 7px 11px; border-radius: 8px; font-size: 0.88rem; font-weight: 700; display: flex; align-items: center; gap: 6px; word-break: break-word; }}
  .card-body {{ padding: 0 16px 16px 16px; border-top: 1px solid #e2e8f0; }}
  .info-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 8px 14px; margin: 12px 0; font-size: 0.84rem; color: #334155; }}
  .info-item b {{ color: #0f172a; }}
  .sec {{ background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; margin-top: 10px; font-size: 0.86rem; white-space: pre-wrap; word-break: break-word; }}
  .sec-title {{ font-size: 0.78rem; font-weight: 700; text-transform: uppercase; color: var(--accent); margin-bottom: 4px; letter-spacing: 0.03em; }}
  .comment {{ background: #ffffff; border: 1px solid #cbd5e1; border-left: 4px solid #10b981; border-radius: 8px; padding: 9px 12px; margin-top: 8px; font-size: 0.85rem; white-space: pre-wrap; word-break: break-word; }}
  .comment-meta {{ font-size: 0.76rem; font-weight: 700; color: #047857; margin-bottom: 4px; }}
  @media print {{
    @page {{ margin: 10mm; }}
    body {{ background: #fff; padding: 0; font-size: 11px; }}
    .container {{ max-width: 100%; }}
    .no-print {{ display: none !important; }}
    details.card {{ box-shadow: none; break-inside: avoid; page-break-inside: avoid; border: 1px solid #94a3b8; margin-bottom: 8px; }}
    details.card.done {{ display: none; }}
  }}
</style>
</head>
<body>
<div class="container">
<div class='header'>
<h1>Relatório de Chamados (<span id='visibleCount'>{total}</span> de {total})</h1>
<div class='meta'>Gerado em {now_br} • Toque em qualquer chamado para expandir/recolher</div>
  <div class="toolbar no-print">
    <input id="quickSearch" type="search" placeholder="Filtrar rápido por localidade, nome, ID ou título..." oninput="filterCards()" />
    <button class="btn" onclick="clearFilters()">Limpar Filtros</button>
    <button class="btn" onclick="toggleAll(true)">Expandir Todos</button>
    <button class="btn" onclick="toggleAll(false)">Recolher Todos</button>
    <button class="btn btn-primary" onclick="window.print()">🖨️ Imprimir / PDF</button>
  </div>
  <div class="filters-row no-print">
    <select id="filterLoc" onchange="filterCards()">
      <option value="">📍 Todas as Localidades</option>
      {loc_opts_html}
    </select>
    <select id="filterUser" onchange="filterCards()">
      <option value="">👤 Todos os Usuários</option>
      {user_opts_html}
    </select>
    <select id="filterBase" onchange="filterCards()">
      <option value="">🖥️ Todos os Sistemas</option>
      {base_opts_html}
    </select>
    <select id="filterTag" onchange="filterCards()">
      <option value="">🏷️ Todas as Tags</option>
      {tag_opts_html}
    </select>
  </div>
</div>
{''.join(cards_html)}
</div>
<script>
function filterCards() {{
  const q = (document.getElementById('quickSearch').value || '').toLowerCase().trim();
  const loc = document.getElementById('filterLoc').value || '';
  const user = document.getElementById('filterUser').value || '';
  const base = document.getElementById('filterBase').value || '';
  const tag = document.getElementById('filterTag').value || '';
  const cards = document.querySelectorAll('details.card');
  let vis = 0;
  cards.forEach(c => {{
    const hay = c.getAttribute('data-search') || '';
    const matchQ = !q || hay.includes(q);
    const matchLoc = !loc || (c.getAttribute('data-loc') || '') === loc;
    const matchUser = !user || (c.getAttribute('data-user') || '') === user;
    const matchBase = !base || (c.getAttribute('data-base') || '') === base;
    const matchTag = !tag || (c.getAttribute('data-tag') || '') === tag;
    const show = matchQ && matchLoc && matchUser && matchBase && matchTag;
    c.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  document.getElementById('visibleCount').innerText = vis;
}}
function clearFilters() {{
  document.getElementById('quickSearch').value = '';
  document.getElementById('filterLoc').value = '';
  document.getElementById('filterUser').value = '';
  document.getElementById('filterBase').value = '';
  document.getElementById('filterTag').value = '';
  filterCards();
}}
function toggleAll(open) {{
  document.querySelectorAll('details.card').forEach(c => c.open = open);
}}
function toggleDone(id, cb) {{
  const card = cb.closest('details.card');
  if (card) card.classList.toggle('done', cb.checked);
}}
</script>
</body>
</html>"""


