import pandas as pd
import streamlit as st
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.metric_cards import render_metric_cards
from src.components.status_banner import render_log_expander
from src.syncs.sync_donations import check_donations_sync_running, read_donations_last_log_lines

import os

@st.dialog("⚙️ Configurar / Enviar Planilha de Doações & Baixa", width="large")
def modal_config_donations():
    """Modal nativo (@st.dialog) para gerenciar link do SharePoint e envio direto da planilha de doações."""
    st.markdown("### 🖥️ Gestão da Planilha de Doação e Baixa de Máquinas")
    st.caption("Você pode consultar a URL da planilha oficial no SharePoint ou fazer o envio direto do arquivo Excel baixado do seu computador.")

    tab_online, tab_upload = st.tabs(["🌐 Link SharePoint / Atualizar Online", "📥 Envio Direto de Planilha"])

    with tab_online:
        from src.config import get_spreadsheet_setting, is_sharepoint_url
        excel_url_env = get_spreadsheet_setting("DONATIONS_EXCEL_RELATIVE_PATH", "").strip()
        st.write("Planilha oficial vinculada no SharePoint:")

        if is_sharepoint_url(excel_url_env):
            st.link_button(
                "🌐 Abrir Planilha no SharePoint (Excel Online) ↗",
                excel_url_env,
                type="secondary",
                width='stretch',
                help="Abre o arquivo original diretamente no SharePoint / Excel Online em uma nova aba."
            )
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        elif excel_url_env:
            st.info(f"📂 **Caminho Local Configurado:** `{excel_url_env}`")

        if st.button("🚀 Sincronizar pelo Link do SharePoint Agora", type="primary", width='stretch'):
            import sys, subprocess
            popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
            subprocess.Popen([sys.executable, "src/syncs/sync_donations.py"], **popen_kwargs)
            st.toast("🚀 Sincronização disparada com sucesso!", icon="🤖")
            st.rerun()

    with tab_upload:
        st.write("Faça o upload manual do arquivo Excel (.xlsx) da planilha de Doação e Baixa:")
        uploaded_donations_excel = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=["xlsx", "xls"], key="modal_up_donations_excel")

        if st.button("⚡ Processar e Gravar no Banco", type="primary", width='stretch'):
            if not uploaded_donations_excel:
                st.warning("Selecione um arquivo Excel primeiro.")
            else:
                from src.database.donations_db import sync_donations_from_excel
                import time
                with st.spinner("Lendo aba 'Equipamentos doados'..."):
                    res = sync_donations_from_excel(uploaded_donations_excel)
                    st.success("🎉 Planilha de doações importada com sucesso!")
                    time.sleep(1.2)
                    st.rerun()


def _pluralize_equip(name: str, count: int) -> str:
    n_lower = name.lower().strip()
    if count == 1:
        return f"1 {n_lower}"
    # Regras simples de plural em português
    if n_lower.endswith(('r', 's', 'z')):
        return f"{count} {n_lower}es"
    elif n_lower.endswith('m'):
        return f"{count} {n_lower[:-1]}ns"
    elif n_lower.endswith('l'):
        return f"{count} {n_lower[:-1]}is"
    elif n_lower.endswith(('a', 'e', 'i', 'o', 'u')):
        return f"{count} {n_lower}s"
    return f"{count} {n_lower}s"


def _format_equip_breakdown(grp_df: pd.DataFrame) -> str:
    raw_items = [str(x).strip() for x in grp_df['equipamento'].fillna('Equipamento').tolist() if str(x).strip()]
    counts = {}
    for item in raw_items:
        counts[item] = counts.get(item, 0) + 1
    sorted_items = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
    parts = [_pluralize_equip(eq_name, c) for eq_name, c in sorted_items if eq_name]
    if not parts:
        return ""
    if len(parts) == 1:
        return f"sendo {parts[0]}"
    return f"sendo {', '.join(parts[:-1])} e {parts[-1]}"


def _build_grouped_report_html(date_str: str, df_date: pd.DataFrame, mode: str = "preparo") -> tuple[str, str, bool]:
    """
    Constrói assunto, corpo HTML e indicador de máquinas para relatórios de movimentação:
      - mode="preparo": saída/preparo da Bancada para o DMP (subtab Visão Geral).
      - mode="entrega_dmp": recebimento na Bancada de equipamentos trazidos pelo DMP.
    """
    from datetime import datetime
    import html as html_lib

    try:
        dt_obj = datetime.strptime(date_str, "%Y-%m-%d")
        formatted_date = dt_obj.strftime("%d/%m/%Y")
    except Exception:
        formatted_date = date_str

    is_entrega = mode == "entrega_dmp"
    obs_col = "observacoes" if is_entrega and "observacoes" in df_date.columns else "motivo_baixa"

    raw_movs = [str(m).strip() for m in df_date.get('tipo_movimentacao', pd.Series(dtype=object)).tolist() if str(m).strip()]
    movs = sorted(list(set(raw_movs)))
    movs_clean = [m.capitalize() for m in movs]
    if len(movs_clean) == 1:
        movs_str = movs_clean[0]
    elif len(movs_clean) > 1:
        movs_str = ", ".join(movs_clean[:-1]) + " e " + movs_clean[-1]
    else:
        movs_str = "Movimentação"

    total_prep = len(df_date)
    equip_word = "equipamento" if total_prep == 1 else "equipamentos"
    if is_entrega:
        subject = f"[DMP -> BANCADA] - Recebimento de {movs_str} de equipamentos ({total_prep} {equip_word}) do dia {formatted_date}"
    else:
        subject = f"[DOAÇÃO] - Preparação de {movs_str} de equipamentos ({total_prep} {equip_word}) do dia {formatted_date}"

    global_breakdown = _format_equip_breakdown(df_date)
    intro_breakdown_str = f" ({global_breakdown})" if global_breakdown else ""

    html_parts = []
    html_parts.append("<div style='font-family: Arial, Helvetica, sans-serif; color: #000000; line-height: 1.5;'>")
    html_parts.append("<p>Prezados, boa tarde.</p>")
    if is_entrega:
        html_parts.append(
            f"<p>Registramos o recebimento na Bancada STI na data de <strong>{formatted_date}</strong> "
            f"de um total de <strong>{total_prep} {equip_word}</strong>{intro_breakdown_str} encaminhado(s) pelo DMP, sendo eles:</p>"
        )
    else:
        html_parts.append(
            f"<p>Na tarde de hoje (<strong>{formatted_date}</strong>), preparamos um total de "
            f"<strong>{total_prep} {equip_word}</strong>{intro_breakdown_str}, sendo eles:</p>"
        )

    has_machines_in_batch = False
    mov_types_with_machines = []
    machine_keywords = ("computador", "notebook", "workstation", "desktop", "microcomputador", "all in one", "pc", "servidor")
    non_machine_types = ("ssd", "hdd", "hd", "leitor biométrico", "leitor biometrico", "telefone", "monitor", "switch", "nobreak", "estabilizador")

    for mov_type, grp in df_date.groupby('tipo_movimentacao'):
        qtd_grp = len(grp)
        qtd_word = "equipamento" if qtd_grp == 1 else "equipamentos"
        grp_breakdown = _format_equip_breakdown(grp)
        grp_breakdown_str = f", {grp_breakdown}" if grp_breakdown else ""

        prefix_grp = "Equipamentos recebidos para" if is_entrega else "Equipamentos para"
        html_parts.append(f"<p style='margin-top: 20px; margin-bottom: 8px;'><strong>🔹 {prefix_grp} {mov_type.upper()} ({qtd_grp} {qtd_word}{grp_breakdown_str}):</strong></p>")

        # Na entrega do DMP, a conferência física de RAM e HD/SSD é feita apenas durante os trabalhos técnicos
        ram_vals = [] if is_entrega else [str(v).strip() for v in grp.get('memoria_ram', pd.Series(dtype=object)).tolist() if v is not None and str(v).strip() not in ("", "nan", "None")]
        ssd_vals = [] if is_entrega else [str(v).strip() for v in grp.get('ssd', pd.Series(dtype=object)).tolist() if v is not None and str(v).strip() not in ("", "nan", "None")]
        origem_vals = [str(v).strip() for v in grp.get('origem_unidade', pd.Series(dtype=object)).tolist() if v is not None and str(v).strip() not in ("", "nan", "None")] if is_entrega else []
        obs_vals = [str(v).strip() for v in grp.get(obs_col, pd.Series(dtype=object)).tolist() if v is not None and str(v).strip() not in ("", "nan", "None")]
        has_ram = len(ram_vals) > 0
        has_ssd = len(ssd_vals) > 0
        has_origem = len(origem_vals) > 0
        has_obs = len(obs_vals) > 0

        table_html = [
            "<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; width: 100%; border: 2px solid #cccccc; font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #000000;'>"
        ]
        th_style = "background-color: #2f5597; border: 2px solid #cccccc; padding: 6px 10px; text-align: left;"

        headers_html = [
            "<tr>",
            f"<th style='{th_style}'><span style=\"color:#ffffff\">#</span></th>",
            f"<th style='{th_style}'><span style=\"color:#ffffff\">Patrimônio</span></th>",
            f"<th style='{th_style}'><span style=\"color:#ffffff\">Modelo</span></th>",
            f"<th style='{th_style}'><span style=\"color:#ffffff\">Serial Number PC</span></th>",
            f"<th style='{th_style}'><span style=\"color:#ffffff\">Equipamento</span></th>"
        ]
        if has_ram:
            headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">Memória RAM</span></th>")
        if has_ssd:
            headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">SSD/HD</span></th>")
        if has_origem:
            headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">Origem / Setor</span></th>")
        if has_obs:
            headers_html.append(f"<th style='{th_style}'><span style=\"color:#ffffff\">Motivo/Obs</span></th>")
        headers_html.append("</tr>")

        table_html.append("".join(headers_html))

        def _clean_cell(val) -> str:
            s = str(val or "").strip()
            return "" if s.lower() in ("nan", "none", "null") else html_lib.escape(s)

        for idx, (_, row) in enumerate(grp.iterrows()):
            item_num = idx + 1
            pat = _clean_cell(row.get('patrimonio', ''))
            mod = _clean_cell(row.get('modelo', ''))
            ser = _clean_cell(row.get('serial_number', ''))
            eqp = _clean_cell(row.get('equipamento', ''))
            ram = _clean_cell(row.get('memoria_ram', ''))
            ssd = _clean_cell(row.get('ssd', ''))
            origem = _clean_cell(row.get('origem_unidade', '')) if is_entrega else ""
            obs = _clean_cell(row.get(obs_col, ''))

            bg_style = "background-color: #d9e1f2;" if idx % 2 == 1 else ""

            row_html = [
                "<tr>",
                f"<td style='border: 2px solid #cccccc; padding: 6px 10px; text-align: center; font-weight: bold; {bg_style}'>{item_num}</td>",
                f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{pat}</td>",
                f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{mod}</td>",
                f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{ser}</td>",
                f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{eqp}</td>"
            ]
            if has_ram:
                row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{ram}</td>")
            if has_ssd:
                row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{ssd}</td>")
            if has_origem:
                row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{origem}</td>")
            if has_obs:
                row_html.append(f"<td style='border: 2px solid #cccccc; padding: 6px 10px; {bg_style}'>{obs}</td>")
            row_html.append("</tr>")

            table_html.append("".join(row_html))

        colspan_extra = int(has_ram) + int(has_ssd) + int(has_origem) + int(has_obs)
        total_cols = 5 + colspan_extra
        footer_text = f"Total {mov_type}: {qtd_grp} {qtd_word}{grp_breakdown_str}"
        footer_html = f"<tr><td colspan='{total_cols}' style='border: 2px solid #cccccc; padding: 6px 10px; background-color: #e9ecef; font-weight: bold; text-align: right;'>{footer_text}</td></tr>"
        table_html.append(footer_html)

        table_html.append("</table>")
        html_parts.append("".join(table_html))

        eq_list_lower = [str(e).strip().lower() for e in grp['equipamento'].fillna('').tolist() if str(e).strip()]
        has_machine_in_grp = any(
            any(k in eq for k in machine_keywords) or (eq and eq not in non_machine_types)
            for eq in eq_list_lower
        ) or not eq_list_lower

        if has_machine_in_grp:
            has_machines_in_batch = True
            label_mov = str(mov_type).strip().capitalize()
            if label_mov not in mov_types_with_machines:
                mov_types_with_machines.append(label_mov)
            if is_entrega:
                html_parts.append(
                    f"<p style='margin-top: 6px; margin-bottom: 14px; font-size: 12.5px; color: #1f2937;'>"
                    f"ℹ️ <em><strong>Observação de Recebimento ({label_mov}):</strong> Os equipamentos relacionados acima foram recebidos do DMP "
                    f"para triagem e preparo; a conferência física de memória RAM e dispositivos de armazenamento (HD/SSD) será realizada durante a execução dos devidos trabalhos técnicos.</em></p>"
                )
            else:
                html_parts.append(
                    f"<p style='margin-top: 6px; margin-bottom: 14px; font-size: 12.5px; color: #1f2937;'>"
                    f"ℹ️ <em><strong>Observação ({label_mov}):</strong> As máquinas relacionadas acima foram conferidas e entregues completas, "
                    f"acompanhadas de seus respectivos dispositivos de armazenamento (HD/SSD) e pentes de memória RAM instalados.</em></p>"
                )

    if has_machines_in_batch:
        tipos_label = " e ".join(mov_types_with_machines)
        if is_entrega:
            html_parts.append(
                f"<div style='margin-top: 18px; padding: 10px 14px; background-color: #f8fafc; border-left: 4px solid #2f5597; border: 1px solid #cbd5e1; font-size: 13px; color: #000000;'>"
                f"<strong>📌 Nota de Recebimento ({tipos_label}):</strong> "
                f"Registramos que os equipamentos destinados a <strong>{tipos_label.lower()}</strong> listados acima foram entregues pelo DMP à Bancada STI "
                f"na data de <strong>{formatted_date}</strong>. A conferência física detalhada de <strong>memória RAM e dispositivos de armazenamento (HD/SSD)</strong> será realizada oportunamente durante a execução dos devidos trabalhos técnicos."
                f"</div>"
            )
        else:
            html_parts.append(
                f"<div style='margin-top: 18px; padding: 10px 14px; background-color: #f8fafc; border-left: 4px solid #2f5597; border: 1px solid #cbd5e1; font-size: 13px; color: #000000;'>"
                f"<strong>📌 Nota de Conferência e Entrega ({tipos_label}):</strong> "
                f"Registramos que todas as máquinas destinadas a <strong>{tipos_label.lower()}</strong> relacionadas neste chamado saíram da bancada técnica "
                f"e foram entregues <strong>completas</strong> ao setor responsável, devidamente montadas com seus respectivos <strong>discos de armazenamento (HD/SSD) e memórias RAM</strong>."
                f"</div>"
            )

    html_parts.append("</div>")
    full_html = "\n".join(html_parts)
    return subject, full_html, has_machines_in_batch


def build_preparo_report_html(date_str: str, df_date: pd.DataFrame) -> tuple[str, str, bool]:
    """
    Constrói o assunto, o corpo HTML do relatório de preparo de equipamentos
    e um booleano indicando se houve inclusão de observação de entrega completa.
    """
    return _build_grouped_report_html(date_str, df_date, mode="preparo")


def build_entrega_dmp_report_html(date_str: str, df_date: pd.DataFrame) -> tuple[str, str, bool]:
    """
    Constrói o assunto, o corpo HTML do relatório de recebimento na Bancada STI
    dos equipamentos entregues pelo DMP.
    """
    return _build_grouped_report_html(date_str, df_date, mode="entrega_dmp")


def render_donations_page():
    """Renderiza a página de Doação & Redistribuição de Máquinas com subtabs Visão Geral e Entrega do DMP."""
    st.title("🖥️ Sistema de Doação & Redistribuição de Máquinas")

    valid_subtabs = ("visao_geral", "entrega_do_dmp")
    qp_subtab = str(st.query_params.get("subtab", "")).strip().lower()
    if qp_subtab in valid_subtabs and st.session_state.get("redistribuicao_subtab") != qp_subtab and "redistribuicao_subtab" not in st.session_state:
        st.session_state["redistribuicao_subtab"] = qp_subtab

    subtab = st.radio(
        "Selecione a visão:",
        options=list(valid_subtabs),
        format_func=lambda x: (
            "📊 Visão Geral (Doação, Baixa & Redistribuição)"
            if x == "visao_geral"
            else "📥 Entrega do DMP (Recebimento na Bancada)"
        ),
        horizontal=True,
        key="redistribuicao_subtab",
        label_visibility="collapsed",
    )

    if st.query_params.get("subtab") != subtab:
        st.query_params["subtab"] = subtab

    if subtab == "entrega_do_dmp":
        _render_entrega_dmp_subtab()
    else:
        _render_visao_geral_subtab()


def _render_visao_geral_subtab():
    """Subtab principal (Visão Geral): menu lateral, preparo e histórico da planilha oficial."""
    from src.database import get_donations_data

    st.write("Acompanhe o inventário de equipamentos destinados a doação, redistribuição, garantia ou baixados.")
    
    donations_ativo = check_donations_sync_running()

    if "was_donations_syncing" not in st.session_state:
        st.session_state["was_donations_syncing"] = False

    if st.session_state["was_donations_syncing"] and not donations_ativo:
        st.session_state["was_donations_syncing"] = False
        st.toast("🎉 Sincronização de doações concluída com sucesso!", icon="✅")
        st.rerun()

    if donations_ativo:
        st.session_state["was_donations_syncing"] = True

    render_log_expander(
        "🤖 Sincronização de Doações em Segundo Plano",
        donations_ativo,
        read_donations_last_log_lines,
        check_donations_sync_running,
        "O robô está lendo a planilha do SharePoint. O painel permanece livre para uso!"
    )
    
    from src.config import resolve_spreadsheet_path, DONATIONS_FILE_PATH
    EXCEL_PATH = str(resolve_spreadsheet_path("DONATIONS_EXCEL_RELATIVE_PATH", "doacoes_redistribuicao.xlsx") or DONATIONS_FILE_PATH)
                    
    df = get_donations_data()
    
    from src.auth import is_admin
    if is_admin():
        st.sidebar.title("🖥️ Painel de Controle")

        if donations_ativo:
            st.sidebar.button("🤖 Atualizando...", type="primary", width='stretch', disabled=True)
        else:
            if st.sidebar.button("🔄 Sincronizar Planilha", type="primary", width='stretch', help="Busca atualizações na planilha do SharePoint em segundo plano."):
                import sys, subprocess
                popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
                subprocess.Popen([sys.executable, "src/syncs/sync_donations.py"], **popen_kwargs)
                st.toast("🚀 Sincronização de doações iniciada em segundo plano!", icon="🤖")
                st.rerun()

        if st.sidebar.button("⚙️ Configurar / Enviar Planilha", width='stretch', help="Gerenciar link do SharePoint ou enviar a planilha de doações manualmente."):
            modal_config_donations()
        st.sidebar.markdown("---")

    if df.empty:
        st.warning("⚠️ Nenhum dado encontrado no cache local. Por favor, clique em 'Sincronizar Planilha' ou use 'Configurar / Enviar Planilha' para carregar os registros.")
        return
        
    df['Ano'] = pd.to_datetime(df['data_movimentacao'], errors='coerce').dt.year
    df['Ano'] = df['Ano'].fillna("Sem Data").astype(str).str.replace(".0", "", regex=False)

    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros de Equipamentos")
    
    mov_options = ["Todos"] + sorted(list(df['tipo_movimentacao'].unique()))
    selected_mov = st.sidebar.selectbox("Tipo de Movimentação", mov_options, key="donations_mov")
    
    equip_options = ["Todos"] + sorted(list(df['equipamento'].unique()))
    selected_equip = st.sidebar.selectbox("Tipo de Equipamento", equip_options, key="donations_equip")
    
    model_options = ["Todos"] + sorted(list(df['modelo'].unique()))
    selected_model = st.sidebar.selectbox("Modelo", model_options, key="donations_model")
    
    year_options = ["Todos"] + sorted(list(df['Ano'].unique()), reverse=True)
    selected_year = st.sidebar.selectbox("Ano da Movimentação", year_options, key="donations_year")

    ssd_options = ["Todos"] + sorted(list(df['ssd'].unique()))
    selected_ssd = st.sidebar.selectbox("SSD / Armazenamento", ssd_options, key="donations_ssd")

    ram_unique = sorted([str(x) for x in df.get('memoria_ram', pd.Series(dtype=object)).fillna('').unique() if str(x).strip() and str(x).strip().lower() not in ('nan', 'none')])
    if ram_unique:
        ram_options = ["Todos"] + ram_unique
        selected_ram = st.sidebar.selectbox("Memória RAM", ram_options, key="donations_ram")
    else:
        selected_ram = "Todos"
    
    search_query = st.sidebar.text_input("🔎 Buscar (Patrimônio, Modelo, Chamado, RAM, SSD)", "", key="donations_search").strip()
    
    items_per_page = render_items_per_page_selector(
        key_prefix="redistribuicao",
        options=[10, 25, 50, 100, "Todos"],
        default_index=1,
        label="📄 Equipamentos por página:"
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("📋 Gerador de Texto (Preparo)")
    
    valid_dates = df[df['data_movimentacao'] != '']['data_movimentacao'].unique()
    valid_dates = sorted(list(valid_dates), reverse=True)
    
    def format_date_br(date_str):
        from datetime import datetime
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            return date_str
            
    selected_date_str = st.sidebar.selectbox("Selecione a Data de Preparo", valid_dates, format_func=format_date_br)
    generate_btn = st.sidebar.button("📝 Gerar Texto do Chamado", width='stretch')

    @st.dialog("📋 Texto de Preparo de Chamado", width="large")
    def show_preparo_text(date_str, df_all):
        df_date = df_all[df_all['data_movimentacao'] == date_str]
        
        if df_date.empty:
            st.warning("Nenhum equipamento encontrado nesta data.")
            return

        subject, full_html, has_doacao_or_baixa_machines = build_preparo_report_html(date_str, df_date)

        st.write("📋 **Assunto do Chamado:**")
        st.code(subject, language="text")
        if has_doacao_or_baixa_machines:
            st.info(
                "🛡️ **Observação de resguardo incluída automaticamente:** As máquinas relacionadas nesta data "
                "já contêm no relatório e no HTML abaixo a observação expressa de entrega **completa (com memórias RAM e dispositivos de armazenamento HD/SSD)**, "
                "atendendo às exigências da administração superior e do DMP."
            )
        st.markdown("---")
        
        st.write("💡 Selecione o texto abaixo com o mouse, copie (Ctrl+C) e cole diretamente no chamado do OTRS:")
        st.markdown(
            f'<div style="background-color: #ffffff; padding: 20px; border-radius: 6px; border: 1px solid #dddddd; max-height: 400px; overflow-y: auto;">{full_html}</div>', 
            unsafe_allow_html=True
        )
        
        st.markdown("---")
        st.write("💻 Ou copie o código-fonte HTML abaixo (clique no botão **'Código-Fonte'** no OTRS e cole):")
        st.code(full_html, language="html")

    if generate_btn:
        show_preparo_text(selected_date_str, df)

    df_filtered = df.copy()
    if selected_mov != "Todos":
        df_filtered = df_filtered[df_filtered['tipo_movimentacao'] == selected_mov]
    if selected_equip != "Todos":
        df_filtered = df_filtered[df_filtered['equipamento'] == selected_equip]
    if selected_model != "Todos":
        df_filtered = df_filtered[df_filtered['modelo'] == selected_model]
    if selected_year != "Todos":
        df_filtered = df_filtered[df_filtered['Ano'] == selected_year]
    if selected_ssd != "Todos":
        df_filtered = df_filtered[df_filtered['ssd'] == selected_ssd]
    if selected_ram != "Todos" and 'memoria_ram' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['memoria_ram'] == selected_ram]

    if search_query:
        query_lower = search_query.lower()
        ram_series = df_filtered['memoria_ram'].astype(str).str.lower() if 'memoria_ram' in df_filtered.columns else pd.Series('', index=df_filtered.index)
        ssd_series = df_filtered['ssd'].astype(str).str.lower() if 'ssd' in df_filtered.columns else pd.Series('', index=df_filtered.index)
        df_filtered = df_filtered[
            df_filtered['patrimonio'].str.lower().str.contains(query_lower) |
            df_filtered['modelo'].str.lower().str.contains(query_lower) |
            df_filtered['chamado'].str.lower().str.contains(query_lower) |
            ram_series.str.contains(query_lower) |
            ssd_series.str.contains(query_lower)
        ]

    st.markdown("---")
    total_equip = len(df_filtered)
    doados = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'doação'])
    redistribuicoes = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'redistribuição'])
    baixas = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'baixa'])
    garantias = len(df_filtered[df_filtered['tipo_movimentacao'].str.lower() == 'garantia'])

    render_metric_cards([
        {
            "title": "EQUIPAMENTOS",
            "value": total_equip,
            "border_color": "#3b82f6",
        },
        {
            "title": "DOAÇÕES",
            "value": doados,
            "border_color": "#10b981",
            "value_color": "#10b981",
        },
        {
            "title": "REDISTRIBUIÇÕES",
            "value": redistribuicoes,
            "border_color": "#8b5cf6",
            "value_color": "#8b5cf6",
        },
        {
            "title": "BAIXAS",
            "value": baixas,
            "border_color": "#ef4444",
            "value_color": "#ef4444",
        },
        {
            "title": "GARANTIAS",
            "value": garantias,
            "border_color": "#f59e0b",
            "value_color": "#f59e0b",
        },
    ])


    st.markdown("---")
    
    g_col1, g_col2 = st.columns(2)
    
    with g_col1:
        st.subheader("📊 Distribuição por Movimentação")
        if not df_filtered.empty:
            mov_counts = df_filtered['tipo_movimentacao'].value_counts().reset_index()
            mov_counts.columns = ['Movimentação', 'Quantidade']
            st.bar_chart(data=mov_counts, x='Movimentação', y='Quantidade', width='stretch')
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    with g_col2:
        st.subheader("📅 Histórico de Movimentações por Ano")
        if not df_filtered.empty:
            df_filtered['Ano'] = pd.to_datetime(df_filtered['data_movimentacao'], errors='coerce').dt.year
            df_filtered['Ano'] = df_filtered['Ano'].fillna("Sem Data").astype(str).str.replace(".0", "", regex=False)
            
            ano_counts = df_filtered.groupby(['Ano', 'tipo_movimentacao']).size().unstack(fill_value=0)
            st.bar_chart(ano_counts, width='stretch')
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    st.markdown("---")
    st.subheader("📋 Detalhamento dos Equipamentos")
    
    df_page, current_page, total_pages, total_items = paginate_items(
        df_filtered,
        page_key="redistribuicao",
        items_per_page=items_per_page
    )

    st.dataframe(
        df_page,
        column_config={
            "patrimonio": st.column_config.TextColumn("Patrimônio"),
            "modelo": st.column_config.TextColumn("Modelo"),
            "serial_number": st.column_config.TextColumn("Número de Série"),
            "equipamento": st.column_config.TextColumn("Equipamento"),
            "tipo_movimentacao": st.column_config.TextColumn("Movimentação"),
            "data_movimentacao": st.column_config.DateColumn("Data da Movimentação", format="DD/MM/YYYY"),
            "chamado": st.column_config.TextColumn("Chamado relacionado"),
            "memoria_ram": st.column_config.TextColumn("Memória RAM"),
            "ssd": st.column_config.TextColumn("SSD / HD"),
            "motivo_baixa": st.column_config.TextColumn("Motivo da Baixa"),
        },
        hide_index=True,
        width='stretch'
    )

    render_pagination_controls(
        page_key="redistribuicao",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )


def enrich_patrimonios_for_entrega_dmp(
    patrimonios: list[str],
    default_mov: str = "Redistribuição",
    default_equip: str = "Auto (AD/SCCM)",
    default_date: str | None = None,
    default_chamado: str = "",
    default_responsavel: str = "",
    default_obs: str = "",
) -> pd.DataFrame:
    """
    Cruza uma lista de patrimônios com AD, SCCM (incluindo RAM/disco) e histórico
    para pré-preencher o lote de conferência de entrega do DMP à Bancada STI.
    """
    from datetime import datetime
    from src.components.dmp_patrimonio_report import query_dmp_patrimonio_report
    from src.database.connection import get_connection

    data_str = default_date or datetime.now().strftime("%d/%m/%Y")
    if len(data_str) == 10 and data_str[4] == "-":
        try:
            data_str = datetime.strptime(data_str, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            pass
    if not patrimonios:
        return pd.DataFrame()

    df_dmp, _ = query_dmp_patrimonio_report(patrimonios)

    # Busca memória RAM e discos no cache SCCM para complementar os dados
    sccm_hw: dict[str, tuple[str, str]] = {}
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT name, memory_ram, disk_drives FROM sccm_cache_devices")
        for name, ram, disk in cur.fetchall():
            if name:
                sccm_hw[str(name).strip().lower()] = (
                    str(ram or "").strip(),
                    str(disk or "").strip(),
                )
        cur.close()
        conn.close()
    except Exception:
        pass

    rows = []
    for _, r in df_dmp.iterrows():
        pat = str(r.get("patrimonio", "")).strip()
        hostname = str(r.get("hostname", "")).strip()
        modelo = str(r.get("modelo", "")).strip()
        serial = str(r.get("serial_number", "")).strip()
        setor = str(r.get("setor", "")).strip()
        localidade = str(r.get("localidade", "")).strip()
        origem = " • ".join([p for p in (localidade, setor) if p and p != "-"])
        usuario = str(r.get("usuario_nome", "") or r.get("usuario_login", "")).strip()

        hw_ram, hw_disk = sccm_hw.get(hostname.lower(), ("", ""))
        if hw_ram.lower() in ("nan", "none", "null"):
            hw_ram = ""
        if hw_disk.lower() in ("nan", "none", "null"):
            hw_disk = ""

        if default_equip != "Auto (AD/SCCM)":
            equip_val = default_equip
        else:
            hint = f"{hostname} {modelo}".lower()
            if any(k in hint for k in ("note", "thinkpad", "latitude", "elitebook", "probook", "inspiron")):
                equip_val = "Notebook"
            else:
                equip_val = "Computador"

        rows.append(
            {
                "patrimonio": pat,
                "modelo": "" if modelo == "Não identificado" else modelo,
                "serial_number": "" if serial == "-" else serial,
                "equipamento": equip_val,
                "tipo_movimentacao": default_mov,
                "data_entrega": data_str,
                "chamado": default_chamado.strip(),
                "memoria_ram": hw_ram,
                "ssd": hw_disk,
                "origem_unidade": origem,
                "usuario_anterior": "" if usuario == "-" else usuario,
                "hostname": hostname,
                "responsavel_entrega": default_responsavel.strip(),
                "observacoes": default_obs.strip(),
            }
        )

    return pd.DataFrame(rows)


def _build_entrega_dmp_excel_bytes(df_export: pd.DataFrame) -> bytes:
    """Gera planilha Excel (.xlsx) formatada com aba 'Entrega do DMP' pronta para arquivamento."""
    import io

    rename_map = {
        "patrimonio": "Patrimônio",
        "modelo": "Modelo",
        "serial_number": "Serial Number PC",
        "equipamento": "Equipamento",
        "tipo_movimentacao": "Tipo de Movimentação",
        "data_entrega": "Data da Entrega",
        "chamado": "Chamado",
        "memoria_ram": "Memória RAM",
        "ssd": "SSD / HD",
        "origem_unidade": "Origem / Setor",
        "usuario_anterior": "Último Usuário",
        "hostname": "Hostname",
        "responsavel_entrega": "Entregue por (DMP)",
        "recebido_por": "Recebido por (Bancada)",
        "observacoes": "Observações",
    }
    cols = [c for c in rename_map if c in df_export.columns]
    df_out = df_export[cols].rename(columns=rename_map)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="Entrega do DMP", index=False)
    buf.seek(0)
    return buf.getvalue()


@st.dialog("📋 Relatório de Recebimento — Entrega do DMP", width="large")
def show_entrega_dmp_report_modal(date_str: str, df_all: pd.DataFrame):
    """Modal com assunto, pré-visualização HTML copiável, código-fonte e exportação Excel/CSV da entrega."""
    df_date = df_all[df_all["data_entrega"] == date_str]
    if df_date.empty:
        st.warning("Nenhum equipamento registrado nesta data de entrega.")
        return

    subject, full_html, has_machines = build_entrega_dmp_report_html(date_str, df_date)

    st.write("📋 **Assunto do Chamado / E-mail:**")
    st.code(subject, language="text")
    if has_machines:
        st.info(
            "🛡️ **Nota de recebimento incluída automaticamente:** O relatório registra que a conferência física de "
            "memória RAM e dispositivos de armazenamento (HD/SSD) será realizada durante a execução dos devidos trabalhos técnicos."
        )
    st.markdown("---")

    st.write("💡 Selecione o texto abaixo com o mouse, copie (Ctrl+C) e cole diretamente no chamado do OTRS ou e-mail:")
    st.markdown(
        f'<div style="background-color: #ffffff; padding: 20px; border-radius: 6px; border: 1px solid #dddddd; max-height: 400px; overflow-y: auto;">{full_html}</div>',
        unsafe_allow_html=True,
    )

    st.markdown("---")
    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        st.download_button(
            "📥 Baixar Aba Excel (.xlsx)",
            data=_build_entrega_dmp_excel_bytes(df_date),
            file_name=f"entrega_dmp_{date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            key=f"dl_entrega_xlsx_{date_str}",
        )
    with col_exp2:
        st.download_button(
            "📄 Baixar CSV (.csv)",
            data=df_date.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"entrega_dmp_{date_str}.csv",
            mime="text/csv",
            width="stretch",
            key=f"dl_entrega_csv_{date_str}",
        )

    with st.expander("💻 Ou copie o código-fonte HTML puro (botão 'Código-Fonte' no OTRS)", expanded=False):
        st.code(full_html, language="html")


def _on_entrega_dmp_barcode_scanned():
    """Callback para leitor de código de barras USB/Bluetooth (HID + Enter): adiciona patrimônio sem duplicar."""
    import re
    from src.components.dmp_patrimonio_report import parse_patrimonios_input

    scanned = st.session_state.get("entrega_dmp_barcode_input", "").strip()
    if not scanned:
        return

    clean_code = re.sub(r"[^a-zA-Z0-9\-_]", "", scanned)
    if clean_code:
        existing_text = st.session_state.get("entrega_dmp_patrimonios_text", "")
        existing_list = parse_patrimonios_input(existing_text)
        if clean_code not in existing_list:
            existing_list.append(clean_code)
            updated_text = "\n".join(existing_list)
            st.session_state["entrega_dmp_patrimonios_text"] = updated_text
            st.session_state["entrega_dmp_textarea"] = updated_text
            st.toast(f"🔫 Patrimônio {clean_code} adicionado à lista!", icon="✅")
        else:
            st.toast(f"ℹ️ Patrimônio {clean_code} já está na lista.", icon="ℹ️")

    st.session_state["entrega_dmp_barcode_input"] = ""


def _render_entrega_dmp_subtab():
    """Subtab Entrega do DMP: registro patrimônio por patrimônio, conferência, gravação no Postgres e relatório."""
    from datetime import date, datetime
    from src.auth import get_current_user, is_admin
    from src.components.dmp_patrimonio_report import parse_patrimonios_input
    from src.database import delete_entrega_dmp, get_entregas_dmp_data, save_entregas_dmp, update_entrega_dmp

    st.write(
        "Registre aqui os equipamentos trazidos pelo **DMP** para a sala da **Bancada STI**, "
        "patrimônio por patrimônio, com conferência automática no AD/SCCM, gravação no PostgreSQL e geração de relatório de recebimento."
    )

    df_entregas = get_entregas_dmp_data()

    # Sidebar dedicada à subtab Entrega do DMP
    st.sidebar.title("📥 Entrega do DMP")
    st.sidebar.subheader("📋 Relatório de Recebimento")
    valid_dates = sorted([d for d in df_entregas["data_entrega"].unique() if str(d).strip()], reverse=True) if not df_entregas.empty else []

    if valid_dates:
        selected_date_str = st.sidebar.selectbox(
            "Selecione a Data da Entrega",
            valid_dates,
            format_func=lambda d: datetime.strptime(d, "%Y-%m-%d").strftime("%d/%m/%Y") if len(str(d)) == 10 else str(d),
            key="entrega_dmp_report_date",
        )
        if st.sidebar.button("📝 Gerar Relatório de Recebimento", type="primary", width="stretch", key="btn_entrega_dmp_report"):
            show_entrega_dmp_report_modal(selected_date_str, df_entregas)
    else:
        st.sidebar.info("Registre ao menos uma entrega para gerar o relatório por data.")

    if not df_entregas.empty:
        st.sidebar.markdown("---")
        st.sidebar.subheader("📤 Exportar Histórico Completo")
        st.sidebar.download_button(
            "📥 Exportar Tudo (Excel .xlsx)",
            data=_build_entrega_dmp_excel_bytes(df_entregas),
            file_name=f"entregas_dmp_completo_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            key="btn_dl_all_entrega_xlsx",
        )
        st.sidebar.download_button(
            "📄 Exportar Tudo (CSV)",
            data=df_entregas.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"entregas_dmp_completo_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            width="stretch",
            key="btn_dl_all_entrega_csv",
        )

    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros do Histórico")
    mov_hist_opts = ["Todos"] + (sorted([m for m in df_entregas["tipo_movimentacao"].unique() if str(m).strip()]) if not df_entregas.empty else [])
    sel_mov_hist = st.sidebar.selectbox("Tipo de Movimentação", mov_hist_opts, key="entrega_dmp_filter_mov")
    search_hist = st.sidebar.text_input("🔎 Buscar Patrimônio / Modelo / Chamado / Serial", "", key="entrega_dmp_filter_search").strip()

    items_per_page = render_items_per_page_selector(
        key_prefix="entrega_dmp",
        options=[10, 25, 50, 100, "Todos"],
        default_index=1,
        label="📄 Registros por página:",
    )

    # Painel de Entrada e Conferência
    with st.expander("➕ Registrar Nova Entrega do DMP (Patrimônio por Patrimônio / Lote)", expanded=True):
        st.text_input(
            "🔫 Leitor de Código de Barras / Bipagem Rápida (bipe a etiqueta ou digite e pressione Enter)",
            placeholder="Aproxime o leitor do código de barras da etiqueta de patrimônio...",
            key="entrega_dmp_barcode_input",
            on_change=_on_entrega_dmp_barcode_scanned,
            help="Compatível com leitores USB/Bluetooth em modo teclado (HID): cada leitura adiciona automaticamente o patrimônio abaixo sem duplicar.",
        )

        tab_txt, tab_csv = st.tabs(["✍️ Digitar / Colar Patrimônios (Textarea)", "📁 Importar Arquivo (CSV / Excel)"])

        raw_pats_text = ""
        with tab_txt:
            raw_pats_text = st.text_area(
                "Patrimônios recebidos (um por linha, vírgula ou tabulação):",
                value=st.session_state.get("entrega_dmp_patrimonios_text", ""),
                placeholder="43188\n43299\n47675\n55477",
                height=140,
                key="entrega_dmp_textarea",
            )
            if raw_pats_text != st.session_state.get("entrega_dmp_patrimonios_text", ""):
                st.session_state["entrega_dmp_patrimonios_text"] = raw_pats_text

        with tab_csv:
            up_file = st.file_uploader("Selecione arquivo CSV ou Excel com os patrimônios:", type=["csv", "xlsx", "xls"], key="entrega_dmp_file_up")
            if up_file is not None:
                try:
                    df_up = pd.read_csv(up_file) if up_file.name.lower().endswith(".csv") else pd.read_excel(up_file)
                    pat_col = next((c for c in df_up.columns if "patrim" in str(c).lower()), df_up.columns[0] if len(df_up.columns) > 0 else None)
                    if pat_col is not None:
                        extracted = [
                            str(v).strip().replace(".0", "")
                            for v in df_up[pat_col].dropna().tolist()
                            if str(v).strip() and not str(v).strip().lower().startswith("patrim")
                        ]
                        if extracted:
                            raw_pats_text = "\n".join(dict.fromkeys(extracted))
                            st.session_state["entrega_dmp_patrimonios_text"] = raw_pats_text
                            st.session_state["entrega_dmp_textarea"] = raw_pats_text
                            st.success(f"✅ {len(extracted)} patrimônio(s) carregado(s) do arquivo!")
                except Exception as exc:
                    st.error(f"Erro ao ler arquivo: {exc}")

        c1, c2, c3 = st.columns(3)
        with c1:
            dt_entrega = st.date_input("Data da Entrega", value=date.today(), format="DD/MM/YYYY", key="entrega_dmp_date")
            tipo_mov_padrao = st.selectbox(
                "Tipo de Movimentação",
                ["Preparo / Misto (A Definir)", "Redistribuição", "Baixa", "Doação", "Garantia", "Devolução / Recebimento DMP"],
                key="entrega_dmp_mov_type",
                help="Use 'Preparo / Misto (A Definir)' quando o DMP entregar um lote variado e a destinação for definida depois na Bancada.",
            )
        with c2:
            equip_padrao = st.selectbox(
                "Tipo de Equipamento",
                ["Auto (AD/SCCM)", "Computador", "Notebook", "Monitor", "Nobreak", "Impressora", "Outro"],
                key="entrega_dmp_equip_type",
                help="Em 'Auto (AD/SCCM)', o sistema identifica automaticamente Computador ou Notebook pelo AD/SCCM. Só altere se o lote inteiro for de um único tipo (ex: só Monitores).",
            )
            chamado_padrao = st.text_input("Chamado OTRS/CitSmart (opcional)", "", key="entrega_dmp_chamado")
        with c3:
            resp_entrega = st.text_input("Entregue por (servidor do DMP)", "", key="entrega_dmp_resp")
            obs_padrao = st.text_input("Observação padrão (opcional)", "", placeholder="Ex: Recebido completo (HD/SSD + RAM)", key="entrega_dmp_obs")

        b1, b2 = st.columns([2, 1])
        with b1:
            btn_conferir = st.button("🔍 Conferir Patrimônios no AD / SCCM", type="primary", width="stretch", key="btn_conferir_entrega_dmp")
        with b2:
            if st.button("🗑️ Limpar Lista", width="stretch", key="btn_limpar_entrega_dmp"):
                st.session_state["entrega_dmp_patrimonios_text"] = ""
                st.session_state["entrega_dmp_textarea"] = ""
                st.session_state.pop("entrega_dmp_preview_df", None)
                st.rerun()

        if btn_conferir:
            pats = parse_patrimonios_input(st.session_state.get("entrega_dmp_patrimonios_text", ""))
            if not pats:
                st.warning("⚠️ Informe ao menos um patrimônio para conferir.")
            else:
                with st.spinner(f"Consultando {len(pats)} patrimônio(s) no AD, SCCM e histórico..."):
                    st.session_state["entrega_dmp_preview_df"] = enrich_patrimonios_for_entrega_dmp(
                        pats,
                        default_mov=tipo_mov_padrao,
                        default_equip=equip_padrao,
                        default_date=dt_entrega.strftime("%Y-%m-%d"),
                        default_chamado=chamado_padrao,
                        default_responsavel=resp_entrega,
                        default_obs=obs_padrao,
                    )

        preview_df = st.session_state.get("entrega_dmp_preview_df")
        if isinstance(preview_df, pd.DataFrame) and not preview_df.empty:
            st.markdown("#### ✏️ Conferência Item a Item (edite caso falte RAM, SSD/HD, Serial ou Observação)")
            edited_df = st.data_editor(
                preview_df,
                num_rows="dynamic",
                width="stretch",
                hide_index=True,
                column_config={
                    "patrimonio": st.column_config.TextColumn("Patrimônio", required=True),
                    "modelo": st.column_config.TextColumn("Modelo"),
                    "serial_number": st.column_config.TextColumn("Serial Number PC"),
                    "equipamento": st.column_config.TextColumn("Equipamento"),
                    "tipo_movimentacao": st.column_config.SelectboxColumn(
                        "Movimentação",
                        options=["Preparo / Misto (A Definir)", "Redistribuição", "Baixa", "Doação", "Garantia", "Devolução / Recebimento DMP"],
                    ),
                    "data_entrega": st.column_config.TextColumn("Data (DD/MM/YYYY)"),
                    "chamado": st.column_config.TextColumn("Chamado"),
                    "memoria_ram": st.column_config.TextColumn("Memória RAM"),
                    "ssd": st.column_config.TextColumn("SSD / HD"),
                    "origem_unidade": st.column_config.TextColumn("Origem / Setor"),
                    "usuario_anterior": st.column_config.TextColumn("Último Usuário"),
                    "hostname": st.column_config.TextColumn("Hostname"),
                    "responsavel_entrega": st.column_config.TextColumn("Entregue por (DMP)"),
                    "observacoes": st.column_config.TextColumn("Observações"),
                },
                key="entrega_dmp_editor",
            )

            col_save1, col_save2 = st.columns([2, 1])
            with col_save1:
                if st.button("💾 Gravar Entrega no PostgreSQL", type="primary", width="stretch", key="btn_save_entrega_dmp"):
                    curr_user = get_current_user() or {}
                    recebedor = curr_user.get("display_name") or curr_user.get("username") or "Bancada STI"
                    recs = edited_df.to_dict(orient="records")
                    for rec in recs:
                        rec["recebido_por"] = recebedor
                    saved_cnt = save_entregas_dmp(recs)
                    st.session_state["entrega_dmp_patrimonios_text"] = ""
                    st.session_state.pop("entrega_dmp_preview_df", None)
                    st.success(f"🎉 {saved_cnt} equipamento(s) gravado(s) no histórico de entregas do DMP!")
                    st.rerun()
            with col_save2:
                st.download_button(
                    "📄 Baixar Lote em CSV",
                    data=edited_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"lote_entrega_dmp_{dt_entrega.strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    width="stretch",
                    key="btn_dl_preview_csv",
                )

    # Histórico Gravado no PostgreSQL
    st.markdown("---")
    df_hist = df_entregas.copy()
    if sel_mov_hist != "Todos" and not df_hist.empty:
        df_hist = df_hist[df_hist["tipo_movimentacao"] == sel_mov_hist]
    if search_hist and not df_hist.empty:
        ql = search_hist.lower()
        mask = (
            df_hist["patrimonio"].str.lower().str.contains(ql, na=False)
            | df_hist["modelo"].str.lower().str.contains(ql, na=False)
            | df_hist["serial_number"].str.lower().str.contains(ql, na=False)
            | df_hist["chamado"].str.lower().str.contains(ql, na=False)
            | df_hist["origem_unidade"].str.lower().str.contains(ql, na=False)
            | df_hist["observacoes"].str.lower().str.contains(ql, na=False)
        )
        df_hist = df_hist[mask]

    total_h = len(df_hist)
    render_metric_cards(
        [
            {"title": "ENTREGAS DMP", "value": total_h, "border_color": "#3b82f6"},
            {
                "title": "REDISTRIBUIÇÕES",
                "value": len(df_hist[df_hist["tipo_movimentacao"].str.lower() == "redistribuição"]) if total_h else 0,
                "border_color": "#8b5cf6",
                "value_color": "#8b5cf6",
            },
            {
                "title": "BAIXAS",
                "value": len(df_hist[df_hist["tipo_movimentacao"].str.lower() == "baixa"]) if total_h else 0,
                "border_color": "#ef4444",
                "value_color": "#ef4444",
            },
            {
                "title": "DOAÇÕES",
                "value": len(df_hist[df_hist["tipo_movimentacao"].str.lower() == "doação"]) if total_h else 0,
                "border_color": "#10b981",
                "value_color": "#10b981",
            },
            {
                "title": "GARANTIAS",
                "value": len(df_hist[df_hist["tipo_movimentacao"].str.lower() == "garantia"]) if total_h else 0,
                "border_color": "#f59e0b",
                "value_color": "#f59e0b",
            },
        ]
    )

    st.markdown("---")
    st.subheader(f"📦 Histórico de Equipamentos Recebidos do DMP ({total_h} registros)")
    if df_hist.empty:
        st.info("Nenhum registro de entrega do DMP encontrado com os filtros atuais.")
        return

    sort_col1, sort_col2 = st.columns([2, 1])
    sort_options = {
        "data_entrega": "Data da Entrega",
        "id": "ID",
        "patrimonio": "Patrimônio",
        "equipamento": "Equipamento",
        "modelo": "Modelo",
        "serial_number": "Serial Number",
        "tipo_movimentacao": "Movimentação",
        "origem_unidade": "Origem / Setor",
        "chamado": "Chamado",
    }
    with sort_col1:
        sel_sort_col = st.selectbox(
            "Ordenar por:",
            options=[c for c in sort_options if c in df_hist.columns],
            format_func=lambda c: sort_options.get(c, c),
            key="entrega_dmp_sort_col",
        )
    with sort_col2:
        sel_sort_dir = st.radio(
            "Ordem:",
            options=["DESC", "ASC"],
            format_func=lambda d: "⬇️ Decrescente" if d == "DESC" else "⬆️ Crescente",
            horizontal=True,
            key="entrega_dmp_sort_dir",
        )

    if sel_sort_col in df_hist.columns:
        asc = sel_sort_dir == "ASC"
        if sel_sort_col == "id":
            df_hist = df_hist.assign(_sort_num=pd.to_numeric(df_hist["id"], errors="coerce").fillna(0)).sort_values(
                by="_sort_num", ascending=asc, kind="mergesort"
            ).drop(columns=["_sort_num"])
        else:
            df_hist = df_hist.assign(_sort_key=df_hist[sel_sort_col].astype(str).str.lower()).sort_values(
                by="_sort_key", ascending=asc, kind="mergesort"
            ).drop(columns=["_sort_key"])

    df_page, current_page, total_pages, total_items = paginate_items(
        df_hist,
        page_key="entrega_dmp",
        items_per_page=items_per_page,
    )

    cols_show = [
        "id", "data_entrega", "patrimonio", "equipamento", "modelo", "serial_number",
        "tipo_movimentacao", "origem_unidade", "chamado",
        # "memoria_ram", "ssd",  # Ocultado temporariamente: conferência feita durante os trabalhos técnicos
        "responsavel_entrega", "recebido_por", "observacoes",
    ]
    cols_show = [c for c in cols_show if c in df_page.columns]
    hist_col_cfg = {
        "id": st.column_config.NumberColumn("ID", width="small", disabled=True),
        "data_entrega": st.column_config.TextColumn("Data Entrega (DD/MM/YYYY)"),
        "patrimonio": st.column_config.TextColumn("Patrimônio"),
        "equipamento": st.column_config.TextColumn("Equipamento"),
        "modelo": st.column_config.TextColumn("Modelo"),
        "serial_number": st.column_config.TextColumn("Serial Number"),
        "tipo_movimentacao": st.column_config.SelectboxColumn(
            "Movimentação",
            options=["Preparo / Misto (A Definir)", "Redistribuição", "Baixa", "Doação", "Garantia", "Devolução / Recebimento DMP"],
        ),
        "memoria_ram": st.column_config.TextColumn("Memória RAM"),
        "ssd": st.column_config.TextColumn("SSD / HD"),
        "origem_unidade": st.column_config.TextColumn("Origem / Setor"),
        "chamado": st.column_config.TextColumn("Chamado"),
        "responsavel_entrega": st.column_config.TextColumn("Entregue por (DMP)"),
        "recebido_por": st.column_config.TextColumn("Recebido por (Bancada)", disabled=True),
        "observacoes": st.column_config.TextColumn("Observações"),
    }

    if is_admin():
        df_hist_view = df_page[cols_show].copy()
        if "data_entrega" in df_hist_view.columns:
            df_hist_view["data_entrega"] = df_hist_view["data_entrega"].apply(
                lambda d: datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
                if len(str(d)) >= 10 and str(d)[4] == "-"
                else str(d)
            )

        edited_hist = st.data_editor(
            df_hist_view,
            num_rows="dynamic",
            column_config=hist_col_cfg,
            hide_index=True,
            width="stretch",
            key=f"entrega_dmp_hist_editor_{current_page}",
        )

        if st.button("💾 Salvar Edições / Exclusões da Tabela", type="primary", key="btn_save_hist_edits"):
            orig_ids = {int(x) for x in df_hist_view["id"].dropna().tolist()}
            curr_ids = {int(x) for x in edited_hist["id"].dropna().tolist()}
            removed_ids = orig_ids - curr_ids
            del_cnt = sum(1 for rid in removed_ids if delete_entrega_dmp(rid))

            upd_cnt = 0
            orig_by_id = df_hist_view.set_index("id").to_dict(orient="index")
            for _, row in edited_hist.iterrows():
                rid_val = row.get("id")
                if pd.isna(rid_val):
                    continue
                rid = int(rid_val)
                orig_row = orig_by_id.get(rid)
                if orig_row is None:
                    continue
                row_dict = row.to_dict()
                if any(str(row_dict.get(k, "")) != str(orig_row.get(k, "")) for k in row_dict if k != "id"):
                    if update_entrega_dmp(rid, row_dict):
                        upd_cnt += 1

            st.toast(f"✅ {upd_cnt} atualizado(s), {del_cnt} removido(s).", icon="💾")
            st.rerun()
    else:
        st.dataframe(
            df_page[cols_show],
            column_config=hist_col_cfg,
            hide_index=True,
            width="stretch",
        )

    render_pagination_controls(
        page_key="entrega_dmp",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page,
    )

