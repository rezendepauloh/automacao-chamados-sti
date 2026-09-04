import pandas as pd
import streamlit as st
from src.database import (
    get_notifications,
    mark_notification_as_read,
    mark_notification_as_unread,
    mark_all_notifications_as_read,
    get_unread_notifications_count
)
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)



def render_notificacoes_page():
    """Renderiza a Central de Notificações da Bancada."""
    st.title("🔔 Central de Notificações da Bancada")
    st.write("Acompanhe alertas automáticos sobre novas portarias publicadas e escalas de plantão (matutino e semanal).")
    st.markdown("---")

    # Contagem não lidas
    unread_count = get_unread_notifications_count()

    c_title, c_act1, c_act2 = st.columns([2, 1, 1])
    with c_title:
        if unread_count > 0:
            st.warning(f"Você possui **{unread_count}** notificação(ões) pendente(s) não lida(s).")
        else:
            st.success("🎉 Todas as notificações estão em dia!")
    with c_act1:
        if st.button("🔄 Verificar Alertas Agora", width='stretch', help="Varre portarias e escalas de plantão em busca de novos alertas para a bancada."):
            with st.spinner("Verificando portarias e escalas de plantão..."):
                try:
                    from src.syncs.sync_plantoes_alerts import check_and_generate_plantao_alerts
                    from src.syncs.sync_portarias import sync_portarias_and_generate_alerts
                    sync_portarias_and_generate_alerts()
                    check_and_generate_plantao_alerts()
                    st.toast("Alertas de portarias e plantões verificados com sucesso!", icon="🔔")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao verificar alertas: {e}")
    with c_act2:
        if st.button("✅ Marcar Todas como Lidas", width='stretch', disabled=(unread_count == 0)):
            mark_all_notifications_as_read()
            st.toast("Todas as notificações foram marcadas como lidas!", icon="✅")
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # --- FILTROS SIDEBAR ---
    st.sidebar.markdown("## 🔍 Filtros de Notificação")
    default_status_idx = 0 if unread_count > 0 else 1
    filter_status = st.sidebar.radio(
        "📌 Status:",
        ["Não Lidas", "Todas", "Lidas"],
        index=default_status_idx,
        key="filter_notif_status"
    )
    
    filter_tipo = st.sidebar.selectbox(
        "🏷️ Tipo de Alerta:",
        ["Todos", "Portaria", "Plantão Matutino", "Plantão Semanal"],
        key="filter_notif_tipo"
    )

    sort_notif = st.sidebar.selectbox(
        "⬆️⬇️ Ordenar por Data:",
        ["Mais recentes primeiro (DESC)", "Mais antigas primeiro (ASC)"],
        key="filter_notif_sort"
    )

    items_per_page = render_items_per_page_selector(
        key_prefix="notificacoes",
        options=[5, 10, 20, 50, 100, "Todos"],
        default_index=1,
        label="📄 Notificações por página:"
    )


    # Carrega notificações do banco
    df_notif = get_notifications(only_unread=False, limit=500)

    if df_notif.empty:
        st.info("Nenhuma notificação registrada no momento.")
        return

    # Aplicação de filtros
    if filter_status == "Não Lidas":
        df_notif = df_notif[df_notif['lida'] == 0]
    elif filter_status == "Lidas":
        df_notif = df_notif[df_notif['lida'] == 1]

    if filter_tipo != "Todos":
        df_notif = df_notif[df_notif['tipo'] == filter_tipo]

    # Ordenação por data/id
    if sort_notif == "Mais antigas primeiro (ASC)":
        df_notif = df_notif.sort_values(by="id", ascending=True)
    else:
        df_notif = df_notif.sort_values(by="id", ascending=False)

    if df_notif.empty:
        st.info("Nenhuma notificação encontrada para os filtros selecionados.")
        return

    # Aplicação da Paginação
    df_page, current_page, total_pages, total_items = paginate_items(
        df_notif,
        page_key="notificacoes",
        items_per_page=items_per_page
    )

    st.markdown(f"**Exibindo {len(df_page)} de {total_items} notificação(ões) filtrada(s)**")
    st.markdown("<br>", unsafe_allow_html=True)

    # Lista de Notificações
    for _, row in df_page.iterrows():
        notif_id = int(row['id'])
        tipo = str(row['tipo'])
        titulo = str(row['titulo'])
        mensagem = str(row['mensagem'])
        try:
            data_criacao = pd.to_datetime(row['data_criacao']).strftime('%d/%m/%Y %H:%M:%S')
        except Exception:
            data_criacao = str(row['data_criacao'])
        link_pagina = str(row.get('link_pagina', ''))
        is_read = int(row['lida']) == 1

        # Definição de ícones e cores por tipo
        if "Portaria" in tipo:
            badge_icon = "📜"
            badge_color = "#3498db"
        elif "Matutino" in tipo:
            badge_icon = "☀️"
            badge_color = "#e67e22"
        else:
            badge_icon = "🌙"
            badge_color = "#8e44ad"

        status_badge = "🟢 Lida" if is_read else "🔴 Nova / Não Lida"

        with st.container(border=True):
            col_b, col_txt, col_btns = st.columns([1, 4, 2])
            
            with col_b:
                st.markdown(f"### {badge_icon}")
                st.caption(f"<span style='color: {badge_color}; font-weight: bold;'>{tipo}</span>", unsafe_allow_html=True)
                st.caption(status_badge)

            with col_txt:
                st.markdown(f"#### {titulo}")
                st.write(mensagem)
                st.caption(f"🕒 Registrado em: {data_criacao}")

            with col_btns:
                st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
                
                # Botão para redirecionar para a página da notificação
                if link_pagina:
                    if st.button("🔗 Acessar Página", key=f"btn_nav_{notif_id}", width='stretch'):
                        from src.components.header import PAGE_TO_SLUG
                        mark_notification_as_read(notif_id)
                        st.session_state["current_page"] = link_pagina
                        if link_pagina in PAGE_TO_SLUG:
                            st.query_params["tab"] = PAGE_TO_SLUG[link_pagina]
                        st.rerun()

                # Botão para marcar como lida ou não lida
                if not is_read:
                    if st.button("✔ Marcar como Lida", key=f"btn_read_n_{notif_id}", width='stretch'):
                        mark_notification_as_read(notif_id)
                        st.toast("Notificação marcada como lida!", icon="✅")
                        st.rerun()
                else:
                    if st.button("↩ Marcar como Não Lida", key=f"btn_unread_n_{notif_id}", width='stretch'):
                        mark_notification_as_unread(notif_id)
                        st.toast("Notificação reativada como não lida!", icon="🔔")
                        st.rerun()

                # Botão de envio manual pelo WhatsApp
                if st.button("📲 Enviar no WhatsApp", key=f"btn_wpp_send_{notif_id}", width='stretch', help="Envia esta notificação via WhatsApp para os integrantes ativos da bancada"):
                    from src.services.evolution_client import send_whatsapp_text, get_connection_status
                    from src.database import get_whatsapp_destinatarios, log_whatsapp_dispatch

                    st_conn = get_connection_status()
                    if not st_conn.get("online") or st_conn.get("state") != "open":
                        st.error("WhatsApp desconectado! Conecte a sessão na aba Configurações > WhatsApp.")
                    else:
                        dests = get_whatsapp_destinatarios(only_active=True)
                        if dests.empty:
                            st.warning("Nenhum destinatário ativo configurado.")
                        else:
                            with st.spinner("Enviando via WhatsApp..."):
                                enviou_algum = False
                                for _, drow in dests.iterrows():
                                    dtel = drow["telefone"]
                                    dnome = drow["nome"]
                                    p_nome = dnome.split()[0]
                                    msg_text = (
                                        f"🔔 *Notificação Bancada STI*\n\n"
                                        f"Olá, *{p_nome}*!\n\n"
                                        f"*{titulo}*\n"
                                        f"{mensagem}\n\n"
                                        f"_Enviado manualmente pela Central de Notificações._"
                                    )
                                    s_res = send_whatsapp_text(dtel, msg_text)
                                    if s_res.get("success"):
                                        log_whatsapp_dispatch(tipo, f"manual_{notif_id}", str(row.get("data_evento", "")), dtel, msg_text, "enviado", str(s_res))
                                        enviou_algum = True
                                if enviou_algum:
                                    st.toast("Notificação enviada com sucesso no WhatsApp!", icon="📲")
                                else:
                                    st.error("Falha ao entregar notificação no WhatsApp.")

    # Controles de Paginação no Rodapé
    render_pagination_controls(
        page_key="notificacoes",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )

