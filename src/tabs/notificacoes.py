import pandas as pd
import streamlit as st
from src.database import (
    get_notifications,
    mark_notification_as_read,
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

    c_title, c_act = st.columns([3, 1])
    with c_title:
        if unread_count > 0:
            st.warning(f"Você possui **{unread_count}** notificação(ões) pendente(s) não lida(s).")
        else:
            st.success("🎉 Todas as notificações estão em dia!")
    with c_act:
        if st.button("✅ Marcar Todas como Lidas", use_container_width=True, disabled=(unread_count == 0)):
            mark_all_notifications_as_read()
            st.toast("Todas as notificações foram marcadas como lidas!", icon="✅")
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # --- FILTROS SIDEBAR ---
    st.sidebar.markdown("## 🔍 Filtros de Notificação")
    filter_status = st.sidebar.radio(
        "📌 Status:",
        ["Não Lidas", "Todas", "Lidas"],
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
                    if st.button("🔗 Acessar Página", key=f"btn_nav_{notif_id}", use_container_width=True):
                        mark_notification_as_read(notif_id)
                        st.session_state["current_page"] = link_pagina
                        st.rerun()

                # Botão para marcar como lida
                if not is_read:
                    if st.button("✔ Marcar como Lida", key=f"btn_read_n_{notif_id}", use_container_width=True):
                        mark_notification_as_read(notif_id)
                        st.toast("Notificação marcada como lida!", icon="✅")
                        st.rerun()

    # Controles de Paginação no Rodapé
    render_pagination_controls(
        page_key="notificacoes",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )

