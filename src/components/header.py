import streamlit as st
from src.database import get_unread_notifications_count, get_notifications

PAGE_TO_SLUG = {
    "📋 Painel de Chamados": "chamados",
    "🏢 Catálogo de Unidades": "unidades",
    "📞 Central Telefônica (OXE)": "central-telefonica",
    "📅 Plantões da Bancada": "plantoes",
    "📅 Calendário Geral": "calendario-geral",
    "📜 Portarias da Bancada": "portarias",
    "📍 Mapa & Localização": "mapa",
    "🖥️ Doação & Redistribuição": "redistribuicao",
    "📜 Fiscalização de Contratos": "fiscalizacao",
    "🛡️ Controle de Garantia": "garantia",
    "🖨️ Impressoras (PaperCut)": "impressoras",
    "⚡ Scripts de Automação": "scripts-automacao",
    "📚 FAQ & Tutoriais": "faq",
    "🔔 Central de Notificações": "notificacoes",
}



SLUG_TO_PAGE = {v: k for k, v in PAGE_TO_SLUG.items()}


def render_header_navigation() -> str:
    """
    Renderiza o menu hambúrguer popover fixado no canto superior direito do header nativo.
    Sincroniza o estado da página ativa com os Query Parameters da URL (?tab=slug).
    Exibe notificações em toast para novos alertas e retorna a página atualmente selecionada.
    """
    # 1. Sincroniza estado inicial a partir do GET parameter na URL (?tab=slug)
    url_tab = st.query_params.get("tab")
    if url_tab and url_tab in SLUG_TO_PAGE:
        st.session_state["current_page"] = SLUG_TO_PAGE[url_tab]
    elif "current_page" not in st.session_state:
        st.session_state["current_page"] = "📋 Painel de Chamados"

    # Garante que a URL reflita o slug da página atual
    current_slug = PAGE_TO_SLUG.get(st.session_state["current_page"], "chamados")
    if st.query_params.get("tab") != current_slug:
        st.query_params["tab"] = current_slug

    def set_page(page_name: str):
        st.session_state["current_page"] = page_name
        st.query_params["tab"] = PAGE_TO_SLUG.get(page_name, "chamados")
        st.rerun()

    # Notificações em Toast ao abrir o app/recarregar
    if "toasted_notif_ids" not in st.session_state:
        st.session_state["toasted_notif_ids"] = set()

    try:
        df_unread = get_notifications(only_unread=True, limit=5)
        if not df_unread.empty:
            for _, row in df_unread.iterrows():
                n_id = int(row['id'])
                if n_id not in st.session_state["toasted_notif_ids"]:
                    st.session_state["toasted_notif_ids"].add(n_id)
                    st.toast(f"🔔 **{row['titulo']}**: {row['mensagem'][:90]}...", icon="📢")
    except Exception:
        pass

    unread_count = get_unread_notifications_count()
    notif_btn_label = f"🔔 Central de Notificações ({unread_count})" if unread_count > 0 else "🔔 Central de Notificações"

    with st.popover("☰ Menu"):
        st.markdown("### 📌 Sistemas / Páginas")
        if st.button("📋 Painel de Chamados", use_container_width=True):
            set_page("📋 Painel de Chamados")
        if st.button("🏢 Catálogo de Unidades", use_container_width=True):
            set_page("🏢 Catálogo de Unidades")
        if st.button("📞 Central Telefônica (OXE)", use_container_width=True):
            set_page("📞 Central Telefônica (OXE)")

        if st.button("📅 Plantões da Bancada", use_container_width=True):
            set_page("📅 Plantões da Bancada")
        if st.button("📅 Calendário Geral", use_container_width=True):
            set_page("📅 Calendário Geral")
        if st.button("📜 Portarias da Bancada", use_container_width=True):
            set_page("📜 Portarias da Bancada")
        if st.button("📍 Mapa & Localização", use_container_width=True):
            set_page("📍 Mapa & Localização")
        if st.button("🖥️ Doação & Redistribuição", use_container_width=True):
            set_page("🖥️ Doação & Redistribuição")
        if st.button("📜 Fiscalização de Contratos", use_container_width=True):
            set_page("📜 Fiscalização de Contratos")
        if st.button("🛡️ Controle de Garantia", use_container_width=True):
            set_page("🛡️ Controle de Garantia")
        if st.button("🖨️ Impressoras (PaperCut)", use_container_width=True):
            set_page("🖨️ Impressoras (PaperCut)")

        if st.button("⚡ Scripts de Automação", use_container_width=True):
            set_page("⚡ Scripts de Automação")
        if st.button("📚 FAQ & Tutoriais", use_container_width=True):
            set_page("📚 FAQ & Tutoriais")
        
        st.markdown("---")
        if st.button(notif_btn_label, use_container_width=True, type="primary" if unread_count > 0 else "secondary"):
            set_page("🔔 Central de Notificações")

    return st.session_state["current_page"]
