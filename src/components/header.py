import streamlit as st
from src.database import get_unread_notifications_count, get_notifications
from src.auth import get_current_user, is_admin, logout_user

PAGE_TO_SLUG = {
    "📋 Painel de Chamados": "chamados",
    "🏢 Catálogo de Unidades": "unidades",
    "📞 Central Telefônica (OXE)": "central-telefonica",
    "📅 Plantões da Bancada": "plantoes",
    "🏖️ Férias da Bancada": "ferias",
    "📅 Calendário Geral": "calendario-geral",
    "📜 Portarias da Bancada": "portarias",
    "📍 Mapa & Localização": "mapa",
    "🖥️ Doação & Redistribuição": "redistribuicao",
    "📜 Fiscalização de Contratos": "fiscalizacao",
    "✈️ Viagens da Bancada": "viagens",
    "🛡️ Controle de Garantia": "garantia",
    "🖨️ Impressoras (PaperCut)": "impressoras",
    "🌳 Active Directory (AD)": "active-directory",
    "💻 Inventário SCCM": "sccm",
    "⚡ Scripts de Automação": "scripts-automacao",
    "📚 FAQ & Tutoriais": "faq",
    "🔐 Cofre de Senhas": "senhas",
    "🔔 Central de Notificações": "notificacoes",
    "⚙️ Configurações": "configuracoes",
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
        if "subtab" in st.query_params:
            del st.query_params["subtab"]
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

    def get_btn_type(page_name: str) -> str:
        return "primary" if st.session_state.get("current_page") == page_name else "secondary"

    # Filtra menus de acordo com privilégios RBAC (Cofre restrito para admins)
    user_is_admin = is_admin()

    MENU_ITEMS = [
        ("📋 Painel de Chamados", "hdr_btn_chamados"),
        ("🏢 Catálogo de Unidades", "hdr_btn_unidades"),
        ("📞 Central Telefônica (OXE)", "hdr_btn_central_telefonica"),
        ("📅 Plantões da Bancada", "hdr_btn_plantoes"),
        ("🏖️ Férias da Bancada", "hdr_btn_ferias"),
        ("📅 Calendário Geral", "hdr_btn_calendario_geral"),
        ("📜 Portarias da Bancada", "hdr_btn_portarias"),
        ("📍 Mapa & Localização", "hdr_btn_mapa"),
        ("🖥️ Doação & Redistribuição", "hdr_btn_redistribuicao"),
        ("📜 Fiscalização de Contratos", "hdr_btn_fiscalizacao"),
        ("✈️ Viagens da Bancada", "hdr_btn_viagens"),
        ("🛡️ Controle de Garantia", "hdr_btn_garantia"),
        ("🖨️ Impressoras (PaperCut)", "hdr_btn_impressoras"),
        ("🌳 Active Directory (AD)", "hdr_btn_ad"),
        ("💻 Inventário SCCM", "hdr_btn_sccm"),
        ("⚡ Scripts de Automação", "hdr_btn_scripts_automacao"),
        ("📚 FAQ & Tutoriais", "hdr_btn_faq"),
    ]

    # Somente administradores têm acesso ao Cofre de Senhas
    if user_is_admin:
        MENU_ITEMS.append(("🔐 Cofre de Senhas", "hdr_btn_senhas"))

    FOOTER_ITEMS = [
        ("⚙️ Configurações", "hdr_btn_configuracoes", "⚙️ Configurações"),
        (notif_btn_label, "hdr_btn_notificacoes", "🔔 Central de Notificações")
    ]

    current_user = get_current_user()

    with st.popover("☰ Menu"):
        # Identificação do usuário logado
        if current_user:
            d_name = current_user.get("display_name", current_user.get("username", "Operador"))
            u_role = "🛡️ Administrador" if user_is_admin else "👁️ Consulta"
            st.markdown(f"**👤 {d_name}**")
            st.caption(f"Perfil: `{u_role}` | Login: `{current_user.get('username')}`")

            # Exibe contas administrativas associadas se for admin
            if user_is_admin and current_user.get("admin_sys"):
                st.caption(f"🔑 Admin Sys: `{current_user.get('admin_sys')}` | Admin AD: `{current_user.get('admin_ad')}`")

            if st.button("🚪 Sair / Logout", key="btn_logout_header", use_container_width=True, type="secondary"):
                logout_user()

            st.markdown("---")

        st.markdown("### 📌 Sistemas / Páginas")

        for page_name, btn_key in MENU_ITEMS:
            col_main, col_newtab = st.columns([5, 1], gap="small")
            slug = PAGE_TO_SLUG.get(page_name, "chamados")
            with col_main:
                if st.button(page_name, key=btn_key, use_container_width=True, type=get_btn_type(page_name)):
                    set_page(page_name)
            with col_newtab:
                st.link_button("↗", url=f"?tab={slug}", use_container_width=True, help=f"Abrir {page_name} em nova aba")

        st.markdown("---")

        for label, btn_key, page_name in FOOTER_ITEMS:
            col_main, col_newtab = st.columns([5, 1], gap="small")
            slug = PAGE_TO_SLUG.get(page_name, "chamados")
            with col_main:
                if st.button(label, key=btn_key, use_container_width=True, type=get_btn_type(page_name)):
                    set_page(page_name)
            with col_newtab:
                st.link_button("↗", url=f"?tab={slug}", use_container_width=True, help=f"Abrir {label} em nova aba")

    return st.session_state["current_page"]

