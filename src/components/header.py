import streamlit as st
from src.database import get_unread_notifications_count, get_notifications


def render_header_navigation() -> str:
    """
    Renderiza o menu hambúrguer popover fixado no canto superior direito do header nativo.
    Exibe notificações em toast para novos alertas e retorna a página atualmente selecionada.
    """
    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "📋 Painel de Chamados"

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
            st.session_state["current_page"] = "📋 Painel de Chamados"
            st.rerun()
        if st.button("📅 Plantões da Bancada", use_container_width=True):
            st.session_state["current_page"] = "📅 Plantões da Bancada"
            st.rerun()
        if st.button("📜 Portarias da Bancada", use_container_width=True):
            st.session_state["current_page"] = "📜 Portarias da Bancada"
            st.rerun()
        if st.button("📍 Mapa & Localização", use_container_width=True):
            st.session_state["current_page"] = "📍 Mapa & Localização"
            st.rerun()
        if st.button("🖥️ Doação & Redistribuição", use_container_width=True):
            st.session_state["current_page"] = "🖥️ Doação & Redistribuição"
            st.rerun()
        if st.button("📜 Fiscalização de Contratos", use_container_width=True):
            st.session_state["current_page"] = "📜 Fiscalização de Contratos"
            st.rerun()
        if st.button("🖨️ Impressoras (PaperCut)", use_container_width=True):
            st.session_state["current_page"] = "🖨️ Impressoras (PaperCut)"
            st.rerun()
        if st.button("📚 FAQ & Tutoriais", use_container_width=True):
            st.session_state["current_page"] = "📚 FAQ & Tutoriais"
            st.rerun()
        
        st.markdown("---")
        if st.button(notif_btn_label, use_container_width=True, type="primary" if unread_count > 0 else "secondary"):
            st.session_state["current_page"] = "🔔 Central de Notificações"
            st.rerun()

    return st.session_state["current_page"]
