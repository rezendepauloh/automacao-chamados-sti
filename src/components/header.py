import streamlit as st

def render_header_navigation() -> str:
    """
    Renderiza o menu hambúrguer popover fixado no canto superior direito do header nativo.
    Retorna a página atualmente selecionada.
    """
    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "📋 Painel de Chamados"

    with st.popover("☰ Menu"):
        st.markdown("### 📌 Sistemas / Páginas")
        if st.button("📋 Painel de Chamados", use_container_width=True):
            st.session_state["current_page"] = "📋 Painel de Chamados"
            st.rerun()
        if st.button("📅 Plantões da Bancada", use_container_width=True):
            st.session_state["current_page"] = "📅 Plantões da Bancada"
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
        if st.button("📚 FAQ & Tutoriais", use_container_width=True):
            st.session_state["current_page"] = "📚 FAQ & Tutoriais"
            st.rerun()

    return st.session_state["current_page"]
