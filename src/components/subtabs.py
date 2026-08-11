import streamlit as st


def render_subtabs(tab_map: dict, default_slug: str, key: str) -> str:
    """
    Renderiza um st.radio estilizado como abas nativas e sincroniza a escolha com st.query_params["subtab"].

    Args:
        tab_map (dict): Mapeamento de slug para título da aba (ex: {"subtab1": "Aba 1"}).
        default_slug (str): Slug padrão a ser utilizado se a URL não possuir um slug válido.
        key (str): Chave única para o componente st.radio.

    Returns:
        str: O título da aba selecionada (valor correspondente do dict).
    """
    subtab_slugs = list(tab_map.keys())

    # Lê o subtab da URL ou define pelo default_slug
    current_subtab_url = st.query_params.get("subtab", default_slug)
    if current_subtab_url not in subtab_slugs:
        current_subtab_url = default_slug

    default_index = subtab_slugs.index(current_subtab_url)

    selected_slug = st.radio(
        label="subtabs_nav",
        options=subtab_slugs,
        format_func=lambda slug: tab_map[slug],
        index=default_index,
        horizontal=True,
        label_visibility="collapsed",
        key=key,
    )

    if selected_slug != current_subtab_url:
        st.query_params["subtab"] = selected_slug

    return tab_map[selected_slug]
