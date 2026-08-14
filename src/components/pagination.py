import math
import pandas as pd
import streamlit as st


def render_items_per_page_selector(
    key_prefix: str,
    options: list = [10, 20, 50, 100, "Todos"],
    default_index: int = 1,
    label: str = "📄 Itens por página:"
) -> int:
    """
    Renderiza um selectbox no sidebar (ItemsPerPage) para controlar a quantidade de registros por página.
    Suporta números inteiros e a opção 'Todos' (retorna 999999).
    """
    per_page_key = f"{key_prefix}_items_per_page"
    selected = st.sidebar.selectbox(
        label,
        options=options,
        index=default_index,
        key=per_page_key
    )
    if str(selected).lower() == "todos":
        return 999999
    try:
        return int(selected)
    except (ValueError, TypeError):
        return 10



def paginate_items(
    items: list | pd.DataFrame,
    page_key: str,
    items_per_page: int = 10
) -> tuple[list | pd.DataFrame, int, int, int]:
    """
    Fatia a lista de itens ou DataFrame para a página atual e gerencia o estado em st.session_state.

    Retorna:
    - items_slice (os registros fatiados da página ativa)
    - current_page (página atual 1-indexed)
    - total_pages (total de páginas calculadas)
    - total_items (quantidade total de elementos)
    """
    total_items = len(items)
    if total_items == 0:
        return (items.iloc[:0] if hasattr(items, 'iloc') else []), 1, 0, 0

    total_pages = max(1, math.ceil(total_items / items_per_page))

    state_key = f"{page_key}_current_page"
    if state_key not in st.session_state:
        st.session_state[state_key] = 1

    current_page = st.session_state[state_key]
    if current_page > total_pages:
        current_page = total_pages
        st.session_state[state_key] = total_pages
    elif current_page < 1:
        current_page = 1
        st.session_state[state_key] = 1

    start_idx = (current_page - 1) * items_per_page
    end_idx = start_idx + items_per_page

    if hasattr(items, 'iloc'):
        items_slice = items.iloc[start_idx:end_idx]
    else:
        items_slice = items[start_idx:end_idx]

    return items_slice, current_page, total_pages, total_items


def render_pagination_controls(
    page_key: str,
    current_page: int,
    total_pages: int,
    total_items: int,
    items_per_page: int
):
    """
    Renderiza a régua de botões de navegação da paginação (Anterior, Páginas, Reticências, Próximo e Resumo).
    Oculta-se automaticamente se total_pages <= 1.
    """
    if total_pages <= 1:
        return

    state_key = f"{page_key}_current_page"

    st.markdown("""
    <style>
    .pagination-summary {
        text-align: center;
        margin-top: 15px;
        margin-bottom: 10px;
        color: #a0a0a0;
        font-size: 0.88rem;
    }
    </style>
    """, unsafe_allow_html=True)

    start_num = (current_page - 1) * items_per_page + 1
    end_num = min(current_page * items_per_page, total_items)

    st.markdown(
        f'<div class="pagination-summary">Exibindo <b>{start_num}</b> até <b>{end_num}</b> de <b>{total_items}</b> registros • Página <b>{current_page}</b> de <b>{total_pages}</b></div>',
        unsafe_allow_html=True
    )

    # Calculamos as páginas visíveis com reticências
    max_visible_neighbors = 1
    pages_to_show = []

    pages_to_show.append(1)

    start_p = max(2, current_page - max_visible_neighbors)
    end_p = min(total_pages - 1, current_page + max_visible_neighbors)

    if start_p > 2:
        pages_to_show.append("...")

    for p in range(start_p, end_p + 1):
        pages_to_show.append(p)

    if end_p < total_pages - 1:
        pages_to_show.append("...")

    if total_pages > 1 and total_pages not in pages_to_show:
        pages_to_show.append(total_pages)

    num_buttons = len(pages_to_show)
    cols = st.columns([1.2] + [0.8] * num_buttons + [1.2])

    with cols[0]:
        if st.button("⬅️ Anterior", key=f"{page_key}_btn_prev", disabled=(current_page == 1), width='stretch'):
            st.session_state[state_key] = current_page - 1
            st.rerun()

    for idx, p in enumerate(pages_to_show, start=1):
        with cols[idx]:
            if p == "...":
                st.markdown("<div style='text-align: center; padding-top: 6px; font-weight: bold; color: #888;'>...</div>", unsafe_allow_html=True)
            else:
                is_active = (p == current_page)
                btn_type = "primary" if is_active else "secondary"
                if st.button(f"{p}", key=f"{page_key}_btn_p_{p}", type=btn_type, width='stretch'):
                    st.session_state[state_key] = p
                    st.rerun()

    with cols[-1]:
        if st.button("Próxima ➡️", key=f"{page_key}_btn_next", disabled=(current_page == total_pages), width='stretch'):
            st.session_state[state_key] = current_page + 1
            st.rerun()
