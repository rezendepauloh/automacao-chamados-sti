"""
Componente reutilizável para exibição de Cards KPI / Métricas no Streamlit.
Padroniza visualmente as métricas do sistema com base nos estilos definidos em assets/css/styles.css.
"""

from typing import Any, Dict, List, Optional, Union
import html
import textwrap
import streamlit as st


def render_metric_card(
    title: str,
    value: Any,
    border_color: str = "#3b82f6",
    value_color: Optional[str] = None,
    title_color: Optional[str] = None,
    subtitle: Optional[str] = None,
    extra_html: Optional[str] = None,
    text_align: str = "left",
) -> None:
    """
    Renderiza um card individual de métrica/KPI com visual padronizado.

    Args:
        title (str): Rótulo ou título da métrica (ex: "TOTAL DE USUÁRIOS").
        value (Any): Valor principal a ser exibido.
        border_color (str): Cor da borda esquerda (ex: "#3b82f6", "#10b981", "#ef4444").
        value_color (str, optional): Cor customizada para o valor numérico. Se None, usa o padrão do tema.
        title_color (str, optional): Cor customizada para o título. Se None, usa o padrão do tema.
        subtitle (str, optional): Texto auxiliar exibido abaixo ou ao lado do valor (ex: "processos").
        extra_html (str, optional): HTML adicional para rodapé ou detalhes do card.
        text_align (str): Alinhamento do conteúdo ("left", "center", "right"). Padrão: "left".
    """
    style_align = f"text-align: {text_align};" if text_align != "left" else ""
    card_style = f'style="border-left-color: {border_color}; {style_align}"'.strip()

    val_style = f'style="color: {value_color};"' if value_color else ""
    tit_style = f'style="color: {title_color};"' if title_color else ""

    subtitle_html = (
        f' <span style="font-size: 0.9rem; opacity: 0.75; font-weight: normal;">{html.escape(str(subtitle))}</span>'
        if subtitle
        else ""
    )

    extra_content = extra_html.strip() if extra_html else ""

    html_content = textwrap.dedent(f"""<div class="metric-card" {card_style}>
<div class="metric-title" {tit_style}>{title}</div>
<div class="metric-value" {val_style}>{value}{subtitle_html}</div>
{extra_content}
</div>""").strip()
    st.markdown(html_content, unsafe_allow_html=True)


def render_metric_cards(
    cards: List[Dict[str, Any]],
    cols: Optional[Union[int, List[int], List[float]]] = None,
) -> None:
    """
    Renderiza múltiplos cards de métrica distribuídos em colunas st.columns.

    Args:
        cards (list of dict): Lista de dicionários contendo os parâmetros de cada card.
            Chaves aceitas em cada dict:
                - title (str) [obrigatório]
                - value (Any) [obrigatório]
                - border_color (str, opcional)
                - value_color (str, opcional)
                - title_color (str, opcional)
                - subtitle (str, opcional)
                - extra_html (str, opcional)
                - text_align (str, opcional)
        cols (int | list, optional): Quantidade de colunas ou proporção das colunas.
            Se None, cria automaticamente `len(cards)` colunas.
    """
    if not cards:
        return

    num_cards = len(cards)
    col_layout = cols if cols is not None else num_cards

    columns = st.columns(col_layout)

    for idx, card in enumerate(cards):
        col_target = columns[idx % len(columns)]
        with col_target:
            render_metric_card(
                title=card.get("title", ""),
                value=card.get("value", ""),
                border_color=card.get("border_color", "#3b82f6"),
                value_color=card.get("value_color"),
                title_color=card.get("title_color"),
                subtitle=card.get("subtitle"),
                extra_html=card.get("extra_html"),
                text_align=card.get("text_align", "left"),
            )
