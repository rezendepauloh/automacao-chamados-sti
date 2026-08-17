import re
import html
import requests
import pandas as pd
import streamlit as st
from urllib.parse import quote
from src.config import setup_logging, DEBUG_DIR_FAQ, ATOS_NORMAS_API_URL, ATOS_NORMAS_DOWNLOAD_URL
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.status_banner import render_log_expander
from src.syncs.sync_portarias import check_portarias_sync_running, read_portarias_last_log_lines

logger = setup_logging(DEBUG_DIR_FAQ / "portarias.log", "portarias")

MEMBROS_BANCADA = [
    "Paulo Henrique Gonçalves Rezende",
    "Reginaldo da Silva Bandeira",
    "Luiz Leonardo Villalba"
]


def clean_text_content(text: str) -> str:
    """Limpa tags HTML, desfaz entidades HTML e corrige caracteres quebrados de unicode."""
    if not text or not isinstance(text, str):
        return ""

    # 1. Desfaz entidades HTML (ex: &quot;, &#39;, &amp;)
    cleaned = html.unescape(text)

    # 2. Remove tags HTML (ex: <strong>, </strong>, <b>, </i>, etc.)
    cleaned = re.sub(r"<[^>]+>", "", cleaned)

    # 3. Tratamento de caracteres de controle e unicode corrompido do MPMS
    cleaned = cleaned.replace("\u0096", "–")
    cleaned = cleaned.replace("\u2013", "–")
    cleaned = cleaned.replace("\u2014", "—")
    cleaned = cleaned.replace("\u00a0", " ")
    cleaned = cleaned.replace("\r", "")

    # Remove quebras de linha acumuladas exageradas
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)

    return cleaned.strip()


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_portarias_bancada():
    """Busca portarias dos membros da bancada na API pública do MPMS."""
    base_url = ATOS_NORMAS_API_URL
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest"
    }

    portarias_dict = {}  # atocod -> portaria_item

    for membro in MEMBROS_BANCADA:
        try:
            params = {
                "atotit": f'"{membro}"',
                "atotipcod[]": "1",
                "atocod": ""
            }
            resp = requests.get(base_url, params=params, headers=headers, timeout=15)
            
            if resp.status_code == 200:
                data = resp.json()
                atos = data.get("atos", [])
                
                for ato in atos:
                    ato_id = ato.get("atocod")
                    if not ato_id:
                        continue
                    
                    titulo_limpo = clean_text_content(ato.get("atotit", ""))
                    texto_limpo = clean_text_content(ato.get("atotxt", ""))
                    
                    # Extrai informações do anexo (PDF)
                    anx_info = ato.get("anxcod") or {}
                    anxcod_val = anx_info.get("anxcod") if isinstance(anx_info, dict) else None
                    anx_file = anx_info.get("anxlin") if isinstance(anx_info, dict) else None

                    # URL de download do PDF usando atocod (ato_id)
                    pdf_url = f"{ATOS_NORMAS_DOWNLOAD_URL}{ato_id}" if ato_id else None

                    # Informações do subtipo e situação
                    subtipo_info = ato.get("atosubtipcod") or {}
                    subtipo_nome = subtipo_info.get("atosubtipnom", "Geral") if isinstance(subtipo_info, dict) else "Geral"

                    origem = ato.get("atoorigem", "Procuradoria-Geral de Justiça")

                    if ato_id not in portarias_dict:
                        portarias_dict[ato_id] = {
                            "id": ato_id,
                            "numero": ato.get("atonum", "S/N"),
                            "diario_num": ato.get("atodjnum", ""),
                            "data_emissao": ato.get("atodta", ""),
                            "data_publicacao": ato.get("atodtapub", ""),
                            "titulo": titulo_limpo,
                            "texto": texto_limpo,
                            "origem": origem,
                            "subtipo": subtipo_nome,
                            "anxcod": anxcod_val,
                            "pdf_nome": anx_file,
                            "pdf_url": pdf_url,
                            "membros": [membro]
                        }
                    else:
                        if membro not in portarias_dict[ato_id]["membros"]:
                            portarias_dict[ato_id]["membros"].append(membro)

        except Exception as e:
            logger.error(f"Erro ao buscar portarias para {membro}: {e}")

    result = list(portarias_dict.values())
    return sorted(result, key=lambda x: x["id"], reverse=True)


def render_portarias_page():
    """Renderiza a página de consulta de Portarias dos membros da Bancada."""
    st.title("📜 Portarias dos Membros da Bancada")
    st.write("Consulta unificada de atos e portarias publicados no diário oficial do MPMS referente aos membros da equipe.")
    st.markdown("---")

    portarias_ativo = check_portarias_sync_running()

    if "was_portarias_syncing" not in st.session_state:
        st.session_state["was_portarias_syncing"] = False

    if st.session_state["was_portarias_syncing"] and not portarias_ativo:
        st.session_state["was_portarias_syncing"] = False
        st.cache_data.clear() # Limpa o cache para forçar a releitura da API
        st.toast("🎉 Sincronização de portarias concluída com sucesso!", icon="✅")
        st.rerun()

    if portarias_ativo:
        st.session_state["was_portarias_syncing"] = True

    render_log_expander(
        "🤖 Sincronização de Portarias em Segundo Plano",
        portarias_ativo,
        read_portarias_last_log_lines,
        check_portarias_sync_running,
        "O robô está consultando a API do Diário Oficial neste momento. O painel permanece livre para uso!"
    )

    with st.sidebar:
        st.markdown("## 🔍 Filtros de Portarias")
        
        selected_membro = st.selectbox(
            "👥 Membro da Bancada:",
            ["Todos"] + MEMBROS_BANCADA,
            key="filter_portaria_membro"
        )

        search_query = st.text_input(
            "🔍 Buscar palavra-chave:",
            "",
            key="filter_portaria_search",
            placeholder="Ex: férias, fiscalização, designar..."
        )

    # Carrega dados com cache de 30 minutos
    with st.spinner("Consultando portarias na API do MPMS..."):
        all_portarias = fetch_portarias_bancada()

    if not all_portarias:
        st.warning("⚠️ Não foi possível obter portarias do MPMS no momento ou nenhuma publicação encontrada.")
        return

    # Extrai anos disponíveis para filtro na sidebar
    anos_disponiveis = sorted(list(set(
        p['data_emissao'].split('/')[-1] for p in all_portarias if p['data_emissao'] and '/' in p['data_emissao']
    )), reverse=True)

    with st.sidebar:
        selected_ano = st.selectbox("📅 Ano:", ["Todos"] + anos_disponiveis, key="filter_portaria_ano")
        sort_order = st.selectbox(
            "⬆️⬇️ Ordenar por Data:",
            ["Mais recentes primeiro (DESC)", "Mais antigas primeiro (ASC)"],
            key="filter_portaria_sort"
        )
        items_per_page = render_items_per_page_selector("portarias", options=[6, 10, 20, 50, 100, "Todos"], default_index=1)

        st.markdown("<br>", unsafe_allow_html=True)
        if portarias_ativo:
            st.button("🤖 Atualizando...", width='stretch', disabled=True)
        else:
            if st.button("🔄 Atualizar Dados (API)", width='stretch', help="Busca novas portarias em segundo plano."):
                import sys, subprocess, time
                creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
                subprocess.Popen([sys.executable, "src/syncs/sync_portarias.py"], creationflags=creationflags)
                time.sleep(0.8)
                st.toast("🚀 Sincronização iniciada em segundo plano!", icon="🤖")
                st.rerun()

    # Aplicação dos filtros
    filtered_portarias = all_portarias.copy()

    if selected_membro != "Todos":
        filtered_portarias = [p for p in filtered_portarias if selected_membro in p['membros']]

    if selected_ano != "Todos":
        filtered_portarias = [p for p in filtered_portarias if p['data_emissao'].endswith(selected_ano)]

    if search_query:
        query_lower = search_query.lower()
        filtered_portarias = [
            p for p in filtered_portarias
            if query_lower in p['titulo'].lower() or query_lower in p['numero'].lower()
        ]

    # Ordenação por Data (ASC / DESC)
    is_desc = (sort_order == "Mais recentes primeiro (DESC)")
    filtered_portarias.sort(key=lambda x: x["id"], reverse=is_desc)


    # Métrica de Resumo em Cards
    m1, m2, m3, m4 = st.columns(4)
    count_paulo = sum(1 for p in all_portarias if "Paulo" in str(p['membros']))
    count_reginaldo = sum(1 for p in all_portarias if "Reginaldo" in str(p['membros']))
    count_luiz = sum(1 for p in all_portarias if "Luiz" in str(p['membros']))

    with m1:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #3b82f6;">
                <div class="metric-title">TOTAL EXIBIDO</div>
                <div class="metric-value">{len(filtered_portarias)}</div>
            </div>
        """, unsafe_allow_html=True)

    with m2:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #10b981;">
                <div class="metric-title">PAULO REZENDE</div>
                <div class="metric-value">{count_paulo}</div>
            </div>
        """, unsafe_allow_html=True)

    with m3:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #f59e0b;">
                <div class="metric-title">REGINALDO BANDEIRA</div>
                <div class="metric-value">{count_reginaldo}</div>
            </div>
        """, unsafe_allow_html=True)

    with m4:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #8b5cf6;">
                <div class="metric-title">LUIZ VILLALBA</div>
                <div class="metric-value">{count_luiz}</div>
            </div>
        """, unsafe_allow_html=True)


    st.markdown("<br>", unsafe_allow_html=True)

    if not filtered_portarias:
        st.info("Nenhuma portaria encontrada para os filtros selecionados.")
        return

    # Modal com Detalhes da Portaria
    @st.dialog("📜 Detalhes da Portaria", width="large")
    def open_portaria_modal(portaria):
        st.subheader(f"Portaria nº {portaria['numero']}")
        st.caption(f"📅 Data da Portaria: **{portaria['data_emissao']}**  |  📰 Diário nº: **{portaria['diario_num']}** ({portaria['data_publicacao']})")
        st.markdown("---")

        st.markdown("### 📌 Envolvido(s) na Portaria")
        badge_html = " ".join([
            f'<span style="background-color: #ff4b4b; color: white; padding: 4px 12px; border-radius: 12px; font-size: 0.85rem; font-weight: bold;">👤 {m}</span>'
            for m in portaria['membros']
        ])
        st.markdown(badge_html, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("### 📝 Ementa / Descrição")
        st.info(portaria['titulo'])

        if portaria['texto']:
            st.markdown("### 📄 Conteúdo Completo")
            st.write(portaria['texto'])

        st.markdown("---")
        c_orig, c_pdf = st.columns([1, 1])
        with c_orig:
            st.caption(f"🏛️ **Origem:** {portaria['origem']} ({portaria['subtipo']})")
        with c_pdf:
            if portaria['pdf_url']:
                st.link_button("📥 Abrir PDF / Anexo MPMS ↗", url=portaria['pdf_url'], type="primary", width='stretch')

    # Paginação dos registros
    page_portarias, current_page, total_pages, total_items = paginate_items(
        filtered_portarias,
        page_key="portarias",
        items_per_page=items_per_page
    )

    # Renderização dos Cards de Portarias fatiados pela página
    cols = st.columns(2)
    for idx, p in enumerate(page_portarias):
        col_target = cols[idx % 2]
        with col_target:
            with st.container(border=True):
                st.caption(f"📜 Portaria nº **{p['numero']}** • Data: {p['data_emissao']}")
                
                membros_str = ", ".join(p['membros'])
                st.markdown(f"**Membro(s):** `{membros_str}`")

                titulo_curto = p['titulo'][:160] + "..." if len(p['titulo']) > 160 else p['titulo']
                st.write(titulo_curto)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                c_b1, c_b2 = st.columns([1, 1])
                with c_b1:
                    if st.button("📖 Ver Detalhes", key=f"btn_port_{p['id']}", width='stretch'):
                        open_portaria_modal(p)
                with c_b2:
                    if p['pdf_url']:
                        st.link_button("📥 PDF MPMS ↗", url=p['pdf_url'], width='stretch')


    # Controles da Paginação ao final da página
    render_pagination_controls(
        page_key="portarias",
        current_page=current_page,
        total_pages=total_pages,
        total_items=total_items,
        items_per_page=items_per_page
    )

