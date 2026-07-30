import re
import html
import requests
import pandas as pd
import streamlit as st
from urllib.parse import quote
from src.config import setup_logging, DEBUG_DIR_FAQ, ATOS_NORMAS_API_URL, ATOS_NORMAS_DOWNLOAD_URL

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

    # Botão para forçar atualização do cache na sidebar
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
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Atualizar Dados (API)", use_container_width=True):
            st.cache_data.clear()
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

    # Métrica de Resumo em Cards
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Total Exibido", f"{len(filtered_portarias)}")
    with m2:
        count_paulo = sum(1 for p in all_portarias if "Paulo" in str(p['membros']))
        st.metric("Paulo Rezende", f"{count_paulo}")
    with m3:
        count_reginaldo = sum(1 for p in all_portarias if "Reginaldo" in str(p['membros']))
        st.metric("Reginaldo Bandeira", f"{count_reginaldo}")
    with m4:
        count_luiz = sum(1 for p in all_portarias if "Luiz" in str(p['membros']))
        st.metric("Luiz Villalba", f"{count_luiz}")

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
                st.markdown(
                    f'<a href="{portaria["pdf_url"]}" target="_blank" style="display: inline-block; background-color: #2a2b36; border: 1px solid #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 8px 16px; border-radius: 6px; float: right;">📥 Abrir PDF / Anexo MPMS ↗</a>',
                    unsafe_allow_html=True
                )

    # Renderização dos Cards de Portarias
    cols = st.columns(2)
    for idx, p in enumerate(filtered_portarias):
        col_target = cols[idx % 2]
        with col_target:
            with st.container(border=True):
                st.caption(f"📜 Portaria nº **{p['numero']}** • Data: {p['data_emissao']}")
                
                # Exibe tags dos membros envolvidos
                membros_str = ", ".join(p['membros'])
                st.markdown(f"**Membro(s):** `{membros_str}`")

                # Resumo do título cortado
                titulo_curto = p['titulo'][:160] + "..." if len(p['titulo']) > 160 else p['titulo']
                st.write(titulo_curto)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                c_b1, c_b2 = st.columns([1, 1])
                with c_b1:
                    if st.button("📖 Ver Detalhes", key=f"btn_port_{p['id']}", use_container_width=True):
                        open_portaria_modal(p)
                with c_b2:
                    if p['pdf_url']:
                        st.markdown(
                            f'<a href="{p["pdf_url"]}" target="_blank" style="display: block; text-align: center; background-color: #2a2b36; border: 1px solid #343541; color: white; text-decoration: none; font-size: 0.85rem; padding: 6px; border-radius: 6px; font-weight: bold;">📥 PDF MPMS ↗</a>',
                            unsafe_allow_html=True
                        )
