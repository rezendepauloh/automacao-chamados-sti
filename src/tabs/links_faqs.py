import os
import json
import sqlite3
from pathlib import Path
import pandas as pd
import streamlit as st
from src.config import setup_logging, DEBUG_DIR_FAQ, VIDEO_FAQ_DIR

logger = setup_logging(DEBUG_DIR_FAQ / "faq.log", "faq")


def format_file_size(size_in_bytes: int) -> str:
    """Formata bytes em string legível (KB, MB, GB)."""
    if size_in_bytes < 1024 * 1024:
        return f"{size_in_bytes / 1024:.1f} KB"
    elif size_in_bytes < 1024 * 1024 * 1024:
        return f"{size_in_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_in_bytes / (1024 * 1024 * 1024):.2f} GB"


def scan_video_faqs(dir_path: Path):
    """Varre recursivamente o diretório em busca de vídeos e organiza por subpastas/categorias."""
    if not dir_path or not dir_path.exists():
        return []

    valid_extensions = {".mp4", ".mkv", ".mov", ".avi", ".webm", ".wmv"}
    videos = []

    try:
        for file in dir_path.rglob("*"):
            if file.is_file() and file.suffix.lower() in valid_extensions:
                try:
                    relative_parent = file.parent.relative_to(dir_path)
                    categoria = str(relative_parent).replace("\\", " > ").replace("/", " > ")
                    if categoria == ".":
                        categoria = "Geral"
                except Exception:
                    categoria = "Geral"

                try:
                    size_bytes = file.stat().st_size
                    tamanho_fmt = format_file_size(size_bytes)
                except Exception:
                    tamanho_fmt = "N/A"

                videos.append({
                    "titulo": file.stem,
                    "nome_arquivo": file.name,
                    "categoria": categoria,
                    "caminho": file,
                    "tamanho": tamanho_fmt,
                    "extensao": file.suffix.lower()
                })
    except Exception as e:
        logger.error(f"Erro ao varrer diretório de vídeos FAQ: {e}")

    return sorted(videos, key=lambda x: (x["categoria"], x["titulo"]))


def render_faq_page():
    """Renderiza a página de FAQs, Tutoriais do SharePoint, Vídeos FAQ e Links Úteis da Bancada."""
    st.title("📚 FAQ, Tutoriais & Links Úteis da Bancada")
    st.write("Base de conhecimento centralizada com tutoriais da equipe e atalhos rápidos para sistemas externos.")
    st.markdown("---")

    root_dir = Path(__file__).parent.parent.parent
    db_path = root_dir / "chamados.db"
    json_faq_path = root_dir / "temp" / "faqs_template.json"
    json_links_path = root_dir / "temp" / "links_uteis_template.json"
    
    # Carrega dados do SQLite (FAQs)
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faqs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                tipo_faq TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                conteudo TEXT,
                data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        cursor.execute("PRAGMA table_info(faqs)")
        cols_db = [col[1] for col in cursor.fetchall()]
        if "conteudo" not in cols_db:
            cursor.execute("ALTER TABLE faqs ADD COLUMN conteudo TEXT")
            conn.commit()

        cursor.execute("SELECT COUNT(*) FROM faqs")
        count = cursor.fetchone()[0]

        if count == 0 and json_faq_path.exists():
            with open(json_faq_path, "r", encoding="utf-8") as f:
                faqs_json = json.load(f)
            
            for item in faqs_json:
                cursor.execute("""
                    INSERT OR IGNORE INTO faqs (titulo, tipo_faq, url, conteudo)
                    VALUES (?, ?, ?, ?)
                """, (item.get("titulo"), item.get("tipo_faq", "Geral"), item.get("url"), item.get("conteudo")))
            conn.commit()

        df_faqs = pd.read_sql_query("SELECT id, titulo, tipo_faq, url, conteudo FROM faqs", conn)
        conn.close()
    except Exception as e:
        st.error(f"Erro ao carregar banco de dados de FAQs: {e}")
        df_faqs = pd.DataFrame()

    # Carrega links úteis
    links_uteis = []
    if json_links_path.exists():
        try:
            with open(json_links_path, "r", encoding="utf-8") as f:
                links_uteis = json.load(f)
        except Exception as e:
            logger.error(f"Erro ao ler links_uteis_template.json: {e}")

    # Varre vídeos FAQ
    videos_list = scan_video_faqs(VIDEO_FAQ_DIR)

    # Injeção de CSS para transformar o st.radio em botões de abas estilizados (sem bolinhas)
    st.markdown("""
    <style>
    /* Oculta os círculos/bolinhas do radio input */
    div[data-testid="stRadio"] input[type="radio"] {
        display: none !important;
    }
    div[data-testid="stRadio"] div[data-testid="stMarkdownContainer"] p {
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
    }
    /* Container horizontal flex de abas */
    div[data-testid="stRadio"] > div[role="radiogroup"] {
        display: flex !important;
        flex-direction: row !important;
        gap: 10px !important;
        border-bottom: 2px solid #2a2b36 !important;
        padding-bottom: 12px !important;
        margin-bottom: 20px !important;
    }
    /* Estilização padrão dos botões de abas */
    div[data-testid="stRadio"] > div[role="radiogroup"] > label {
        background-color: #1e1f29 !important;
        border: 1px solid #343541 !important;
        border-radius: 8px !important;
        padding: 10px 22px !important;
        color: #b0b0b0 !important;
        cursor: pointer !important;
        transition: all 0.2s ease-in-out !important;
        margin: 0 !important;
    }
    /* Efeito ao passar o mouse */
    div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
        background-color: #2a2b36 !important;
        color: #ffffff !important;
        border-color: #ff4b4b !important;
    }
    /* Estilo para a aba ativa / selecionada */
    div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
        background-color: #ff4b4b !important;
        color: #ffffff !important;
        border-color: #ff4b4b !important;
        font-weight: bold !important;
        box-shadow: 0 4px 12px rgba(255, 75, 75, 0.35) !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # Navegação superior estilo Abas
    active_tab = st.radio(
        "Navegação:",
        [
            "📚 FAQs & Tutoriais (SharePoint)",
            "🎥 Vídeos FAQ (Tutoriais)",
            "🔗 Links Úteis da Bancada"
        ],
        horizontal=True,
        label_visibility="collapsed",
        key="faq_nav_radio"
    )
    st.markdown("<br>", unsafe_allow_html=True)


    # Roteamento dinâmico das Abas com Filtros Específicos na Sidebar
    if active_tab == "📚 FAQs & Tutoriais (SharePoint)":
        st.sidebar.markdown("## 🔍 Filtros do FAQ")
        search_query = st.sidebar.text_input("Buscar por palavra-chave:", "", key="faq_search")
        
        tipos_disponiveis = ["Todos"] + sorted(df_faqs['tipo_faq'].dropna().unique().tolist()) if not df_faqs.empty else ["Todos"]
        selected_tipo = st.sidebar.selectbox("📂 Categoria:", tipos_disponiveis, key="faq_cat")

        if df_faqs.empty:
            st.info("Nenhum FAQ cadastrado no momento.")
        else:
            @st.dialog("📖 Leitor de FAQ", width="large")
            def open_faq_modal(faq_id):
                faq_item = df_faqs[df_faqs['id'] == faq_id].iloc[0]
                
                st.markdown("""
                <style>
                .faq-container {
                    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
                    color: #e0e0e0;
                    line-height: 1.6;
                }
                .faq-container h1, .faq-container h2, .faq-container h3, .faq-container h4 {
                    color: #ffffff !important;
                    margin-top: 1.5rem;
                    margin-bottom: 0.75rem;
                    font-weight: 600;
                }
                .faq-container p {
                    margin-bottom: 1rem;
                    font-size: 0.95rem;
                }
                .faq-container img {
                    max-width: 100% !important;
                    height: auto !important;
                    border-radius: 8px !important;
                    margin: 16px 0 !important;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.4) !important;
                    border: 1px solid #343541 !important;
                }
                .faq-container ol, .faq-container ul {
                    padding-left: 1.5rem;
                    margin-bottom: 1rem;
                }
                .faq-container li {
                    margin-bottom: 0.4rem;
                }
                .faq-container code {
                    background-color: #2a2b36;
                    color: #ff4b4b;
                    padding: 2px 6px;
                    border-radius: 4px;
                    font-size: 0.9rem;
                }
                </style>
                """, unsafe_allow_html=True)

                st.subheader(faq_item['titulo'])
                st.caption(f"Categoria: **{faq_item['tipo_faq']}**")
                st.markdown("---")
                
                if faq_item['conteudo'] and str(faq_item['conteudo']).strip():
                    st.markdown(f'<div class="faq-container">{faq_item["conteudo"]}</div>', unsafe_allow_html=True)
                else:
                    st.info("O conteúdo detalhado deste FAQ ainda não foi sincronizado localmente.")
                    st.write("Você pode visualizar o tutorial completo diretamente no SharePoint pelo botão abaixo.")
                    
                st.markdown("---")
                st.markdown(f'<a href="{faq_item["url"]}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 8px 16px; border-radius: 6px;">🔗 Abrir no SharePoint (Nova Aba) ↗</a>', unsafe_allow_html=True)

            filtered_df = df_faqs.copy()
            if search_query:
                filtered_df = filtered_df[filtered_df['titulo'].str.contains(search_query, case=False, na=False)]
            if selected_tipo != "Todos":
                filtered_df = filtered_df[filtered_df['tipo_faq'] == selected_tipo]

            st.markdown(f"**Exibindo {len(filtered_df)} de {len(df_faqs)} FAQs / Tutoriais**")
            st.markdown("<br>", unsafe_allow_html=True)

            cols = st.columns(2)
            for index, row in filtered_df.iterrows():
                col_target = cols[index % 2]
                with col_target:
                    with st.container(border=True):
                        st.caption(f"📌 {row['tipo_faq']}")
                        st.subheader(row['titulo'])
                        
                        c_btn1, c_btn2 = st.columns([1, 1])
                        with c_btn1:
                            if st.button("📖 Ler Tutorial", key=f"btn_read_{row['id']}", use_container_width=True):
                                open_faq_modal(row['id'])
                        with c_btn2:
                            st.markdown(f'<a href="{row["url"]}" target="_blank" style="display: block; text-align: center; background-color: #2a2b36; border: 1px solid #343541; color: white; text-decoration: none; font-size: 0.85rem; padding: 6px; border-radius: 6px; font-weight: bold;">🔗 SharePoint ↗</a>', unsafe_allow_html=True)

    elif active_tab == "🎥 Vídeos FAQ (Tutoriais)":
        st.sidebar.markdown("## 🔍 Filtros de Vídeos FAQ")
        search_vid = st.sidebar.text_input("Pesquisar vídeo por palavra-chave:", "", key="search_video_input")
        
        categorias_vid = ["Todas"] + sorted(list(set(v['categoria'] for v in videos_list))) if videos_list else ["Todas"]
        selected_cat_vid = st.sidebar.selectbox("📂 Categoria / Pasta:", categorias_vid, key="select_cat_video")

        st.subheader("🎥 Vídeos de FAQ & Tutoriais da Bancada")
        st.write("Vídeos demonstrativos armazenados localmente e sincronizados via SharePoint.")
        st.markdown("<br>", unsafe_allow_html=True)

        if not VIDEO_FAQ_DIR.exists():
            st.warning(f"⚠️ O diretório de Vídeos FAQ não foi encontrado em:\n`{VIDEO_FAQ_DIR}`\n\nVerifique se a pasta existe ou ajuste a variável `VIDEO_FAQ_PATH` no seu arquivo `.env`.")
        elif not videos_list:
            st.info(f"Nenhum vídeo de FAQ encontrado na pasta:\n`{VIDEO_FAQ_DIR}`")
        else:
            @st.dialog("🎥 Reproduzir Vídeo FAQ", width="large")
            def open_video_modal(video_item):
                st.subheader(video_item['titulo'])
                st.caption(f"📂 Categoria / Pasta: **{video_item['categoria']}**  |  💾 Tamanho: **{video_item['tamanho']}**")
                st.markdown("---")

                st.markdown("""
                <style>
                div[data-testid="stDialog"] video, video {
                    max-height: 620px !important;
                    max-width: 100% !important;
                    object-fit: contain !important;
                    margin: 0 auto !important;
                    display: block !important;
                    border-radius: 8px !important;
                    box-shadow: 0 4px 14px rgba(0,0,0,0.5) !important;
                }
                </style>
                """, unsafe_allow_html=True)

                c_left, c_main, c_right = st.columns([0.1, 3.8, 0.1])
                with c_main:
                    try:
                        ext = video_item['extensao'].lower().replace('.', '')
                        mime_map = {
                            'mp4': 'video/mp4',
                            'webm': 'video/webm',
                            'mov': 'video/mp4',
                            'mkv': 'video/mp4',
                            'avi': 'video/x-msvideo',
                            'wmv': 'video/x-ms-wmv'
                        }
                        mime_type = mime_map.get(ext, 'video/mp4')

                        with open(video_item['caminho'], 'rb') as f:
                            video_bytes = f.read()

                        st.video(video_bytes, format=mime_type)
                    except Exception as e_vid:
                        try:
                            st.video(str(video_item['caminho']), format="video/mp4")
                        except Exception as e_fb:
                            st.error(f"Erro ao carregar reprodução do vídeo: {e_fb}")

                st.markdown("---")
                
                c_info, c_act = st.columns([2, 1])
                with c_info:
                    st.caption(f"📁 **Caminho do Arquivo:** `{video_item['caminho']}`")
                with c_act:
                    if st.button("🖥️ Abrir no Player do Windows", key=f"btn_win_open_{hash(video_item['titulo'])}", use_container_width=True):
                        try:
                            os.startfile(str(video_item['caminho']))
                            st.toast("Vídeo aberto no player nativo do Windows!", icon="🎬")
                        except Exception as e_start:
                            st.error(f"Erro ao abrir arquivo: {e_start}")

            filtered_videos = videos_list
            if search_vid:
                filtered_videos = [v for v in filtered_videos if search_vid.lower() in v['titulo'].lower()]
            if selected_cat_vid != "Todas":
                filtered_videos = [v for v in filtered_videos if v['categoria'] == selected_cat_vid]

            st.markdown(f"**Exibindo {len(filtered_videos)} de {len(videos_list)} vídeo(s)**")
            st.markdown("<br>", unsafe_allow_html=True)

            if not filtered_videos:
                st.info("Nenhum vídeo corresponde aos filtros selecionados.")
            else:
                vid_cols = st.columns(2)
                for idx, vid in enumerate(filtered_videos):
                    col_target = vid_cols[idx % 2]
                    with col_target:
                        with st.container(border=True):
                            st.caption(f"📂 {vid['categoria']}")
                            st.markdown(f"### 🎬 {vid['titulo']}")
                            st.caption(f"📁 `{vid['nome_arquivo']}` • {vid['tamanho']}")
                            st.markdown("<br>", unsafe_allow_html=True)

                            if st.button("🎥 Assistir Vídeo", key=f"btn_vid_{idx}_{hash(vid['titulo'])}", use_container_width=True):
                                open_video_modal(vid)

    elif active_tab == "🔗 Links Úteis da Bancada":
        st.sidebar.markdown("## 🔍 Filtros de Links")
        search_link = st.sidebar.text_input("Pesquisar por nome ou URL:", "", key="search_link_input")

        st.subheader("🌐 Links e Atalhos Rápidos da Bancada")
        st.write("Acesso direto aos sistemas operacionais, filas de atendimento e ferramentas externas.")
        st.markdown("<br>", unsafe_allow_html=True)

        if not links_uteis:
            st.info("Nenhum link útil cadastrado em `temp/links_uteis_template.json`.")
        else:
            filtered_links = links_uteis
            if search_link:
                filtered_links = [l for l in filtered_links if search_link.lower() in l.get("titulo", "").lower()]

            st.markdown(f"**Exibindo {len(filtered_links)} de {len(links_uteis)} link(s)**")
            st.markdown("<br>", unsafe_allow_html=True)

            link_cols = st.columns(3)
            for idx, item in enumerate(filtered_links):
                col_lk = link_cols[idx % 3]
                with col_lk:
                    with st.container(border=True):
                        st.markdown(f"#### 🚀 {item.get('titulo', 'Sem Título')}")
                        st.caption(item.get("url", ""))
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(
                            f'<a href="{item.get("url")}" target="_blank" style="display: block; text-align: center; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 10px; border-radius: 6px;">🔗 Acessar Sistema ↗</a>', 
                            unsafe_allow_html=True
                        )
