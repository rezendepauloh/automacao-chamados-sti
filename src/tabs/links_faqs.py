import os
import re
import json
import sqlite3
from pathlib import Path
import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from src.config import (
    setup_logging, DEBUG_DIR_FAQ, VIDEO_FAQ_DIR, IMAGE_FAQ_DIR,
    VIDEO_FAQ_URL, IMAGE_FAQ_URL
)
from src.components.subtabs import render_subtabs
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)

logger = setup_logging(DEBUG_DIR_FAQ / "faq.log", "faq")


def parse_sharepoint_content(html_str: str) -> str:
    """
    Traduz elementos de imagens e vídeos do SharePoint, remove bloco de autoria e ícones quebrados.
    """
    if not html_str or not isinstance(html_str, str):
        return ""

    # 1. Limpeza do Bloco de Autoria via Regex
    html_str = re.sub(r'Paulo Henrique.*?Published \d{2}/\d{2}/\d{4}', '', html_str, flags=re.DOTALL | re.IGNORECASE)

    soup = BeautifulSoup(html_str, "html.parser")

    # 2. Remoção de Ícones Quebrados (tags <i>)
    for tag in soup.find_all("i"):
        tag.decompose()

    # 3. Processamento de Imagens (<div class="imagePlugin" data-imageurl="...">)
    image_divs = soup.find_all("div", class_="imagePlugin")
    for div in image_divs:
        img_url = div.get("data-imageurl")
        if img_url:
            if img_url.startswith("/"):
                img_url = f"https://ministeriopublicoms.sharepoint.com{img_url}"
            new_img = soup.new_tag("img", src=img_url, attrs={"class": "sp-image"})
            div.replace_with(new_img)

    # 4. Restauração de Vídeos do SharePoint (data-sp-controldata ou DocumentEmbedWebPart)
    controldata_divs = soup.find_all(lambda t: t.name == "div" and any(k.endswith("controldata") for k in t.attrs))
    for div in controldata_divs:
        raw_control = None
        for k, v in div.attrs.items():
            if k.endswith("controldata"):
                raw_control = v
                break

        if not raw_control:
            continue

        try:
            cdata = json.loads(raw_control)
            file_url = None
            
            # Tenta buscar em properties.file ou properties.serverRelativeUrl
            props = cdata.get("properties", {})
            if isinstance(props, dict):
                file_url = props.get("file") or props.get("serverRelativeUrl") or props.get("url")

            # Tenta buscar em serverProcessedContent.links.serverRelativeUrl se não encontrou
            if not file_url:
                sp_content = cdata.get("serverProcessedContent", {})
                if isinstance(sp_content, dict):
                    links = sp_content.get("links", {})
                    if isinstance(links, dict):
                        file_url = links.get("serverRelativeUrl") or links.get("baseUrl")

            if file_url and any(str(file_url).lower().endswith(ext) for ext in [".mp4", ".mov", ".webm", ".avi"]):
                if str(file_url).startswith("/"):
                    file_url = f"https://ministeriopublicoms.sharepoint.com{file_url}"
                new_video = soup.new_tag(
                    "video",
                    controls="",
                    src=file_url,
                    style="width: 100%; max-height: 500px; border-radius: 8px; margin: 20px 0;"
                )
                div.replace_with(new_video)
        except Exception:
            pass

    return str(soup)



def format_file_size(size_in_bytes: int) -> str:
    """Formata bytes em string legível (KB, MB, GB)."""
    if size_in_bytes < 1024 * 1024:
        return f"{size_in_bytes / 1024:.1f} KB"
    elif size_in_bytes < 1024 * 1024 * 1024:
        return f"{size_in_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_in_bytes / (1024 * 1024 * 1024):.2f} GB"


def open_in_vlc_player(target_path_or_url: str):
    """Abre o arquivo de vídeo ou URL diretamente no VLC Player ou no player padrão do sistema."""
    import subprocess
    target_str = str(target_path_or_url)
    
    # 1. Se estiver no WSL/Linux e puder chamar o VLC do Windows
    vlc_windows_paths = [
        "/mnt/c/Program Files/VideoLAN/VLC/vlc.exe",
        "/mnt/c/Program Files (x86)/VideoLAN/VLC/vlc.exe"
    ]
    for vlc_path in vlc_windows_paths:
        if os.path.exists(vlc_path):
            try:
                # Converte caminho WSL para Windows se for arquivo local
                if target_str.startswith("/"):
                    try:
                        res = subprocess.run(["wslpath", "-w", target_str], capture_output=True, text=True, check=True)
                        win_target = res.stdout.strip()
                    except Exception:
                        win_target = target_str
                else:
                    win_target = target_str
                subprocess.Popen([vlc_path, win_target])
                return True
            except Exception as e:
                logger.error(f"Erro ao abrir no VLC via WSL: {e}")

    # 2. Se tiver comando 'vlc' no PATH Linux
    import shutil
    if shutil.which("vlc"):
        try:
            subprocess.Popen(["vlc", target_str])
            return True
        except Exception as e:
            logger.error(f"Erro ao abrir VLC no Linux: {e}")

    # 3. Se estiver nativo no Windows
    if sys.platform == "win32":
        try:
            os.startfile(target_str)
            return True
        except Exception as e:
            logger.error(f"Erro ao abrir arquivo no Windows: {e}")

    return False


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


def scan_image_faqs(dir_path: Path):
    """Varre recursivamente o diretório em busca de imagens e organiza por subpastas."""
    if not dir_path or not dir_path.exists():
        return []

    valid_extensions = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
    imagens = []

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

                imagens.append({
                    "titulo": file.stem,
                    "nome_arquivo": file.name,
                    "categoria": categoria,
                    "caminho": file,
                    "tamanho": tamanho_fmt,
                    "extensao": file.suffix.lower()
                })
    except Exception as e:
        logger.error(f"Erro ao varrer diretório de imagens FAQ: {e}")

    return sorted(imagens, key=lambda x: (x["categoria"], x["titulo"]))


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

    # Varre vídeos e imagens FAQ
    videos_list = scan_video_faqs(VIDEO_FAQ_DIR)
    imagens_list = scan_image_faqs(IMAGE_FAQ_DIR)

    # Navegação superior estilo Abas com suporte a query parameter (?subtab=slug)
    FAQ_SUBTAB_MAP = {
        "sharepoint": "📚 FAQs & Tutoriais (SharePoint)",
        "videos": "🎥 Vídeos FAQ (Tutoriais)",
        "imagens": "🖼️ Imagens FAQ (Galeria)",
        "links": "🔗 Links Úteis da Bancada"
    }

    active_tab = render_subtabs(FAQ_SUBTAB_MAP, default_slug="sharepoint", key="faq_nav_radio")

    st.markdown("<br>", unsafe_allow_html=True)



    # Roteamento dinâmico das Abas com Filtros Específicos na Sidebar
    if active_tab == "📚 FAQs & Tutoriais (SharePoint)":
        st.sidebar.markdown("## 🔍 Filtros do FAQ")
        search_query = st.sidebar.text_input("Buscar por palavra-chave:", "", key="faq_search")
        
        tipos_disponiveis = ["Todos"] + sorted(df_faqs['tipo_faq'].dropna().unique().tolist()) if not df_faqs.empty else ["Todos"]
        selected_tipo = st.sidebar.selectbox("📂 Categoria:", tipos_disponiveis, key="faq_cat")
        items_per_page_faq = render_items_per_page_selector("faq_sp", options=[6, 10, 20, 50], default_index=1)

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
                    line-height: 1.8;
                    font-size: 1.05rem;
                }
                .faq-container h1, .faq-container h2, .faq-container h3, .faq-container h4 {
                    color: #ffffff !important;
                    margin-top: 1.5rem;
                    margin-bottom: 0.75rem;
                    font-weight: 600;
                }
                .faq-container p {
                    margin-bottom: 1rem;
                    font-size: 1.05rem;
                    line-height: 1.8;
                }
                .faq-container img, .faq-container .sp-image {
                    display: block;
                    margin: 20px auto;
                    max-width: 100%;
                    height: auto;
                    border-radius: 8px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                    border: 1px solid #343541;
                }
                .faq-container ol, .faq-container ul {
                    padding-left: 1.5rem;
                    margin-bottom: 1.5rem;
                }
                .faq-container li {
                    margin-bottom: 0.6rem;
                    line-height: 1.8;
                }
                .faq-container strong, .faq-container b {
                    color: #f8f9fa !important;
                    font-weight: 700;
                }
                .faq-container code {
                    background-color: #2a2b36;
                    color: #ff4b4b;
                    padding: 2px 6px;
                    border-radius: 4px;
                    font-size: 0.95rem;
                }
                </style>
                """, unsafe_allow_html=True)

                st.subheader(faq_item['titulo'])
                st.caption(f"Categoria: **{faq_item['tipo_faq']}**")
                st.markdown("---")
                
                if faq_item['conteudo'] and str(faq_item['conteudo']).strip():
                    parsed_html = parse_sharepoint_content(faq_item['conteudo'])
                    st.markdown(f'<div class="faq-container">{parsed_html}</div>', unsafe_allow_html=True)
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

            # Paginação dos registros do FAQ
            page_faqs, cur_p_faq, tot_p_faq, tot_i_faq = paginate_items(
                filtered_df,
                page_key="faq_sp",
                items_per_page=items_per_page_faq
            )

            cols = st.columns(2)
            for index, row in page_faqs.iterrows():
                col_target = cols[index % 2]
                with col_target:
                    with st.container(border=True):
                        st.caption(f"📌 {row['tipo_faq']}")
                        st.subheader(row['titulo'])
                        
                        c_btn1, c_btn2 = st.columns([1, 1])
                        with c_btn1:
                            if st.button("📖 Ler Tutorial", key=f"btn_read_{row['id']}", width='stretch'):
                                open_faq_modal(row['id'])
                        with c_btn2:
                            st.link_button("🔗 SharePoint ↗", url=row["url"], width='stretch')


            render_pagination_controls("faq_sp", cur_p_faq, tot_p_faq, tot_i_faq, items_per_page_faq)

    elif active_tab == "🎥 Vídeos FAQ (Tutoriais)":
        st.sidebar.markdown("## 🔍 Filtros de Vídeos FAQ")
        search_vid = st.sidebar.text_input("Pesquisar vídeo por palavra-chave:", "", key="search_video_input")
        
        categorias_vid = ["Todas"] + sorted(list(set(v['categoria'] for v in videos_list))) if videos_list else ["Todas"]
        selected_cat_vid = st.sidebar.selectbox("📂 Categoria / Pasta:", categorias_vid, key="select_cat_video")
        items_per_page_vid = render_items_per_page_selector("faq_vid", options=[6, 10, 20, 50], default_index=1)

        @st.dialog("📤 Enviar Novo Vídeo FAQ (.mp4 / .mkv / .webm)")
        def modal_upload_video():
            st.markdown("### 📤 Upload de Vídeo Tutorial")
            st.caption(f"Os arquivos enviados serão salvos diretamente na pasta de tutoriais do app: `{VIDEO_FAQ_DIR}`")
            cat_input = st.text_input("📂 Nome da Categoria / Subpasta:", value="Geral", help="Cria ou organiza o vídeo na pasta correspondente.")
            up_vid = st.file_uploader("Selecione o arquivo de vídeo:", type=["mp4", "webm", "mkv", "mov", "avi", "wmv"], key="uploader_faq_video")
            
            if st.button("⚡ Salvar Vídeo no Servidor", type="primary", width='stretch'):
                if not up_vid:
                    st.warning("Selecione um arquivo de vídeo primeiro.")
                else:
                    dest_dir = VIDEO_FAQ_DIR / (cat_input.strip() if cat_input.strip() else "Geral")
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest_file = dest_dir / up_vid.name
                    with open(dest_file, "wb") as f_v:
                        f_v.write(up_vid.read())
                    st.success(f"🎉 Vídeo '{up_vid.name}' salvo com sucesso!")
                    import time
                    time.sleep(1)
                    st.rerun()

        st.sidebar.markdown("---")
        st.sidebar.markdown("## ⚙️ Ações e SharePoint")
        if VIDEO_FAQ_URL:
            st.sidebar.link_button("🌐 Abrir Pasta no SharePoint ↗", VIDEO_FAQ_URL, width='stretch', help="Abre a pasta oficial de vídeos no SharePoint em nova aba.")
        if st.sidebar.button("📤 Enviar Vídeo (.mp4)", width='stretch', help="Fazer upload de vídeo para a biblioteca local."):
            modal_upload_video()

        col_head1, col_head2 = st.columns([3, 1])
        with col_head1:
            st.subheader("🎥 Vídeos de FAQ & Tutoriais da Bancada")
            st.write("Vídeos demonstrativos da equipe. Reproduza diretamente no navegador ou abra no SharePoint.")
        with col_head2:
            if VIDEO_FAQ_URL:
                st.link_button("🌐 SharePoint ↗", VIDEO_FAQ_URL, width='stretch', help="Acessar pasta no SharePoint Online")

        st.markdown("<br>", unsafe_allow_html=True)

        @st.dialog("🎥 Reproduzir Vídeo FAQ", width="large")
        def open_video_modal(video_item):
            st.subheader(video_item.get('titulo', 'Vídeo'))
            st.caption(f"📂 Categoria / Pasta: **{video_item.get('categoria', 'Geral')}**  |  💾 Tamanho: **{video_item.get('tamanho', 'N/A')}**")
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
                caminho = video_item.get('caminho')
                video_rendered = False
                if caminho and Path(caminho).exists():
                    try:
                        ext = video_item.get('extensao', '.mp4').lower().replace('.', '')
                        mime_map = {
                            'mp4': 'video/mp4',
                            'webm': 'video/webm',
                            'mov': 'video/mp4',
                            'mkv': 'video/mp4',
                            'avi': 'video/x-msvideo',
                            'wmv': 'video/x-ms-wmv'
                        }
                        mime_type = mime_map.get(ext, 'video/mp4')
                        with open(caminho, 'rb') as f:
                            video_bytes = f.read()
                        st.video(video_bytes, format=mime_type)
                        video_rendered = True
                    except Exception as e_vid:
                        try:
                            st.video(str(caminho), format="video/mp4")
                            video_rendered = True
                        except Exception as e_fb:
                            st.error(f"Erro ao reproduzir arquivo local: {e_fb}")
                elif video_item.get('url') and any(video_item.get('url', '').lower().endswith(ext) for ext in ['.mp4', '.webm', '.mov', '.mkv']):
                    try:
                        st.video(video_item['url'])
                        video_rendered = True
                    except Exception:
                        video_rendered = False

                if not video_rendered:
                    st.info("💡 Este item é um tutorial/vídeo hospedado na nuvem do SharePoint. Você pode assistir online no navegador ou abrir a mídia diretamente no player externo VLC.")
                    if VIDEO_FAQ_URL:
                        st.markdown(f"""
                        <div style="background-color: #1e1f29; border: 1px solid #343541; padding: 18px; border-radius: 8px; text-align: center; margin: 10px 0;">
                            <h4 style="color: #ffffff; margin-bottom: 8px;">🎥 Assistir Vídeo no SharePoint Stream</h4>
                            <p style="color: #a0a0b0; font-size: 0.95rem; margin-bottom: 14px;">Abra a gravação diretamente na plataforma corporativa com suporte completo a todos os codecs.</p>
                            <a href="{video_item.get('url') or VIDEO_FAQ_URL}" target="_blank" style="background-color: #ff4b4b; color: white; text-decoration: none; padding: 8px 18px; border-radius: 6px; font-weight: bold; display: inline-block;">🌐 Assistir no SharePoint Stream ↗</a>
                        </div>
                        """, unsafe_allow_html=True)

            st.markdown("---")
            c_info, c_act1, c_act2 = st.columns([2, 1, 1])
            with c_info:
                if video_item.get('caminho'):
                    st.caption(f"📁 **Arquivo Local:** `{video_item['caminho']}`")
                elif video_item.get('url'):
                    st.caption(f"🔗 **URL:** `{video_item['url']}`")
            with c_act1:
                target_media = str(video_item.get('caminho') or video_item.get('url') or VIDEO_FAQ_URL)
                if st.button("🎬 Abrir no VLC", key=f"btn_vlc_modal_{hash(video_item.get('titulo'))}", width='stretch', help="Executa o player VLC com o arquivo ou link do vídeo."):
                    opened = open_in_vlc_player(target_media)
                    if opened:
                        st.toast("Vídeo enviado para o VLC Player!", icon="🎬")
                    else:
                        st.warning("VLC Player não localizado. Abra diretamente pelo link do SharePoint.")
            with c_act2:
                target_sp = video_item.get('url') or VIDEO_FAQ_URL
                if target_sp:
                    st.link_button("🌐 SharePoint ↗", target_sp, width='stretch')

        # Constrói lista de vídeos combinando arquivos locais com tutoriais do SharePoint
        all_videos_display = []
        if videos_list:
            all_videos_display.extend(videos_list)
        else:
            # Fallback de itens conhecidos do FAQ para exibição em cards estilo foto 2
            if not df_faqs.empty:
                for _, f_row in df_faqs.iterrows():
                    all_videos_display.append({
                        "titulo": f_row['titulo'],
                        "nome_arquivo": f"{f_row['titulo']}.mp4",
                        "categoria": f_row['tipo_faq'],
                        "caminho": None,
                        "url": f_row['url'],
                        "tamanho": "Nuvem",
                        "extensao": ".mp4"
                    })

        filtered_videos = all_videos_display
        if search_vid:
            filtered_videos = [v for v in filtered_videos if search_vid.lower() in v['titulo'].lower()]
        if selected_cat_vid != "Todas":
            filtered_videos = [v for v in filtered_videos if v['categoria'] == selected_cat_vid]

        st.markdown(f"**Exibindo {len(filtered_videos)} de {len(all_videos_display)} vídeo(s) / tutoriais**")
        st.markdown("<br>", unsafe_allow_html=True)

        if not filtered_videos:
            st.info("Nenhum vídeo ou tutorial corresponde aos filtros selecionados.")
        else:
            page_videos, cur_p_vid, tot_p_vid, tot_i_vid = paginate_items(
                filtered_videos,
                page_key="faq_vid",
                items_per_page=items_per_page_vid
            )

            vid_cols = st.columns(3)
            for idx, vid in enumerate(page_videos):
                col_target = vid_cols[idx % 3]
                with col_target:
                    with st.container(border=True):
                        st.caption(f"📂 {vid.get('categoria', 'Geral')}")
                        st.markdown(f"#### 🎬 {vid['titulo']}")
                        st.caption(f"📁 `{vid['nome_arquivo']}` • {vid.get('tamanho', 'N/A')}")
                        st.markdown("<br>", unsafe_allow_html=True)

                        c_btn_v1, c_btn_v2, c_btn_v3 = st.columns([1, 1, 1])
                        with c_btn_v1:
                            if st.button("🎥 Assistir", key=f"btn_vid_card_{idx}_{hash(vid['titulo'])}", width='stretch', help="Abrir modal de reprodução"):
                                open_video_modal(vid)
                        with c_btn_v2:
                            target_media = str(vid.get('caminho') or vid.get('url') or VIDEO_FAQ_URL)
                            if st.button("🎬 VLC", key=f"btn_vlc_card_{idx}_{hash(vid['titulo'])}", width='stretch', help="Abrir no VLC Player"):
                                opened = open_in_vlc_player(target_media)
                                if opened:
                                    st.toast("Iniciando VLC...", icon="🎬")
                                else:
                                    st.warning("VLC não localizado.")
                        with c_btn_v3:
                            sp_url = vid.get('url') or VIDEO_FAQ_URL
                            st.link_button("🌐 Nuvem ↗", url=sp_url, width='stretch', help="Abrir no SharePoint")

            render_pagination_controls("faq_vid", cur_p_vid, tot_p_vid, tot_i_vid, items_per_page_vid)

    elif active_tab == "🖼️ Imagens FAQ (Galeria)":
        st.sidebar.markdown("## 🔍 Filtros de Imagens FAQ")
        search_img = st.sidebar.text_input("Pesquisar por palavra-chave:", "", key="search_img_input")

        @st.dialog("📤 Enviar Novas Imagens de FAQ (.png / .jpg / .webp)")
        def modal_upload_imagem():
            st.markdown("### 📤 Upload de Imagens / Tutoriais")
            st.caption(f"As imagens enviadas serão organizadas na pasta de tutoriais do app: `{IMAGE_FAQ_DIR}`")
            cat_input = st.text_input("📂 Nome da Pasta / Tutorial:", value="Geral", help="Cria ou organiza as imagens na subpasta correspondente.")
            up_imgs = st.file_uploader("Selecione uma ou mais imagens:", type=["png", "jpg", "jpeg", "webp", "gif", "bmp"], accept_multiple_files=True, key="uploader_faq_images")

            if st.button("⚡ Salvar Imagens no Servidor", type="primary", width='stretch'):
                if not up_imgs:
                    st.warning("Selecione ao menos uma imagem primeiro.")
                else:
                    dest_dir = IMAGE_FAQ_DIR / (cat_input.strip() if cat_input.strip() else "Geral")
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    for up_item in up_imgs:
                        dest_file = dest_dir / up_item.name
                        with open(dest_file, "wb") as f_i:
                            f_i.write(up_item.read())
                    st.success(f"🎉 {len(up_imgs)} imagem(ns) enviada(s) com sucesso para a pasta '{cat_input}'!")
                    import time
                    time.sleep(1)
                    st.rerun()

        st.sidebar.markdown("---")
        st.sidebar.markdown("## ⚙️ Ações e SharePoint")
        if IMAGE_FAQ_URL:
            st.sidebar.link_button("🌐 Abrir Pasta no SharePoint ↗", IMAGE_FAQ_URL, width='stretch', help="Abre a pasta oficial de imagens no SharePoint em nova aba.")
        if st.sidebar.button("📤 Enviar Imagens", width='stretch', help="Fazer upload de imagens para a galeria local."):
            modal_upload_imagem()

        # Constrói pastas de imagens combinando pastas locais e tutoriais do SharePoint
        folders_dict = {}
        for img in imagens_list:
            cat = img['categoria']
            if cat not in folders_dict:
                folders_dict[cat] = []
            folders_dict[cat].append(img)

        folders_list = [
            {
                "categoria": cat,
                "imagens": sorted(imgs, key=lambda x: x["titulo"]),
                "total": len(imgs),
                "url": IMAGE_FAQ_URL
            }
            for cat, imgs in folders_dict.items()
        ]

        if not folders_list and not df_faqs.empty:
            # Fallback com os tutoriais do SharePoint catalogados no FAQ
            for _, f_row in df_faqs.iterrows():
                folders_list.append({
                    "categoria": f_row['tipo_faq'],
                    "titulo": f_row['titulo'],
                    "imagens": [],
                    "total": 1,
                    "url": f_row['url']
                })

        folders_list.sort(key=lambda x: x.get("categoria", "Geral"))

        categorias_img = ["Todas"] + sorted(list(set(f.get("categoria", "Geral") for f in folders_list)))
        selected_cat_img = st.sidebar.selectbox("📂 Categoria / Pasta:", categorias_img, key="select_cat_img")
        items_per_page_img = render_items_per_page_selector("faq_img", options=[6, 12, 24, 50], default_index=1)

        col_h_img1, col_h_img2 = st.columns([3, 1])
        with col_h_img1:
            st.subheader("🖼️ Galeria de Imagens de FAQ (Tutoriais)")
            st.write("Capturas de tela e diagramas da equipe. Visualize no carrossel ou acesse no SharePoint.")
        with col_h_img2:
            if IMAGE_FAQ_URL:
                st.link_button("🌐 SharePoint ↗", IMAGE_FAQ_URL, width='stretch', help="Acessar pasta de imagens no SharePoint Online")

        st.markdown("<br>", unsafe_allow_html=True)

        filtered_folders = folders_list
        if search_img:
            s_lower = search_img.lower()
            filtered_folders = [
                f for f in filtered_folders
                if s_lower in f.get('categoria', '').lower()
                or s_lower in f.get('titulo', '').lower()
                or any(s_lower in img['titulo'].lower() for img in f.get('imagens', []))
            ]
        if selected_cat_img != "Todas":
            filtered_folders = [f for f in filtered_folders if f.get('categoria') == selected_cat_img]

        st.markdown(f"**Exibindo {len(filtered_folders)} de {len(folders_list)} pasta(s) de tutoriais**")
        st.markdown("<br>", unsafe_allow_html=True)

        @st.dialog("🖼️ Visualizador de Galeria de Fotos (Carrossel)", width="large")
        def open_image_modal():
            folder_name = st.session_state.get('active_img_folder', '')
            matching_folder = next((f for f in folders_list if f.get('categoria') == folder_name or f.get('titulo') == folder_name), None)

            if not matching_folder:
                st.info("Tutorial ou pasta não encontrada.")
                return

            folder_imgs = matching_folder.get('imagens', [])
            if not folder_imgs:
                st.subheader(f"📂 {matching_folder.get('titulo') or matching_folder.get('categoria')}")
                st.info("As imagens deste tutorial ainda não foram baixadas localmente. Acesse a pasta completa no SharePoint:")
                target_u = matching_folder.get('url') or IMAGE_FAQ_URL
                if target_u:
                    st.link_button("🌐 Abrir no SharePoint Online ↗", target_u, width='stretch')
                return

            idx = st.session_state.get('current_img_idx', 0)
            if idx < 0 or idx >= len(folder_imgs):
                idx = 0
                st.session_state['current_img_idx'] = 0

            img_item = folder_imgs[idx]

            st.subheader(f"📂 {folder_name}")
            st.markdown(f"**{img_item['titulo']}**  *(Imagem {idx + 1} de {len(folder_imgs)})*")
            st.markdown("---")

            st.markdown("""
            <style>
            div[data-testid="stDialog"] img {
                max-height: 480px !important;
                max-width: 100% !important;
                object-fit: contain !important;
                margin: 0 auto !important;
                display: block !important;
                border-radius: 8px !important;
                box-shadow: 0 4px 14px rgba(0,0,0,0.4) !important;
            }
            div[data-testid="stDialog"] div[data-testid="stHorizontalBlock"] {
                align-items: center !important;
            }
            div[data-testid="stDialog"] div[data-testid="stColumn"] {
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
            }
            </style>
            """, unsafe_allow_html=True)

            # Navegação do Carrossel de Fotos (Passar e Voltar)
            c_prev, c_img, c_next = st.columns([1, 6, 1])
            with c_prev:
                if st.button("⬅️ Anterior", key="btn_prev_img", width='stretch', disabled=(idx == 0)):
                    st.session_state['current_img_idx'] = idx - 1
                    st.rerun()

            with c_img:
                try:
                    st.image(str(img_item['caminho']), width='stretch')
                except Exception as e:
                    st.error(f"Erro ao carregar a imagem: {e}")

            with c_next:
                if st.button("Próximo ➡️", key="btn_next_img", width='stretch', disabled=(idx == len(folder_imgs) - 1)):
                    st.session_state['current_img_idx'] = idx + 1
                    st.rerun()

            st.markdown("---")
            c_info, c_act = st.columns([3, 1])
            with c_info:
                st.caption(f"💾 **Tamanho:** `{img_item['tamanho']}`")
                st.caption(f"📁 **Arquivo:** `{img_item['caminho']}`")
            with c_act:
                target_url = matching_folder.get('url') or IMAGE_FAQ_URL
                if target_url:
                    st.link_button("🌐 Abrir no SharePoint", target_url, width='stretch')

        if st.session_state.get('active_img_folder'):
            open_image_modal()

        if not filtered_folders:
            st.info("Nenhuma pasta ou tutorial corresponde aos filtros selecionados.")
        else:
            page_folders, cur_p_img, tot_p_img, tot_i_img = paginate_items(
                filtered_folders,
                page_key="faq_img",
                items_per_page=items_per_page_img
            )

            img_cols = st.columns(3)
            for idx, folder in enumerate(page_folders):
                col_target = img_cols[idx % 3]
                with col_target:
                    with st.container(border=True):
                        display_title = folder.get('titulo') or folder.get('categoria')
                        display_cat = folder.get('categoria') if folder.get('titulo') else "Galeria"
                        st.caption(f"📂 {display_cat}")
                        st.markdown(f"#### 📁 {display_title}")
                        st.caption(f"🖼️ **{folder.get('total', 1)}** item(ns) neste tutorial")
                        st.markdown("<br>", unsafe_allow_html=True)

                        col_f1, col_f2 = st.columns(2)
                        with col_f1:
                            if st.button("🖼️ Ver Galeria", key=f"btn_folder_view_{idx}_{hash(display_title)}", width='stretch'):
                                st.session_state['active_img_folder'] = display_title
                                st.session_state['current_img_idx'] = 0
                                st.rerun()
                        with col_f2:
                            sp_url = folder.get('url') or IMAGE_FAQ_URL
                            st.link_button("🌐 SharePoint ↗", url=sp_url, width='stretch')

            render_pagination_controls("faq_img", cur_p_img, tot_p_img, tot_i_img, items_per_page_img)

    elif active_tab == "🔗 Links Úteis da Bancada":
        st.sidebar.markdown("## 🔍 Filtros de Links")
        search_link = st.sidebar.text_input("Pesquisar por nome ou URL:", "", key="search_link_input")
        items_per_page_link = render_items_per_page_selector("faq_links", options=[6, 12, 24, 48], default_index=1)

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

            page_links, cur_p_link, tot_p_link, tot_i_link = paginate_items(
                filtered_links,
                page_key="faq_links",
                items_per_page=items_per_page_link
            )

            link_cols = st.columns(3)
            for idx, item in enumerate(page_links):
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

            render_pagination_controls("faq_links", cur_p_link, tot_p_link, tot_i_link, items_per_page_link)

