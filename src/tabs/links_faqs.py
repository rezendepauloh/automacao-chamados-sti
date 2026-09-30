import os
import sys
import re
import json
import sqlite3
import subprocess
import shutil
import time
import requests
import base64
import hashlib
from pathlib import Path
import pandas as pd
import streamlit as st
try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

from src.config import (
    setup_logging, DEBUG_DIR_FAQ, VIDEO_FAQ_DIR, IMAGE_FAQ_DIR,
    VIDEO_FAQ_URL, IMAGE_FAQ_URL, CITSMART_EMAIL, PASSWORD,
    HEADLESS, EXPLICIT_WAIT, get_chrome_driver
)
from src.terminal import log, CYAN, GREEN, YELLOW, RED, WHITE
from src.components.subtabs import render_subtabs
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)

logger = setup_logging(DEBUG_DIR_FAQ / "faq.log", "faq")


from urllib.parse import unquote

SHAREPOINT_COOKIES_FILE = Path(__file__).resolve().parent.parent.parent / "uploads" / "faq" / "sharepoint_cookies.json"


def get_sharepoint_cookies() -> dict:
    """
    Obtém cookies autenticados do SharePoint.
    Tenta ler do cache em uploads/faq/sharepoint_cookies.json ou extrair do Firefox no host WSL.
    """
    cookies = {}
    if SHAREPOINT_COOKIES_FILE.exists():
        try:
            with open(SHAREPOINT_COOKIES_FILE, "r", encoding="utf-8") as f:
                cookies = json.load(f)
                if cookies and isinstance(cookies, dict):
                    return cookies
        except Exception as e:
            logger.warning(f"Erro ao ler cache de cookies: {e}")

    # Tenta extrair diretamente do perfil do Firefox do usuário (Windows via WSL)
    ff_profiles_dir = Path("/mnt/c/Users/paulogoncalves/AppData/Roaming/Mozilla/Firefox/Profiles")
    if ff_profiles_dir.exists():
        for p in ff_profiles_dir.glob("*.default-release"):
            cookie_db = p / "cookies.sqlite"
            if cookie_db.exists():
                try:
                    tmp_db = Path("/tmp/ff_faq_cookies.sqlite")
                    shutil.copy2(cookie_db, tmp_db)
                    conn = sqlite3.connect(tmp_db)
                    cursor = conn.cursor()
                    cursor.execute("SELECT name, value FROM moz_cookies WHERE host LIKE '%sharepoint.com'")
                    for name, value in cursor.fetchall():
                        cookies[name] = value
                    conn.close()
                    try:
                        tmp_db.unlink(missing_ok=True)
                    except Exception:
                        pass
                    if cookies:
                        SHAREPOINT_COOKIES_FILE.parent.mkdir(parents=True, exist_ok=True)
                        with open(SHAREPOINT_COOKIES_FILE, "w", encoding="utf-8") as f:
                            json.dump(cookies, f, indent=2)
                        return cookies
                except Exception as err:
                    logger.warning(f"Falha ao extrair cookies do Firefox: {err}")
    return cookies


def get_image_as_base64(file_path: Path) -> str:
    """Codifica a imagem local para formato Data URI base64 seguro."""
    try:
        if not file_path.exists():
            return ""
        with open(file_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        ext = file_path.suffix.lower()
        mime = "image/png"
        if ext in [".jpg", ".jpeg"]:
            mime = "image/jpeg"
        elif ext == ".gif":
            mime = "image/gif"
        elif ext == ".webp":
            mime = "image/webp"
        elif ext == ".svg":
            mime = "image/svg+xml"
        return f"data:{mime};base64,{encoded}"
    except Exception as e:
        logger.error(f"Erro ao converter imagem em base64: {e}")
        return ""


def slugify_faq_title(title: str) -> str:
    """Gera um slug de diretório seguro e legível a partir do título do FAQ."""
    import unicodedata
    if not title:
        return "Geral"
    text = unicodedata.normalize('NFKD', title).encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'[^\w\s-]', '', text).strip()
    slug = re.sub(r'[-\s]+', '_', text)
    return slug if slug else "Geral"


def ensure_sharepoint_image_cached(img_url: str, faq_slug: str = "Geral") -> str:
    """
    Verifica se a imagem do SharePoint já está salva localmente em uploads/faq/imagens/<faq_slug>/.
    Caso exista em uploads/faq/imagens/ raiz (legado), migra ou carrega dela.
    Se não estiver em disco, baixa usando os cookies autenticados da intranet.
    Retorna o Data URI em base64 da imagem para exibição direta no navegador.
    """
    if not img_url:
        return ""

    if img_url.startswith("/"):
        img_url = f"https://ministeriopublicoms.sharepoint.com{img_url}"

    clean_filename = unquote(img_url.split("/")[-1].split("?")[0])
    clean_filename = re.sub(r'[\\/*?:"<>|]', '_', clean_filename)
    if not clean_filename:
        clean_filename = f"img_{hashlib.md5(img_url.encode('utf-8')).hexdigest()[:12]}.jpg"

    slug = slugify_faq_title(faq_slug)
    folder_dir = IMAGE_FAQ_DIR / slug
    dest_file = folder_dir / clean_filename
    root_legacy_file = IMAGE_FAQ_DIR / clean_filename

    # Se já existir na pasta organizada
    if dest_file.exists() and dest_file.stat().st_size > 0:
        return get_image_as_base64(dest_file)

    # Se existir na raiz (legado), migra para a pasta do tutorial
    if root_legacy_file.exists() and root_legacy_file.is_file() and root_legacy_file.stat().st_size > 0:
        try:
            folder_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(root_legacy_file, dest_file)
            return get_image_as_base64(dest_file)
        except Exception:
            return get_image_as_base64(root_legacy_file)

    # Tenta baixar com cookies para a subpasta do tutorial
    cookies = get_sharepoint_cookies()
    if cookies:
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",
                "Accept": "image/avif,image/webp,image/png,image/svg+xml,image/*;q=0.8,*/*;q=0.5"
            }
            resp = requests.get(img_url, cookies=cookies, headers=headers, timeout=12)
            if resp.status_code == 200 and len(resp.content) > 100:
                folder_dir.mkdir(parents=True, exist_ok=True)
                with open(dest_file, "wb") as f_out:
                    f_out.write(resp.content)
                logger.info(f"✅ Imagem FAQ baixada em '{slug}/{clean_filename}'")
                return get_image_as_base64(dest_file)
            else:
                logger.warning(f"Download de imagem retornou status {resp.status_code} para {clean_filename}")
        except Exception as e:
            logger.error(f"Erro ao baixar imagem {img_url}: {e}")

    return ""


def ensure_sharepoint_video_cached(video_url: str, faq_slug: str = "Geral") -> str:
    """
    Verifica se o vídeo do SharePoint já está salvo localmente em uploads/faq/videos/<faq_slug>/.
    Se não estiver, baixa usando os cookies autenticados da intranet.
    Retorna o Data URI em base64 (para vídeos leves) ou caminho local.
    """
    if not video_url:
        return ""

    if video_url.startswith("/"):
        video_url = f"https://ministeriopublicoms.sharepoint.com{video_url}"

    clean_filename = unquote(video_url.split("/")[-1].split("?")[0])
    clean_filename = re.sub(r'[\\/*?:"<>|]', '_', clean_filename)
    if not clean_filename:
        clean_filename = f"video_{hashlib.md5(video_url.encode('utf-8')).hexdigest()[:12]}.mp4"

    slug = slugify_faq_title(faq_slug)
    folder_dir = VIDEO_FAQ_DIR / slug
    dest_file = folder_dir / clean_filename

    # Verifica se o arquivo já existe no destino
    if dest_file.exists() and dest_file.stat().st_size > 1000:
        return get_video_as_base64(dest_file)

    # Verifica se já existe em outra subpasta do VIDEO_FAQ_DIR
    if VIDEO_FAQ_DIR.exists():
        for existing in VIDEO_FAQ_DIR.rglob(clean_filename):
            if existing.is_file() and existing.stat().st_size > 1000:
                try:
                    folder_dir.mkdir(parents=True, exist_ok=True)
                    if existing != dest_file:
                        shutil.copy2(existing, dest_file)
                    return get_video_as_base64(dest_file)
                except Exception:
                    return get_video_as_base64(existing)

    # Tenta baixar com cookies para a subpasta do tutorial
    cookies = get_sharepoint_cookies()
    if cookies:
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",
                "Accept": "*/*"
            }
            resp = requests.get(video_url, cookies=cookies, headers=headers, stream=True, timeout=35)
            if resp.status_code == 200:
                folder_dir.mkdir(parents=True, exist_ok=True)
                with open(dest_file, "wb") as f_out:
                    for chunk in resp.iter_content(chunk_size=65536):
                        if chunk:
                            f_out.write(chunk)
                if dest_file.exists() and dest_file.stat().st_size > 1000:
                    logger.info(f"✅ Vídeo FAQ baixado em '{slug}/{clean_filename}'")
                    return get_video_as_base64(dest_file)
        except Exception as e:
            logger.error(f"Erro ao baixar vídeo {video_url}: {e}")

    return ""


def get_video_as_base64(file_path: Path) -> str:
    """Codifica o arquivo de vídeo local em base64 Data URI."""
    try:
        if not file_path.exists():
            return ""
        # Limite de segurança de 80MB para base64 direto
        if file_path.stat().st_size > 80 * 1024 * 1024:
            return ""
        with open(file_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        ext = file_path.suffix.lower()
        mime = "video/mp4"
        if ext == ".webm":
            mime = "video/webm"
        elif ext == ".mov":
            mime = "video/mp4"
        return f"data:{mime};base64,{encoded}"
    except Exception as e:
        logger.error(f"Erro ao converter vídeo em base64: {e}")
        return ""


def parse_sharepoint_content(html_str: str, faq_slug: str = "Geral") -> str:
    """
    Traduz elementos de imagens e vídeos do SharePoint, remove bloco de autoria,
    ícones quebrados e renderiza imagens com legenda e visualização direta em alta qualidade,
    organizando a mídia por tutorial (subpasta).
    """
    if not html_str or not isinstance(html_str, str):
        return ""

    # 1. Limpeza do Bloco de Autoria via Regex
    html_str = re.sub(r'Paulo Henrique.*?Published \d{2}/\d{2}/\d{4}', '', html_str, flags=re.DOTALL | re.IGNORECASE)

    if BeautifulSoup is None:
        return html_str

    soup = BeautifulSoup(html_str, "html.parser")

    # 2. Remoção de Ícones Quebrados (tags <i>)
    for tag in soup.find_all("i"):
        tag.decompose()

    # 3. Processamento de Imagens (<div class="imagePlugin" data-imageurl="...">)
    image_divs = soup.find_all("div", class_="imagePlugin")
    for div in image_divs:
        img_url = div.get("data-imageurl")
        caption = (div.get("data-captiontext") or "").strip()
        if img_url:
            if img_url.startswith("/"):
                img_url = f"https://ministeriopublicoms.sharepoint.com{img_url}"

            # Tenta obter a imagem em base64 salva localmente na subpasta do tutorial
            base64_src = ensure_sharepoint_image_cached(img_url, faq_slug=faq_slug)

            if base64_src:
                caption_html = f'<figcaption class="sp-caption" style="text-align: center; color: #94a3b8; font-size: 0.88rem; margin-top: 6px; font-style: italic;">{caption}</figcaption>' if caption else ''
                img_tag_html = f'''
                <figure class="sp-img-card" style="margin: 20px auto; max-width: 95%; text-align: center;">
                    <img class="sp-image" src="{base64_src}" style="max-width: 100%; height: auto; border-radius: 8px; border: 1px solid #3b3c4a; box-shadow: 0 4px 14px rgba(0,0,0,0.35); display: block; margin: 0 auto;" />
                    {caption_html}
                </figure>
                '''
                div.replace_with(BeautifulSoup(img_tag_html, "html.parser"))
            else:
                label_text = caption if caption else "Visualizar Imagem do Tutorial"
                card_html = f'''
                <div class="sp-img-card" style="margin: 20px auto; max-width: 720px; background: #1e1f29; border: 1px solid #3b3c4a; border-radius: 10px; padding: 18px 22px; text-align: center; box-shadow: 0 4px 14px rgba(0,0,0,0.35);">
                    <div style="font-size: 1.05rem; font-weight: 600; color: #f8f9fa; margin-bottom: 6px;">
                        🖼️ {label_text}
                    </div>
                    <div style="font-size: 0.85rem; color: #94a3b8; margin-bottom: 14px;">
                        Esta captura de tela requer sincronização de cookies ou abertura direta na intranet.
                    </div>
                    <a href="{img_url}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; font-size: 0.92rem; padding: 10px 22px; border-radius: 6px; box-shadow: 0 2px 8px rgba(255, 75, 75, 0.4); transition: background-color 0.2s;">
                        🔗 Abrir Imagem no SharePoint ↗
                    </a>
                </div>
                '''
                div.replace_with(BeautifulSoup(card_html, "html.parser"))

    # 3b. Processamento de tags <img> nativas remanescentes
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-sp-originalimgsrc") or ""
        if not src or img.get("data-automation-id") == "lowQualityImagePlaceholder":
            img.decompose()
            continue

        if src.startswith("/"):
            src = f"https://ministeriopublicoms.sharepoint.com{src}"

        if "sharepoint.com" in src:
            base64_src = ensure_sharepoint_image_cached(src, faq_slug=faq_slug)
            alt_txt = img.get("alt") or ""
            if base64_src:
                img["src"] = base64_src
                img["class"] = img.get("class", []) + ["sp-image"]
                img["style"] = "max-width: 100%; height: auto; border-radius: 8px; border: 1px solid #3b3c4a; margin: 16px auto; display: block; box-shadow: 0 4px 14px rgba(0,0,0,0.35);"
            else:
                alt_display = alt_txt if alt_txt else "Visualizar Imagem"
                link_card = f'''
                <div class="sp-img-card" style="margin: 16px auto; max-width: 600px; background: #1e1f29; border: 1px solid #3b3c4a; border-radius: 8px; padding: 14px 18px; text-align: center;">
                    <div style="font-size: 0.95rem; font-weight: 600; color: #f8f9fa; margin-bottom: 8px;">
                        🖼️ {alt_display}
                    </div>
                    <a href="{src}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; font-size: 0.85rem; padding: 8px 16px; border-radius: 6px;">
                        🔗 Abrir no SharePoint ↗
                    </a>
                </div>
                '''
                img.replace_with(BeautifulSoup(link_card, "html.parser"))
        else:
            img["src"] = src
            if "sp-image" not in img.get("class", []):
                img["class"] = img.get("class", []) + ["sp-image"]



    # 4. Restauração e Renderização de Vídeos do SharePoint (div.sp-video-embed, controldata ou video tags)
    # 4a. Tags de vídeo enriquecidas (div.sp-video-embed)
    for v_embed in soup.find_all("div", class_="sp-video-embed"):
        v_url = v_embed.get("data-videourl") or ""
        v_title = v_embed.get("data-videotitle") or "Vídeo Tutorial"
        if v_url:
            cached_src = ensure_sharepoint_video_cached(v_url, faq_slug=faq_slug)
            if cached_src:
                v_html = f'''
                <div class="sp-video-card" style="margin: 24px auto; max-width: 95%; text-align: center;">
                    <div style="font-size: 0.95rem; font-weight: 600; color: #f8f9fa; margin-bottom: 8px;">
                        🎬 {v_title}
                    </div>
                    <video controls src="{cached_src}" style="width: 100%; max-height: 520px; border-radius: 8px; border: 1px solid #3b3c4a; box-shadow: 0 4px 14px rgba(0,0,0,0.5);"></video>
                </div>
                '''
                v_embed.replace_with(BeautifulSoup(v_html, "html.parser"))
            else:
                card_v = f'''
                <div class="sp-img-card" style="margin: 20px auto; max-width: 720px; background: #1e1f29; border: 1px solid #3b3c4a; border-radius: 10px; padding: 18px 22px; text-align: center; box-shadow: 0 4px 14px rgba(0,0,0,0.35);">
                    <div style="font-size: 1.05rem; font-weight: 600; color: #f8f9fa; margin-bottom: 6px;">
                        🎬 {v_title}
                    </div>
                    <div style="font-size: 0.85rem; color: #94a3b8; margin-bottom: 14px;">
                        Vídeo corporativo do tutorial disponível no SharePoint Stream.
                    </div>
                    <a href="{v_url}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; font-size: 0.92rem; padding: 10px 22px; border-radius: 6px; box-shadow: 0 2px 8px rgba(255, 75, 75, 0.4);">
                        🔗 Assistir no SharePoint Stream ↗
                    </a>
                </div>
                '''
                v_embed.replace_with(BeautifulSoup(card_v, "html.parser"))

    # 4b. Restauração de controldata de vídeos (caso ainda contenha tags cruas)
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
            v_title = "Vídeo do Tutorial"

            props = cdata.get("properties", {})
            if isinstance(props, dict):
                file_url = props.get("file") or props.get("serverRelativeUrl") or props.get("url")
                v_title = props.get("title") or v_title

            if not file_url:
                sp_content = cdata.get("serverProcessedContent", {})
                if isinstance(sp_content, dict):
                    links = sp_content.get("links", {})
                    if isinstance(links, dict):
                        file_url = links.get("serverRelativeUrl") or links.get("baseUrl")
                    text_dict = sp_content.get("searchablePlainTexts", {})
                    if isinstance(text_dict, dict) and text_dict.get("title"):
                        v_title = text_dict["title"]

            if file_url and any(str(file_url).lower().endswith(ext) for ext in [".mp4", ".mov", ".webm", ".avi"]):
                if str(file_url).startswith("/"):
                    file_url = f"https://ministeriopublicoms.sharepoint.com{file_url}"
                
                cached_src = ensure_sharepoint_video_cached(file_url, faq_slug=faq_slug)
                if cached_src:
                    v_html = f'''
                    <div class="sp-video-card" style="margin: 24px auto; max-width: 95%; text-align: center;">
                        <div style="font-size: 0.95rem; font-weight: 600; color: #f8f9fa; margin-bottom: 8px;">
                            🎬 {v_title}
                        </div>
                        <video controls src="{cached_src}" style="width: 100%; max-height: 520px; border-radius: 8px; border: 1px solid #3b3c4a; box-shadow: 0 4px 14px rgba(0,0,0,0.5);"></video>
                    </div>
                    '''
                    div.replace_with(BeautifulSoup(v_html, "html.parser"))
                else:
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
        logger.info(f"Diretório local de vídeos FAQ não encontrado: {dir_path}")
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
                    "extensao": file.suffix.lower(),
                    "origem": "OneDrive Local" if "OneDrive" in str(file) else "Upload Local"
                })
        
        logger.info(f"Varredura de vídeos concluída em '{dir_path}': {len(videos)} arquivo(s) local(is) indexado(s).")
    except Exception as e:
        logger.error(f"Erro ao varrer diretório de vídeos FAQ: {e}")

    return sorted(videos, key=lambda x: (x["categoria"], x["titulo"]))


def scan_image_faqs(dir_path: Path):
    """Varre recursivamente o diretório em busca de imagens e organiza por subpastas."""
    if not dir_path or not dir_path.exists():
        logger.info(f"Diretório local de imagens FAQ não encontrado: {dir_path}")
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
                    "extensao": file.suffix.lower(),
                    "origem": "OneDrive Local" if "OneDrive" in str(file) else "Upload Local"
                })
        
        logger.info(f"Varredura de imagens concluída em '{dir_path}': {len(imagens)} arquivo(s) local(is) indexado(s).")
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
    json_videos_path = root_dir / "temp" / "videos_faq_template.json"
    json_imagens_path = root_dir / "temp" / "imagens_faq_template.json"
    
    # Carrega dados do SQLite (FAQs, Vídeos FAQ e Imagens FAQ)
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 1. Tabela de FAQs / Artigos
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
        count_faqs = cursor.fetchone()[0]
        if count_faqs == 0 and json_faq_path.exists():
            with open(json_faq_path, "r", encoding="utf-8") as f:
                faqs_json = json.load(f)
            for item in faqs_json:
                cursor.execute("""
                    INSERT OR IGNORE INTO faqs (titulo, tipo_faq, url, conteudo)
                    VALUES (?, ?, ?, ?)
                """, (item.get("titulo"), item.get("tipo_faq", "Geral"), item.get("url"), item.get("conteudo")))
            conn.commit()

        # 2. Tabela de Vídeos FAQ (SharePoint Online como Fonte Primária)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faq_videos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                categoria TEXT NOT NULL,
                nome_arquivo TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                caminho_relativo TEXT,
                tamanho_bytes INTEGER DEFAULT 0,
                data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cursor.execute("SELECT COUNT(*) FROM faq_videos")
        count_vids = cursor.fetchone()[0]
        if count_vids == 0 and json_videos_path.exists():
            with open(json_videos_path, "r", encoding="utf-8") as f:
                vids_json = json.load(f)
            for item in vids_json:
                cursor.execute("""
                    INSERT OR IGNORE INTO faq_videos (titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (item.get("titulo"), item.get("categoria", "Geral"), item.get("nome_arquivo"), item.get("url"), item.get("caminho_relativo"), item.get("tamanho_bytes", 0)))
            conn.commit()

        # 3. Tabela de Imagens FAQ (SharePoint Online como Fonte Primária)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS faq_imagens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                categoria TEXT NOT NULL,
                nome_arquivo TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                caminho_relativo TEXT,
                tamanho_bytes INTEGER DEFAULT 0,
                data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cursor.execute("SELECT COUNT(*) FROM faq_imagens")
        count_imgs = cursor.fetchone()[0]
        if count_imgs == 0 and json_imagens_path.exists():
            with open(json_imagens_path, "r", encoding="utf-8") as f:
                imgs_json = json.load(f)
            for item in imgs_json:
                cursor.execute("""
                    INSERT OR IGNORE INTO faq_imagens (titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (item.get("titulo"), item.get("categoria", "Geral"), item.get("nome_arquivo"), item.get("url"), item.get("caminho_relativo"), item.get("tamanho_bytes", 0)))
            conn.commit()

        # Leitura dos dados a partir do SQLite
        df_faqs = pd.read_sql_query("SELECT id, titulo, tipo_faq, url, conteudo FROM faqs", conn)
        df_vids = pd.read_sql_query("SELECT id, titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes FROM faq_videos", conn)
        df_imgs = pd.read_sql_query("SELECT id, titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes FROM faq_imagens", conn)
        conn.close()
    except Exception as e:
        logger.error(f"Erro ao carregar banco de dados de FAQs/Vídeos/Imagens: {e}")
        df_faqs = pd.DataFrame()
        df_vids = pd.DataFrame()
        df_imgs = pd.DataFrame()

    # Carrega links úteis
    links_uteis = []
    if json_links_path.exists():
        try:
            with open(json_links_path, "r", encoding="utf-8") as f:
                links_uteis = json.load(f)
        except Exception as e:
            logger.error(f"Erro ao ler links_uteis_template.json: {e}")

    # Catálogos primários a partir do SQLite
    catalog_videos = df_vids.to_dict('records') if not df_vids.empty else []
    catalog_imagens = df_imgs.to_dict('records') if not df_imgs.empty else []

    # Varre vídeos e imagens locais (se existirem na pasta de upload ou cache)
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

        st.sidebar.markdown("---")
        st.sidebar.markdown("---")
        st.sidebar.markdown("## 📥 Mídia Offline")
        if st.sidebar.button("🔄 Sincronizar Mídias dos FAQs", width="stretch", help="Baixa e atualiza as imagens e vídeos de todos os FAQs do SharePoint usando os cookies da sua sessão ativa, organizados por pasta de cada tutorial."):
            with st.spinner("Sincronizando mídias dos FAQs autenticados..."):
                cookies = get_sharepoint_cookies()
                if not cookies:
                    st.sidebar.error("Nenhum cookie de sessão encontrado no navegador.")
                else:
                    sucessos_img = 0
                    sucessos_vid = 0
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",
                        "Accept": "*/*"
                    }
                    IMAGE_FAQ_DIR.mkdir(parents=True, exist_ok=True)
                    VIDEO_FAQ_DIR.mkdir(parents=True, exist_ok=True)
                    
                    for _, r in df_faqs.iterrows():
                        content = r.get("conteudo") or ""
                        faq_tit = r.get("titulo") or "Geral"
                        faq_slug = slugify_faq_title(faq_tit)
                        target_img_folder = IMAGE_FAQ_DIR / faq_slug
                        target_vid_folder = VIDEO_FAQ_DIR / faq_slug
                        target_img_folder.mkdir(parents=True, exist_ok=True)
                        target_vid_folder.mkdir(parents=True, exist_ok=True)
                        
                        # 1. Sincronização de Imagens
                        found_img_urls = re.findall(r'data-imageurl="([^"]+)"', content) + re.findall(r'src="([^"]+sharepoint\.com[^"]+)"', content)
                        for u in set(found_img_urls):
                            if u.startswith("/"):
                                u = f"https://ministeriopublicoms.sharepoint.com{u}"
                            clean_fn = unquote(u.split("/")[-1].split("?")[0])
                            clean_fn = re.sub(r'[\\/*?:"<>|]', '_', clean_fn)
                            if not clean_fn:
                                clean_fn = f"img_{hashlib.md5(u.encode('utf-8')).hexdigest()[:12]}.jpg"
                            dest = target_img_folder / clean_fn
                            
                            # Se já existe na raiz legada, copia para a subpasta
                            root_legacy = IMAGE_FAQ_DIR / clean_fn
                            if not dest.exists() and root_legacy.exists() and root_legacy.is_file():
                                try:
                                    shutil.copy2(root_legacy, dest)
                                    sucessos_img += 1
                                    continue
                                except Exception:
                                    pass

                            if not dest.exists():
                                try:
                                    resp = requests.get(u, cookies=cookies, headers=headers, timeout=12)
                                    if resp.status_code == 200 and len(resp.content) > 100:
                                        with open(dest, "wb") as f_out:
                                            f_out.write(resp.content)
                                        sucessos_img += 1
                                except Exception:
                                    pass
                            else:
                                sucessos_img += 1

                        # 2. Sincronização de Vídeos
                        found_vid_urls = re.findall(r'data-videourl="([^"]+)"', content)
                        for vu in set(found_vid_urls):
                            if vu.startswith("/"):
                                vu = f"https://ministeriopublicoms.sharepoint.com{vu}"
                            clean_vfn = unquote(vu.split("/")[-1].split("?")[0])
                            clean_vfn = re.sub(r'[\\/*?:"<>|]', '_', clean_vfn)
                            if not clean_vfn:
                                clean_vfn = f"vid_{hashlib.md5(vu.encode('utf-8')).hexdigest()[:12]}.mp4"
                            dest_v = target_vid_folder / clean_vfn

                            if not dest_v.exists():
                                try:
                                    resp_v = requests.get(vu, cookies=cookies, headers=headers, stream=True, timeout=35)
                                    if resp_v.status_code == 200:
                                        with open(dest_v, "wb") as f_out_v:
                                            for chunk in resp_v.iter_content(chunk_size=65536):
                                                if chunk:
                                                    f_out_v.write(chunk)
                                        if dest_v.exists() and dest_v.stat().st_size > 1000:
                                            sucessos_vid += 1
                                except Exception:
                                    pass
                            else:
                                sucessos_vid += 1

                    st.sidebar.success(f"🎉 Sincronização concluída: {sucessos_img} imagem(ns) e {sucessos_vid} vídeo(s) organizados por pasta!")
                    time.sleep(1)
                    st.rerun()

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
                    margin: 16px auto;
                    max-width: 100%;
                    height: auto;
                    border-radius: 8px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                    border: 1px solid #343541;
                }
                .faq-container .sp-img-card {
                    margin: 20px auto;
                    max-width: 100%;
                    text-align: center;
                }
                .faq-container .sp-caption {
                    margin-top: 6px;
                    color: #94a3b8;
                    font-size: 0.88rem;
                    font-style: italic;
                }
                .faq-container .sp-fallback-msg {
                    border-radius: 8px;
                    background: rgba(255, 75, 75, 0.08);
                    border: 1px dashed rgba(255, 75, 75, 0.4);
                    padding: 12px;
                    margin: 12px auto;
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
                    parsed_html = parse_sharepoint_content(faq_item['conteudo'], faq_slug=faq_item.get('titulo', 'Geral'))
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
                
                # 1. Se já existe o arquivo local em cache/upload/OneDrive
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

                # 2. Se for link em nuvem do SharePoint e ainda não tiver arquivo local
                if not video_rendered:
                    st.info("💡 Este vídeo está na nuvem corporativa do SharePoint Online. Clique no botão abaixo para autenticar, baixar para o sistema e reproduzir diretamente no player:")
                    
                    target_sp_video = video_item.get('url') or VIDEO_FAQ_URL
                    
                    c_dl1, c_dl2 = st.columns([2, 1])
                    with c_dl1:
                        if st.button("📥 Baixar & Assistir no Sistema", type="primary", key=f"btn_dl_play_{hash(video_item.get('titulo'))}", width='stretch', help="Faz o download autenticado do vídeo para o servidor e reproduz nativamente no sistema."):
                            with st.spinner("Autenticando no SharePoint corporativo e baixando o vídeo para o player..."):
                                try:
                                    dest_dir = VIDEO_FAQ_DIR / video_item.get('categoria', 'Geral')
                                    dest_dir.mkdir(parents=True, exist_ok=True)
                                    dest_file = dest_dir / video_item.get('nome_arquivo', f"{video_item.get('titulo')}.mp4")
                                    
                                    # 1. Tentativa via HTTP direto
                                    headers = {
                                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                                    }
                                    dl_url = target_sp_video + ("&download=1" if "?e=" in target_sp_video else "?download=1")
                                    resp = requests.get(dl_url, headers=headers, timeout=20, allow_redirects=True)
                                    
                                    download_success = False
                                    if resp.status_code == 200 and len(resp.content) > 100000 and not resp.content.startswith(b'<!DOCTYPE') and not resp.content.startswith(b'<html'):
                                        with open(dest_file, "wb") as f_out:
                                            f_out.write(resp.content)
                                        download_success = True
                                    
                                    # 2. Se necessitar de autenticação institucional, usa Selenium com as credenciais corporativas
                                    if not download_success:
                                        driver = get_chrome_driver(headless=HEADLESS)
                                        try:
                                            driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
                                            driver.execute("send_command", {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': str(dest_dir)}})
                                            driver.get(target_sp_video)
                                            time.sleep(3)

                                            from selenium.webdriver.common.by import By
                                            if "login" in driver.current_url.lower():
                                                user_in = driver.find_elements(By.XPATH, "//input[@type='email' or @name='loginfmt' or @type='text']")
                                                if user_in and CITSMART_EMAIL:
                                                    user_in[0].clear()
                                                    user_in[0].send_keys(CITSMART_EMAIL)
                                                    sub_btn = driver.find_elements(By.XPATH, "//input[@type='submit'] | //button[@type='submit']")
                                                    if sub_btn:
                                                        sub_btn[0].click()
                                                        time.sleep(3)
                                                if PASSWORD:
                                                    pass_in = driver.find_elements(By.XPATH, "//input[@type='password']")
                                                    if pass_in:
                                                        pass_in[0].clear()
                                                        pass_in[0].send_keys(PASSWORD)
                                                        sub_btn = driver.find_elements(By.XPATH, "//input[@type='submit'] | //button[@type='submit']")
                                                        if sub_btn:
                                                            sub_btn[0].click()
                                                            time.sleep(4)
                                                stay_in = driver.find_elements(By.XPATH, "//input[@id='idSIButton9'] | //input[@value='Sim' or @value='Yes']")
                                                if stay_in:
                                                    stay_in[0].click()
                                                    time.sleep(4)

                                            # Extrai cookies de sessão autenticados
                                            s_auth = requests.Session()
                                            s_auth.headers.update(headers)
                                            for ck in driver.get_cookies():
                                                s_auth.cookies.set(name=ck['name'], value=ck['value'], domain=ck.get('domain'))
                                            
                                            resp_auth = s_auth.get(dl_url, timeout=60, stream=True)
                                            if resp_auth.status_code == 200:
                                                with open(dest_file, "wb") as f_out:
                                                    for chunk in resp_auth.iter_content(chunk_size=65536):
                                                        if chunk:
                                                            f_out.write(chunk)
                                                if dest_file.exists() and dest_file.stat().st_size > 100000:
                                                    download_success = True
                                        finally:
                                            try:
                                                driver.quit()
                                            except Exception:
                                                pass

                                    if download_success and dest_file.exists():
                                        video_item['caminho'] = dest_file
                                        st.success(f"🎉 Vídeo baixado com sucesso ({format_file_size(dest_file.stat().st_size)})! Iniciando player...")
                                        time.sleep(1)
                                        st.rerun()
                                    else:
                                        st.warning("Não foi possível transferir o arquivo direto. Utilize a opção de assistir no SharePoint Stream:")
                                        st.link_button("🌐 Assistir no SharePoint Stream ↗", target_sp_video, width='stretch')
                                except Exception as e_dl:
                                    st.error(f"Erro durante autenticação/download: {e_dl}")
                    with c_dl2:
                        st.link_button("🌐 Assistir no Stream ↗", target_sp_video, width='stretch')

            st.markdown("---")
            c_info, c_act1, c_act2 = st.columns([1.8, 1.1, 1.1])
            with c_info:
                origem_label = "💾 OneDrive Local" if video_item.get('origem') == "OneDrive Local" else ("📂 Servidor Local" if video_item.get('caminho') else "🌐 SharePoint Online (Nuvem)")
                st.caption(f"🏷️ **Origem:** `{origem_label}`")
                if video_item.get('caminho'):
                    st.caption(f"📁 **Arquivo Local:** `{video_item['caminho']}`")
                elif video_item.get('url'):
                    st.caption(f"🔗 **URL SharePoint:** `{video_item['url']}`")
            with c_act1:
                target_media = str(video_item.get('caminho') or video_item.get('url') or VIDEO_FAQ_URL)
                if st.button("🎬 Abrir no VLC", key=f"btn_vlc_modal_direct_{hash(video_item.get('titulo'))}", width='stretch', help="Executa o player VLC local no seu Windows."):
                    opened = open_in_vlc_player(target_media)
                    if opened:
                        st.toast("Vídeo enviado para o VLC Player com áudio e vídeo!", icon="🎬")
                    else:
                        st.warning("VLC Player não localizado. Abra diretamente pelo SharePoint ou baixe o vídeo.")
            with c_act2:
                target_sp = video_item.get('url') or VIDEO_FAQ_URL
                if target_sp:
                    st.link_button("🌐 SharePoint ↗", target_sp, width='stretch')

        # Constrói lista de vídeos priorizando o catálogo oficial do SharePoint com vínculo local
        all_videos_display = []
        
        # Mapeia arquivos locais por título/nome para associar rapidamente
        local_vids_map = {v['nome_arquivo']: v for v in videos_list}
        local_vids_by_title = {v['titulo']: v for v in videos_list}

        if catalog_videos:
            for item in catalog_videos:
                fname = item.get("nome_arquivo", "")
                ftitle = item.get("titulo", "")
                matched_local = local_vids_map.get(fname) or local_vids_by_title.get(ftitle)
                
                caminho_local = matched_local['caminho'] if matched_local else None
                tamanho_fmt = matched_local['tamanho'] if matched_local else format_file_size(item.get("tamanho_bytes", 0))
                origem_item = matched_local.get('origem', 'OneDrive Local') if matched_local else "SharePoint Online"

                all_videos_display.append({
                    "titulo": ftitle,
                    "nome_arquivo": fname,
                    "categoria": item.get("categoria", "Geral"),
                    "caminho": caminho_local,
                    "url": item.get("url", VIDEO_FAQ_URL),
                    "tamanho": tamanho_fmt,
                    "extensao": Path(fname).suffix.lower() if fname else ".mp4",
                    "origem": origem_item
                })
        elif videos_list:
            all_videos_display.extend(videos_list)
        else:
            if not df_faqs.empty:
                for _, f_row in df_faqs.iterrows():
                    all_videos_display.append({
                        "titulo": f_row['titulo'],
                        "nome_arquivo": f"{f_row['titulo']}.mp4",
                        "categoria": f_row['tipo_faq'],
                        "caminho": None,
                        "url": f_row['url'],
                        "tamanho": "Nuvem",
                        "extensao": ".mp4",
                        "origem": "SharePoint Online"
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

                        c_btn_v1, c_btn_v2 = st.columns(2)
                        with c_btn_v1:
                            if st.button("🎥 Assistir", key=f"btn_vid_card_{idx}_{hash(vid['titulo'])}", width='stretch', help="Abrir modal de reprodução"):
                                open_video_modal(vid)
                        with c_btn_v2:
                            sp_url = vid.get('url') or VIDEO_FAQ_URL
                            st.link_button("🌐 SharePoint ↗", url=sp_url, width='stretch', help="Abrir vídeo direto no SharePoint Online / Stream")

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

        # Constrói pastas de imagens espelhando cada página do SharePoint com sua subpasta
        folders_dict = {}
        local_by_folder = {}
        for img in imagens_list:
            cat = img['categoria']
            local_by_folder.setdefault(cat, []).append(img)

        folders_list = []
        if not df_faqs.empty:
            for _, f_row in df_faqs.iterrows():
                f_titulo = f_row['titulo']
                f_slug = slugify_faq_title(f_titulo)
                imgs = local_by_folder.get(f_slug, [])
                folders_list.append({
                    "titulo": f_titulo,
                    "categoria": f_row['tipo_faq'],
                    "slug": f_slug,
                    "imagens": sorted(imgs, key=lambda x: x["titulo"]),
                    "total": len(imgs),
                    "url": f_row['url']
                })
        else:
            for cat, imgs in local_by_folder.items():
                folders_list.append({
                    "titulo": cat.replace("_", " "),
                    "categoria": "Geral",
                    "slug": cat,
                    "imagens": sorted(imgs, key=lambda x: x["titulo"]),
                    "total": len(imgs),
                    "url": IMAGE_FAQ_URL
                })

        folders_list.sort(key=lambda x: (x.get("categoria", "Geral"), x.get("titulo", "")))

        categorias_img = ["Todas"] + sorted(list(set(f.get("categoria", "Geral") for f in folders_list)))
        selected_cat_img = st.sidebar.selectbox("📂 Categoria / Grupo:", categorias_img, key="select_cat_img")
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

