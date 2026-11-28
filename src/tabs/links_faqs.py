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
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import pandas as pd
import streamlit as st
try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

from src.config import (
    setup_logging, DEBUG_DIR_FAQ, VIDEO_FAQ_DIR, IMAGE_FAQ_DIR, MARKDOWN_FAQ_DIR,
    VIDEO_FAQ_URL, IMAGE_FAQ_URL, MARKDOWN_FAQ_URL, CITSMART_EMAIL, PASSWORD,
    HEADLESS, EXPLICIT_WAIT, get_chrome_driver, get_spreadsheet_setting
)
from src.services.faq_markdown import (
    DEFAULT_SP_HOST,
    deduplicate_markdown_faqs,
    encode_sharepoint_url_path,
    extract_markdown_media_references,
    extract_markdown_title,
    extract_sharepoint_folder_rel_path,
    extract_sharepoint_site_url,
    format_category_name,
    is_html_faq_content,
    mirror_markdown_faqs_to_disk,
    persist_markdown_faqs_to_db,
    render_markdown_for_streamlit,
    scan_local_markdown_faqs,
)
from src.terminal import log, CYAN, GREEN, YELLOW, RED, WHITE
from src.components.subtabs import render_subtabs
from src.components.pagination import (
    render_items_per_page_selector,
    paginate_items,
    render_pagination_controls
)
from src.components.status_banner import render_log_expander
from src.syncs.sync_sharepoint_catalog import (
    check_sharepoint_catalog_sync_running,
    read_sharepoint_catalog_last_log_lines
)

logger = setup_logging(DEBUG_DIR_FAQ / "faq.log", "faq")


from urllib.parse import unquote, quote

SHAREPOINT_COOKIES_FILE = Path(__file__).resolve().parent.parent.parent / "uploads" / "faq" / "sharepoint_cookies.json"


def normalize_sharepoint_url(url: str) -> str:
    """Garante que a URL do SharePoint esteja absoluta e com caracteres especiais/espaços codificados."""
    if not url or not isinstance(url, str):
        return ""
    u = url.strip()
    if u.startswith("/"):
        u = f"https://ministeriopublicoms.sharepoint.com{u}"
    # Se já tem protocolo http/https, codifica cada segmento após o domínio preservando barras
    if "://" in u:
        scheme, rest = u.split("://", 1)
        if "/" in rest:
            domain, path = rest.split("/", 1)
            # Separa query string se houver
            query = ""
            if "?" in path:
                path, query = path.split("?", 1)
                query = "?" + query
            path_segments = [quote(unquote(seg)) for seg in path.split("/")]
            return f"{scheme}://{domain}/{'/'.join(path_segments)}{query}"
    return u


def get_sharepoint_cookies() -> dict:
    """
    Obtém cookies autenticados do SharePoint em ordem de prioridade:
    1. Cofre cifrado no banco de dados (SHAREPOINT_SESSION_COOKIES via sharepoint_downloader).
    2. Cache JSON legado em uploads/faq/sharepoint_cookies.json.
    3. Extração legada do perfil Firefox local (quando disponível no host Windows).
    """
    cookies = {}

    # 1. Cookies persistidos no banco de dados (gerenciados por sharepoint_downloader)
    try:
        from src.scrapers.sharepoint_downloader import get_cached_cookies
        db_cookies = get_cached_cookies()
        if db_cookies:
            for c in db_cookies:
                name = c.get("name")
                val = c.get("value")
                domain = (c.get("domain") or "").lower()
                if name and val and (not domain or "sharepoint.com" in domain):
                    cookies[name] = val
            if cookies:
                return cookies
    except Exception as e:
        logger.debug(f"Falha ao ler cookies do banco via sharepoint_downloader: {e}")

    # 2. Fallback: arquivo JSON local legado
    if SHAREPOINT_COOKIES_FILE.exists():
        try:
            with open(SHAREPOINT_COOKIES_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if loaded and isinstance(loaded, dict):
                    return loaded
        except Exception as e:
            logger.warning(f"Erro ao ler cache de cookies: {e}")

    # 3. Fallback legado: perfil do Firefox do usuário se existir no host
    appdata_dir = os.getenv("APPDATA")
    ff_profiles_dir = Path(appdata_dir) / "Mozilla/Firefox/Profiles" if appdata_dir else None
    if ff_profiles_dir and ff_profiles_dir.exists():
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


def _renew_sharepoint_cookies_dict(target_url: str) -> dict:
    """Renova os cookies do SharePoint via Playwright (sharepoint_downloader) e devolve como dict."""
    try:
        from src.scrapers.sharepoint_downloader import renew_cookies_with_playwright
        new_list = renew_cookies_with_playwright(target_url)
        if not new_list:
            return {}
        cookies = {}
        for c in new_list:
            name = c.get("name")
            val = c.get("value")
            domain = (c.get("domain") or "").lower()
            if name and val and (not domain or "sharepoint.com" in domain):
                cookies[name] = val
        return cookies
    except Exception as e:
        logger.warning(f"Falha ao renovar cookies SharePoint via Playwright: {e}")
        return {}


@st.cache_data(ttl=600, max_entries=500, show_spinner=False)
def _read_image_b64_cached(file_path_str: str, mtime_ns: int) -> str:
    """Lê e codifica uma imagem em base64 com cache em memória invalidado pelo mtime do arquivo."""
    fp = Path(file_path_str)
    with open(fp, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("utf-8")
    ext = fp.suffix.lower()
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


def get_image_as_base64(file_path: Path) -> str:
    """Codifica a imagem local para formato Data URI base64 seguro (com cache em memória)."""
    try:
        if not file_path.exists():
            return ""
        stat = file_path.stat()
        return _read_image_b64_cached(str(file_path.resolve()), stat.st_mtime_ns)
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

def _persist_markdown_faqs_to_db(cursor, conn, md_items: list) -> int:
    """Alias retrocompatível para `persist_markdown_faqs_to_db`."""
    return persist_markdown_faqs_to_db(cursor, conn, md_items)


def _fetch_markdown_faqs_from_sharepoint(
    site_url: str,
    md_base_rel_path: str,
    cookies: dict,
    headers: dict,
    max_workers: int = 8,
) -> tuple:
    """
    Varre recursivamente a pasta `Manuais Markdown` no SharePoint via API REST com `requests.Session`
    e baixa o conteúdo dos arquivos `.md` em paralelo (ignorando README.md).
    Retorna `(items, auth_failed)`.
    """
    discovered = []
    auth_failed = False
    sp_host = "/".join(site_url.split("/")[:3]) if "://" in site_url else DEFAULT_SP_HOST
    base_norm = md_base_rel_path.rstrip("/")

    session = requests.Session()
    session.cookies.update(cookies)

    def _quote_odata_path(path_str: str) -> str:
        return path_str.replace("'", "''")

    # 1. Descobre todos os arquivos .md nas subpastas
    def _discover_md_folder(folder_rel_url: str, top_folder: str = ""):
        nonlocal auth_failed
        if auth_failed:
            return

        escaped_folder = _quote_odata_path(folder_rel_url)
        r_files = session.get(
            f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{escaped_folder}')/Files",
            headers=headers,
            timeout=20,
        )
        if r_files.status_code in (401, 403):
            auth_failed = True
            return

        if r_files.status_code == 200:
            for f_obj in r_files.json().get("d", {}).get("results", []):
                fn = f_obj.get("Name", "")
                if not fn.lower().endswith(".md") or fn.lower() == "readme.md":
                    continue
                f_rel = f_obj.get("ServerRelativeUrl", "")
                if f_rel:
                    discovered.append((fn, f_rel, top_folder or "geral"))

        r_fold = session.get(
            f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{escaped_folder}')/Folders",
            headers=headers,
            timeout=20,
        )
        if r_fold.status_code in (401, 403):
            auth_failed = True
            return

        if r_fold.status_code == 200:
            for subf in r_fold.json().get("d", {}).get("results", []):
                s_name = subf.get("Name")
                if not s_name or s_name in ("Forms",):
                    continue
                sub_rel = subf.get("ServerRelativeUrl")
                if not sub_rel:
                    continue
                _discover_md_folder(sub_rel, top_folder or s_name)

    _discover_md_folder(base_norm, "")
    if auth_failed or not discovered:
        return [], auth_failed

    # 2. Baixa os arquivos .md em paralelo reutilizando a sessão HTTP (keep-alive)
    dl_headers = {"User-Agent": headers.get("User-Agent", "Mozilla/5.0")}

    def _download_single_md(entry: tuple) -> dict:
        fn, f_rel, folder_key = entry
        escaped_file = _quote_odata_path(f_rel)
        raw_content = ""
        try:
            r_val = session.get(
                f"{site_url}/_api/web/GetFileByServerRelativeUrl('{escaped_file}')/$value",
                headers=dl_headers,
                timeout=20,
            )
            if r_val.status_code == 200:
                r_val.encoding = "utf-8"
                raw_content = r_val.text.lstrip("\ufeff")
            else:
                full_dl = encode_sharepoint_url_path(f_rel, sp_host=sp_host)
                r_fb = session.get(full_dl, headers=dl_headers, timeout=20)
                if r_fb.status_code == 200 and "text/html" not in (r_fb.headers.get("content-type") or "").lower():
                    r_fb.encoding = "utf-8"
                    raw_content = r_fb.text.lstrip("\ufeff")
        except Exception as exc:
            logger.warning(f"Falha ao baixar manual '{fn}': {exc}")

        full_url = encode_sharepoint_url_path(f_rel, sp_host=sp_host)
        return {
            "filename": fn,
            "folder": folder_key,
            "titulo": extract_markdown_title(raw_content, fn),
            "tipo_faq": format_category_name(folder_key),
            "url": full_url,
            "server_rel_url": f_rel,
            "conteudo": raw_content,
        }

    collected = []
    workers = max(1, min(max_workers, len(discovered)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_download_single_md, entry) for entry in discovered]
        for fut in as_completed(futures):
            try:
                collected.append(fut.result())
            except Exception as exc:
                logger.warning(f"Erro em worker de download Markdown: {exc}")

    return collected, auth_failed


def sync_sharepoint_catalog_via_api() -> dict:
    """
    Sincroniza o catálogo completo diretamente da API REST do SharePoint Online:
    1. 'Manuais Markdown' (arquivos .md deduplicados por categoria) -> tabela faqs (substitui Site Pages)
    2. 'Vídeos FAQ' (vídeos mp4/mov por pasta) -> tabela faq_videos
    3. 'Imagens FAQ' (imagens jpg/png recursivas) -> tabela faq_imagens
    Renova cookies automaticamente via Playwright se expirados e suporta fallback local.
    """
    md_setting = get_spreadsheet_setting("MARKDOWN_FAQ_PATH", "") or MARKDOWN_FAQ_URL
    site_url = extract_sharepoint_site_url(md_setting)
    md_base_rel = extract_sharepoint_folder_rel_path(md_setting)
    sp_host = "/".join(site_url.split("/")[:3]) if "://" in site_url else DEFAULT_SP_HOST

    cookies = get_sharepoint_cookies()
    if not cookies:
        logger.info("🔄 Cookies SharePoint ausentes. Tentando autenticação automática via Playwright...")
        cookies = _renew_sharepoint_cookies_dict(md_setting or site_url)

    headers = {
        "Accept": "application/json;odata=verbose",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0"
    }

    from src.database.connection import get_connection
    conn = get_connection()
    cursor = conn.cursor()

    stats = {"faqs": 0, "videos": 0, "imagens": 0}

    try:
        # -------------------------------------------------------------
        # 1. Sincroniza MANUAIS MARKDOWN (.md) — substitui Site Pages
        # -------------------------------------------------------------
        md_raw_items = []
        if cookies:
            md_raw_items, auth_failed = _fetch_markdown_faqs_from_sharepoint(
                site_url=site_url,
                md_base_rel_path=md_base_rel,
                cookies=cookies,
                headers=headers,
            )
            if auth_failed:
                logger.info("🔄 Sessão SharePoint expirada (401/403). Renovando cookies via Playwright...")
                renewed = _renew_sharepoint_cookies_dict(md_setting or site_url)
                if renewed:
                    cookies = renewed
                    md_raw_items, _ = _fetch_markdown_faqs_from_sharepoint(
                        site_url=site_url,
                        md_base_rel_path=md_base_rel,
                        cookies=cookies,
                        headers=headers,
                    )

        if md_raw_items:
            deduped_md = deduplicate_markdown_faqs(md_raw_items)
            mirrored = mirror_markdown_faqs_to_disk(MARKDOWN_FAQ_DIR, deduped_md)
            stats["faqs"] = _persist_markdown_faqs_to_db(cursor, conn, deduped_md)
            logger.info(
                f"📘 Manuais Markdown sincronizados do SharePoint: {stats['faqs']} únicos "
                f"(de {len(md_raw_items)} arquivos encontrados; {mirrored} espelhados em '{MARKDOWN_FAQ_DIR}')."
            )
        else:
            # Fallback local (pasta cache uploads/faq/manuais_markdown ou OneDrive local)
            local_md_items = scan_local_markdown_faqs(
                MARKDOWN_FAQ_DIR, sp_base_rel_path=md_base_rel, sp_host=sp_host
            )
            if local_md_items:
                stats["faqs"] = _persist_markdown_faqs_to_db(cursor, conn, local_md_items)
                logger.info(
                    f"📂 Manuais Markdown carregados da pasta local '{MARKDOWN_FAQ_DIR}': {stats['faqs']} únicos."
                )
            elif not cookies:
                return {
                    "success": False,
                    "error": "Cookies de autenticação do SharePoint não encontrados e nenhum manual Markdown local disponível.",
                }

        if not cookies:
            return {"success": True, "stats": stats}

        # -------------------------------------------------------------
        # 2. Sincroniza VÍDEOS FAQ (Pastas e Arquivos)
        # -------------------------------------------------------------
        v_base = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('/sites/dit-manutencao/Shared Documents/Tutoriais-FAQs/Vídeos FAQ')"
        resp_vf = requests.get(f"{v_base}/Folders", cookies=cookies, headers=headers, timeout=15)
        if resp_vf.status_code == 200:
            v_folders = resp_vf.json().get('d', {}).get('results', [])
            for vf in v_folders:
                v_cat = vf.get('Name')
                if not v_cat or v_cat in ['Forms']:
                    continue
                vf_rel = vf.get('ServerRelativeUrl')
                resp_vfiles = requests.get(
                    f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{vf_rel}')/Files",
                    cookies=cookies, headers=headers, timeout=15
                )
                if resp_vfiles.status_code == 200:
                    v_files = resp_vfiles.json().get('d', {}).get('results', [])
                    for file_obj in v_files:
                        fn = file_obj.get('Name', '')
                        if not any(fn.lower().endswith(ext) for ext in ['.mp4', '.mov', '.webm', '.avi']):
                            continue
                        f_rel = file_obj.get('ServerRelativeUrl', '')
                        full_v_url = f"{sp_host}{f_rel}"
                        f_len = int(file_obj.get('Length', 0) or 0)
                        f_title = Path(fn).stem

                        cursor.execute("""
                            INSERT INTO faq_videos (titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes, data_atualizacao)
                            VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                            ON CONFLICT(url) DO UPDATE SET
                                titulo = excluded.titulo,
                                categoria = excluded.categoria,
                                nome_arquivo = excluded.nome_arquivo,
                                tamanho_bytes = excluded.tamanho_bytes,
                                data_atualizacao = CURRENT_TIMESTAMP
                        """, (f_title, v_cat, fn, full_v_url, f_rel, f_len))
                        stats["videos"] += 1
            conn.commit()

        # -------------------------------------------------------------
        # 3. Sincroniza IMAGENS FAQ (Pastas e Subpastas Recursivas)
        # -------------------------------------------------------------
        def harvest_images_folder(folder_rel_url: str, current_category: str = ""):
            r_sub = requests.get(
                f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{folder_rel_url}')/Files",
                cookies=cookies, headers=headers, timeout=15
            )
            if r_sub.status_code == 200:
                for ifile in r_sub.json().get('d', {}).get('results', []):
                    ifn = ifile.get('Name', '')
                    if not any(ifn.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp']):
                        continue
                    i_rel = ifile.get('ServerRelativeUrl', '')
                    full_i_url = f"{sp_host}{i_rel}"
                    i_len = int(ifile.get('Length', 0) or 0)
                    i_title = Path(ifn).stem

                    cursor.execute("""
                        INSERT INTO faq_imagens (titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes, data_atualizacao)
                        VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(url) DO UPDATE SET
                            titulo = excluded.titulo,
                            categoria = excluded.categoria,
                            nome_arquivo = excluded.nome_arquivo,
                            tamanho_bytes = excluded.tamanho_bytes,
                            data_atualizacao = CURRENT_TIMESTAMP
                    """, (i_title, current_category or "Geral", ifn, full_i_url, i_rel, i_len))
                    stats["imagens"] += 1

            # Subpastas
            r_fold = requests.get(
                f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{folder_rel_url}')/Folders",
                cookies=cookies, headers=headers, timeout=15
            )
            if r_fold.status_code == 200:
                for subf in r_fold.json().get('d', {}).get('results', []):
                    s_name = subf.get('Name')
                    if not s_name or s_name in ['Forms']:
                        continue
                    sub_rel = subf.get('ServerRelativeUrl')
                    sub_cat = f"{current_category} > {s_name}" if current_category else s_name
                    harvest_images_folder(sub_rel, sub_cat)

        harvest_images_folder("/sites/dit-manutencao/Shared Documents/Tutoriais-FAQs/Imagens FAQ", "")
        conn.commit()

        logger.info(f"✨ Sincronização SharePoint API concluída com sucesso: {stats}")
        return {"success": True, "stats": stats}

    except Exception as e:
        logger.error(f"Erro na sincronização SharePoint API: {e}")
        return {"success": False, "error": str(e)}
    finally:
        conn.close()


def ensure_sharepoint_image_cached(img_url: str, faq_slug: str = "Geral") -> str:
    """
    Verifica se a imagem do SharePoint já está salva localmente em uploads/faq/imagens/<faq_slug>/.
    Caso exista em uploads/faq/imagens/ raiz (legado), migra ou carrega dela.
    Se não estiver em disco, baixa usando os cookies autenticados da intranet.
    Retorna o Data URI em base64 da imagem para exibição direta no navegador.
    """
    if not img_url:
        return ""

    img_url = normalize_sharepoint_url(img_url)

    clean_filename = unquote(img_url.split("/")[-1].split("?")[0])
    clean_filename = re.sub(r'[\\/*?:"<>|]', '_', clean_filename)
    if not clean_filename:
        clean_filename = f"img_{hashlib.md5(img_url.encode('utf-8')).hexdigest()[:12]}.jpg"

    slug = slugify_faq_title(faq_slug)
    folder_dir = IMAGE_FAQ_DIR / slug
    dest_file = folder_dir / clean_filename
    root_legacy_file = IMAGE_FAQ_DIR / clean_filename

    # Se já existir na pasta organizada por slug
    if dest_file.exists() and dest_file.stat().st_size > 0:
        return get_image_as_base64(dest_file)

    # Verifica caminho relativo espelhado dentro de IMAGE_FAQ_DIR (ex.: Imagens FAQ/Formatação/SCCM 01.png)
    decoded_url = unquote(img_url.split("?")[0])
    if "Imagens FAQ/" in decoded_url:
        rel_after_imgs = decoded_url.split("Imagens FAQ/", 1)[1].lstrip("/")
        if rel_after_imgs and ".." not in Path(rel_after_imgs).parts:
            mirrored_file = IMAGE_FAQ_DIR / rel_after_imgs
            if mirrored_file.exists() and mirrored_file.is_file() and mirrored_file.stat().st_size > 0:
                return get_image_as_base64(mirrored_file)

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
            c_type = resp.headers.get("content-type", "").lower()
            is_real_img = "image" in c_type or resp.content.startswith(b'\xff\xd8') or resp.content.startswith(b'\x89PNG') or resp.content.startswith(b'GIF') or resp.content.startswith(b'RIFF')
            if resp.status_code == 200 and len(resp.content) > 100 and is_real_img:
                folder_dir.mkdir(parents=True, exist_ok=True)
                with open(dest_file, "wb") as f_out:
                    f_out.write(resp.content)
                logger.info(f"✅ Imagem FAQ baixada em '{slug}/{clean_filename}'")
                return get_image_as_base64(dest_file)
            else:
                logger.warning(f"Download de imagem retornou status {resp.status_code} ({c_type}) para {clean_filename}")
        except Exception as e:
            logger.error(f"Erro ao baixar imagem {img_url}: {e}")

    return ""


def prefetch_markdown_images_parallel(
    md_text: str,
    faq_slug: str = "Geral",
    md_server_rel_url: str = "",
    max_workers: int = 6,
) -> int:
    """
    Pré-baixa em paralelo (com `requests.Session` compartilhada) apenas as imagens
    referenciadas no Markdown que ainda não existem no cache local em `uploads/faq/imagens/`.
    Reduz drasticamente o tempo de abertura de tutoriais com muitas capturas de tela.
    """
    img_urls, _ = extract_markdown_media_references(md_text, md_server_rel_url=md_server_rel_url)
    if not img_urls:
        return 0

    slug = slugify_faq_title(faq_slug)
    folder_dir = IMAGE_FAQ_DIR / slug
    missing = []

    for raw_u in img_urls:
        u = normalize_sharepoint_url(raw_u)
        if not u:
            continue
        clean_fn = unquote(u.split("/")[-1].split("?")[0])
        clean_fn = re.sub(r'[\\/*?:"<>|]', '_', clean_fn)
        if not clean_fn:
            clean_fn = f"img_{hashlib.md5(u.encode('utf-8')).hexdigest()[:12]}.jpg"

        dest = folder_dir / clean_fn
        if dest.exists() and dest.stat().st_size > 0:
            continue

        # Verifica caminho espelhado em IMAGE_FAQ_DIR
        decoded_url = unquote(u.split("?")[0])
        if "Imagens FAQ/" in decoded_url:
            rel_after = decoded_url.split("Imagens FAQ/", 1)[1].lstrip("/")
            if rel_after and ".." not in Path(rel_after).parts:
                mirrored = IMAGE_FAQ_DIR / rel_after
                if mirrored.exists() and mirrored.is_file() and mirrored.stat().st_size > 0:
                    continue

        missing.append((u, dest))

    if not missing:
        return 0

    cookies = get_sharepoint_cookies()
    if not cookies:
        return 0

    folder_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.cookies.update(cookies)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",
        "Accept": "image/avif,image/webp,image/png,image/svg+xml,image/*;q=0.8,*/*;q=0.5",
    }

    def _dl_one(pair: tuple) -> bool:
        u, dest_file = pair
        try:
            resp = session.get(u, headers=headers, timeout=12)
            c_type = (resp.headers.get("content-type") or "").lower()
            is_real = (
                "image" in c_type
                or resp.content.startswith(b"\xff\xd8")
                or resp.content.startswith(b"\x89PNG")
                or resp.content.startswith(b"GIF")
                or resp.content.startswith(b"RIFF")
            )
            if resp.status_code == 200 and len(resp.content) > 100 and is_real:
                with open(dest_file, "wb") as f_out:
                    f_out.write(resp.content)
                return True
        except Exception:
            pass
        return False

    downloaded = 0
    workers = max(1, min(max_workers, len(missing)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_dl_one, item) for item in missing]
        for fut in as_completed(futures):
            try:
                if fut.result():
                    downloaded += 1
            except Exception:
                pass
    return downloaded


def ensure_sharepoint_video_cached(video_url: str, faq_slug: str = "Geral") -> str:
    """
    Verifica se o vídeo do SharePoint já está salvo localmente em uploads/faq/videos/<faq_slug>/.
    Se não estiver, baixa usando os cookies autenticados da intranet.
    Retorna o Data URI em base64 (para vídeos leves) ou caminho local.
    """
    if not video_url:
        return ""

    video_url = normalize_sharepoint_url(video_url)

    clean_filename = unquote(video_url.split("/")[-1].split("?")[0])
    clean_filename = re.sub(r'[\\/*?:"<>|]', '_', clean_filename)
    if not clean_filename:
        clean_filename = f"video_{hashlib.md5(video_url.encode('utf-8')).hexdigest()[:12]}.mp4"

    slug = slugify_faq_title(faq_slug)
    folder_dir = VIDEO_FAQ_DIR / slug
    dest_file = folder_dir / clean_filename

    # Verifica se o arquivo já existe no destino
    if dest_file.exists() and dest_file.stat().st_size > 1000:
        return get_video_as_base64(dest_file), dest_file

    # Verifica se já existe em outra subpasta do VIDEO_FAQ_DIR
    if VIDEO_FAQ_DIR.exists():
        for existing in VIDEO_FAQ_DIR.rglob(clean_filename):
            if existing.is_file() and existing.stat().st_size > 1000:
                try:
                    folder_dir.mkdir(parents=True, exist_ok=True)
                    if existing != dest_file:
                        shutil.copy2(existing, dest_file)
                    return get_video_as_base64(dest_file), dest_file
                except Exception:
                    return get_video_as_base64(existing), existing

    # Tenta baixar com cookies para a subpasta do tutorial
    cookies = get_sharepoint_cookies()
    if cookies:
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",
                "Accept": "*/*"
            }
            resp = requests.get(video_url, cookies=cookies, headers=headers, stream=True, timeout=35)
            c_type = resp.headers.get("content-type", "").lower()
            if resp.status_code == 200 and "text/html" not in c_type:
                folder_dir.mkdir(parents=True, exist_ok=True)
                with open(dest_file, "wb") as f_out:
                    for chunk in resp.iter_content(chunk_size=65536):
                        if chunk:
                            f_out.write(chunk)
                if dest_file.exists() and dest_file.stat().st_size > 1000:
                    logger.info(f"✅ Vídeo FAQ baixado em '{slug}/{clean_filename}'")
                    return get_video_as_base64(dest_file), dest_file
            else:
                logger.warning(f"Download de vídeo retornou status {resp.status_code} ({c_type}) para {clean_filename}")
        except Exception as e:
            logger.error(f"Erro ao baixar vídeo {video_url}: {e}")

    return "", None


def get_video_as_base64(file_path: Path) -> str:
    """Codifica o arquivo de vídeo local em base64 Data URI se for leve (até 75MB)."""
    try:
        if not file_path.exists():
            return ""
        # Limite ideal para reprodução no navegador sem travar ou sobrecarregar memória DOM
        if file_path.stat().st_size > 75 * 1024 * 1024:
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
    html_str = re.sub(r'[\w\s]+Published \d{2}/\d{2}/\d{4}', '', html_str, flags=re.DOTALL | re.IGNORECASE)

    if BeautifulSoup is None:
        return html_str

    soup = BeautifulSoup(html_str, "html.parser")

    # 2. Remoção de Ícones Quebrados (tags <i>) e avatares/placeholders irrelevantes
    for tag in soup.find_all(["i", "span"]):
        if tag.name == "i" or "ms-Icon" in tag.get("class", []):
            tag.decompose()

    for img in list(soup.find_all("img")):
        src = img.get("src") or img.get("data-sp-originalimgsrc") or ""
        if "userphoto.jpg" in src or "fabric-cdn-prod" in src or img.get("data-automation-id") == "lowQualityImagePlaceholder":
            # Remove o container do avatar do autor se for imagem pequena
            parent = img.parent
            img.decompose()
            if parent and parent.name in ["div", "span", "figure"] and not parent.get_text(strip=True) and not parent.find_all(["img", "video"]):
                parent.decompose()

    # 3. Processamento de Imagens (<div class="imagePlugin" data-imageurl="...">)
    image_divs = soup.find_all("div", class_="imagePlugin")
    for div in image_divs:
        img_url = div.get("data-imageurl")
        caption = (div.get("data-captiontext") or "").strip()
        if img_url:
            img_url = normalize_sharepoint_url(img_url)

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
        if not src:
            img.decompose()
            continue

        src = normalize_sharepoint_url(src)

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
            cached_src, local_file = ensure_sharepoint_video_cached(v_url, faq_slug=faq_slug)
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
            elif local_file and local_file.exists():
                file_size_fmt = format_file_size(local_file.stat().st_size)
                clean_name = local_file.name
                card_v = f'''
                <div class="sp-img-card" style="margin: 20px auto; max-width: 720px; background: #1e1f29; border: 1px solid #3b3c4a; border-radius: 10px; padding: 18px 22px; text-align: center; box-shadow: 0 4px 14px rgba(0,0,0,0.35);">
                    <div style="font-size: 1.05rem; font-weight: 600; color: #f8f9fa; margin-bottom: 6px;">
                        🎬 {v_title}
                    </div>
                    <div style="font-size: 0.88rem; color: #94a3b8; margin-bottom: 12px;">
                        Vídeo em alta definição disponível localmente ({file_size_fmt}) e na intranet:
                    </div>
                    <div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">
                        <a href="bancada://run?tool=vlc&target={local_file}" style="display: inline-block; background-color: #3b82f6; color: white; text-decoration: none; font-weight: bold; font-size: 0.9rem; padding: 10px 18px; border-radius: 6px; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.4);">
                            ▶️ Abrir no Player VLC / Sistema
                        </a>
                        <a href="{v_url}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; font-size: 0.9rem; padding: 10px 18px; border-radius: 6px; box-shadow: 0 2px 8px rgba(255, 75, 75, 0.4);">
                            🔗 Assistir no SharePoint Stream ↗
                        </a>
                    </div>
                </div>
                '''
                v_embed.replace_with(BeautifulSoup(card_v, "html.parser"))
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
                
                cached_src, local_file = ensure_sharepoint_video_cached(file_url, faq_slug=faq_slug)
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
                elif local_file and local_file.exists():
                    file_size_fmt = format_file_size(local_file.stat().st_size)
                    card_v = f'''
                    <div class="sp-img-card" style="margin: 20px auto; max-width: 720px; background: #1e1f29; border: 1px solid #3b3c4a; border-radius: 10px; padding: 18px 22px; text-align: center; box-shadow: 0 4px 14px rgba(0,0,0,0.35);">
                        <div style="font-size: 1.05rem; font-weight: 600; color: #f8f9fa; margin-bottom: 6px;">
                            🎬 {v_title}
                        </div>
                        <div style="font-size: 0.88rem; color: #94a3b8; margin-bottom: 12px;">
                            Vídeo em alta definição disponível localmente ({file_size_fmt}) e na intranet:
                        </div>
                        <div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">
                            <a href="bancada://run?tool=vlc&target={local_file}" style="display: inline-block; background-color: #3b82f6; color: white; text-decoration: none; font-weight: bold; font-size: 0.9rem; padding: 10px 18px; border-radius: 6px; box-shadow: 0 2px 8px rgba(59, 130, 246, 0.4);">
                                ▶️ Abrir no Player VLC / Sistema
                            </a>
                            <a href="{file_url}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; font-size: 0.9rem; padding: 10px 18px; border-radius: 6px; box-shadow: 0 2px 8px rgba(255, 75, 75, 0.4);">
                                🔗 Assistir no SharePoint Stream ↗
                            </a>
                        </div>
                    </div>
                    '''
                    div.replace_with(BeautifulSoup(card_v, "html.parser"))
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


@st.cache_data(ttl=120, show_spinner=False)
def _load_faqs_cached() -> pd.DataFrame:
    """
    Carrega o catálogo de FAQs priorizando os arquivos `.md` espelhados em disco
    (`uploads/faq/manuais_markdown/<subpasta>/<arquivo>.md`) para abertura instantânea,
    com fallback automático para a tabela `faqs` no banco de dados.
    """
    # 1. Caminho ultrarrápido: leitura direta dos .md cacheados em uploads/faq/manuais_markdown/
    local_md = scan_local_markdown_faqs(MARKDOWN_FAQ_DIR)
    if local_md:
        rows = [
            {
                "id": idx + 1,
                "titulo": item["titulo"],
                "tipo_faq": item["tipo_faq"],
                "url": item["url"],
                "conteudo": item["conteudo"],
            }
            for idx, item in enumerate(local_md)
        ]
        return pd.DataFrame(rows)

    # 2. Fallback: consulta tabela `faqs` no banco de dados
    try:
        from src.database.bootstrap import setup_faq_tables
        setup_faq_tables()
    except Exception:
        pass

    root_dir = Path(__file__).parent.parent.parent
    json_faq_path = root_dir / "temp" / "faqs_template.json"
    try:
        from src.database.connection import get_connection
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM faqs")
        if cursor.fetchone()[0] == 0 and json_faq_path.exists():
            with open(json_faq_path, "r", encoding="utf-8") as f:
                faqs_json = json.load(f)
            for item in faqs_json:
                cursor.execute(
                    """
                    INSERT INTO faqs (titulo, tipo_faq, url, conteudo)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT (url) DO NOTHING
                    """,
                    (item.get("titulo"), item.get("tipo_faq", "Geral"), item.get("url"), item.get("conteudo")),
                )
            conn.commit()
        df = pd.read_sql_query("SELECT id, titulo, tipo_faq, url, conteudo FROM faqs ORDER BY tipo_faq, titulo", conn)
        conn.close()
        return df
    except Exception as e:
        logger.error(f"Erro ao carregar FAQs do banco de dados: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=120, show_spinner=False)
def _load_videos_catalog_cached() -> list:
    """Carrega o catálogo de vídeos do banco (cacheado por 120s)."""
    root_dir = Path(__file__).parent.parent.parent
    json_videos_path = root_dir / "temp" / "videos_faq_template.json"
    try:
        from src.database.connection import get_connection
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM faq_videos")
        if cursor.fetchone()[0] == 0 and json_videos_path.exists():
            with open(json_videos_path, "r", encoding="utf-8") as f:
                vids_json = json.load(f)
            for item in vids_json:
                cursor.execute(
                    """
                    INSERT INTO faq_videos (titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT (url) DO NOTHING
                    """,
                    (item.get("titulo"), item.get("categoria", "Geral"), item.get("nome_arquivo"), item.get("url"), item.get("caminho_relativo"), item.get("tamanho_bytes", 0)),
                )
            conn.commit()
        df_vids = pd.read_sql_query("SELECT id, titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes FROM faq_videos", conn)
        conn.close()
        return df_vids.to_dict("records") if not df_vids.empty else []
    except Exception as e:
        logger.error(f"Erro ao carregar vídeos FAQ: {e}")
        return []


@st.cache_data(ttl=120, show_spinner=False)
def _load_imagens_catalog_cached() -> list:
    """Carrega o catálogo de imagens do banco (cacheado por 120s)."""
    root_dir = Path(__file__).parent.parent.parent
    json_imagens_path = root_dir / "temp" / "imagens_faq_template.json"
    try:
        from src.database.connection import get_connection
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM faq_imagens")
        if cursor.fetchone()[0] == 0 and json_imagens_path.exists():
            with open(json_imagens_path, "r", encoding="utf-8") as f:
                imgs_json = json.load(f)
            for item in imgs_json:
                cursor.execute(
                    """
                    INSERT INTO faq_imagens (titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT (url) DO NOTHING
                    """,
                    (item.get("titulo"), item.get("categoria", "Geral"), item.get("nome_arquivo"), item.get("url"), item.get("caminho_relativo"), item.get("tamanho_bytes", 0)),
                )
            conn.commit()
        df_imgs = pd.read_sql_query("SELECT id, titulo, categoria, nome_arquivo, url, caminho_relativo, tamanho_bytes FROM faq_imagens", conn)
        conn.close()
        return df_imgs.to_dict("records") if not df_imgs.empty else []
    except Exception as e:
        logger.error(f"Erro ao carregar imagens FAQ: {e}")
        return []


@st.cache_data(ttl=120, show_spinner=False)
def _scan_video_faqs_cached() -> list:
    return scan_video_faqs(VIDEO_FAQ_DIR)


@st.cache_data(ttl=120, show_spinner=False)
def _scan_image_faqs_cached() -> list:
    return scan_image_faqs(IMAGE_FAQ_DIR)


@st.cache_data(ttl=300, show_spinner=False)
def _load_links_uteis_cached() -> list:
    json_links_path = Path(__file__).parent.parent.parent / "temp" / "links_uteis_template.json"
    if json_links_path.exists():
        try:
            with open(json_links_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Erro ao ler links_uteis_template.json: {e}")
    return []


def render_faq_page():
    """Renderiza a página de FAQs, Tutoriais do SharePoint, Vídeos FAQ e Links Úteis da Bancada."""
    st.title("📚 FAQ, Tutoriais & Links Úteis da Bancada")
    st.write("Base de conhecimento centralizada com tutoriais da equipe e atalhos rápidos para sistemas externos.")
    st.markdown("---")

    # Controle de processo em segundo plano (Sincronização SharePoint)
    sp_sync_ativo = check_sharepoint_catalog_sync_running()

    if "was_sp_syncing" not in st.session_state:
        st.session_state["was_sp_syncing"] = False

    if st.session_state["was_sp_syncing"] and not sp_sync_ativo:
        st.session_state["was_sp_syncing"] = False
        st.cache_data.clear()
        st.toast("🎉 Sincronização do catálogo SharePoint concluída com sucesso!", icon="✅")
        st.rerun()

    if sp_sync_ativo:
        st.session_state["was_sp_syncing"] = True

    render_log_expander(
        "🤖 Sincronização de Catálogo SharePoint em Segundo Plano",
        sp_sync_ativo,
        read_sharepoint_catalog_last_log_lines,
        check_sharepoint_catalog_sync_running,
        "O worker está consultando a API REST do SharePoint Online (Artigos, Vídeos e Imagens) em segundo plano. O painel permanece totalmente livre para uso!"
    )

    # Navegação superior estilo Abas com suporte a query parameter (?subtab=slug)
    FAQ_SUBTAB_MAP = {
        "sharepoint": "📚 FAQs & Tutoriais (SharePoint)",
        "videos": "🎥 Vídeos FAQ (Tutoriais)",
        "imagens": "🖼️ Imagens FAQ (Galeria)",
        "links": "🔗 Links Úteis da Bancada"
    }

    active_tab = render_subtabs(FAQ_SUBTAB_MAP, default_slug="sharepoint", key="faq_nav_radio")

    st.markdown("<br>", unsafe_allow_html=True)

    # Roteamento dinâmico das Abas (lazy-loading: cada subaba carrega apenas seus próprios dados cacheados)
    if active_tab == "📚 FAQs & Tutoriais (SharePoint)":
        df_faqs = _load_faqs_cached()
        st.sidebar.markdown("## 🔍 Filtros do FAQ")
        search_query = st.sidebar.text_input("Buscar por palavra-chave:", "", key="faq_search")
        
        tipos_disponiveis = ["Todos"] + sorted(df_faqs['tipo_faq'].dropna().unique().tolist()) if not df_faqs.empty else ["Todos"]
        selected_tipo = st.sidebar.selectbox("📂 Categoria:", tipos_disponiveis, key="faq_cat")
        items_per_page_faq = render_items_per_page_selector("faq_sp", options=[6, 10, 20, 50], default_index=1)

        st.sidebar.markdown("---")
        st.sidebar.markdown("## 🌐 SharePoint Online")
        if sp_sync_ativo:
            st.sidebar.button("🤖 Sincronizando Catálogo...", width='stretch', disabled=True)
        else:
            if st.sidebar.button("🔄 Sincronizar Catálogo do SharePoint", width="stretch", help="Consulta as APIs REST do SharePoint Online em segundo plano e atualiza os artigos, vídeos e imagens sem bloquear a página."):
                popen_kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if sys.platform == "win32" else {}
                subprocess.Popen([sys.executable, "src/syncs/sync_sharepoint_catalog.py"], **popen_kwargs)
                time.sleep(0.8)
                st.toast("🚀 Sincronização do catálogo iniciada em segundo plano!", icon="🤖")
                st.rerun()

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
                        
                        # 1. Sincronização de Imagens e Vídeos (suporte nativo a Markdown .md + fallback HTML legado)
                        img_candidates = []
                        vid_candidates = []
                        if not is_html_faq_content(content):
                            md_imgs, md_vids = extract_markdown_media_references(
                                content, md_server_rel_url=r.get("url") or ""
                            )
                            img_candidates.extend(md_imgs)
                            vid_candidates.extend(md_vids)
                        else:
                            soup_sync = BeautifulSoup(content, "html.parser") if BeautifulSoup else None
                            if soup_sync:
                                for div in soup_sync.find_all("div", class_="imagePlugin"):
                                    u = div.get("data-imageurl")
                                    if u:
                                        img_candidates.append(u)
                                for img in soup_sync.find_all("img"):
                                    src = img.get("src") or img.get("data-sp-originalimgsrc")
                                    if src and "userphoto.jpg" not in src and "fabric-cdn-prod" not in src and img.get("data-automation-id") != "lowQualityImagePlaceholder":
                                        img_candidates.append(src)
                                for v_emb in soup_sync.find_all("div", class_="sp-video-embed"):
                                    vu = v_emb.get("data-videourl")
                                    if vu:
                                        vid_candidates.append(vu)
                                for div in soup_sync.find_all(lambda t: t.name == "div" and any(k.endswith("controldata") for k in t.attrs)):
                                    for k, val in div.attrs.items():
                                        if k.endswith("controldata"):
                                            try:
                                                c_obj = json.loads(val)
                                                props = c_obj.get("properties", {})
                                                if isinstance(props, dict):
                                                    furl = props.get("file") or props.get("serverRelativeUrl") or props.get("url")
                                                    if furl and any(str(furl).lower().endswith(ext) for ext in [".mp4", ".mov", ".webm", ".avi"]):
                                                        vid_candidates.append(furl)
                                            except Exception:
                                                pass
                            else:
                                img_candidates = re.findall(r'data-imageurl="([^"]+)"', content) + re.findall(r'src="([^"]+sharepoint\.com[^"]+)"', content)
                                vid_candidates = re.findall(r'data-videourl="([^"]+)"', content)

                        for u in set(img_candidates):
                            u = normalize_sharepoint_url(u)
                            if not u:
                                continue
                            clean_fn = unquote(u.split("/")[-1].split("?")[0])
                            clean_fn = re.sub(r'[\\/*?:"<>|]', '_', clean_fn)
                            if not clean_fn:
                                clean_fn = f"img_{hashlib.md5(u.encode('utf-8')).hexdigest()[:12]}.jpg"
                            dest = target_img_folder / clean_fn

                            root_legacy = IMAGE_FAQ_DIR / clean_fn
                            if not dest.exists() and root_legacy.exists() and root_legacy.is_file() and root_legacy.stat().st_size > 0:
                                try:
                                    shutil.copy2(root_legacy, dest)
                                    sucessos_img += 1
                                    continue
                                except Exception:
                                    pass

                            if not dest.exists() or dest.stat().st_size == 0:
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

                        for vu in set(vid_candidates):
                            vu = normalize_sharepoint_url(vu)
                            if not vu:
                                continue
                            clean_vfn = unquote(vu.split("/")[-1].split("?")[0])
                            clean_vfn = re.sub(r'[\\/*?:"<>|]', '_', clean_vfn)
                            if not clean_vfn:
                                clean_vfn = f"vid_{hashlib.md5(vu.encode('utf-8')).hexdigest()[:12]}.mp4"
                            dest_v = target_vid_folder / clean_vfn

                            # Verifica se já existe em qualquer subpasta de vídeo
                            if not dest_v.exists() and VIDEO_FAQ_DIR.exists():
                                for existing in VIDEO_FAQ_DIR.rglob(clean_vfn):
                                    if existing.is_file() and existing.stat().st_size > 1000:
                                        try:
                                            if existing != dest_v:
                                                shutil.copy2(existing, dest_v)
                                            break
                                        except Exception:
                                            pass

                            if not dest_v.exists() or dest_v.stat().st_size < 1000:
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
                .faq-container img, .faq-container .sp-image,
                div[data-testid="stDialog"] .sp-image {
                    display: block;
                    margin: 16px auto;
                    max-width: 100%;
                    height: auto;
                    border-radius: 8px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                    border: 1px solid #343541;
                }
                .faq-container .sp-img-card,
                div[data-testid="stDialog"] .sp-img-card {
                    margin: 20px auto;
                    max-width: 100%;
                    text-align: center;
                }
                .faq-container .sp-caption,
                div[data-testid="stDialog"] .sp-caption {
                    margin-top: 6px;
                    color: #94a3b8;
                    font-size: 0.88rem;
                    font-style: italic;
                }
                .faq-container .sp-fallback-msg,
                div[data-testid="stDialog"] .sp-fallback-msg {
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
                    raw_content = str(faq_item['conteudo'])
                    if is_html_faq_content(raw_content):
                        parsed_html = parse_sharepoint_content(raw_content, faq_slug=faq_item.get('titulo', 'Geral'))
                        st.markdown(f'<div class="faq-container">{parsed_html}</div>', unsafe_allow_html=True)
                    else:
                        prefetch_markdown_images_parallel(
                            raw_content,
                            faq_slug=faq_item.get('titulo', 'Geral'),
                            md_server_rel_url=faq_item.get('url', ''),
                        )
                        rendered_md = render_markdown_for_streamlit(
                            raw_content,
                            faq_title=faq_item.get('titulo', 'Geral'),
                            md_server_rel_url=faq_item.get('url', ''),
                            image_loader=ensure_sharepoint_image_cached,
                            video_loader=ensure_sharepoint_video_cached,
                        )
                        st.markdown(rendered_md, unsafe_allow_html=True)
                else:
                    st.info("O conteúdo detalhado deste FAQ ainda não foi sincronizado localmente.")
                    st.write("Você pode visualizar o tutorial completo diretamente no SharePoint pelo botão abaixo.")
                    
                st.markdown("---")
                st.markdown(f'<a href="{faq_item["url"]}" target="_blank" style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; font-weight: bold; padding: 8px 16px; border-radius: 6px;">🔗 Abrir no SharePoint (Nova Aba) ↗</a>', unsafe_allow_html=True)

            filtered_df = df_faqs.copy()
            if search_query:
                q_mask = (
                    filtered_df['titulo'].str.contains(search_query, case=False, na=False)
                    | filtered_df['tipo_faq'].str.contains(search_query, case=False, na=False)
                    | filtered_df['conteudo'].str.contains(search_query, case=False, na=False)
                )
                filtered_df = filtered_df[q_mask]
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
        catalog_videos = _load_videos_catalog_cached()
        videos_list = _scan_video_faqs_cached()
        df_faqs = pd.DataFrame()
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
        from src.auth import is_admin
        if is_admin():
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
                target_sp_video = video_item.get('url') or VIDEO_FAQ_URL
                video_rendered = False
                
                # 1. Se já existe o arquivo local em disco
                if caminho and Path(caminho).exists():
                    file_p = Path(caminho)
                    file_size = file_p.stat().st_size
                    # Se for até 75MB, reproduz diretamente no navegador via base64 / st.video
                    if file_size <= 75 * 1024 * 1024:
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
                            with open(file_p, 'rb') as f:
                                video_bytes = f.read()
                            st.video(video_bytes, format=mime_type)
                            video_rendered = True
                        except Exception as e_vid:
                            try:
                                st.video(str(file_p), format="video/mp4")
                                video_rendered = True
                            except Exception as e_fb:
                                st.error(f"Erro ao reproduzir arquivo local: {e_fb}")
                    else:
                        # Vídeo grande (>75MB): exibe aviso e botão direto bancada:// para VLC para não estourar memória do navegador
                        st.info(f"💡 Este tutorial é um vídeo em alta definição ({format_file_size(file_size)}). Para evitar lentidão no navegador, assista-o diretamente pelo player VLC:")
                        c_bancada_play = f'''
                        <div style="margin: 18px auto; text-align: center;">
                            <a href="bancada://run?tool=vlc&target={file_p}" style="display: inline-block; background-color: #3b82f6; color: white; text-decoration: none; font-weight: bold; font-size: 1.05rem; padding: 12px 24px; border-radius: 8px; box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);">
                                ▶️ Abrir no Player VLC Local (Windows)
                            </a>
                        </div>
                        '''
                        st.markdown(c_bancada_play, unsafe_allow_html=True)
                        video_rendered = True

                # 2. Se for link em nuvem do SharePoint e ainda não tiver arquivo local
                if not video_rendered:
                    st.info("💡 Este vídeo está na nuvem corporativa do SharePoint Online. Clique no botão abaixo para baixar usando seus cookies corporativos ou assista no SharePoint Stream:")
                    
                    c_dl1, c_dl2 = st.columns([2, 1])
                    with c_dl1:
                        if st.button("📥 Baixar Vídeo para o Sistema", type="primary", key=f"btn_dl_play_{hash(video_item.get('titulo'))}", width='stretch', help="Faz o download autenticado do vídeo para o servidor usando a sessão ativa."):
                            with st.spinner("Baixando vídeo tutorial com cookies corporativos autenticados..."):
                                try:
                                    _, local_downloaded = ensure_sharepoint_video_cached(target_sp_video, faq_slug=video_item.get('categoria', 'Geral'))
                                    if local_downloaded and local_downloaded.exists() and local_downloaded.stat().st_size > 1000:
                                        video_item['caminho'] = local_downloaded
                                        st.success(f"🎉 Vídeo baixado com sucesso ({format_file_size(local_downloaded.stat().st_size)})! Atualizando tela...")
                                        time.sleep(1)
                                        st.rerun()
                                    else:
                                        st.warning("Não foi possível transferir o arquivo direto com a sessão atual. Utilize a opção de assistir no SharePoint Stream:")
                                        st.link_button("🌐 Assistir no SharePoint Stream ↗", target_sp_video, width='stretch')
                                except Exception as e_dl:
                                    st.error(f"Erro durante download: {e_dl}")
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
                if video_item.get('caminho') and Path(video_item['caminho']).exists():
                    st.markdown(
                        f'<a href="bancada://run?tool=vlc&target={video_item["caminho"]}" style="display: block; text-align: center; background-color: #3b82f6; color: white; text-decoration: none; font-weight: bold; padding: 9px; border-radius: 6px; box-shadow: 0 2px 6px rgba(59,130,246,0.4);">🎬 Abrir no VLC</a>',
                        unsafe_allow_html=True
                    )
                else:
                    target_media = str(video_item.get('url') or VIDEO_FAQ_URL)
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
        
        # Mapeia arquivos locais por múltiplos critérios (nome exato, slug e stem)
        local_vids_map = {v['nome_arquivo']: v for v in videos_list}
        local_vids_by_title = {v['titulo']: v for v in videos_list}

        def find_local_video(filename, title):
            if filename in local_vids_map:
                return local_vids_map[filename]
            if title in local_vids_by_title:
                return local_vids_by_title[title]
            # Busca por correspondência aproximada ou por arquivo em disco
            clean_fn = unquote(filename).strip()
            for v in videos_list:
                if v['nome_arquivo'].lower() == clean_fn.lower() or v['titulo'].lower() == title.lower():
                    return v
            # Varredura em disco por segurança
            if VIDEO_FAQ_DIR.exists() and clean_fn:
                for match_f in VIDEO_FAQ_DIR.rglob(clean_fn):
                    if match_f.is_file() and match_f.stat().st_size > 1000:
                        return {
                            "titulo": match_f.stem,
                            "nome_arquivo": match_f.name,
                            "categoria": match_f.parent.name,
                            "caminho": match_f,
                            "tamanho": format_file_size(match_f.stat().st_size),
                            "extensao": match_f.suffix.lower(),
                            "origem": "Upload Local"
                        }
            return None

        if catalog_videos:
            for item in catalog_videos:
                fname = item.get("nome_arquivo", "")
                ftitle = item.get("titulo", "")
                matched_local = find_local_video(fname, ftitle)
                
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
                            if vid.get('caminho') and Path(vid['caminho']).exists():
                                st.markdown(
                                    f'<a href="bancada://run?tool=vlc&target={vid["caminho"]}" style="display: block; text-align: center; background-color: #3b82f6; color: white; text-decoration: none; font-weight: bold; padding: 7px; border-radius: 6px; font-size: 0.88rem; box-shadow: 0 2px 6px rgba(59,130,246,0.35);">▶️ VLC ↗</a>',
                                    unsafe_allow_html=True
                                )
                            else:
                                sp_url = vid.get('url') or VIDEO_FAQ_URL
                                st.link_button("🌐 Stream ↗", url=sp_url, width='stretch', help="Abrir vídeo direto no SharePoint Online / Stream")

            render_pagination_controls("faq_vid", cur_p_vid, tot_p_vid, tot_i_vid, items_per_page_vid)

    elif active_tab == "🖼️ Imagens FAQ (Galeria)":
        catalog_imagens = _load_imagens_catalog_cached()
        imagens_list = _scan_image_faqs_cached()
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
        from src.auth import is_admin
        if is_admin():
            st.sidebar.markdown("## ⚙️ Ações e SharePoint")
            if IMAGE_FAQ_URL:
                st.sidebar.link_button("🌐 Abrir Pasta no SharePoint ↗", IMAGE_FAQ_URL, width='stretch', help="Abre a pasta oficial de imagens no SharePoint em nova aba.")
            if st.sidebar.button("📤 Enviar Imagens", width='stretch', help="Fazer upload de imagens para a galeria local."):
                modal_upload_imagem()

        # Constrói pastas de imagens unificando imagens locais com o catálogo oficial do SQLite
        folders_dict = {}
        local_by_folder = {}
        for img in imagens_list:
            cat = img['categoria']
            local_by_folder.setdefault(cat, []).append(img)

        # Mapeia catálogo de imagens do SQLite por categoria/pasta
        catalog_by_cat = {}
        if catalog_imagens:
            for c_img in catalog_imagens:
                c_cat = c_img.get('categoria', 'Geral')
                catalog_by_cat.setdefault(c_cat, []).append(c_img)

        # Constrói pastas de imagens unificando o catálogo oficial da biblioteca 'Imagens FAQ' (SQLite)
        # com os arquivos baixados localmente na pasta uploads/faq/imagens/
        folders_list = []
        seen_slugs = set()

        def build_sp_folder_url(folder_name: str) -> str:
            """Gera a URL oficial de navegação dentro da biblioteca Imagens FAQ no SharePoint."""
            if not folder_name:
                return IMAGE_FAQ_URL
            folder_part = folder_name.replace(" > ", "/")
            encoded_path = quote(f"/sites/dit-manutencao/Shared Documents/Tutoriais-FAQs/Imagens FAQ/{folder_part}")
            return f"https://ministeriopublicoms.sharepoint.com/sites/dit-manutencao/Shared%20Documents/Forms/AllItems.aspx?id={encoded_path}&viewid=fa6fa6a3%2Dc035%2D4588%2Da4c8%2D6df7e448e42d"

        # 1. Pastas do Catálogo Oficial da biblioteca 'Imagens FAQ'
        for cat_name, c_imgs in catalog_by_cat.items():
            cat_slug = slugify_faq_title(cat_name)
            seen_slugs.add(cat_slug)
            seen_slugs.add(cat_name)

            # Localiza imagens em disco correspondentes (por slug ou nome direto)
            local_imgs = local_by_folder.get(cat_slug, []) or local_by_folder.get(cat_name, [])
            combined_imgs = list(local_imgs)
            local_names = {
                Path(x['caminho']).name.lower() if x.get('caminho') else x['nome_arquivo'].lower()
                for x in local_imgs
            }

            for ci in c_imgs:
                fn = ci.get('nome_arquivo', '')
                if fn.lower() not in local_names:
                    combined_imgs.append({
                        "titulo": ci.get('titulo') or Path(fn).stem,
                        "nome_arquivo": fn,
                        "categoria": cat_name,
                        "caminho": None,
                        "url": ci.get('url', IMAGE_FAQ_URL),
                        "tamanho": format_file_size(ci.get('tamanho_bytes', 0)),
                        "extensao": Path(fn).suffix.lower() if fn else ".png",
                        "origem": "SharePoint Online"
                    })

            folder_group = cat_name.split(" > ")[0] if " > " in cat_name else "Imagens FAQ"
            folders_list.append({
                "titulo": cat_name,
                "categoria": folder_group,
                "slug": cat_slug,
                "imagens": sorted(combined_imgs, key=lambda x: x["titulo"]),
                "total": len(combined_imgs),
                "url": build_sp_folder_url(cat_name)
            })

        # 2. Pastas extras locais que não estejam no catálogo
        for cat, imgs in local_by_folder.items():
            cat_slug = slugify_faq_title(cat)
            if cat not in seen_slugs and cat_slug not in seen_slugs:
                folders_list.append({
                    "titulo": cat.replace("_", " "),
                    "categoria": "Upload Local",
                    "slug": cat_slug,
                    "imagens": sorted(imgs, key=lambda x: x["titulo"]),
                    "total": len(imgs),
                    "url": build_sp_folder_url(cat.replace("_", " "))
                })
                seen_slugs.add(cat)
                seen_slugs.add(cat_slug)

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
            matching_folder = next((f for f in folders_list if f.get('categoria') == folder_name or f.get('titulo') == folder_name or f.get('slug') == folder_name), None)

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

            st.subheader(f"📂 {matching_folder.get('titulo', folder_name)}")
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
                img_path = img_item.get('caminho')
                img_url = img_item.get('url')

                # Se o arquivo não existir localmente, tenta baixar via cookies sob demanda
                if (not img_path or not Path(img_path).exists()) and img_url:
                    with st.spinner("Baixando imagem com sessão corporativa..."):
                        b64_res = ensure_sharepoint_image_cached(img_url, faq_slug=matching_folder.get('slug', 'Geral'))
                        if b64_res:
                            # Procura o arquivo recém-baixado em disco para salvar na estrutura
                            clean_fn = unquote(img_url.split("/")[-1].split("?")[0])
                            dest_f = IMAGE_FAQ_DIR / matching_folder.get('slug', 'Geral') / clean_fn
                            if dest_f.exists():
                                img_item['caminho'] = dest_f
                                img_path = dest_f

                try:
                    if img_path and Path(img_path).exists():
                        b64_src = get_image_as_base64(Path(img_path))
                        if b64_src:
                            st.markdown(f'<img src="{b64_src}" style="max-height: 480px; max-width: 100%; object-fit: contain; margin: 0 auto; display: block; border-radius: 8px; box-shadow: 0 4px 14px rgba(0,0,0,0.4);" />', unsafe_allow_html=True)
                        else:
                            st.image(str(img_path), width='stretch')
                    elif img_url:
                        st.image(img_url, width='stretch')
                    else:
                        st.warning("Imagem não disponível para visualização.")
                except Exception as e:
                    st.error(f"Erro ao carregar a imagem: {e}")

            with c_next:
                if st.button("Próximo ➡️", key="btn_next_img", width='stretch', disabled=(idx == len(folder_imgs) - 1)):
                    st.session_state['current_img_idx'] = idx + 1
                    st.rerun()

            st.markdown("---")
            c_info, c_act = st.columns([2.5, 1.5])
            with c_info:
                st.caption(f"💾 **Tamanho:** `{img_item['tamanho']}`")
                if img_item.get('caminho'):
                    st.caption(f"📁 **Arquivo:** `{img_item['caminho']}`")
                elif img_item.get('url'):
                    st.caption(f"🔗 **URL SharePoint:** `{img_item['url']}`")
            with c_act:
                target_url = img_item.get('url') or matching_folder.get('url') or IMAGE_FAQ_URL
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
        links_uteis = _load_links_uteis_cached()
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

