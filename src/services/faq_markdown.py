# -*- coding: utf-8 -*-
"""
Serviço de leitura, deduplicação, resolução de mídias e renderização de FAQs em Markdown (.md)
armazenados na biblioteca 'Manuais Markdown' do SharePoint Online (ou pasta local em fallback).
"""
import html
import posixpath
import re
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple
from urllib.parse import parse_qs, quote, unquote, urlparse


DEFAULT_SP_HOST = "https://ministeriopublicoms.sharepoint.com"
DEFAULT_SP_SITE_PATH = "/sites/dit-manutencao"
DEFAULT_SP_SITE_URL = f"{DEFAULT_SP_HOST}{DEFAULT_SP_SITE_PATH}"
DEFAULT_MARKDOWN_SERVER_REL_PATH = (
    "/sites/dit-manutencao/Shared Documents/Tutoriais-FAQs/Manuais Markdown"
)

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg"}
VIDEO_EXTENSIONS = {".mp4", ".mov", ".webm", ".avi", ".mkv", ".wmv"}

# Mapeamento amigável de nomes de subpastas para categorias exibidas na UI
CATEGORY_DISPLAY_NAMES: Dict[str, str] = {
    "acesso_remoto_e_teletrabalho": "Acesso Remoto e Teletrabalho",
    "administrativo": "Administrativo",
    "cerimonial": "Cerimonial",
    "formatacao": "Formatação",
    "formatacao_e_bancada": "Formatação e Bancada",
    "garantia": "Garantia",
    "geral": "Geral",
    "impressoras": "Impressoras",
    "impressoras_e_rede": "Impressoras e Rede",
    "instalacao": "Instalação",
    "manuais": "Manuais",
    "notebooks_e_perifericos": "Notebooks e Periféricos",
    "onedrive_e_backup": "OneDrive e Backup",
    "sistemas_e_portais": "Sistemas e Portais",
    "telefonia": "Telefonia",
    "telefonia_e_oxe": "Telefonia e OXE",
}

_LOWERCASE_PT_WORDS = {"e", "de", "da", "do", "das", "dos", "em", "no", "na", "nos", "nas", "para", "com", "por", "a", "o", "as", "os"}

# Regex para imagens Markdown: ![alt](<url>) ou ![alt](url)
_MD_IMAGE_RE = re.compile(r'!\[([^\]]*)\]\((?:<([^>]+)>|([^)\s]+))\)')

# Regex para links Markdown: [label](<url>) ou [label](url) (não precedido por '!')
_MD_LINK_RE = re.compile(r'(?<!\!)\[([^\]]+)\]\((?:<([^>]+)>|([^)\s]+))\)')

# Regex para callouts estilo GitHub: > [!TIP] / > [!IMPORTANT] / etc.
_CALLOUT_STYLES: Dict[str, Tuple[str, str, str, str]] = {
    "NOTE": ("ℹ️ Nota", "#3b82f6", "rgba(59, 130, 246, 0.12)", "#93c5fd"),
    "TIP": ("💡 Dica", "#10b981", "rgba(16, 185, 129, 0.12)", "#6ee7b7"),
    "IMPORTANT": ("📌 Importante", "#8b5cf6", "rgba(139, 92, 246, 0.14)", "#c4b5fd"),
    "WARNING": ("⚠️ Atenção", "#f59e0b", "rgba(245, 158, 11, 0.14)", "#fcd34d"),
    "CAUTION": ("🚨 Cuidado", "#ef4444", "rgba(239, 68, 68, 0.14)", "#fca5a5"),
}


def encode_sharepoint_url_path(server_rel_path: str, sp_host: str = DEFAULT_SP_HOST) -> str:
    """Monta URL absoluta do SharePoint codificando cada segmento de caminho."""
    if not server_rel_path:
        return ""
    clean = server_rel_path.strip()
    if not clean.startswith("/"):
        clean = "/" + clean
    segments = [quote(unquote(seg)) for seg in clean.split("/")]
    return f"{sp_host.rstrip('/')}{'/'.join(segments)}"


def extract_sharepoint_folder_rel_path(url_or_path: str) -> str:
    """
    Extrai o ServerRelativeUrl de uma pasta do SharePoint a partir de:
    - URL de visualização AllItems.aspx?id=%2Fsites%2F...
    - URL direta https://.../sites/.../Shared Documents/...
    - Caminho relativo ao servidor /sites/...
    Retorna DEFAULT_MARKDOWN_SERVER_REL_PATH caso não seja uma URL/caminho SharePoint válido.
    """
    if not url_or_path or not isinstance(url_or_path, str):
        return DEFAULT_MARKDOWN_SERVER_REL_PATH

    raw = url_or_path.strip()
    if raw.startswith("/sites/"):
        return unquote(raw).rstrip("/")

    if raw.lower().startswith(("http://", "https://")):
        parsed = urlparse(raw)
        qs = parse_qs(parsed.query)
        if "id" in qs and qs["id"]:
            cand = unquote(qs["id"][0]).strip().rstrip("/")
            if cand.startswith("/"):
                return cand
        path_decoded = unquote(parsed.path).strip().rstrip("/")
        if "/Forms/" not in path_decoded and path_decoded.startswith("/sites/"):
            return path_decoded

    return DEFAULT_MARKDOWN_SERVER_REL_PATH


def extract_sharepoint_site_url(url_or_path: str) -> str:
    """Extrai a URL base do site SharePoint (ex.: https://host/sites/dit-manutencao)."""
    if not url_or_path or not isinstance(url_or_path, str):
        return DEFAULT_SP_SITE_URL

    raw = url_or_path.strip()
    if raw.lower().startswith(("http://", "https://")):
        parsed = urlparse(raw)
        host = f"{parsed.scheme}://{parsed.netloc}"
        m = re.match(r"^(/sites/[^/]+)", unquote(parsed.path))
        if m:
            return f"{host}{m.group(1)}"
        rel = extract_sharepoint_folder_rel_path(raw)
        m2 = re.match(r"^(/sites/[^/]+)", rel)
        if m2:
            return f"{host}{m2.group(1)}"
    return DEFAULT_SP_SITE_URL


def format_category_name(folder_name: str) -> str:
    """Converte o nome da subpasta (ex.: 'formatacao_e_bancada') em rótulo amigável."""
    if not folder_name or folder_name.strip() in ("", "."):
        return "Geral"

    clean_key = folder_name.strip().lower()
    if clean_key in CATEGORY_DISPLAY_NAMES:
        return CATEGORY_DISPLAY_NAMES[clean_key]

    words = re.split(r"[_\-\s]+", folder_name.strip())
    formatted = []
    for idx, w in enumerate(words):
        if not w:
            continue
        lw = w.lower()
        if idx > 0 and lw in _LOWERCASE_PT_WORDS:
            formatted.append(lw)
        elif w.isupper() and len(w) <= 5:
            formatted.append(w)
        else:
            formatted.append(w.capitalize())
    return " ".join(formatted) if formatted else "Geral"


def _strip_leading_decorations(title: str) -> str:
    """Remove emojis e símbolos decorativos no início do título para exibição e ordenação limpas."""
    if not title:
        return ""
    cleaned = re.sub(r'^[^\w\[\(]+', '', title.strip(), flags=re.UNICODE).strip()
    return cleaned if cleaned else title.strip()


def extract_markdown_title(md_text: str, filename: str = "") -> str:
    """
    Extrai o título principal do documento Markdown (primeiro heading '# ...').
    Se não houver H1, deriva do nome do arquivo.
    """
    if md_text and isinstance(md_text, str):
        for line in md_text.splitlines():
            stripped = line.strip()
            if stripped.startswith("# ") and not stripped.startswith("##"):
                raw_title = stripped[2:].strip()
                clean = _strip_leading_decorations(raw_title)
                if clean:
                    return clean

    stem = Path(filename).stem if filename else "Tutorial"
    return format_category_name(stem)


def _normalize_dedup_key(text: str) -> str:
    """Normaliza nome de arquivo ou título para detecção de duplicidades entre 'geral/' e categorias."""
    import unicodedata
    norm = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
    norm = norm.lower().strip()
    norm = re.sub(r"\.md$", "", norm)
    norm = re.sub(r"[^a-z0-9]+", "_", norm).strip("_")
    return norm


def deduplicate_markdown_faqs(items: List[Dict]) -> List[Dict]:
    """
    Remove duplicatas entre a pasta 'geral/' (espelho legado) e subpastas temáticas.

    Regras:
    1. Se um manual existir tanto em uma pasta temática (folder != 'geral') quanto em 'geral/',
       prioriza sempre a versão da pasta temática e descarta a de 'geral/'.
    2. Se o mesmo arquivo existir em múltiplas pastas temáticas, mantém a versão com maior
       conteúdo (mais completa/refatorada).
    3. Itens exclusivos de 'geral/' são preservados normalmente na categoria 'Geral'.
    """
    by_file_key: Dict[str, Dict] = {}

    for item in items:
        filename = item.get("filename") or Path(item.get("server_rel_url", "")).name
        folder = (item.get("folder") or "geral").strip().lower()
        is_geral = folder in ("", "geral", ".")
        content_len = len(item.get("conteudo") or "")

        file_key = _normalize_dedup_key(filename)
        if not file_key:
            file_key = _normalize_dedup_key(item.get("titulo") or "")

        existing = by_file_key.get(file_key)
        if existing is None:
            by_file_key[file_key] = item
            continue

        existing_folder = (existing.get("folder") or "geral").strip().lower()
        existing_is_geral = existing_folder in ("", "geral", ".")
        existing_len = len(existing.get("conteudo") or "")

        # Pasta temática sempre vence 'geral'
        if existing_is_geral and not is_geral:
            by_file_key[file_key] = item
        elif not existing_is_geral and is_geral:
            continue
        elif content_len > existing_len:
            by_file_key[file_key] = item

    # Segunda passagem: remove duplicatas por título normalizado quando uma delas está em 'geral'
    thematic_titles = {
        _normalize_dedup_key(it.get("titulo") or "")
        for it in by_file_key.values()
        if (it.get("folder") or "geral").strip().lower() not in ("", "geral", ".")
    }

    final_items: List[Dict] = []
    for it in by_file_key.values():
        folder = (it.get("folder") or "geral").strip().lower()
        title_key = _normalize_dedup_key(it.get("titulo") or "")
        if folder in ("", "geral", ".") and title_key and title_key in thematic_titles:
            continue
        final_items.append(it)

    return sorted(final_items, key=lambda x: (x.get("tipo_faq", ""), x.get("titulo", "").lower()))


def resolve_markdown_media_url(
    raw_ref: str,
    md_server_rel_url: str = "",
    sp_host: str = DEFAULT_SP_HOST,
) -> str:
    """
    Resolve caminhos relativos de imagens/vídeos referenciados em um arquivo .md
    (ex.: '<../../Imagens FAQ/Formatação/SCCM 01.png>') para a URL absoluta do SharePoint.
    """
    if not raw_ref or not isinstance(raw_ref, str):
        return ""

    cleaned = raw_ref.strip().strip("<>").strip()
    if not cleaned:
        return ""

    if cleaned.lower().startswith(("http://", "https://", "data:", "mailto:", "#")):
        return cleaned

    if cleaned.startswith("/"):
        norm_path = posixpath.normpath(unquote(cleaned))
        return encode_sharepoint_url_path(norm_path, sp_host=sp_host)

    base_rel = md_server_rel_url or f"{DEFAULT_MARKDOWN_SERVER_REL_PATH}/geral/manual.md"
    if base_rel.lower().startswith(("http://", "https://")):
        base_rel = unquote(urlparse(base_rel).path)

    base_dir = posixpath.dirname(unquote(base_rel))
    joined = posixpath.normpath(posixpath.join(base_dir, unquote(cleaned)))

    # Garante que a normalização não escape para fora do site SharePoint
    if not joined.startswith("/sites/"):
        return ""

    return encode_sharepoint_url_path(joined, sp_host=sp_host)


def extract_markdown_media_references(
    md_text: str,
    md_server_rel_url: str = "",
    sp_host: str = DEFAULT_SP_HOST,
) -> Tuple[List[str], List[str]]:
    """Extrai listas únicas de URLs de imagens e vídeos referenciados em um documento Markdown."""
    if not md_text:
        return [], []

    images: List[str] = []
    videos: List[str] = []
    seen_img = set()
    seen_vid = set()

    for match in _MD_IMAGE_RE.finditer(md_text):
        raw_target = match.group(2) or match.group(3) or ""
        resolved = resolve_markdown_media_url(raw_target, md_server_rel_url, sp_host=sp_host)
        if resolved and resolved not in seen_img:
            seen_img.add(resolved)
            images.append(resolved)

    for match in _MD_LINK_RE.finditer(md_text):
        raw_target = match.group(2) or match.group(3) or ""
        ext = Path(unquote(raw_target.strip("<>")).split("?")[0]).suffix.lower()
        if ext in VIDEO_EXTENSIONS:
            resolved = resolve_markdown_media_url(raw_target, md_server_rel_url, sp_host=sp_host)
            if resolved and resolved not in seen_vid:
                seen_vid.add(resolved)
                videos.append(resolved)

    return images, videos


def is_html_faq_content(content: str) -> bool:
    """Detecta se o conteúdo salvo no banco é HTML legado do SharePoint (Site Pages) ou Markdown."""
    if not content or not isinstance(content, str):
        return False
    sample = content.lstrip()[:400].lower()
    html_markers = (
        "data-sp-controldata",
        'class="imageplugin"',
        "canvaszone",
        "<div ",
        "<div>",
        "<p>",
        "<p ",
    )
    return any(marker in sample for marker in html_markers)


def _convert_github_callouts(md_text: str) -> str:
    """
    Converte blocos de alerta estilo GitHub (> [!TIP], > [!IMPORTANT], etc.)
    em caixas HTML estilizadas compatíveis com o tema escuro do Streamlit.
    """
    lines = md_text.splitlines()
    out: List[str] = []
    i = 0

    while i < len(lines):
        line = lines[i]
        m = re.match(r"^\s*>\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$", line, re.IGNORECASE)
        if m:
            kind = m.group(1).upper()
            label, border_color, bg_color, title_color = _CALLOUT_STYLES[kind]
            body_lines: List[str] = []
            i += 1
            while i < len(lines) and re.match(r"^\s*>", lines[i]):
                body_lines.append(re.sub(r"^\s*>\s?", "", lines[i]))
                i += 1

            # Converte formatação inline simples dentro do callout (**bold**, *italic*, `code`)
            body_html_lines = []
            for bl in body_lines:
                escaped = html.escape(bl)
                escaped = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escaped)
                escaped = re.sub(r"\*(.+?)\*", r"<em>\1</em>", escaped)
                escaped = re.sub(r"`([^`]+)`", r"<code>\1</code>", escaped)
                body_html_lines.append(escaped)
            body_html = "<br>".join(body_html_lines).strip()

            callout_html = (
                f'<div class="md-callout" style="margin: 16px 0; padding: 14px 18px; '
                f'border-left: 4px solid {border_color}; background: {bg_color}; '
                f'border-radius: 0 8px 8px 0;">'
                f'<div style="font-weight: 700; color: {title_color}; margin-bottom: 6px; font-size: 0.95rem;">'
                f'{label}</div>'
                f'<div style="color: #e2e8f0; line-height: 1.65; font-size: 0.96rem;">{body_html}</div>'
                f'</div>'
            )
            out.append(callout_html)
            continue

        out.append(line)
        i += 1

    return "\n".join(out)


def render_markdown_for_streamlit(
    md_text: str,
    faq_title: str = "Geral",
    md_server_rel_url: str = "",
    image_loader: Optional[Callable[[str, str], str]] = None,
    video_loader: Optional[Callable[[str, str], Tuple[str, Optional[Path]]]] = None,
    sp_host: str = DEFAULT_SP_HOST,
) -> str:
    """
    Prepara o texto Markdown para renderização limpa via `st.markdown(..., unsafe_allow_html=True)`:
    - Remove o primeiro `# H1` duplicado (já exibido no cabeçalho do modal);
    - Converte callouts `[!TIP]`, `[!IMPORTANT]`, `[!NOTE]`, `[!WARNING]`, `[!CAUTION]`;
    - Resolve `![alt](<../../Imagens FAQ/...>)` em imagens base64 embutidas (ou fallback com link);
    - Resolve `[label](<../../Vídeos FAQ/...>)` em player de vídeo ou card VLC/SharePoint Stream;
    - Mantém tabelas, listas, títulos e blocos de código nativos do Streamlit Markdown.
    """
    if not md_text or not isinstance(md_text, str):
        return ""

    # 1. Remove apenas o primeiro heading H1 no início do documento (evita título duplicado no modal)
    lines = md_text.splitlines()
    for idx, ln in enumerate(lines[:5]):
        stripped = ln.strip()
        if stripped.startswith("# ") and not stripped.startswith("##"):
            lines.pop(idx)
            break
    processed = "\n".join(lines).lstrip("\n")

    # 2. Converte callouts GitHub
    processed = _convert_github_callouts(processed)

    # 3. Substitui imagens Markdown por tags <img> em base64 (com legenda)
    def _replace_image(match: re.Match) -> str:
        alt_text = (match.group(1) or "").strip()
        raw_target = match.group(2) or match.group(3) or ""
        img_url = resolve_markdown_media_url(raw_target, md_server_rel_url, sp_host=sp_host)
        safe_alt = html.escape(alt_text)

        b64_src = ""
        if image_loader and img_url:
            try:
                b64_src = image_loader(img_url, faq_title) or ""
            except Exception:
                b64_src = ""

        if b64_src:
            caption_html = f'<div class="sp-caption">{safe_alt}</div>' if safe_alt else ""
            return (
                f'\n<div class="sp-img-card">'
                f'<img src="{b64_src}" alt="{safe_alt}" class="sp-image" />'
                f'{caption_html}'
                f'</div>\n'
            )

        if img_url:
            label = safe_alt or html.escape(Path(unquote(raw_target.strip("<>"))).name)
            return (
                f'\n<div class="sp-fallback-msg" style="text-align: center;">'
                f'<p style="margin: 0 0 8px 0; color: #f87171;">🖼️ Imagem não disponível em cache local: <strong>{label}</strong></p>'
                f'<a href="{img_url}" target="_blank" style="color: #60a5fa; font-weight: 600; text-decoration: underline;">'
                f'🔗 Abrir imagem no SharePoint ↗</a>'
                f'</div>\n'
            )
        return match.group(0)

    processed = _MD_IMAGE_RE.sub(_replace_image, processed)

    # 4. Substitui links para vídeos (.mp4, .mov, .webm, etc.) por player/card integrado
    def _replace_video_link(match: re.Match) -> str:
        label = (match.group(1) or "Vídeo Tutorial").strip()
        raw_target = match.group(2) or match.group(3) or ""
        ext = Path(unquote(raw_target.strip("<>")).split("?")[0]).suffix.lower()
        if ext not in VIDEO_EXTENSIONS:
            return match.group(0)

        vid_url = resolve_markdown_media_url(raw_target, md_server_rel_url, sp_host=sp_host)
        safe_label = html.escape(label)

        b64_vid = ""
        local_file: Optional[Path] = None
        if video_loader and vid_url:
            try:
                b64_vid, local_file = video_loader(vid_url, faq_title)
            except Exception:
                b64_vid, local_file = "", None

        if b64_vid:
            return (
                f'\n<div class="sp-img-card">'
                f'<video controls preload="metadata" style="width: 100%; max-height: 500px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.4);">'
                f'<source src="{b64_vid}" type="video/mp4" />'
                f'</video>'
                f'<div class="sp-caption">🎬 {safe_label}</div>'
                f'</div>\n'
            )

        vlc_btn = ""
        if local_file and Path(local_file).exists():
            vlc_btn = (
                f'<a href="bancada://run?tool=vlc&target={local_file}" '
                f'style="display: inline-block; background-color: #3b82f6; color: white; text-decoration: none; '
                f'font-weight: bold; font-size: 0.9rem; padding: 9px 16px; border-radius: 6px;">'
                f'▶️ Abrir no VLC</a>'
            )

        stream_btn = ""
        if vid_url:
            stream_btn = (
                f'<a href="{vid_url}" target="_blank" '
                f'style="display: inline-block; background-color: #ff4b4b; color: white; text-decoration: none; '
                f'font-weight: bold; font-size: 0.9rem; padding: 9px 16px; border-radius: 6px;">'
                f'🔗 Assistir no SharePoint ↗</a>'
            )

        return (
            f'\n<div class="sp-img-card" style="margin: 18px auto; max-width: 720px; background: #1e1f29; '
            f'border: 1px solid #3b3c4a; border-radius: 10px; padding: 16px 20px; text-align: center;">'
            f'<div style="font-size: 1rem; font-weight: 600; color: #f8f9fa; margin-bottom: 10px;">🎬 {safe_label}</div>'
            f'<div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">'
            f'{vlc_btn}{stream_btn}'
            f'</div>'
            f'</div>\n'
        )

    processed = _MD_LINK_RE.sub(_replace_video_link, processed)
    return processed


def scan_local_markdown_faqs(
    base_dir: Path,
    sp_base_rel_path: str = DEFAULT_MARKDOWN_SERVER_REL_PATH,
    sp_host: str = DEFAULT_SP_HOST,
) -> List[Dict]:
    """
    Varre recursivamente uma pasta local de Manuais Markdown (cache ou OneDrive)
    e retorna os registros deduplicados prontos para inserção na tabela `faqs`.
    """
    if not base_dir or not base_dir.exists():
        return []

    collected: List[Dict] = []
    for md_file in sorted(base_dir.rglob("*.md")):
        if not md_file.is_file():
            continue
        if md_file.name.lower() == "readme.md":
            continue

        try:
            rel_path = md_file.relative_to(base_dir)
            folder = rel_path.parts[0] if len(rel_path.parts) > 1 else "geral"
            content = md_file.read_text(encoding="utf-8-sig", errors="replace")
        except Exception:
            continue

        rel_posix = "/".join(rel_path.parts)
        server_rel_url = f"{sp_base_rel_path.rstrip('/')}/{rel_posix}"
        full_url = encode_sharepoint_url_path(server_rel_url, sp_host=sp_host)

        collected.append({
            "filename": md_file.name,
            "folder": folder,
            "titulo": extract_markdown_title(content, md_file.name),
            "tipo_faq": format_category_name(folder),
            "url": full_url,
            "server_rel_url": server_rel_url,
            "conteudo": content,
        })

    return deduplicate_markdown_faqs(collected)


def persist_markdown_faqs_to_db(cursor, conn, md_items: List[Dict]) -> int:
    """
    Faz UPSERT dos Manuais Markdown deduplicados na tabela `faqs` e remove registros antigos
    (como páginas `.aspx` legadas de Site Pages ou duplicatas removidas de `geral/`).
    """
    if not md_items:
        return 0

    active_urls: List[str] = []
    for item in md_items:
        url = item["url"]
        active_urls.append(url)
        cursor.execute(
            """
            INSERT INTO faqs (titulo, tipo_faq, url, conteudo, data_atualizacao)
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(url) DO UPDATE SET
                titulo = excluded.titulo,
                tipo_faq = excluded.tipo_faq,
                conteudo = CASE WHEN excluded.conteudo != '' THEN excluded.conteudo ELSE faqs.conteudo END,
                data_atualizacao = CURRENT_TIMESTAMP
            """,
            (item["titulo"], item["tipo_faq"], url, item["conteudo"]),
        )

    placeholders = ",".join("?" for _ in active_urls)
    cursor.execute(f"DELETE FROM faqs WHERE url NOT IN ({placeholders})", tuple(active_urls))
    conn.commit()
    return len(active_urls)


def mirror_markdown_faqs_to_disk(base_dir: Path, md_items: List[Dict]) -> int:
    """
    Salva os Manuais Markdown deduplicados em `base_dir/<subpasta>/<arquivo>.md`
    e remove arquivos `.md` antigos/obsoletos para manter o cache local espelhado e rápido.
    """
    if not base_dir or not md_items:
        return 0

    base_dir.mkdir(parents=True, exist_ok=True)
    written_paths = set()

    for item in md_items:
        content = item.get("conteudo") or ""
        if not content:
            continue
        folder = re.sub(r"[^a-zA-Z0-9_\-]", "_", (item.get("folder") or "geral").strip()) or "geral"
        filename = Path(item.get("filename") or "manual.md").name
        if not filename.lower().endswith(".md"):
            filename = f"{filename}.md"

        dest = base_dir / folder / filename
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(content, encoding="utf-8")
        written_paths.add(dest.resolve())

    # Remove .md antigos que não fazem mais parte do catálogo deduplicado
    for old_md in base_dir.rglob("*.md"):
        try:
            if old_md.is_file() and old_md.resolve() not in written_paths:
                old_md.unlink(missing_ok=True)
        except Exception:
            pass

    # Remove subdiretórios vazios
    for sub in sorted(base_dir.rglob("*"), reverse=True):
        try:
            if sub.is_dir() and not any(sub.iterdir()):
                sub.rmdir()
        except Exception:
            pass

    return len(written_paths)


