import json
import logging
import time
import urllib3
import requests
from pathlib import Path
from typing import Optional, List, Dict

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger("sharepoint_downloader")

def _get_credentials():
    """Obtém credenciais institucionais para autenticação na Microsoft."""
    try:
        from src.config import CITSMART_EMAIL, PASSWORD
        return CITSMART_EMAIL, PASSWORD
    except Exception as e:
        logger.warning(f"Não foi possível obter credenciais de src.config: {e}")
        return "", ""

def get_cached_cookies() -> Optional[List[Dict]]:
    """Carrega os cookies de sessão do SharePoint armazenados no banco de dados."""
    try:
        from src.database.settings_db import get_setting
        raw = get_setting("SHAREPOINT_SESSION_COOKIES")
        if raw and raw.strip():
            cookies = json.loads(raw)
            if isinstance(cookies, list) and len(cookies) > 0:
                return cookies
    except Exception as e:
        logger.debug(f"Erro ao ler cookies em cache: {e}")
    return None

def save_cached_cookies(cookies: List[Dict]) -> bool:
    """Salva os cookies de sessão do SharePoint de forma cifrada no banco de dados."""
    try:
        from src.database.settings_db import set_setting
        raw = json.dumps(cookies)
        return set_setting(
            key="SHAREPOINT_SESSION_COOKIES",
            value=raw,
            is_secret=True,
            category="sharepoint",
            description="Cookies de autenticação persistentes do Microsoft SharePoint"
        )
    except Exception as e:
        logger.warning(f"Erro ao salvar cookies de sessão no banco: {e}")
        return False

def build_download_url(url: str) -> str:
    """Garante que a URL compartilhada do SharePoint contenha o parâmetro de download direto."""
    u = url.strip()
    if "download=1" in u:
        return u
    if "?" in u:
        return u + "&download=1"
    return u + "?download=1"

MAX_DOWNLOAD_BYTES = 100 * 1024 * 1024  # 100 MB limite defensivo contra DoS / Zip-bomb

def download_via_requests(url: str, output_path: Path, cookies: List[Dict], timeout: int = 25) -> Optional[Path]:
    """
    Tenta baixar a planilha diretamente via HTTP GET utilizando cookies de sessão em cache.
    Implementa controle de tamanho máximo de streaming (DoS/Zip-bomb protection) e validação PK.
    """
    download_url = build_download_url(url)
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    })

    for c in cookies:
        name = c.get("name")
        val = c.get("value")
        domain = c.get("domain", "")
        path = c.get("path", "/")
        if name and val:
            session.cookies.set(name=name, value=val, domain=domain, path=path)

    try:
        logger.info("⚡ [Requests Fast] Tentando download direto via HTTP com sessão em cache...")
        t0 = time.time()
        resp = session.get(download_url, verify=False, timeout=timeout, allow_redirects=True, stream=True)
        dur = time.time() - t0

        if resp.status_code != 200:
            logger.info(f"ℹ️ Download via requests retornou status {resp.status_code}. Sessão expirada ou redirecionada.")
            return None

        # Validação do cabeçalho Content-Length
        cl_header = resp.headers.get("Content-Length")
        if cl_header:
            try:
                if int(cl_header) > MAX_DOWNLOAD_BYTES:
                    logger.error(f"❌ Download abortado: tamanho anunciado ({cl_header} bytes) excede o limite de segurança de 100 MB.")
                    return None
            except ValueError:
                pass

        # Leitura segura com streaming limitado (com fallback para resp.content em mocks de teste)
        downloaded = bytearray()
        if hasattr(resp, "iter_content"):
            try:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        downloaded.extend(chunk)
                        if len(downloaded) > MAX_DOWNLOAD_BYTES:
                            logger.error(f"❌ Download abortado: tamanho transferido excedeu o limite de segurança de {MAX_DOWNLOAD_BYTES} bytes.")
                            return None
            except Exception:
                pass

        if not downloaded and hasattr(resp, "content") and resp.content:
            downloaded = bytearray(resp.content[:MAX_DOWNLOAD_BYTES])

        content_bytes = bytes(downloaded)
        # Verifica se o arquivo retornado é um Excel válido (formato ZIP / OpenXML com assinatura PK)
        is_excel = bool(content_bytes and (content_bytes.startswith(b'PK\x03\x04') or content_bytes.startswith(b'PK')))

        if is_excel:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, "wb") as f:
                f.write(content_bytes)
            logger.info(f"🎉 SUCESSO FAST-DOWNLOAD: Planilha salva em {dur:.2f}s ({len(content_bytes)} bytes) em '{output_path}'.")
            return output_path
        else:
            logger.info(f"ℹ️ Conteúdo baixado ({len(content_bytes)} bytes) não é arquivo Excel válido (PK signature ausente).")
    except Exception as e:
        logger.warning(f"⚠️ Tentativa via requests com cache falhou: {e}")

    return None

def renew_cookies_with_playwright(initial_url: str) -> Optional[List[Dict]]:
    """
    Inicializa o Chromium headless via Playwright, realiza o login corporativo
    no Microsoft Entra ID (Azure AD) e salva os novos cookies de sessão no banco.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.error("Playwright não instalado. Execute: pip install playwright && playwright install chromium")
        return None

    email, password = _get_credentials()
    if not email or not password:
        logger.error("Credenciais corporativas (CITSMART_EMAIL ou PASSWORD) ausentes para login no SharePoint.")
        return None

    logger.info("🌐 [Playwright] Inicializando Chromium headless para renovação de sessão Microsoft...")
    t0 = time.time()

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--ignore-certificate-errors", "--no-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()

        try:
            logger.info(f"Acessando link compartilhado do SharePoint...")
            page.goto(initial_url, wait_until="domcontentloaded", timeout=35000)

            # Se redirecionou para tela de login da Microsoft
            if "login.microsoftonline.com" in page.url:
                logger.info("🔑 Tela de login corporativo da Microsoft detectada. Preenchendo e-mail...")
                page.fill('input[name="loginfmt"], input[type="email"]', email)
                page.click('input[type="submit"], button[type="submit"], #idSIButton9')

                page.wait_for_selector('input[name="passwd"], input[type="password"]', timeout=12000)
                logger.info("🔑 Preenchendo senha...")
                page.fill('input[name="passwd"], input[type="password"]', password)
                page.click('input[type="submit"], button[type="submit"], #idSIButton9')

                # Confirmação KMSI ("Manter conectado?")
                try:
                    page.wait_for_selector('#KmsiCheckboxField, input[id="idSIButton9"], button:has-text("Sim"), input[value="Sim"]', timeout=10000)
                    logger.info("🖱️ Confirmando 'Manter conectado' (KMSI)...")
                    page.click('input[id="idSIButton9"], button:has-text("Sim"), input[value="Sim"]')
                except Exception:
                    pass

                # Aguarda redirecionamento para o SharePoint
                try:
                    page.wait_for_url(lambda u: "sharepoint.com" in u, timeout=18000)
                    logger.info("✅ Redirecionamento para o SharePoint concluído com sucesso!")
                except Exception:
                    logger.info(f"URL pós-login: {page.url[:80]}")

            cookies = context.cookies()
            sp_cookies = [c for c in cookies if "sharepoint.com" in c.get("domain", "") or "microsoftonline.com" in c.get("domain", "")]
            dur = time.time() - t0
            logger.info(f"🎉 Sessão Microsoft capturada em {dur:.2f}s ({len(cookies)} cookies totais, {len(sp_cookies)} do domínio SharePoint).")

            if cookies:
                save_cached_cookies(cookies)
                return cookies

        except Exception as e:
            logger.error(f"❌ Erro durante renovação de sessão via Playwright: {e}", exc_info=True)
        finally:
            browser.close()

    return None

def download_sharepoint_file(url: str, output_path: Path, force_renew: bool = False) -> Optional[Path]:
    """
    Interface unificada de alta performance para download de planilhas do SharePoint:
    1. Tenta baixar via requests puro com cookies em cache (~0.3s).
    2. Se os cookies expiraram ou não existem, renova a sessão via Playwright (~8s).
    3. Re-executa o download via requests com os novos cookies e grava no banco.
    """
    if not url or not (url.startswith("http://") or url.startswith("https://")):
        logger.error(f"URL inválida para download do SharePoint: '{url}'")
        return None

    # Etapa 1: Tentativa com cookies em cache (99% dos casos em produção)
    if not force_renew:
        cached_cookies = get_cached_cookies()
        if cached_cookies:
            file_res = download_via_requests(url, output_path, cached_cookies)
            if file_res:
                return file_res
            logger.info("🔄 Cookies em cache expirados ou inválidos. Iniciando renovação via Playwright...")
        else:
            logger.info("ℹ️ Nenhum cookie de sessão em cache encontrado. Inicializando primeira autenticação...")

    # Etapa 2: Renovação via Playwright
    new_cookies = renew_cookies_with_playwright(url)
    if new_cookies:
        file_res = download_via_requests(url, output_path, new_cookies)
        if file_res:
            return file_res

    logger.error(f"❌ Falha ao baixar planilha do SharePoint a partir de '{url}'")
    return None
