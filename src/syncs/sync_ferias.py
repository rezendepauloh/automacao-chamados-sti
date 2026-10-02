# -*- coding: utf-8 -*-
"""
Worker assíncrono para sincronização da planilha de Previsão de Férias da Bancada
a partir do SharePoint Online ou arquivo local com persistência no SQLite.
"""
import os
import sys
import tempfile
import time
from pathlib import Path

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from dotenv import load_dotenv
load_dotenv(override=True)

import requests
from src.components.status_banner import check_process_running, read_log_lines
from src.config import (
    setup_logging, DEBUG_DIR_FERIAS, FERIAS_EXCEL_RELATIVE_PATH,
    CITSMART_EMAIL, PASSWORD, HEADLESS, EXPLICIT_WAIT, get_chrome_driver
)
from src.database import sync_ferias_from_excel
from terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_FERIAS / "sync_ferias.log", "sync_ferias")

LOCK_FILE = Path(tempfile.gettempdir()) / "ferias_sync.lock"
LOG_FILE = DEBUG_DIR_FERIAS / "sync_ferias.log"


def check_ferias_sync_running() -> bool:
    """Verifica se o processo de sincronização de férias está em execução."""
    return check_process_running(LOCK_FILE)


def read_ferias_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log de férias."""
    return read_log_lines(LOG_FILE, n)


def download_sharepoint_ferias_file(url: str) -> Path | None:
    """Faz o download da planilha de férias a partir da URL do SharePoint com suporte a cookies e Selenium."""
    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "Previsao_de_Ferias-Manutencao.xlsx"

    # 1. Tenta carregar cookies corporativos salvos
    cookies_file = Path("uploads/faq/sharepoint_cookies.json")
    cookies = {}
    if cookies_file.exists():
        try:
            import json
            with open(cookies_file, "r", encoding="utf-8") as f:
                cookies = json.load(f)
        except Exception:
            pass

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    # Tenta download via API ou link direto se URL for SharePoint
    download_urls = []
    if "Doc.aspx" in url:
        sep = "&" if "?" in url else "?"
        download_urls.append(url + f"{sep}download=1")
    else:
        download_urls.append(url)

    logger.info("🌐 [1/2] Tentando download via requisição HTTP autenticada...")
    for dl_url in download_urls:
        try:
            resp = requests.get(dl_url, cookies=cookies, headers=headers, timeout=15, allow_redirects=True)
            if resp.status_code == 200 and resp.content and (resp.content.startswith(b'PK') or resp.content.startswith(b'\x50\x4b\x03\x04')):
                with open(file_path, "wb") as f:
                    f.write(resp.content)
                logger.info(f"🎉 SUCESSO NO DOWNLOAD HTTP: Planilha salva ({len(resp.content)} bytes) em '{file_path}'.")
                return file_path
        except Exception as e_http:
            logger.warning(f"⚠️ Tentativa HTTP direto falhou: {e_http}")

    # Fallback 1.5: Procura arquivo no OneDrive sincronizado no Windows/WSL
    onedrive_win = Path("/mnt/c/Users/paulogoncalves/OneDrive - Ministerio Público do Estado de Mato Grosso do Sul/Documentos SharePoint DIT-Manutenção/Previsão de Férias-Manutencao.xlsx")
    if onedrive_win.exists():
        logger.info(f"📂 Encontrado arquivo sincronizado no OneDrive local: {onedrive_win}")
        try:
            shutil.copy2(str(onedrive_win), str(file_path))
            return file_path
        except Exception as e_copy:
            logger.warning(f"Erro ao copiar do OneDrive local: {e_copy}")

    # Fallback 2: Selenium
    driver = None
    try:
        logger.info("🌐 [2/2] Inicializando Selenium Chrome Driver para autenticação...")
        driver = get_chrome_driver(headless=HEADLESS)

        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': str(output_dir)}}
        driver.execute("send_command", params)

        logger.info("Navegando até a URL compartilhada do SharePoint...")
        driver.get(url)
        time.sleep(4)

        from selenium.webdriver.common.by import By

        if "login.microsoftonline.com" in driver.current_url or "login.live.com" in driver.current_url or "login" in driver.current_url:
            logger.info("🔑 Tela de login institucional detectada. Preenchendo credenciais AD...")
            user_inputs = driver.find_elements(By.XPATH, "//input[@type='email' or @name='loginfmt' or @type='text']")
            if user_inputs:
                user_inputs[0].clear()
                user_inputs[0].send_keys(CITSMART_EMAIL)
                submits = driver.find_elements(By.XPATH, "//input[@type='submit'] | //button[@type='submit']")
                if submits:
                    submits[0].click()
                    time.sleep(3)

            if PASSWORD:
                pass_inputs = driver.find_elements(By.XPATH, "//input[@type='password']")
                if pass_inputs:
                    pass_inputs[0].clear()
                    pass_inputs[0].send_keys(PASSWORD or "")
                    submits = driver.find_elements(By.XPATH, "//input[@type='submit'] | //button[@type='submit']")
                    if submits:
                        submits[0].click()
                        time.sleep(4)

            try:
                kmsi = driver.find_elements(By.XPATH, "//input[@id='idSIButton9'] | //input[@value='Sim' or @value='Yes'] | //button[contains(text(),'Sim') or contains(text(),'Yes')]")
                if kmsi:
                    kmsi[0].click()
                    time.sleep(5)
            except Exception:
                pass

        if "sharepoint.com" in driver.current_url:
            logger.info("🔗 Sessão autenticada no SharePoint! Extraindo cookies...")
            try:
                session_auth = requests.Session()
                session_auth.headers.update(headers)
                for cookie in driver.get_cookies():
                    session_auth.cookies.set(name=cookie['name'], value=cookie['value'], domain=cookie.get('domain'))
                
                resp_auth = session_auth.get(download_urls[0], timeout=30, allow_redirects=True)
                if resp_auth.status_code == 200 and resp_auth.content and (resp_auth.content.startswith(b'PK') or resp_auth.content.startswith(b'\x50\x4b\x03\x04')):
                    with open(file_path, "wb") as f:
                        f.write(resp_auth.content)
                    logger.info(f"🎉 SUCESSO NO DOWNLOAD VIA SESSÃO SELENIUM: Arquivo salvo em '{file_path}'.")
                    return file_path
            except Exception as e_cook:
                logger.warning(f"Erro no download via cookies Selenium: {e_cook}")

            driver.get(download_urls[0])
            for _ in range(15):
                time.sleep(1)
                for f_item in output_dir.glob("*.xlsx"):
                    if "Previs" in f_item.name and not f_item.name.endswith(".crdownload") and f_item.stat().st_size > 5000:
                        logger.info(f"🎉 SUCESSO NO DOWNLOAD BROWSER: '{f_item.name}' ({f_item.stat().st_size} bytes).")
                        return f_item

    except Exception as e_sel:
        logger.error(f"❌ Erro durante download Selenium no SharePoint: {e_sel}", exc_info=True)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    return None


def run_ferias_sync() -> bool:
    """Executa a sincronização da planilha de férias."""
    print_header("WORKER - SINCRONIZAÇÃO DE FÉRIAS DA BANCADA", color=CYAN)
    logger.info("Iniciando rotina de sincronização de férias...")
    from src.config import _cfg
    excel_path_env = (_cfg("FERIAS_EXCEL_RELATIVE_PATH") or os.getenv("FERIAS_EXCEL_RELATIVE_PATH", "")).strip()
    target_file = None

    if excel_path_env.startswith("http://") or excel_path_env.startswith("https://"):
        target_file = download_sharepoint_ferias_file(excel_path_env)
    elif excel_path_env:
        local_file = Path.home() / excel_path_env
        if local_file.is_file():
            target_file = local_file

    # Fallbacks locais
    if not target_file or not Path(target_file).is_file():
        default_up = Path("uploads/Previsao_de_Ferias-Manutencao.xlsx")
        if default_up.is_file():
            target_file = default_up

    if not target_file or not Path(target_file).is_file():
        onedrive_win = Path("/mnt/c/Users/paulogoncalves/OneDrive - Ministerio Público do Estado de Mato Grosso do Sul/Documentos SharePoint DIT-Manutenção/Previsão de Férias-Manutencao.xlsx")
        if onedrive_win.is_file():
            target_file = onedrive_win

    if not target_file or not Path(target_file).is_file():
        logger.error(f"Planilha de Férias não localizada nem baixada (Configuração: {excel_path_env})")
        return False

    try:
        success = sync_ferias_from_excel(str(target_file))
        if success:
            logger.info("✅ Sincronização de férias finalizada com sucesso!")
            return True
        return False
    except Exception as e:
        logger.error(f"Erro durante a sincronização de férias: {e}")
        raise e


if __name__ == "__main__":
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_ferias_sync()
    finally:
        if LOCK_FILE.exists():
            try:
                LOCK_FILE.unlink()
            except Exception:
                pass
