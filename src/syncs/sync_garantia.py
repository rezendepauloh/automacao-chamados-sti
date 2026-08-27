import sys
import tempfile
import os
from pathlib import Path

root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from dotenv import load_dotenv
load_dotenv(override=True)

import time
import requests
from src.components.status_banner import check_process_running, read_log_lines
from src.config import (
    WARRANTY_FILE_PATH, setup_logging, DEBUG_DIR_GARANTIA,
    CITSMART_EMAIL, PASSWORD, HEADLESS, EXPLICIT_WAIT, get_chrome_driver
)
from src.database import sync_garantia_from_excel
from terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_GARANTIA / "garantia.log", "sync_garantia")

def check_garantia_sync_running() -> bool:
    """Verifica se o processo de sincronização de garantias está em execução."""
    lock_file = Path(tempfile.gettempdir()) / "garantia_sync.lock"
    return check_process_running(lock_file)

def read_garantia_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log de garantias."""
    log_path = DEBUG_DIR_GARANTIA / "garantia.log"
    return read_log_lines(log_path, n)

def download_sharepoint_garantia_file(url: str) -> Path | None:
    """Faz o download da planilha de garantia a partir da URL do SharePoint com suporte a autenticação institucional."""
    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "garantia_contratos.xlsx"

    download_url = url + "&download=1" if "?e=" in url else url + "?download=1"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    logger.info(f"🌐 [1/2] Tentando download direto via HTTP GET: {download_url}")
    try:
        resp = requests.get(download_url, headers=headers, timeout=15, allow_redirects=True)
        if resp.status_code == 200 and resp.content.startswith(b'PK'):
            with open(file_path, "wb") as f:
                f.write(resp.content)
            logger.info(f"🎉 SUCESSO NO DOWNLOAD HTTP: Arquivo '.xlsx' salvo ({len(resp.content)} bytes) em '{file_path}'.")
            return file_path
        else:
            logger.info(f"ℹ️ Download HTTP direto retornou página/status {resp.status_code} ({len(resp.content)} bytes). Iniciando autenticação com Selenium...")
    except Exception as e_http:
        logger.warning(f"⚠️ Tentativa HTTP direto falhou: {e_http}")

    # Fallback: Autenticação via Selenium no SharePoint / Microsoft Online
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
                logger.info(f"👤 Email digitado: {CITSMART_EMAIL}")

                submits = driver.find_elements(By.XPATH, "//input[@type='submit'] | //button[@type='submit']")
                if submits:
                    submits[0].click()
                    time.sleep(3)

            if PASSWORD:
                pass_inputs = driver.find_elements(By.XPATH, "//input[@type='password']")
                if pass_inputs:
                    pass_inputs[0].clear()
                    pass_inputs[0].send_keys(PASSWORD or "")
                    logger.info("🔑 Senha da rede preenchida.")

                    submits = driver.find_elements(By.XPATH, "//input[@type='submit'] | //button[@type='submit']")
                    if submits:
                        submits[0].click()
                        logger.info("🖱️ Botão de login acionado. Aguardando resposta da Microsoft...")
                        time.sleep(4)

            try:
                stay_btns = driver.find_elements(By.XPATH, "//input[@id='idSIButton9'] | //input[@value='Sim' or @value='Yes'] | //button[contains(text(),'Sim') or contains(text(),'Yes')]")
                if stay_btns:
                    stay_btns[0].click()
                    logger.info("🖱️ Resposta 'Sim' confirmada na tela 'Manter Sessão Iniciada'.")
                    time.sleep(5)
            except Exception:
                pass

        logger.info(f"📍 URL após autenticação: {driver.current_url}")

        if "sharepoint.com" in driver.current_url:
            logger.info("🔗 Sessão autenticada no SharePoint! Extraindo cookies para download direto...")
            try:
                session_auth = requests.Session()
                session_auth.headers.update(headers)
                for cookie in driver.get_cookies():
                    session_auth.cookies.set(
                        name=cookie['name'],
                        value=cookie['value'],
                        domain=cookie.get('domain')
                    )

                resp_auth = session_auth.get(download_url, timeout=30, allow_redirects=True)
                if resp_auth.status_code == 200 and resp_auth.content and (resp_auth.content.startswith(b'PK') or resp_auth.content.startswith(b'\x50\x4b\x03\x04')):
                    with open(file_path, "wb") as f:
                        f.write(resp_auth.content)
                    logger.info(f"🎉 SUCESSO NO DOWNLOAD VIA SESSÃO: Arquivo Excel salvo ({len(resp_auth.content)} bytes) em '{file_path}'.")
                    return file_path
                else:
                    logger.info(f"ℹ️ Download via cookies retornou status {resp_auth.status_code} ({len(resp_auth.content)} bytes). Disparando download via Selenium...")
            except Exception as e_cookies:
                logger.warning(f"Tentativa de download via cookies autenticados: {e_cookies}")

            driver.get(download_url)
            for _ in range(20):
                time.sleep(1)
                for f_item in output_dir.glob("*.xlsx"):
                    if not f_item.name.endswith(".crdownload") and f_item.stat().st_size > 5000:
                        logger.info(f"🎉 SUCESSO NO DOWNLOAD SELENIUM: Arquivo '{f_item.name}' baixado com {f_item.stat().st_size} bytes.")
                        return f_item
        else:
            logger.warning("⚠️ O navegador ainda não alcançou o SharePoint. Tentando forçar o acesso direto...")
            driver.get(download_url)
            for _ in range(15):
                time.sleep(1)
                for f_item in output_dir.glob("*.xlsx"):
                    if not f_item.name.endswith(".crdownload") and f_item.stat().st_size > 5000:
                        return f_item

        for cr_file in output_dir.glob("*.crdownload"):
            if cr_file.stat().st_size > 5000:
                try:
                    with open(cr_file, "rb") as cr_f:
                        cr_bytes = cr_f.read()
                    if cr_bytes.startswith(b'PK') or cr_bytes.startswith(b'\x50\x4b\x03\x04'):
                        with open(file_path, "wb") as f_out:
                            f_out.write(cr_bytes)
                        logger.info(f"🎉 Recuperado arquivo Excel completo a partir de '{cr_file.name}' ({len(cr_bytes)} bytes).")
                        try:
                            cr_file.unlink()
                        except Exception:
                            pass
                        return file_path
                except Exception as e_cr:
                    logger.warning(f"Erro ao ler .crdownload: {e_cr}")

    except Exception as e_sel:
        logger.error(f"❌ Erro durante download Selenium no SharePoint: {e_sel}", exc_info=True)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    return None

def run_garantia_sync():
    """Executa a leitura da planilha de garantias e salva no banco de dados SQLite."""
    print_header("WORKER - SINCRONIZAÇÃO DE GARANTIAS", color=CYAN)
    logger.info("Iniciando sincronização de garantias em segundo plano...")
    excel_path_env = os.getenv("WARRANTY_EXCEL_RELATIVE_PATH", "").strip()
    target_file = None

    if excel_path_env.startswith("http://") or excel_path_env.startswith("https://"):
        target_file = download_sharepoint_garantia_file(excel_path_env)
    else:
        local_file = Path.home() / excel_path_env if excel_path_env else None
        if local_file and local_file.exists():
            target_file = local_file
        elif WARRANTY_FILE_PATH.exists():
            target_file = WARRANTY_FILE_PATH

    if not target_file or not Path(target_file).exists():
        logger.error(f"Planilha de Garantias não localizada nem baixada (Configuração: {excel_path_env})")
        return

    try:
        sync_garantia_from_excel(str(target_file))
        logger.info("Sincronização de garantias finalizada com sucesso!")
    except Exception as e:
        logger.error(f"Erro durante a sincronização de garantias: {e}")
        raise e

if __name__ == "__main__":
    lock_path = Path(tempfile.gettempdir()) / "garantia_sync.lock"
    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_garantia_sync()
    finally:
        if lock_path.exists():
            try:
                lock_path.unlink()
            except Exception:
                pass
