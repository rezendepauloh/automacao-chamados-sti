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
    VIAGENS_FILE_PATH, setup_logging, DEBUG_DIR_VIAGENS,
    CITSMART_EMAIL, PASSWORD, HEADLESS, EXPLICIT_WAIT, get_chrome_driver
)
from src.database import sync_viagens_from_excel
from terminal import print_header, CYAN

logger = setup_logging(DEBUG_DIR_VIAGENS / "viagens_sync.log", "sync_viagens")

def check_viagens_sync_running() -> bool:
    """Verifica se o processo de sincronização de viagens está em execução."""
    lock_file = Path(tempfile.gettempdir()) / "viagens_sync.lock"
    return check_process_running(lock_file)

def read_viagens_last_log_lines(n: int = 15) -> str:
    """Lê as últimas N linhas do arquivo de log de viagens."""
    log_path = DEBUG_DIR_VIAGENS / "viagens_sync.log"
    return read_log_lines(log_path, n)

def download_sharepoint_viagens_file(url: str) -> Path | None:
    """Faz o download da planilha de viagens a partir da URL do SharePoint com suporte a autenticação institucional."""
    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "planilha_viagens.xlsx"

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
            logger.info(f"ℹ️ Download HTTP direto retornou status {resp.status_code} ({len(resp.content)} bytes). Iniciando autenticação com Selenium...")
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
                        time.sleep(4)

        # Trata persistência 'Mantenha-se conectado'
        try:
            kmsi = driver.find_elements(By.XPATH, "//input[@type='submit' and (@id='idSIButton9' or @value='Sim')]")
            if kmsi:
                kmsi[0].click()
                time.sleep(3)
        except Exception:
            pass

        # Aguarda download ou navegação
        time.sleep(5)
        # Verifica se arquivo baixado apareceu
        files = list(output_dir.glob("*.xlsx"))
        if files:
            newest = max(files, key=lambda f: f.stat().st_mtime)
            return newest
    except Exception as e_sel:
        logger.error(f"❌ Erro durante download Selenium no SharePoint: {e_sel}", exc_info=True)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    return None

def run_viagens_sync():
    """Executa a sincronização da planilha de viagens."""
    print_header("WORKER - SINCRONIZAÇÃO DE VIAGENS DA BANCADA", color=CYAN)
    logger.info("Iniciando sincronização de viagens...")
    from src.config import _cfg
    excel_path_env = (_cfg("VIAGENS_EXCEL_RELATIVE_PATH") or os.getenv("VIAGENS_EXCEL_RELATIVE_PATH", "")).strip()
    target_file = None

    if excel_path_env.startswith("http://") or excel_path_env.startswith("https://"):
        target_file = download_sharepoint_viagens_file(excel_path_env)
    elif excel_path_env:
        local_file = Path.home() / excel_path_env
        if local_file.is_file():
            target_file = local_file
        elif VIAGENS_FILE_PATH.is_file():
            target_file = VIAGENS_FILE_PATH

    # Fallback para uploads/temp_viagens.xlsx se disponível
    if not target_file or not Path(target_file).is_file():
        temp_up = Path("uploads/temp_viagens.xlsx")
        if temp_up.is_file():
            target_file = temp_up

    # Fallback para caminho do OneDrive do usuário no Windows se no WSL
    if not target_file or not Path(target_file).is_file():
        onedrive_win = Path("/mnt/c/Users/paulogoncalves/OneDrive - Ministerio Público do Estado de Mato Grosso do Sul/Documentos SharePoint DIT-Manutenção/Viagens/Planilha de Viagens.xlsx")
        if onedrive_win.is_file():
            target_file = onedrive_win

    if not target_file or not Path(target_file).is_file():
        logger.error(f"Planilha de Viagens não localizada nem baixada (Configuração: {excel_path_env})")
        return False

    try:
        success = sync_viagens_from_excel(str(target_file))
        if success:
            logger.info("Sincronização de viagens finalizada com sucesso!")
            return True
        return False
    except Exception as e:
        logger.error(f"Erro durante a sincronização de viagens: {e}")
        raise e

if __name__ == "__main__":
    lock_path = Path(tempfile.gettempdir()) / "viagens_sync.lock"
    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))

    try:
        run_viagens_sync()
    finally:
        if lock_path.exists():
            try:
                lock_path.unlink()
            except Exception:
                pass
