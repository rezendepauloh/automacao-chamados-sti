import os
import re
import sys
import time
import argparse
import tempfile
import ctypes
import logging
from pathlib import Path
from datetime import datetime

# Adiciona o diretório raiz e o diretório src ao sys.path para suportar importações diretas
root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

import pandas as pd
import requests
import keyring

from src.config import (
    setup_logging, DEBUG_DIR_PLANTOES,
    SHAREPOINT_MATUTINO_URL, CITSMART_EMAIL, PASSWORD,
    HEADLESS, EXPLICIT_WAIT, USERNAME, get_chrome_driver, cleanup_old_files
)
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

logger = setup_logging(DEBUG_DIR_PLANTOES / "plantoes.log", "plantoes")


logging.getLogger('selenium.webdriver.remote.remote_connection').setLevel(logging.WARNING)
logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

MESES_PT = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
    "outubro": 10, "novembro": 11, "dezembro": 12
}

def check_plantoes_sync_running() -> bool:
    lock_file = Path(tempfile.gettempdir()) / "automated_plantoes_sync.lock"
    if not lock_file.exists():
        return False
        
    try:
        with open(lock_file, "r") as f:
            pid = int(f.read().strip())
            
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            exit_code = ctypes.c_ulong()
            if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                kernel32.CloseHandle(handle)
                return exit_code.value == 259
            kernel32.CloseHandle(handle)
    except Exception:
        pass
    return False

def create_plantoes_lock():
    lock_file = Path(tempfile.gettempdir()) / "automated_plantoes_sync.lock"
    try:
        with open(lock_file, "w") as f:
            f.write(str(os.getpid()))
    except Exception as e:
        logger.warning(f"Erro ao criar arquivo de lock: {e}")

def remove_plantoes_lock():
    lock_file = Path(tempfile.gettempdir()) / "automated_plantoes_sync.lock"
    try:
        if lock_file.exists():
            lock_file.unlink()
    except Exception as e:
        logger.warning(f"Erro ao remover arquivo de lock: {e}")

def read_plantoes_last_log_lines(n: int = 15) -> str:
    log_path = DEBUG_DIR_PLANTOES / "plantoes.log"
    if not log_path.exists():
        return "Nenhum log de plantões gerado ainda. Aguardando início..."
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
            return "".join(lines[-n:])
    except Exception as e:
        return f"Erro ao ler log de plantões: {e}"

def parse_data_matutino_string(data_raw: str, ano: int) -> tuple[str, str]:
    if not data_raw or pd.isna(data_raw):
        return "", ""
        
    s = str(data_raw).strip()
    match = re.search(r'([A-Za-zçáéíóúâêôãõ\-]+)\s*-\s*(\d{1,2})\s+de\s+([A-Za-zçáéíóúâêôãõ]+)', s, re.IGNORECASE)
    if match:
        dia_sem = match.group(1).strip().capitalize()
        dia_num = int(match.group(2))
        mes_nome = match.group(3).strip().lower()
        mes_num = MESES_PT.get(mes_nome, 1)
        try:
            dt_obj = datetime(ano, mes_num, dia_num)
            dt_iso = dt_obj.strftime("%Y-%m-%d")
            return dt_iso, dia_sem
        except Exception as ex:
            logger.warning(f"⚠️ Erro ao converter data '{s}' para ano {ano}: {ex}")
            
    return "", s

def parse_simp_periodo(periodo_raw: str, ano: int) -> tuple[str, str]:
    if not periodo_raw or pd.isna(periodo_raw):
        return "", ""
        
    s = str(periodo_raw).strip()
    match = re.findall(r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})', s)
    if len(match) >= 2:
        try:
            dt_ini = datetime.strptime(match[0], "%d/%m/%Y %H:%M").strftime("%Y-%m-%d %H:%M:%S")
            dt_fim = datetime.strptime(match[1], "%d/%m/%Y %H:%M").strftime("%Y-%m-%d %H:%M:%S")
            return dt_ini, dt_fim
        except Exception as ex:
            logger.warning(f"Erro ao converter período SIMP '{s}': {ex}")
            
    return "", ""

def sync_matutino_from_excel(excel_path_or_buffer) -> int:
    from src.database import save_plantoes_matutino
    
    logger.info("🚀 Iniciando parsing do arquivo Excel de Plantão Matutino DIT...")
    try:
        excel_file = pd.ExcelFile(excel_path_or_buffer)
        logger.info(f"📂 Sucesso ao abrir Excel. Abas encontradas: {excel_file.sheet_names}")
    except Exception as e:
        logger.error(f"❌ Erro de formato ao ler planilha Excel de plantão matutino: {e}", exc_info=True)
        return 0

    records = []
    for sheet_name in excel_file.sheet_names:
        sheet_clean = str(sheet_name).strip()
        logger.info(f"📊 Analisando aba: '{sheet_name}'")
        
        if not sheet_clean.isdigit():
            logger.warning(f"⚠️ Aba '{sheet_name}' ignorada pois não é um ano numérico (ex: 2026).")
            continue
            
        ano = int(sheet_clean)
        df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
        logger.info(f"Dimensões da aba {ano}: {df_sheet.shape[0]} linhas x {df_sheet.shape[1]} colunas")
        
        header_idx = None
        for i, row in df_sheet.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)]).upper()
            if "DATA" in row_str and ("SERVIDOR" in row_str or "NOME" in row_str):
                header_idx = i
                logger.info(f"📌 Cabeçalho identificado na linha {i}: {row_str[:80]}")
                break
                
        if header_idx is not None:
            df_sheet.columns = [str(val).strip() for val in df_sheet.iloc[header_idx].values]
            df_data = df_sheet.iloc[header_idx + 1:].copy()
        else:
            df_data = df_sheet.copy()
            
        col_data = next((c for c in df_data.columns if "DATA" in c.upper()), None)
        col_serv = next((c for c in df_data.columns if "SERVIDOR" in c.upper() or "NOME" in c.upper()), None)
        col_tel = next((c for c in df_data.columns if "TELEFONE" in c.upper() or "TEL" in c.upper() or "CONTATO" in c.upper()), None)
        
        logger.info(f"Mapeamento de colunas na aba {ano}: Data='{col_data}', Servidor='{col_serv}', Telefone='{col_tel}'")
        
        if not col_data or not col_serv:
            logger.warning(f"⚠️ Colunas obrigatórias (Data/Servidor) não encontradas na aba {ano}.")
            continue
            
        count_sheet = 0
        for _, r in df_data.iterrows():
            d_raw = r.get(col_data)
            serv = str(r.get(col_serv, '')).strip()
            tel = str(r.get(col_tel, '')).strip() if col_tel else ""
            
            if not d_raw or pd.isna(d_raw) or not serv or serv.lower() in ["nan", "none", "", "sem expediente"]:
                continue
                
            dt_iso, dia_sem = parse_data_matutino_string(str(d_raw), ano)
            
            if dt_iso:
                records.append({
                    "ano": ano,
                    "data_iso": dt_iso,
                    "dia_semana": dia_sem,
                    "servidor": serv,
                    "telefone": tel if tel.lower() not in ["nan", "none"] else ""
                })
                count_sheet += 1
                
        logger.info(f"✅ {count_sheet} plantões matutinos extraídos da aba {ano}.")

    if records:
        save_plantoes_matutino(records)
        logger.info(f"💾 TOTAL: {len(records)} registros de plantão matutino salvos no banco SQLite.")
    else:
        logger.warning("⚠️ Nenhum registro de plantão matutino foi extraído do arquivo.")
        
    return len(records)

def download_sharepoint_matutino_file() -> Path | None:
    url = SHAREPOINT_MATUTINO_URL
    logger.info(f"🌐 Iniciando tentativa de obtenção da planilha via URL do SharePoint DIT: {url}")
    
    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "escala_periodo_matutino.xlsx"
    
    download_url = url + "&download=1" if "?e=" in url else url + "?download=1"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    logger.info(f"📥 Tentando download direto via HTTP GET: {download_url}")
    try:
        resp = requests.get(download_url, headers=headers, timeout=15, allow_redirects=True)
        if resp.status_code == 200 and resp.content.startswith(b'PK'):
            with open(file_path, "wb") as f:
                f.write(resp.content)
            logger.info(f"🎉 SUCESSO NO DOWNLOAD HTTP: Arquivo '.xlsx' salvo ({len(resp.content)} bytes) em '{file_path}'.")
            return file_path
        else:
            logger.info(f"ℹ️ Download HTTP direto retornou página de autenticação ({len(resp.content)} bytes). Iniciando navegador Selenium...")
    except Exception as e_http:
        logger.warning(f"⚠️ Tentativa HTTP direto falhou: {e_http}")

    driver = None
    try:
        logger.info("🌐 Inicializando Selenium Chrome Driver...")
        driver = get_chrome_driver(headless=HEADLESS)
        
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': str(output_dir)}}
        driver.execute("send_command", params)
        
        logger.info(f"Navegando até a URL compartilhada do SharePoint...")
        driver.get(url)
        time.sleep(4)
        
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        
        wait = WebDriverWait(driver, EXPLICIT_WAIT)
        
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
            logger.info("🔗 Sessão autenticada no SharePoint! Extraindo cookies para download direto via HTTP...")
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
            
            # Aguarda a finalização do download do navegador
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
            
        # Fallback defensivo: se restou um .crdownload com bytes de Excel válidos
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
                    logger.warning(f"Erro ao verificar crdownload: {e_cr}")

        if file_path.exists() and file_path.stat().st_size > 0:
            logger.info(f"🎉 SUCESSO NO SELENIUM: Planilha baixada ({file_path.stat().st_size} bytes) em '{file_path}'.")
            return file_path
        else:
            img_path = DEBUG_DIR_PLANTOES / "debug_sharepoint_screen.png"
            driver.save_screenshot(str(img_path))
            logger.warning(f"⚠️ Planilha não baixada automaticamente. Print da tela salvo em: '{img_path}'.")
            
    except Exception as e_sel:
        logger.error(f"❌ Erro durante a navegação via Selenium: {e_sel}", exc_info=True)
    finally:
        if driver:
            try:
                driver.quit()
                logger.info("🔒 Navegador Selenium encerrado com sucesso.")
            except:
                pass
                
    if file_path.exists() and file_path.stat().st_size > 0:
        return file_path
        
    return None

def sync_matutino_from_sharepoint() -> int:
    print_header("WORKER - SINCRONIZAÇÃO DE PLANTÃO MATUTINO", color=CYAN)
    create_plantoes_lock()
    logger.info("🔍 Iniciando processo de sincronização do Plantão Matutino DIT...")
    
    file_path = download_sharepoint_matutino_file()
    
    if not file_path or not file_path.exists():
        backup_path = Path("uploads/escala_periodo_matutino.xlsx")
        template_path = Path("uploads/matutino_TEMPLATE.xlsx")
        
        if backup_path.exists():
            file_path = backup_path
            logger.info(f"📁 Usando cópia local de backup encontrada em: {file_path}")
        elif template_path.exists():
            file_path = template_path
            logger.info(f"📁 Usando modelo template local em: {file_path}")
        else:
            logger.error("❌ Não foi possível baixar a planilha do SharePoint DIT e nenhuma cópia local foi encontrada.")
            remove_plantoes_lock()
            return 0
            
    count = sync_matutino_from_excel(file_path)
    remove_plantoes_lock()
    return count

def scrape_simp_plantoes(ano: int = 2026):
    print_header("SCRAPER PLANTÕES - SIMP STI", color=CYAN)
    if check_plantoes_sync_running():
        logger.warning("⚠️ Sincronização de plantões já em execução em outra instância.")
        return []

    create_plantoes_lock()
    logger.info(f"🤖 === INICIANDO RASPAGEM SIMP PLANTÕES STI (ANO {ano}) ===")
    
    driver = None
    records = []
    
    try:
        user = USERNAME
        pwd = PASSWORD
        logger.info(f"👤 Usuário de rede sAMAccountName: '{user}'")
        
        driver = get_chrome_driver(headless=HEADLESS)
        logger.info("🌐 Selenium Chrome Driver inicializado com sucesso.")
        
        login_url = "https://simp.mpms.mp.br/"
        logger.info(f"🔗 Conectando à página de login: {login_url}")
        driver.get(login_url)
        time.sleep(3)
        
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        
        wait = WebDriverWait(driver, EXPLICIT_WAIT)
        
        try:
            inputs = driver.find_elements(By.XPATH, "//input[@type='text' or @type='email' or @name='username' or @id='username' or @name='user']")
            if inputs and pwd:
                inputs[0].clear()
                inputs[0].send_keys(user)
                logger.info(f"👤 Campo de usuário preenchido com: '{user}'")
                
                pass_fields = driver.find_elements(By.XPATH, "//input[@type='password']")
                if pass_fields:
                    pass_fields[0].clear()
                    pass_fields[0].send_keys(pwd)
                    logger.info("🔑 Campo de senha preenchido.")
                    
                    submits = driver.find_elements(By.XPATH, "//button[@type='submit'] | //input[@type='submit'] | //button[contains(text(),'Entrar')]")
                    if submits:
                        submits[0].click()
                        logger.info("🖱️ Botão de login clicado. Aguardando autenticação...")
                        time.sleep(5)
        except Exception as e_login:
            logger.info(f"ℹ️ Etapa de login concluída ou pré-existente: {e_login}")

        plantao_url = "https://simp.mpms.mp.br/sistemas/Plantao-STI"
        logger.info(f"🔗 Navegando até o módulo Plantão STI: {plantao_url}")
        driver.get(plantao_url)
        time.sleep(4)
        
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            logger.info(f"🖼️ IFRAME DETECTADO! Alternando contexto para o iframe 'frame-sistema-Plantao-STI' ({len(iframes)} iframe(s)...)")
            driver.switch_to.frame(iframes[0])
            time.sleep(4)
        else:
            logger.info("ℹ️ Nenhum iframe direto. Navegando diretamente para a URL interna: https://simp.mpms.mp.br/sistemas/plantao")
            driver.get("https://simp.mpms.mp.br/sistemas/plantao")
            time.sleep(4)
            
        logger.info(f"📍 URL/Contexto ativo no navegador.")
        
        img_iframe = DEBUG_DIR_PLANTOES / "debug_simp_iframe.png"
        driver.save_screenshot(str(img_iframe))
        logger.info(f"📸 Print do interior do IFRAME salvo em: '{img_iframe}'")

        month_buttons = driver.find_elements(By.XPATH, "//div[contains(@class,'p-selectbutton')]//div[@role='button'] | //div[contains(@class,'p-selectbutton')]//div[contains(@class,'p-button')]")
        logger.info(f"🗓️ Total de {len(month_buttons)} botões de meses encontrados DENTRO DO IFRAME!")

        meses_validos = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

        if month_buttons:
            for btn in month_buttons:
                try:
                    mes_nome = btn.text.strip()
                    if not mes_nome:
                        mes_nome = btn.get_attribute("aria-label") or ""
                    mes_nome = mes_nome.strip()
                    
                    if not any(m.lower() in mes_nome.lower() for m in meses_validos):
                        continue
                        
                    logger.info(f"🖱️ Clicando no mês: '{mes_nome}'")
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1.5)
                    
                    rows = driver.find_elements(By.XPATH, "//table[contains(@class,'p-datatable-table')]//tbody/tr | //div[contains(@class,'p-datatable')]//tbody/tr")
                    logger.info(f"📊 Mês '{mes_nome}': {len(rows)} linhas de plantão encontradas.")
                    
                    for r in rows:
                        cols = r.find_elements(By.TAG_NAME, "td")
                        if len(cols) >= 5:
                            periodo_raw = cols[0].text.strip().replace("\n", " ")
                            if not periodo_raw or "Período" in periodo_raw:
                                continue
                                
                            sdesk = cols[1].text.strip().replace("\n", " | ")
                            manut = cols[2].text.strip().replace("\n", " | ")
                            infra = cols[3].text.strip().replace("\n", " | ")
                            dev = cols[4].text.strip().replace("\n", " | ")
                            
                            dt_ini, dt_fim = parse_simp_periodo(periodo_raw, ano)
                            rec = {
                                "ano": ano,
                                "mes": mes_nome,
                                "periodo_str": periodo_raw,
                                "data_inicio": dt_ini,
                                "data_fim": dt_fim,
                                "service_desk": sdesk,
                                "manutencao": manut,
                                "infraestrutura": infra,
                                "desenvolvimento": dev
                            }
                            records.append(rec)
                            logger.info(f"✨ SIMP [{mes_nome}]: {periodo_raw} -> Manutenção: '{manut}'")
                except Exception as ex_m:
                    logger.warning(f"⚠️ Aviso ao processar mês '{mes_nome}': {ex_m}")
        else:
            logger.warning("⚠️ Botões de meses não localizados no iframe. Extraindo linhas da tabela visível...")
            rows = driver.find_elements(By.XPATH, "//table//tbody/tr | //div[contains(@class,'p-datatable')]//tbody/tr")
            for r in rows:
                cols = r.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 5:
                    periodo_raw = cols[0].text.strip().replace("\n", " ")
                    if not periodo_raw or "Período" in periodo_raw:
                        continue
                    dt_ini, dt_fim = parse_simp_periodo(periodo_raw, ano)
                    rec = {
                        "ano": ano,
                        "mes": "",
                        "periodo_str": periodo_raw,
                        "data_inicio": dt_ini,
                        "data_fim": dt_fim,
                        "service_desk": cols[1].text.strip(),
                        "manutencao": cols[2].text.strip(),
                        "infraestrutura": cols[3].text.strip(),
                        "desenvolvimento": cols[4].text.strip()
                    }
                    records.append(rec)
                    logger.info(f"✨ SIMP: {periodo_raw} -> Manutenção: '{rec['manutencao']}'")
                    
        logger.info(f"✅ TOTAL: {len(records)} plantões semanais extraídos com sucesso do SIMP.")
        
    except Exception as ex:
        logger.error(f"❌ Falha durante a raspagem do SIMP: {ex}", exc_info=True)
    finally:
        if driver:
            try:
                driver.quit()
                logger.info("🔒 Navegador fechado com sucesso.")
            except:
                pass
        remove_plantoes_lock()
                
    if records:
        from src.database import save_plantoes_semanal
        save_plantoes_semanal(records)
        logger.info(f"💾 {len(records)} plantões semanais salvos no banco de dados SQLite.")
        
    return records

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scraper & Sincronizador de Plantões STI")
    parser.add_argument("--sync-simp", action="store_true", help="Executa raspagem do SIMP em background")
    parser.add_argument("--sync-matutino", action="store_true", help="Executa sincronização da planilha Matutino DIT")
    parser.add_argument("--ano", type=int, default=2026, help="Ano dos plantões")
    
    args = parser.parse_args()
    
    if args.sync_simp:
        logger.info("⚡ Executando modo CLI: Sincronização SIMP")
        scrape_simp_plantoes(args.ano)
    elif args.sync_matutino:
        logger.info("⚡ Executando modo CLI: Sincronização Matutino DIT")
        sync_matutino_from_sharepoint()
    else:
        logger.info("⚡ Executando sincronização completa de ambos os plantões...")
        sync_matutino_from_sharepoint()
        scrape_simp_plantoes(2026)
