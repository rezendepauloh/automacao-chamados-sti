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
    HEADLESS, EXPLICIT_WAIT, USERNAME, SIMP_URL, cleanup_old_files
)
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

logger = setup_logging(DEBUG_DIR_PLANTOES / "plantoes.log", "plantoes")

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

def download_sharepoint_matutino_file(url: str = None) -> Path | None:
    from src.config import get_spreadsheet_setting
    from src.scrapers.sharepoint_downloader import download_sharepoint_file

    if not url:
        url = get_spreadsheet_setting("SHAREPOINT_MATUTINO_URL", SHAREPOINT_MATUTINO_URL)

    output_dir = Path("uploads").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = output_dir / "escala_periodo_matutino.xlsx"
    return download_sharepoint_file(url, file_path)

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
    
    records = []
    
    try:
        user = USERNAME
        pwd = PASSWORD
        logger.info(f"👤 Usuário de rede sAMAccountName: '{user}'")
        
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=HEADLESS,
                args=["--no-sandbox", "--disable-dev-shm-usage", "--ignore-certificate-errors"]
            )
            context = browser.new_context(ignore_https_errors=True)
            page = context.new_page()

            simp_base = (SIMP_URL or "").rstrip('/')
            login_url = f"{simp_base}/"
            logger.info(f"🔗 Conectando à página de login: {login_url}")
            page.goto(login_url, wait_until="domcontentloaded", timeout=60000)
            time.sleep(2)

            try:
                user_input = page.locator("xpath=//input[@type='text' or @type='email' or @name='username' or @id='username' or @name='user']").first
                if user_input.count() > 0 and pwd:
                    user_input.fill(user)
                    logger.info(f"👤 Campo de usuário preenchido com: '{user}'")
                    
                    pass_input = page.locator("xpath=//input[@type='password']").first
                    if pass_input.count() > 0:
                        pass_input.fill(pwd)
                        logger.info("🔑 Campo de senha preenchido.")
                        
                        submit_btn = page.locator("xpath=//button[@type='submit'] | //input[@type='submit'] | //button[contains(text(),'Entrar')]").first
                        if submit_btn.count() > 0:
                            submit_btn.click()
                            logger.info("🖱️ Botão de login clicado. Aguardando autenticação...")
                            time.sleep(5)
            except Exception as e_login:
                logger.info(f"ℹ️ Etapa de login concluída ou pré-existente: {e_login}")

            plantao_url = f"{simp_base}/sistemas/Plantao-STI"
            logger.info(f"🔗 Navegando até o módulo Plantão STI: {plantao_url}")
            page.goto(plantao_url, wait_until="domcontentloaded", timeout=60000)
            time.sleep(4)
            
            iframe_el = page.locator("iframe").first
            frame = page
            if iframe_el.count() > 0:
                logger.info("🖼️ IFRAME DETECTADO! Alternando contexto para o iframe...")
                frame = page.frame_locator("iframe").first
            else:
                fallback_plantao = f"{simp_base}/sistemas/plantao"
                logger.info(f"ℹ️ Nenhum iframe direto. Navegando para URL interna: {fallback_plantao}")
                page.goto(fallback_plantao, wait_until="domcontentloaded", timeout=60000)
                time.sleep(4)

            month_buttons = frame.locator("xpath=//div[contains(@class,'p-selectbutton')]//div[@role='button'] | //div[contains(@class,'p-selectbutton')]//div[contains(@class,'p-button')]")
            btn_count = month_buttons.count()
            logger.info(f"🗓️ Total de {btn_count} botões de meses encontrados!")

            meses_validos = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

            if btn_count > 0:
                for i in range(btn_count):
                    try:
                        btn = month_buttons.nth(i)
                        mes_nome = (btn.text_content() or btn.get_attribute("aria-label") or "").strip()
                        if not any(m.lower() in mes_nome.lower() for m in meses_validos):
                            continue
                            
                        logger.info(f"🖱️ Clicando no mês: '{mes_nome}'")
                        btn.click()
                        time.sleep(1.5)
                        
                        rows = frame.locator("xpath=//table[contains(@class,'p-datatable-table')]//tbody/tr | //div[contains(@class,'p-datatable')]//tbody/tr")
                        r_count = rows.count()
                        logger.info(f"📊 Mês '{mes_nome}': {r_count} linhas de plantão encontradas.")
                        
                        for r_idx in range(r_count):
                            row = rows.nth(r_idx)
                            cols = row.locator("td")
                            if cols.count() >= 5:
                                periodo_raw = (cols.nth(0).text_content() or "").strip().replace("\n", " ")
                                if not periodo_raw or "Período" in periodo_raw:
                                    continue
                                    
                                sdesk = (cols.nth(1).text_content() or "").strip().replace("\n", " | ")
                                manut = (cols.nth(2).text_content() or "").strip().replace("\n", " | ")
                                infra = (cols.nth(3).text_content() or "").strip().replace("\n", " | ")
                                dev = (cols.nth(4).text_content() or "").strip().replace("\n", " | ")
                                
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
                    except Exception as ex_m:
                        logger.warning(f"⚠️ Aviso ao processar mês: {ex_m}")
            browser.close()
            logger.info("🔒 Navegador fechado com sucesso.")

        logger.info(f"✅ TOTAL: {len(records)} plantões semanais extraídos com sucesso do SIMP.")
        
    except Exception as ex:
        logger.error(f"❌ Falha durante a raspagem do SIMP: {ex}", exc_info=True)
    finally:
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
