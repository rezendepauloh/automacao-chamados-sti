import os
import re
import io
import sys
import logging
import urllib3
import requests
import pdfplumber
from pathlib import Path
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

# Configura o path do projeto
root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from src.database import save_ramais_to_db
from src.config import USERNAME, PASSWORD, setup_logging, DEBUG_DIR_RAMAIS

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = setup_logging(DEBUG_DIR_RAMAIS / "ramais_scraper.log", __name__)

logging.getLogger('selenium.webdriver.remote.remote_connection').setLevel(logging.WARNING)
logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

def check_ramais_sync_running() -> bool:
    import ctypes
    import tempfile
    lock_file = Path(tempfile.gettempdir()) / "automated_ramais_sync.lock"
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

def read_ramais_last_log_lines(n: int = 15) -> str:
    log_path = DEBUG_DIR_RAMAIS / "ramais_scraper.log"
    if not log_path.exists():
        return "Nenhum log gerado ainda. Aguardando início..."
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
            return "".join(lines[-n:])
    except Exception as e:
        return f"Erro ao ler arquivo de log: {e}"

def is_header_or_footer(line_str: str) -> bool:
    if not line_str or not line_str.strip():
        return True
    s = line_str.strip().lower()
    
    ignore_patterns = [
        "ministério público", "mato grosso do sul", "procuradoria-geral de justiça",
        "assessoria de cerimonial", "ramais", "sumário", "obs: pedimos que qualquer",
        "mudança de gabinete", "departamento de eventos", "cerimonial@mpms.mp.br",
        "última atualização em", "página", "prefixo 3318", "prefixo 3316", "prefixo 3357"
    ]
    for pat in ignore_patterns:
        if pat in s:
            return True
    if re.match(r'^\d{1,3}$', line_str.strip()):
        return True
    return False

def extract_ramais_from_pdf(pdf_bytes: bytes, tipo_ramal: str) -> list[dict]:
    records = []
    current_localidade = "Geral"

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    headers = [str(cell).strip() if cell else "" for cell in table[0]]
                    
                    if any("MEMBRO" in h.upper() for h in headers) or any("GABINETE" in h.upper() for h in headers):
                        for row in table[1:]:
                            if not row:
                                continue
                            row_str_list = [str(c).strip() for c in row if c]
                            if not row_str_list:
                                continue
                            
                            row_dict = {}
                            for idx, h_text in enumerate(headers):
                                if idx < len(row):
                                    val = str(row[idx]).strip() if row[idx] else ""
                                    if val and val != "None":
                                        row_dict[h_text.upper()] = val
                            
                            membro = row_dict.get("MEMBRO", "")
                            pj = row_dict.get("PJ", "")
                            
                            if not membro and len(row_str_list) >= 2:
                                membro = row_str_list[0]
                                
                            nome_base = f"{pj} {membro}".strip() if pj else membro
                            
                            for col_name, val_ramal in row_dict.items():
                                if col_name in ["PJ", "MEMBRO"]:
                                    continue
                                if val_ramal and re.search(r'\d', val_ramal):
                                    records.append({
                                        "localidade": current_localidade,
                                        "setor_nome": f"{nome_base} ({col_name.title()})".strip() if nome_base else col_name.title(),
                                        "telefone_ramal": val_ramal,
                                        "tipo": tipo_ramal
                                    })
                        continue

            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")
            for line in lines:
                line_clean = line.strip()
                if is_header_or_footer(line_clean):
                    continue

                is_title = False
                if any(kw in line_clean.upper() for kw in [
                    "PROMOTORIA DE JUSTIÇA", "PROCURADORIA", "UNIDADE", "SECRETARIA",
                    "DEPARTAMENTO", "GAECO", "OUVIDORIA", "CENTRO DE APOIO", "TÉRREO",
                    "1º ANDAR", "2º ANDAR", "3º ANDAR", "4º ANDAR"
                ]):
                    is_title = True
                elif line_clean.isupper() and not re.search(r'\d{4,}', line_clean) and len(line_clean) > 4:
                    is_title = True

                if is_title:
                    current_localidade = line_clean
                    continue

                phone_match = re.search(r'(\+?55\s*)?\(?\d{2}\)?\s*9?\d{4}[-\s]?\d{4}|\b\d{4}\b|\b\d{4}/\d{4}\b', line_clean)
                if phone_match:
                    parts = re.split(r'(\b\d{4}(?:/\d{4})*\b|\b\d{4}-\d{4}\b|\b33\d{2}-\d{4}\b)', line_clean)
                    if len(parts) >= 2:
                        setor_str = parts[0].strip()
                        ramal_str = "".join(parts[1:]).strip()
                        if not setor_str:
                            setor_str = current_localidade
                        if ramal_str:
                            records.append({
                                "localidade": current_localidade,
                                "setor_nome": setor_str,
                                "telefone_ramal": ramal_str,
                                "tipo": tipo_ramal
                            })

    return records

def run_ramais_scraper():
    load_dotenv()
    
    otrs_user = USERNAME or os.getenv("OTRS_USER")
    otrs_pass = PASSWORD or os.getenv("OTRS_PASS")

    if not otrs_user or not otrs_pass:
        logger.error("Credenciais USERNAME / PASSWORD não encontradas no config.py ou .env!")
        return

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7"
    })

    logger.info("🔑 Efetuando login na Intranet MPMS...")

    login_url = "https://www.mpms.mp.br/intranet/login"
    login_data = {
        "user": otrs_user,
        "password": otrs_pass
    }

    try:
        resp_login = session.post(login_url, data=login_data, verify=False, timeout=15)
        resp_login.raise_for_status()
        logger.info("Login efetuado com sucesso!")
    except Exception as e:
        logger.error(f"Falha ao realizar login na intranet: {e}")

    intranet_url = "https://www.mpms.mp.br/intranet"
    pdf_links = {}

    try:
        resp_intra = session.get(intranet_url, verify=False, timeout=15)
        soup = BeautifulSoup(resp_intra.text, "html.parser")

        for a_tag in soup.find_all("a", href=True):
            text = a_tag.get_text().strip()
            title = a_tag.get("title", "").strip()
            combined_text = f"{text} {title}".lower()

            if "ramais das comarcas do interior" in combined_text or "interior" in combined_text and "ramais" in combined_text:
                pdf_links["Interior"] = a_tag["href"]
            elif "ramais pgj" in combined_text or "campo grande" in combined_text and "ramais" in combined_text:
                pdf_links["Capital / PGJ"] = a_tag["href"]

    except Exception as e:
        logger.error(f"Erro ao buscar links de ramais na página da Intranet: {e}")

    fallback_links = {
        "Interior": "/anexo/MTAxMDYxNDE2ODMyMWQ0MjFjZmEyZTJkODEzYzI5ZDUzZWUzNTllMDJkM2VhLTAyMg",
        "Capital / PGJ": "/anexo/MTAxMDYxNDI3NTAyMWVmZWY2MjA4MjRhN2FiZDY2ZDU5ZTQxMmUyZDVjODJmLTAyMg"
    }

    for k, v in fallback_links.items():
        if k not in pdf_links:
            pdf_links[k] = v

    all_records = []

    for tipo, path_or_url in pdf_links.items():
        full_url = path_or_url if path_or_url.startswith("http") else f"https://www.mpms.mp.br{path_or_url if path_or_url.startswith('/') else '/' + path_or_url}"
        
        logger.info(f"Baixando PDF de ramais ({tipo}): {full_url}")

        try:
            resp_pdf = session.get(full_url, verify=False, timeout=30)
            if resp_pdf.status_code == 200:
                recs = extract_ramais_from_pdf(resp_pdf.content, tipo)
                logger.info(f"Extraídos {len(recs)} registros do PDF [{tipo}]")
                all_records.extend(recs)
            else:
                logger.error(f"HTTP {resp_pdf.status_code} ao baixar {full_url}")
        except Exception as e_pdf:
            logger.error(f"Erro ao baixar/processar PDF ({tipo}): {e_pdf}")

    if all_records:
        df_ramais = pd.DataFrame(all_records)
        df_ramais.drop_duplicates(subset=["localidade", "setor_nome", "telefone_ramal"], inplace=True)
        save_ramais_to_db(df_ramais)
        logger.info(f"Sincronização concluída com {len(df_ramais)} ramais salvos no SQLite!")
    else:
        logger.warning("Nenhum registro de ramal foi extraído dos PDFs!")

if __name__ == "__main__":
    import tempfile
    lock_file = Path(tempfile.gettempdir()) / "automated_ramais_sync.lock"
    with open(lock_file, "w") as f:
        f.write(str(os.getpid()))
    try:
        run_ramais_scraper()
    finally:
        if lock_file.exists():
            try:
                lock_file.unlink()
            except Exception:
                pass
