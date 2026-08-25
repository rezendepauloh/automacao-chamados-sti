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
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

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
    print_header("SCRAPER RAMAIS - TELEFONIA MPMS", color=CYAN)
    logger.info("🤖 Iniciando coleta e extração de ramais da Intranet...")
    logger.info("=== INICIANDO SCRAPER DE RAMAIS ===")

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })

    # Tenta autenticação se usuário e senha estiverem disponíveis (silencia 404 se a rota não existir)
    if USERNAME and PASSWORD:
        login_url = "https://www.mpms.mp.br/intranet"
        login_data = {
            "usuario": USERNAME,
            "senha": PASSWORD
        }
        try:
            logger.info("🔑 Tentando autenticação opcional na Intranet...")
            resp_login = session.post(login_url, data=login_data, verify=False, timeout=10)
            if resp_login.status_code == 200:
                logger.info("Sessão autenticada ou mantida com sucesso!")
            else:
                logger.warning(f"Aviso de login (HTTP {resp_login.status_code}). Prosseguindo em modo público...")
        except Exception as e:
            logger.warning(f"Não foi possível autenticar na intranet ({e}). Prosseguindo com download público...")

    intranet_url = "https://www.mpms.mp.br/intranet"
    pdf_links = {}

    try:
        resp_intra = session.get(intranet_url, verify=False, timeout=15)
        if resp_intra.status_code == 200:
            soup = BeautifulSoup(resp_intra.text, "html.parser")
            for a_tag in soup.find_all("a", href=True):
                text = a_tag.get_text().strip()
                title = a_tag.get("title", "").strip()
                combined_text = f"{text} {title}".lower()

                if "lista de ramais das comarcas do interior" in combined_text or "ramais das comarcas do interior" in combined_text or ("interior" in combined_text and "ramais" in combined_text):
                    pdf_links["Interior"] = a_tag["href"]
                elif "lista de ramais pgj" in combined_text or "ramais pgj" in combined_text or ("campo grande" in combined_text and "ramais" in combined_text):
                    pdf_links["Capital / PGJ"] = a_tag["href"]
    except Exception as e:
        logger.warning(f"Aviso ao buscar links dinâmicos de ramais na página da Intranet: {e}")

    db_config = {}
    try:
        from src.database import get_ramais_config
        db_config = get_ramais_config()
    except Exception as e_cfg:
        logger.debug(f"Não foi possível obter ramais_config do banco: {e_cfg}")

    fallback_links = {
        "Interior": (db_config.get("Interior") or "/anexo/MTMzMDYxNDI3NTAwODYzMjkwNDNmYmI5MGYwYjU2ZGE5ZWI5M2ZmN2EwMTQxLTA0MQ").strip(),
        "Capital / PGJ": (db_config.get("Capital / PGJ") or "/anexo/MTMzMDYxNDE2ODMwOGI3MjcxZWQ2YzhkYjYyODkwOGFlMDRjNTUzYWFmY2ZhLTA0MQ").strip()
    }

    for k, v in fallback_links.items():
        if k not in pdf_links and v:
            pdf_links[k] = v

    all_records = []

    for tipo, path_or_url in pdf_links.items():
        full_url = path_or_url if path_or_url.startswith("http") else f"https://www.mpms.mp.br{path_or_url if path_or_url.startswith('/') else '/' + path_or_url}"
        
        logger.info(f"📥 Baixando PDF de ramais ({tipo}): {full_url}")

        try:
            resp_pdf = session.get(full_url, verify=False, timeout=30)
            if resp_pdf.status_code == 200:
                content = resp_pdf.content
                # Validação rigorosa dos bytes do arquivo PDF (%PDF-)
                if content and content.startswith(b"%PDF-"):
                    recs = extract_ramais_from_pdf(content, tipo)
                    logger.info(f"📄 Extraídos {len(recs)} registros do PDF [{tipo}]")
                    all_records.extend(recs)
                else:
                    logger.warning(f"⚠️ O link retornado para {tipo} não é um PDF válido (resposta HTML/redirecionamento retornado).")
            else:
                logger.warning(f"⚠️ HTTP {resp_pdf.status_code} ao baixar {full_url}")
        except Exception as e_pdf:
            logger.warning(f"⚠️ Erro ao baixar/processar PDF ({tipo}): {e_pdf}")

    if all_records:
        df_ramais = pd.DataFrame(all_records)
        df_ramais.drop_duplicates(subset=["localidade", "setor_nome", "telefone_ramal"], inplace=True)
        save_ramais_to_db(df_ramais)
        logger.info(f"✅ Sincronização concluída com {len(df_ramais)} ramais salvos no SQLite!")
    else:
        logger.warning("⚠️ Nenhum registro de ramal pôde ser extraído dos PDFs nesta rodada.")


def process_uploaded_pdf_files(files_dict: dict) -> int:
    """
    Processa arquivos PDF enviados manualmente via upload no Streamlit.
    files_dict: {"Interior": bytes, "Capital / PGJ": bytes}
    Retorna o número total de ramais salvos.
    """
    all_records = []
    for tipo, pdf_bytes in files_dict.items():
        if pdf_bytes and pdf_bytes.startswith(b"%PDF-"):
            recs = extract_ramais_from_pdf(pdf_bytes, tipo)
            logger.info(f"📄 [Upload] Extraídos {len(recs)} registros do PDF [{tipo}]")
            all_records.extend(recs)

    if all_records:
        df_ramais = pd.DataFrame(all_records)
        df_ramais.drop_duplicates(subset=["localidade", "setor_nome", "telefone_ramal"], inplace=True)
        save_ramais_to_db(df_ramais)
        logger.info(f"✅ [Upload] {len(df_ramais)} ramais salvos com sucesso no SQLite!")
        return len(df_ramais)
    return 0

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
