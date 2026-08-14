#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path

# Adiciona o diretório raiz e o diretório src ao sys.path para suportar importações diretas
root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

import re
import time
import unidecode
import argparse
import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import ctypes
import tempfile
import logging
import urllib3

from src.database import get_unidades_manuais, save_unidades_to_db
from src.config import *

urllib3.disable_warnings()

logger = setup_logging(DEBUG_DIR_UNIDADES / "unidades_scraper.log", __name__)

logging.getLogger('selenium.webdriver.remote.remote_connection').setLevel(logging.WARNING)
logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

def check_unidades_sync_running() -> bool:
    lock_file = Path(tempfile.gettempdir()) / "automated_unidades_sync.lock"
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

def read_unidades_last_log_lines(n: int = 15) -> str:
    log_path = DEBUG_DIR_UNIDADES / "unidades_scraper.log"
    if not log_path.exists():
        return "Nenhum log gerado ainda. Aguardando início..."
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
            return "".join(lines[-n:])
    except Exception as e:
        return f"Erro ao ler arquivo de log: {e}"

BASE_DOMAIN = "https://www.mpms.mp.br"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

def clean_url(url):
    if not url:
        return ""
    if url.startswith("/"):
        return BASE_DOMAIN + url
    return url

def load_manual_entries_from_db():
    df_db = get_unidades_manuais()
    if df_db.empty:
        return []
    
    entries = []
    for _, row in df_db.iterrows():
        cidade = str(row.get('cidade', '')).strip()
        u_predio = str(row.get('unidade_predio', '')).strip()
        unidade_formatada = f"{cidade} - {u_predio}" if cidade and u_predio and not u_predio.startswith(cidade) else u_predio

        entries.append({
            "Cidade": cidade,
            "Tipo": str(row.get('tipo', '')).strip(),
            "Setor": str(row.get('setor', '')).strip(),
            "Sigla": str(row.get('sigla', '')).strip(),
            "Titular": str(row.get('titular', '')).strip(),
            "Unidade (Prédio)": unidade_formatada if unidade_formatada else cidade,
            "Telefone": str(row.get('telefone', '')).strip(),
            "URL": str(row.get('url', '')).strip()
        })
    return entries

def slugify(text: str) -> str:
    s = unidecode.unidecode(text).lower()
    return s.replace(" ", "-")

def get_cities():
    r = requests.get(PROMOTORIAS_URL, headers=HEADERS, verify=False, timeout=15)
    soup = BeautifulSoup(r.text, 'html.parser')
    innerpage = soup.find(class_="innerpage")
    if not innerpage:
        return []
        
    elems = innerpage.find_all("a")
    seen, out = set(), []
    for a in elems:
        text = a.text.strip()
        href = a.get("href") or ""
        if not text or "/promotorias/" not in href:
            continue
        href = clean_url(href)
        slug = href.rstrip("/").split("/")[-1]
        if slug in seen or href.rstrip("/") == PROMOTORIAS_URL.rstrip("/"):
            continue
        seen.add(slug)
        out.append((text, href, slug))
    return out

def get_promotoria_urls(city_url, slug):
    r = requests.get(city_url, headers=HEADERS, verify=False, timeout=15)
    soup = BeautifulSoup(r.text, 'html.parser')
    innerpage = soup.find(class_="innerpage")
    if not innerpage:
        return []
    
    urls = []
    for a in innerpage.find_all("a"):
        href = a.get("href") or ""
        if href and f"/promotorias/{slug}/" in href and clean_url(href).rstrip("/") != city_url.rstrip("/"):
            absolute_href = clean_url(href).rstrip("/")
            if absolute_href not in urls:
                urls.append(absolute_href)
    return sorted(set(urls))

def scrape_promotoria(city_name, promo_url):
    r = requests.get(promo_url, headers=HEADERS, verify=False, timeout=15)
    soup = BeautifulSoup(r.text, 'html.parser')
    
    root = soup.find(id="promotorias")
    if not root:
        return {
            "Cidade": city_name,
            "Tipo": "Promotoria",
            "Setor": "Não encontrada",
            "Titular": "",
            "Unidade (Prédio)": "",
            "Telefone": "",
            "URL": promo_url
        }
    
    nome = root.find("h2").text.strip()

    try:
        titular = ""
        titular_p = root.find(class_="titular")
        if titular_p:
            name_span = titular_p.find(class_="name")
            if name_span:
                titular = name_span.text.replace("Titular:", "").strip()
    except:
        titular = ""

    try:
        address_text = ""
        address_el = root.find("address")
        if address_el:
            address_text = address_el.text.strip()
            
        m = re.search(r"-\s*([^–-]+)\s*-\s*CEP", address_text)
        raw_building = m.group(1).strip() if m else ""
        city_key = slugify(city_name)
        building = raw_building

        if city_key == "campo-grande":
            if "Parque dos Poderes" in address_text or "Jardim Veraneio" in address_text:
                building = f"{city_name} - PGJ"
            elif "Rua da Paz" in address_text:
                building = f"{city_name} - Rua da Paz"
            elif re.search(r"Ch[áa]c[áa]r[áa] Cachoeira", address_text):
                building = f"{city_name} - Chácara Cachoeira"
            elif "Itanhangá Park" in address_text:
                building = f"{city_name} - Ricardo Brandão"
            elif "Jardim Imá" in address_text:
                building = f"{city_name} - Casa da Mulher Brasileira"
        elif city_key == "corumba":
            if "Centro" in address_text:
                building = f"{city_name} - Sede"
            elif "Dom Bosco" in address_text:
                building = f"{city_name} - Fórum"
        else:
            building = f"{city_name} - Sede"
    except:
        building = ""

    try:
        tel = ""
        phone_p = root.find(class_="phone")
        if phone_p:
            tel = phone_p.text.replace("Telefone:", "").strip()
    except:
        tel = ""

    return {
        "Cidade": city_name,
        "Tipo": "Promotoria",
        "Setor": nome,
        "Titular": titular,
        "Unidade (Prédio)": building,
        "Telefone": tel,
        "URL": promo_url
    }

def get_procuradorias():
    r = requests.get(PROCURADORIAS_URL, headers=HEADERS, verify=False, timeout=15)
    soup = BeautifulSoup(r.text, 'html.parser')
    innerpage = soup.find(class_="innerpage")
    if not innerpage:
        return []
    
    links = []
    for a in innerpage.find_all("a"):
        href = a.get("href") or ""
        if "/procuradorias/" in href and clean_url(href).rstrip("/") != PROCURADORIAS_URL.rstrip("/"):
            links.append(clean_url(href).rstrip("/"))
    return sorted(set(links))

def scrape_procuradoria(url):
    r = requests.get(url, headers=HEADERS, verify=False, timeout=15)
    soup = BeautifulSoup(r.text, 'html.parser')
    
    root = soup.find(id="procuradorias")
    if not root:
        return {
            "Cidade": "Campo Grande",
            "Tipo": "Procuradoria",
            "Setor": "Não encontrada",
            "Titular": "",
            "Unidade (Prédio)": "Campo Grande - PGJ",
            "Telefone": "",
            "URL": url
        }
        
    nome = root.find("h2").text.strip()
    try:
        titular = ""
        titular_p = root.find(class_="titular")
        if titular_p:
            name_span = titular_p.find(class_="name")
            if name_span:
                titular = name_span.text.strip()
    except:
        titular = ""
    try:
        tel = ""
        phone_p = root.find(class_="phone")
        if phone_p:
            tel = phone_p.text.replace("Telefone:", "").strip()
    except:
        tel = ""
        
    return {
        "Cidade": "Campo Grande",
        "Tipo": "Procuradoria",
        "Setor": nome,
        "Titular": titular,
        "Unidade (Prédio)": "Campo Grande - PGJ",
        "Telefone": tel,
        "URL": url
    }

def make_sigla(row: pd.Series) -> str:
    tipo     = row["Tipo"]
    city     = row["Cidade"]
    building = row["Unidade (Prédio)"]
    setor    = row["Setor"]
    
    m = re.match(r"^(\d+(?:ª|a|º))", setor)
    ordinal = m.group(1) if m else ""
    if tipo == "Promotoria":
        if city == "Campo Grande":
            code_map = {
                f"{city} - Chácara Cachoeira":         "PJCHA",
                f"{city} - Rua da Paz":                "PJCGR",
                f"{city} - Ricardo Brandão":           "PJESP",
                f"{city} - Casa da Mulher Brasileira": "PJ Casa da Mulher",
                f"{city} - PGJ":                       "PJCGR"
            }
            code = code_map.get(building, "PJ")
            return f"{ordinal} {code}"
        else:
            return f"{ordinal} PJ de {city}"
    elif tipo == "Procuradoria":
        spec = setor.split()[-1]
        return f"{ordinal} PJ {spec}"
    
    return ""

def save_final_excel(df: pd.DataFrame, output_path: Path):
    colunas_desejadas = [
        "Cidade", "Tipo", "Sigla", "Setor",
        "Titular", "Unidade (Prédio)", "Telefone", "URL"
    ]
    cols_to_use = [c for c in colunas_desejadas if c in df.columns]
    df = df.reindex(columns=cols_to_use)

    try:
        save_unidades_to_db(df)
        logger.info("[OK] Dados de unidades atualizados com sucesso no banco de dados SQLite!")
    except Exception as e:
        logger.error(f"Erro ao salvar unidades no SQLite: {e}")

    logger.info(f"Salvando arquivo de legado Excel em: {output_path}...")
    
    widths = {
        'Cidade':20, 'Tipo':15, 'Setor':50, 'Titular':40,
        'Unidade (Prédio)':25, 'Sigla':20, 'Telefone':30, 'URL':50
    }
    save_df_to_excel_formatted(
        df, output_path, sheet_name="Unidades",
        widths=widths, wrap_cols=['Setor','URL'], height_col='Setor'
    )
    
    logger.info("Concluído!")

def main():
    parser = argparse.ArgumentParser(description="Scraper de Unidades do MPMS")
    parser.add_argument(
        "--only-manual", "-m", 
        action="store_true", 
        help="Atualiza apenas as entradas manuais (sem rodar Selenium)."
    )
    args = parser.parse_args()

    out_file = INPUT_DIR_BRUTOS / "Unidades_MPMS.xlsx"

    if args.only_manual:
        logger.info("\n=== MODO RÁPIDO: ATUALIZANDO APENAS ENTRADAS MANUAIS ===")
        
        if not out_file.exists():
            logger.error(f"ERRO: O arquivo {out_file} não existe.")
            logger.error("Execute o script sem parâmetros primeiro para criar a base.")
            return

        try:
            df_existing = pd.read_excel(out_file, sheet_name="Unidades")
        except Exception as e:
            logger.error(f"Erro ao ler o Excel existente: {e}")
            return

        logger.info(f"Lidos {len(df_existing)} registros do arquivo atual.")

        tipos_selenium = ["Promotoria", "Procuradoria"]
        df_web = df_existing[df_existing["Tipo"].isin(tipos_selenium)].copy()
        
        logger.info(f"Mantendo {len(df_web)} registros obtidos via Web (Promotorias/Procuradorias).")

        manual = load_manual_entries_from_db()
        df_manual = pd.DataFrame(manual)
        
        df_final = pd.concat([df_web, df_manual], ignore_index=True)
        
        save_final_excel(df_final, out_file)
        return

    logger.info("\n=== MODO COMPLETO: INICIANDO SCRAPER (WEB REQUESTS EM TEMPO REAL) ===")
    
    all_data = []

    cities = get_cities()
    logger.info(f"Encontradas {len(cities)} comarcas.")
    
    for city, link, slug in cities:
        urls = get_promotoria_urls(link, slug)
        label = "promotoria" if len(urls)==1 else "promotorias"
        logger.info(f"  {city}: {len(urls)} {label}")
        for u in urls:
            all_data.append(scrape_promotoria(city, u))
            logger.info(f"    [OK] {all_data[-1]['Setor']}")
            time.sleep(0.05)
        
    proc_urls = get_procuradorias()
    logger.info(f"\nEncontradas {len(proc_urls)} procuradorias.")
    
    for u in proc_urls:
        all_data.append(scrape_procuradoria(u))
        logger.info(f"    [OK] {all_data[-1]['Setor']}")
        time.sleep(0.05)    

    df = pd.DataFrame(all_data)
    df["Sigla"] = df.apply(make_sigla, axis=1)

    manual = load_manual_entries_from_db()
    df_manual = pd.DataFrame(manual)
    
    df = pd.concat([df, df_manual], ignore_index=True)

    save_final_excel(df, out_file)

if __name__ == "__main__":
    lock_file = Path(tempfile.gettempdir()) / "automated_unidades_sync.lock"
    with open(lock_file, "w") as f:
        f.write(str(os.getpid()))
    try:
        logger.info("Iniciando execução do unidades_scraper.py...")
        main()
        logger.info("Execução do unidades_scraper.py concluída com sucesso!")
    finally:
        if lock_file.exists():
            try:
                lock_file.unlink()
            except Exception:
                pass
