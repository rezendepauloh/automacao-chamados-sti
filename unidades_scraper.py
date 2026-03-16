#!/usr/bin/env python3
# promotorias_procuradorias_scraper.py

# Para rodar apenas manuais, use:
# python .\unidades_scraper.py --only-manual ou python .\unidades_scraper.py -m

# Para rodar Selenium + manuais, use:
# python .\unidades_scraper.py

import re
import time
import unidecode # type: ignore
import argparse  # [NOVO] Para ler os comandos do terminal
import pandas as pd
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from manual_entries import get_manual_entries, set_city_into_unidade
from config import *
from typing import cast
from xlsxwriter.workbook import Workbook as XlsxWorkbook

# ----------------------------------------
# SLUG
# ----------------------------------------
def slugify(text: str) -> str:
    """Remove acentos, coloca lowercase e troca espaços por hífen."""
    s = unidecode.unidecode(text).lower()
    return s.replace(" ", "-")

# ----------------------------------------
# DRIVER
# ----------------------------------------
def init_driver(headless=True):
    opts = webdriver.ChromeOptions() # type: ignore
    opts.add_argument("--disable-infobars")
    opts.add_argument("--disable-extensions")
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(options=opts) # type: ignore

# ----------------------------------------
# SCRAPE PROMOTORIAS
# ----------------------------------------
def get_cities(driver):
    driver.get(PROMOTORIAS_URL)
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.innerpage"))
    )
    elems = driver.find_elements(By.CSS_SELECTOR, "div.innerpage a")
    seen, out = set(), []
    for a in elems:
        text = a.text.strip()
        href = a.get_attribute("href") or ""
        if not text or "/promotorias/" not in href:
            continue
        slug = href.rstrip("/").split("/")[-1]
        if slug in seen or href.rstrip("/") == PROMOTORIAS_URL.rstrip("/"):
            continue
        seen.add(slug)
        out.append((text, href.rstrip("/"), slug))
    return out

def get_promotoria_urls(driver, city_url, slug):
    driver.get(city_url)
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.innerpage"))
    )
    urls = []
    for a in driver.find_elements(By.CSS_SELECTOR, "div.innerpage a"):
        try:
            href = a.get_attribute("href")
        except StaleElementReferenceException:
            continue
        if href and f"/promotorias/{slug}/" in href and href.rstrip("/") != city_url:
            if href not in urls:
                urls.append(href)
    return sorted(set(urls))

def scrape_promotoria(driver, city_name, promo_url):
    driver.get(promo_url)
    wait = WebDriverWait(driver, EXPLICIT_WAIT)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div#promotorias")))

    root = driver.find_element(By.CSS_SELECTOR, "div#promotorias")
    nome = root.find_element(By.TAG_NAME, "h2").text.strip()

    try:
        titular = root.find_element(By.CSS_SELECTOR, "p.titular span.name").text.replace("Titular:", "").strip()
    except:
        titular = ""

    try:
        address_el = driver.find_element(By.CSS_SELECTOR, "#promotorias address")
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
        tel = root.find_element(By.CSS_SELECTOR, "p.phone").text.replace("Telefone:", "").strip()
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

# ----------------------------------------
# SCRAPE PROCURADORIAS
# ----------------------------------------
def get_procuradorias(driver):
    driver.get(PROCURADORIAS_URL)
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.innerpage"))
    )
    links = []
    for a in driver.find_elements(By.CSS_SELECTOR, "div.innerpage a"):
        href = a.get_attribute("href") or ""
        if "/procuradorias/" in href and href.rstrip("/") != PROCURADORIAS_URL.rstrip("/"):
            links.append(href.rstrip("/"))
    return sorted(set(links))

def scrape_procuradoria(driver, url):
    driver.get(url)
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div#procuradorias"))
    )
    root = driver.find_element(By.CSS_SELECTOR, "div#procuradorias")
    nome = root.find_element(By.TAG_NAME, "h2").text.strip()
    try:
        titular = root.find_element(By.CSS_SELECTOR, "p.titular span.name").text.strip()
    except:
        titular = ""
    try:
        tel = root.find_element(By.CSS_SELECTOR, "p.phone").text.replace("Telefone:", "").strip()
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

# ----------------------------------------
# CALCULA SIGLA
# ----------------------------------------
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

# ----------------------------------------
# FUNÇÃO AUXILIAR: SALVAR EXCEL (Para não repetir código)
# ----------------------------------------
def save_final_excel(df: pd.DataFrame, output_path: Path):
    """
    Recebe o DataFrame pronto e salva com a formatação correta.
    """
    # Reordenar as colunas para garantir consistência
    colunas_desejadas = [
        "Cidade", "Tipo", "Sigla", "Setor",
        "Titular", "Unidade (Prédio)", "Telefone", "URL"
    ]
    # Filtra apenas colunas que existem no DF (para evitar erros se algo faltar)
    cols_to_use = [c for c in colunas_desejadas if c in df.columns]
    df = df.reindex(columns=cols_to_use)

    print(f"Salvando arquivo em: {output_path}...")
    
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Unidades", index=False)
        wb = cast(XlsxWorkbook, writer.book) # type: ignore
        wrap = wb.add_format({'text_wrap': True}) 
        ws  = writer.sheets['Unidades']
        
        widths = {
            'Cidade':20, 'Tipo':15, 'Setor':50, 'Titular':40,
            'Unidade (Prédio)':25, 'Sigla':20, 'Telefone':30, 'URL':50
        }
        
        for i, col in enumerate(df.columns):
            fmt = wrap if col in ('Setor','URL') else None
            ws.set_column(i, i, widths.get(col,20), fmt)
        
        # ajusta altura
        for r, cell in enumerate(df['Setor'], start=1):
            # Garante que cell é string antes de contar
            cell_str = str(cell) if pd.notna(cell) else ""
            lines = cell_str.count("\n")+1
            ws.set_row(r, 15*lines)
    
    print("Concluído!")

# ----------------------------------------
# FLUXO PRINCIPAL
# ----------------------------------------
def main():
    # 1. Configura os argumentos de linha de comando
    parser = argparse.ArgumentParser(description="Scraper de Unidades do MPMS")
    parser.add_argument(
        "--only-manual", "-m", 
        action="store_true", 
        help="Atualiza apenas as entradas manuais (sem rodar Selenium)."
    )
    args = parser.parse_args()

    # Caminho do arquivo de saída
    out_file = INPUT_DIR_BRUTOS / "Unidades_MPMS.xlsx"

    # =========================================================
    # MODO 1: ATUALIZAÇÃO RÁPIDA (SÓ MANUAIS)
    # =========================================================
    if args.only_manual:
        print("\n=== MODO RÁPIDO: ATUALIZANDO APENAS ENTRADAS MANUAIS ===")
        
        if not out_file.exists():
            print(f"ERRO: O arquivo {out_file} não existe.")
            print("Execute o script sem parâmetros primeiro para criar a base.")
            return

        # Carrega o Excel existente
        try:
            df_existing = pd.read_excel(out_file, sheet_name="Unidades")
        except Exception as e:
            print(f"Erro ao ler o Excel existente: {e}")
            return

        print(f"Lidos {len(df_existing)} registros do arquivo atual.")

        # Filtra para manter APENAS o que veio do Selenium
        # Lógica: O Selenium traz apenas 'Promotoria' e 'Procuradoria'.
        # Tudo que for diferente disso assumimos que é manual e vamos substituir.
        tipos_selenium = ["Promotoria", "Procuradoria"]
        df_web = df_existing[df_existing["Tipo"].isin(tipos_selenium)].copy()
        
        print(f"Mantendo {len(df_web)} registros obtidos via Web (Promotorias/Procuradorias).")

        # Carrega as novas entradas manuais
        manual = get_manual_entries()
        manual = set_city_into_unidade(manual)
        df_manual = pd.DataFrame(manual)
        
        # Junta Web Antigo + Manual Novo
        df_final = pd.concat([df_web, df_manual], ignore_index=True)
        
        # Salva
        save_final_excel(df_final, out_file)
        return

    # =========================================================
    # MODO 2: COMPLETO (SELENIUM + MANUAIS)
    # =========================================================
    print("\n=== MODO COMPLETO: INICIANDO SCRAPER (WEB) ===")
    
    driver = init_driver(headless=True)
    all_data = []

    # 1) Promotorias
    cities = get_cities(driver)
    print(f"Encontradas {len(cities)} comarcas.")
    
    for city, link, slug in cities:
        urls = get_promotoria_urls(driver, link, slug)
        label = "promotoria" if len(urls)==1 else "promotorias"
        print(f"  {city}: {len(urls)} {label}")
        for u in urls:
            all_data.append(scrape_promotoria(driver, city, u))
            print(f"    ✓ {all_data[-1]['Setor']}")
            time.sleep(0.3)
        
    # 2) Procuradorias
    proc_urls = get_procuradorias(driver)
    print(f"\nEncontradas {len(proc_urls)} procuradorias.")
    
    for u in proc_urls:
        all_data.append(scrape_procuradoria(driver, u))
        print(f"    ✓ {all_data[-1]['Setor']}")
        time.sleep(0.3)

    driver.quit()    

    # 3) Processa dados da Web
    df = pd.DataFrame(all_data)
    # Calcula sigla para os dados da Web
    df["Sigla"] = df.apply(make_sigla, axis=1)

    # 4) Adiciona manuais
    manual = get_manual_entries()
    manual = set_city_into_unidade(manual)
    df_manual = pd.DataFrame(manual)
    
    # Unifica
    df = pd.concat([df, df_manual], ignore_index=True)

    # Salva
    save_final_excel(df, out_file)

if __name__ == "__main__":
    main()