#!/usr/bin/env python3
# promotorias_procuradorias_scraper.py

# Para rodar apenas manuais, use:
# python .\unidades_scraper.py --only-manual ou python .\unidades_scraper.py -m

# Para rodar a extração completa em tempo real + manuais, use:
# python .\unidades_scraper.py

import re
import time
import unidecode
import argparse
import pandas as pd
from pathlib import Path
import requests
from bs4 import BeautifulSoup
import urllib3
from manual_entries import get_manual_entries, set_city_into_unidade
from config import *

urllib3.disable_warnings()

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

# ----------------------------------------
# SLUG
# ----------------------------------------
def slugify(text: str) -> str:
    """Remove acentos, coloca lowercase e troca espaços por hífen."""
    s = unidecode.unidecode(text).lower()
    return s.replace(" ", "-")

# ----------------------------------------
# SCRAPE PROMOTORIAS
# ----------------------------------------
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

# ----------------------------------------
# SCRAPE PROCURADORIAS
# ----------------------------------------
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
    
    widths = {
        'Cidade':20, 'Tipo':15, 'Setor':50, 'Titular':40,
        'Unidade (Prédio)':25, 'Sigla':20, 'Telefone':30, 'URL':50
    }
    save_df_to_excel_formatted(
        df, output_path, sheet_name="Unidades",
        widths=widths, wrap_cols=['Setor','URL'], height_col='Setor'
    )
    
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
    # MODO 2: COMPLETO (REAL-TIME VIA REQUESTS)
    # =========================================================
    print("\n=== MODO COMPLETO: INICIANDO SCRAPER (WEB REQUESTS EM TEMPO REAL) ===")
    
    all_data = []

    # 1) Promotorias
    cities = get_cities()
    print(f"Encontradas {len(cities)} comarcas.")
    
    for city, link, slug in cities:
        urls = get_promotoria_urls(link, slug)
        label = "promotoria" if len(urls)==1 else "promotorias"
        print(f"  {city}: {len(urls)} {label}")
        for u in urls:
            all_data.append(scrape_promotoria(city, u))
            print(f"    [OK] {all_data[-1]['Setor']}")
            time.sleep(0.05)
        
    # 2) Procuradorias
    proc_urls = get_procuradorias()
    print(f"\nEncontradas {len(proc_urls)} procuradorias.")
    
    for u in proc_urls:
        all_data.append(scrape_procuradoria(u))
        print(f"    [OK] {all_data[-1]['Setor']}")
        time.sleep(0.05)    

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