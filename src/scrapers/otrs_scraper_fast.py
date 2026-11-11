import sys
from pathlib import Path

# Adiciona o diretório raiz e o diretório src ao sys.path
root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

import requests
import urllib3
import re
import json
import os
import shutil
import time
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed

from config import (
    OTRS_URL, USERNAME, PASSWORD, DEBUG_DIR_OTRS,
    INPUT_DIR_BRUTOS, BACKUP_PATH_OTRS, TEMP_PATH_OTRS,
    setup_logging, save_df_to_excel_formatted,
    setup_ad_connection, fetch_ad_department,
    fetch_ip_from_sccm, fetch_hostname_from_sccm
)
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HEADERS = [
    'Chamado#', 'Data Criação', 'Título', 'Cidade - Prédio',
    'Unidade', 'Nome do Usuário', 'ID do Cliente', 'Descrição',
    'IP_Origem', 'Hostname', 'Link', 'Comentários'
]

logger = setup_logging(DEBUG_DIR_OTRS / "otrs_scraper_fast.log", __name__)
ad_conn = setup_ad_connection()

def get_timestamp():
    return datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

def fetch_unidade(username):
    try:
        return fetch_ad_department(ad_conn, username, is_username=True)
    except Exception as e:
        logger.warning(f"Erro ao buscar unidade no AD para {username}: {e}")
        return "N/A"

class OTRSClient:
    def __init__(self, base_url=None, username=None, password=None):
        self.base_url = (base_url or OTRS_URL).rstrip("/")
        self.username = username or USERNAME
        self.password = password or PASSWORD
        self.session = requests.Session()
        self.session.verify = False
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7"
        })
        self.authenticated = False

    def login(self) -> bool:
        login_url = f"{self.base_url}/otrs/index.pl"
        payload = {
            "Action": "Login",
            "RequestedURL": "",
            "Lang": "pt_BR",
            "TimeZoneOffset": "-240",
            "User": self.username,
            "Password": self.password or ""
        }
        
        logger.info(f"🔑 Realizando autenticação HTTP para {self.username}...")
        resp = self.session.post(login_url, data=payload, allow_redirects=True, timeout=15)
        
        if resp.status_code == 200 and "Action=Login" not in resp.url and "LoginButton" not in resp.text:
            self.authenticated = True
            logger.info("✅ Login HTTP efetuado com sucesso no OTRS!")
            return True
        else:
            soup = BeautifulSoup(resp.text, "html.parser")
            err_elem = soup.select_one("p.Error, div.ErrorBox, .ErrorMessage")
            err_msg = err_elem.get_text(strip=True) if err_elem else "Credenciais recusadas ou timeout."
            logger.error(f"❌ Falha de login no OTRS: {err_msg}")
            raise RuntimeError(f"Falha de Autenticação no OTRS: {err_msg}")

    def get_ticket_zoom_details(self, zoom_link: str) -> tuple[str, list]:
        """Extrai descrição inicial e comentários de um chamado via HTTP."""
        if not zoom_link:
            return "", []

        full_url = zoom_link if zoom_link.startswith("http") else f"{self.base_url}{zoom_link}"
        try:
            resp = self.session.get(full_url, timeout=12)
            if resp.status_code != 200:
                logger.warning(f"Status {resp.status_code} ao buscar detalhes de {zoom_link}")
                return "", []

            soup = BeautifulSoup(resp.text, "html.parser")
            widgets = soup.select("#ArticleItems div.WidgetSimple")
            if not widgets:
                widgets = soup.select("div.WidgetSimple")

            desc = ""
            comments = []

            for idx, widget in enumerate(widgets):
                data_envio = ""
                autor = "Sistema"
                
                h2 = widget.find("h2")
                if h2:
                    date_span = h2.select_one('span[title*="Criado"], span[title*="Created"]')
                    if date_span:
                        raw_title = date_span.get("title", "")
                        data_envio = raw_title.split(":", 1)[1].strip() if ":" in raw_title else date_span.get_text(strip=True)
                    
                    sender_span = h2.select_one("span.Hidden") or h2.select_one("span:not(.Hidden)")
                    if sender_span:
                        autor = sender_span.get_text(strip=True).replace('"', '')
                    elif "por " in h2.get_text():
                        autor = h2.get_text().split("por ")[-1].strip().replace('"', '')

                texto_nota = ""
                iframe = widget.find("iframe")
                if iframe and iframe.get("src"):
                    src = iframe["src"]
                    iframe_url = src if src.startswith("http") else f"{self.base_url}{src}"
                    try:
                        iframe_resp = self.session.get(iframe_url, timeout=10)
                        if iframe_resp.status_code == 200:
                            soup_iframe = BeautifulSoup(iframe_resp.text, "html.parser")
                            texto_nota = soup_iframe.get_text("\n", strip=True)
                    except Exception as e_if:
                        logger.debug(f"Falha ao carregar iframe {iframe_url}: {e_if}")

                if not texto_nota:
                    content_div = widget.select_one("div.ArticleMailContent, div.Content")
                    if content_div:
                        texto_nota = content_div.get_text("\n", strip=True)

                clean_text = "\n".join(line.strip() for line in texto_nota.splitlines() if line.strip())

                if idx == 0:
                    desc = clean_text
                else:
                    comments.append({
                        "data": data_envio or datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        "autor": autor or "Sistema",
                        "texto": clean_text
                    })

            return desc, comments
        except Exception as e:
            logger.error(f"Erro ao extrair zoom de {zoom_link}: {e}")
            return "", []

    def fetch_queue_tickets(self, max_pages: int = None, cache: dict = None, max_workers: int = 5) -> list[dict]:
        """Navega pela fila do OTRS e extrai os chamados em alta velocidade com concorrência."""
        if not self.authenticated:
            self.login()

        cache = cache or {}
        all_data = []
        page = 1
        queue_url = f"{self.base_url}/otrs/index.pl?Action=AgentTicketQueue"

        while True:
            logger.info(f"📄 Buscando página {page} da fila OTRS...")
            resp = self.session.get(queue_url, timeout=15)
            if resp.status_code != 200:
                logger.error(f"Erro ao acessar fila: HTTP {resp.status_code}")
                break

            soup = BeautifulSoup(resp.text, "html.parser")
            table = soup.select_one("table#OverviewControl, table.TableSmall, table.DataTable")
            if not table:
                logger.warning("Nenhuma tabela de chamados encontrada nesta página.")
                break

            rows = table.select("tbody tr.MasterAction, tbody tr")
            if not rows:
                rows = [tr for tr in table.select("tr") if tr.select("td")]

            logger.info(f"Encontrados {len(rows)} chamados na página {page}.")
            page_items = []

            for row in rows:
                cells = row.find_all("td")
                if len(cells) < 5:
                    continue

                item = {h: "" for h in HEADERS}
                
                # Link e Chamado#
                a_tag = row.find("a", href=lambda h: h and "Action=AgentTicketZoom" in h)
                if a_tag:
                    item['Chamado#'] = a_tag.get_text(strip=True)
                    item['Link'] = a_tag["href"]
                else:
                    continue

                # Data Criação
                if len(cells) > 4:
                    div_c = cells[4].find("div")
                    item['Data Criação'] = div_c.get("title", "").strip() if div_c else cells[4].get_text(strip=True)

                # Título
                if len(cells) > 6:
                    div_t = cells[6].find("div")
                    item['Título'] = div_t.get("title", "").strip() if div_t else cells[6].get_text(strip=True)

                # Cidade - Prédio
                if len(cells) > 8:
                    div_city = cells[8].find("div")
                    item['Cidade - Prédio'] = div_city.get("title", "").strip() if div_city else cells[8].get_text(strip=True)

                # Nome do Usuário
                if len(cells) > 9:
                    div_user = cells[9].find("div")
                    item['Nome do Usuário'] = div_user.get("title", "").strip() if div_user else cells[9].get_text(strip=True)

                # ID do Cliente
                if len(cells) > 10:
                    span_client = cells[10].find("span")
                    client_id = span_client.get("title", "").strip() if span_client else cells[10].get_text(strip=True)
                    item['ID do Cliente'] = client_id
                    if client_id:
                        item['Unidade'] = fetch_unidade(client_id)
                    else:
                        item['Unidade'] = "N/A"

                page_items.append(item)

            # Buscar detalhes (Zoom) ao vivo para todos os chamados da fila
            # (garantindo que novas notas/comentários em chamados abertos nunca sejam perdidos)
            tickets_to_fetch = page_items
            for item in page_items:
                cid = item['Chamado#']
                # Aproveita apenas IP e Hostname do cache caso já tenham sido descobertos antes
                if cid in cache:
                    if cache[cid].get('IP_Origem') and not item.get('IP_Origem'):
                        item['IP_Origem'] = cache[cid]['IP_Origem']
                    if cache[cid].get('Hostname') and not item.get('Hostname'):
                        item['Hostname'] = cache[cid]['Hostname']

            if tickets_to_fetch:
                logger.info(f"⚡ Baixando detalhes de {len(tickets_to_fetch)} chamados em paralelo (workers={max_workers})...")
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_item = {
                        executor.submit(self.get_ticket_zoom_details, it['Link']): it
                        for it in tickets_to_fetch
                    }
                    for future in as_completed(future_to_item):
                        it = future_to_item[future]
                        try:
                            desc, comments = future.result()
                            it['Descrição'] = desc
                            it['Comentários'] = json.dumps(comments, ensure_ascii=False)
                            
                            # Extração de IP da descrição
                            ip_encontrado = ""
                            if desc:
                                ip_match = re.search(r'IP:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', desc)
                                if ip_match:
                                    ip_encontrado = ip_match.group(1)
                                    it['IP_Origem'] = ip_encontrado

                            # SCCM Lookup se necessário
                            client_id = it.get('ID do Cliente')
                            if client_id:
                                if not it.get('IP_Origem'):
                                    it['IP_Origem'] = fetch_ip_from_sccm(client_id) or ""
                                if not it.get('Hostname'):
                                    it['Hostname'] = fetch_hostname_from_sccm(client_id) or ""
                        except Exception as e:
                            logger.error(f"Erro ao processar detalhes do chamado {it.get('Chamado#')}: {e}")

            all_data.extend(page_items)

            if max_pages and page >= max_pages:
                break

            # Verifica próxima página na paginação
            pag_span = soup.select_one("span.Pagination, div.OverviewActions span.Pagination")
            if not pag_span:
                break

            links = pag_span.find_all("a")
            selected = pag_span.select_one("a.Selected")
            if not selected or selected not in links:
                break

            curr_idx = links.index(selected)
            if curr_idx + 1 >= len(links):
                # Última página alcançada
                break

            next_link = links[curr_idx + 1]
            next_href = next_link.get("href")
            if not next_href:
                break

            queue_url = next_href if next_href.startswith("http") else f"{self.base_url}{next_href}"
            page += 1

        return all_data

def brute_data(data):
    df = pd.DataFrame(data, columns=HEADERS).dropna(subset=['Chamado#'], how='all')
    out_dir = INPUT_DIR_BRUTOS
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = get_timestamp()
    file = out_dir / f"Chamados_OTRS_{ts}.xlsx"

    widths = {
        'Chamado#': 15,
        'Data Criação': 20,
        'Título': 40,
        'Cidade - Prédio': 25,
        'Unidade': 40,
        'Nome do Usuário': 25,
        'ID do Cliente': 15,
        'Descrição': 100,
        'IP_Origem': 15,
        'Hostname': 20,
        'Link': 40,
        'Comentários': 50
    }
    df = df.fillna("")
    save_df_to_excel_formatted(
        df, file, sheet_name="Sheet1",
        widths=widths, wrap_cols=['Descrição', 'Comentários'], height_col='Descrição'
    )
    logger.info(f"💾 SUCESSO! Total de {len(df)} chamados salvos em: {file}")
    return df

def scrape_otrs_fast(max_pages: int = None, max_workers: int = 5) -> bool:
    """Função de entrada que substitui o scrape_otrs do Selenium."""
    print_header("SCRAPER OTRS RÁPIDO (HTTP / API-LIKE)", color=CYAN)
    start_time = time.time()
    
    # Carrega cache existente
    cache = {}
    out_dir = INPUT_DIR_BRUTOS
    out_dir.mkdir(parents=True, exist_ok=True)
    existing_files = sorted(out_dir.glob("Chamados_OTRS_*.xlsx"))
    if existing_files:
        latest = existing_files[-1]
        try:
            df_old = pd.read_excel(latest, dtype=str)
            for _, r in df_old.iterrows():
                cid = str(r.get('Chamado#', '')).strip()
                if cid:
                    cache[cid] = {
                        'Descrição': str(r.get('Descrição', '')).strip() if pd.notna(r.get('Descrição')) else '',
                        'Comentários': str(r.get('Comentários', '[]')),
                        'IP_Origem': str(r.get('IP_Origem', '')).strip() if pd.notna(r.get('IP_Origem')) else '',
                        'Hostname': str(r.get('Hostname', '')).strip() if pd.notna(r.get('Hostname')) else ''
                    }
        except Exception as e:
            logger.warning(f"Erro ao carregar cache do Excel: {e}")

    try:
        client = OTRSClient()
        tickets_data = client.fetch_queue_tickets(max_pages=max_pages, cache=cache, max_workers=max_workers)
        
        elapsed = time.time() - start_time
        logger.info(f"✨ Concluído em {elapsed:.2f}s! Total de {len(tickets_data)} chamados coletados.")

        brute_data(tickets_data)
        
        from config import cleanup_old_files
        cleanup_old_files(INPUT_DIR_BRUTOS, "Chamados_OTRS_*.xlsx", keep_count=10)
        
        logger.info("✅ Coleta rápida do OTRS finalizada com SUCESSO!")
        return True
    except Exception as e:
        logger.error(f"❌ Erro na coleta rápida do OTRS: {e}")
        return False

if __name__ == "__main__":
    success = scrape_otrs_fast()
    sys.exit(0 if success else 1)

