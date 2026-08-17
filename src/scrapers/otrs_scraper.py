import sys
from pathlib import Path

# Adiciona o diretório raiz e o diretório src ao sys.path para suportar importações diretas
root_dir = Path(__file__).resolve().parent.parent.parent
src_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException
)
import pandas as pd
import os
import json
import shutil
from datetime import datetime
import logging
from ldap3 import SUBTREE
from config import (
    OTRS_URL, PASSWORD, IMPLICIT_WAIT,
    HEADLESS, EXPLICIT_WAIT, DEBUG_DIR_OTRS,
    DOMINIO_MMC, USERNAME, INPUT_DIR_BRUTOS,
    BACKUP_PATH_OTRS, TEMP_PATH_OTRS, setup_logging, save_df_to_excel_formatted,
    setup_ad_connection, get_chrome_driver, fetch_ad_department, cleanup_old_files
)
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

HEADERS = ['Chamado#', 'Data Criação', 'Título', 'Cidade - Prédio', 'Unidade', 'Nome do Usuário', 'ID do Cliente', 'Descrição', 'IP_Origem', 'Hostname', 'Link', 'Comentários']

logger = setup_logging(DEBUG_DIR_OTRS / "otrs_scraper.log", __name__)


logging.getLogger('selenium.webdriver.remote.remote_connection').setLevel(logging.WARNING)
logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

def get_timestamp():
    return datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

ad_conn = setup_ad_connection()

def fetch_unidade(username):
    return fetch_ad_department(ad_conn, username, is_username=True)

def backup_master():
    if INPUT_DIR_BRUTOS.exists():
        shutil.copy2(INPUT_DIR_BRUTOS, BACKUP_PATH_OTRS)
        logging.debug(f"Backup created at {BACKUP_PATH_OTRS}")

def restore_master():
    if BACKUP_PATH_OTRS.exists():
        shutil.copy2(BACKUP_PATH_OTRS, INPUT_DIR_BRUTOS)
        logging.debug("Master restored from backup.")

def cleanup_backup():
    if BACKUP_PATH_OTRS.exists():
        BACKUP_PATH_OTRS.unlink()
        logging.debug("Backup file removed.")

def write_master(df: pd.DataFrame):
    df.to_excel(TEMP_PATH_OTRS, index=False)
    tmp = pd.read_excel(TEMP_PATH_OTRS, dtype=str)
    if tmp.shape == df.shape:
        os.replace(TEMP_PATH_OTRS, INPUT_DIR_BRUTOS)
        logging.debug(f"Master updated: {INPUT_DIR_BRUTOS}")
    else:
        raise ValueError("Integrity check failed: row/column mismatch.")

def merge_data(new_df: pd.DataFrame) -> pd.DataFrame:
    if not INPUT_DIR_BRUTOS.exists():
        logging.debug("No master found, using new data.")
        return new_df.copy()

    old_df = pd.read_excel(INPUT_DIR_BRUTOS, dtype=str)
    old_ids = set(old_df['Chamado#'])
    new_ids = set(new_df['Chamado#'])

    to_add = new_ids - old_ids
    to_drop = old_ids - new_ids
    to_update = new_ids & old_ids

    df_add = new_df[new_df['Chamado#'].isin(to_add)]
    df_update = new_df[new_df['Chamado#'].isin(to_update)]
    df_keep = old_df[~old_df['Chamado#'].isin(to_drop)]

    df_keep = df_keep[~df_keep['Chamado#'].isin(to_update)]

    merged = pd.concat([df_keep, df_update, df_add], ignore_index=True)
    logging.debug(f"Merged data - added: {len(df_add)}, updated: {len(df_update)}, dropped: {len(to_drop)}")
    return merged

def get_ticket_details(driver):
    desc = ""
    comments = []
    
    try:
        container = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.ID, "ArticleItems"))
        )
        
        widgets = container.find_elements(By.CSS_SELECTOR, "div.WidgetSimple")
        logger.info(f"Detectados {len(widgets)} artigos/notas no chamado zoom.")
        
        for idx, widget in enumerate(widgets):
            data_envio = ""
            autor = "Desconhecido"
            
            try:
                h2 = widget.find_element(By.TAG_NAME, "h2")
                
                try:
                    date_span = h2.find_element(By.CSS_SELECTOR, 'span[title*="Criado"], span[title*="Created"]')
                    raw_title = date_span.get_attribute("title") or ""
                    if ":" in raw_title:
                        data_envio = raw_title.split(":", 1)[1].strip()
                    else:
                        data_envio = date_span.text.strip()
                except:
                    pass
                
                autor = ""
                
                try:
                    sender_span = h2.find_element(By.CSS_SELECTOR, 'span.Hidden')
                    autor = sender_span.get_attribute("textContent").strip()
                except:
                    pass
                    
                if not autor:
                    try:
                        sender_span = h2.find_element(By.CSS_SELECTOR, 'span:not(.Hidden)')
                        autor = sender_span.get_attribute("textContent").strip()
                    except:
                        pass
                        
                if not autor:
                    try:
                        h2_text = h2.get_attribute("textContent")
                        if "por " in h2_text:
                            autor = h2_text.split("por ")[-1].strip()
                    except:
                        pass
                        
                if autor:
                    autor = autor.replace('"', '').strip()
                else:
                    autor = "Sistema"
            except Exception as meta_err:
                logger.warning(f"Erro ao extrair metadados do artigo {idx+1}: {meta_err}")
            
            texto_nota = ""
            try:
                iframes = widget.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    driver.switch_to.frame(iframes[0])
                    body = WebDriverWait(driver, EXPLICIT_WAIT).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    texto_nota = body.text
                    driver.switch_to.default_content()
                else:
                    content_div = widget.find_element(By.CSS_SELECTOR, "div.ArticleMailContent, div.Content")
                    texto_nota = content_div.text
            except Exception as iframe_err:
                driver.switch_to.default_content()
                logger.debug(f"Nota {idx+1} sem iframe padrão ou erro de extração direta: {iframe_err}")
                try:
                    content_div = widget.find_element(By.CSS_SELECTOR, "div.ArticleMailContent, div.Content")
                    texto_nota = content_div.text
                except:
                    pass

            clean_text = "\n".join(line.strip() for line in texto_nota.splitlines() if line.strip())
            
            if idx == 0:
                desc = clean_text
            else:
                comments.append({
                    "data": data_envio or datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    "autor": autor or "Sistema",
                    "texto": clean_text
                })
                
    except Exception as e:
        logger.error(f"Erro geral em get_ticket_details: {e}")
        
    return desc, comments

def process_ticket(driver, row):
    ticket_id = row.get_attribute('id')
    current_url = driver.current_url

    link = WebDriverWait(row, EXPLICIT_WAIT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "td a.MasterActionLink"))
    )
    driver.execute_script("arguments[0].click();", link)

    WebDriverWait(driver, EXPLICIT_WAIT).until(
        lambda d: d.current_url != current_url and EC.presence_of_element_located((By.ID, "ArticleItems"))(d)
    )

    desc, comments = get_ticket_details(driver)

    driver.execute_script("window.history.go(-1);")
    
    try:
        WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.ID, ticket_id))
        )
    except:
        WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.TableSmall"))
        )

    return desc, comments

def check_pagination(driver):
    try:
        pagination_span = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.OverviewActions span.Pagination"))
        )

        page_links = pagination_span.find_elements(By.TAG_NAME, "a")
        
        if len(page_links) > 0:
            logger.info("Passo 8.1: Paginação detectada: Verdadeiro")
            return True
        else:
            logger.info("Passo 8.1: Paginação detectada: Falso (única página)")
            return False
            
    except Exception as e:
        logger.error(f"Passo 8.1: Erro na verificação de paginação: {str(e)}")
        return False

def extract_row_data(driver, row, cache=None):
    data = {h: '' for h in HEADERS}
    try:
        current = WebDriverWait(driver, EXPLICIT_WAIT).until(EC.visibility_of(row))
        cells = current.find_elements(By.TAG_NAME, 'td')

        try:
            a_elem = cells[3].find_element(By.TAG_NAME, 'a')
            data['Chamado#'] = a_elem.text.strip()
            data['Link'] = a_elem.get_attribute('href') or ""
        except Exception as e:
            logger.error(f"Erro Chamado# / Link: {e}")

        try:
            raw_date = cells[4].find_element(By.TAG_NAME, 'div').get_attribute('title')
            data['Data Criação'] = (raw_date or "").strip()
        except Exception as e:
            logger.error(f"Erro Data Criação: {e}")

        try:
            raw_title = cells[6].find_element(By.TAG_NAME, 'div').get_attribute('title')
            data['Título'] = (raw_title or "").strip()
        except Exception as e:
            logger.error(f"Erro Título: {e}")

        try:
            raw_city = cells[8].find_element(By.TAG_NAME, 'div').get_attribute('title')
            data['Cidade - Prédio'] = (raw_city or "").strip()
        except Exception as e:
            logger.error(f"Erro Cidade - Prédio: {e}")

        try:
            raw_user = cells[9].find_element(By.TAG_NAME, 'div').get_attribute('title')
            client_user = (raw_user or "").strip()
            data['Nome do Usuário'] = client_user
        except Exception as e:
            logger.error(f"Erro Nome do Usuário: {e}")

        try:
            raw_client_id = cells[10].find_element(By.TAG_NAME, 'span').get_attribute('title')
            client_id = (raw_client_id or "").strip()
            data['ID do Cliente'] = client_id
            
            if client_id:
                data['Unidade'] = fetch_unidade(client_id)
            else:
                data['Unidade'] = "N/A"
        except Exception as e:
            logger.error(f"Erro ID do Cliente ou lookup AD: {e}")

        try:
            cid = data['Chamado#']
            desc, comments = process_ticket(driver, current)
            data['Descrição'] = desc
            data['Comentários'] = json.dumps(comments, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Erro Descrição e Comentários: {e}")
 
        desc = data.get('Descrição', '')
        ip_encontrado = ""
        hostname_encontrado = ""
        
        if cache and cid in cache:
            if cache[cid].get('IP_Origem'):
                ip_encontrado = cache[cid]['IP_Origem']
                data['IP_Origem'] = ip_encontrado
            if cache[cid].get('Hostname'):
                hostname_encontrado = cache[cid]['Hostname']
                data['Hostname'] = hostname_encontrado
            if ip_encontrado or hostname_encontrado:
                logger.info(f"⚡ [CACHE MATCH] IP/Hostname do chamado {cid} recuperado do cache anterior: IP={ip_encontrado}, Hostname={hostname_encontrado}")
            
        if not ip_encontrado and desc:
            import re
            ip_match = re.search(r'IP:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', desc)
            if ip_match:
                ip_encontrado = ip_match.group(1)
                data['IP_Origem'] = ip_encontrado
                logger.info(f"IP encontrado na descrição do OTRS: {ip_encontrado}")
                
        if (not ip_encontrado or not hostname_encontrado) and data.get('ID do Cliente'):
            from config import fetch_ip_from_sccm, fetch_hostname_from_sccm
            client_id = data['ID do Cliente']
            
            if not ip_encontrado:
                sccm_ip = fetch_ip_from_sccm(client_id)
                if sccm_ip:
                    data['IP_Origem'] = sccm_ip
                    ip_encontrado = sccm_ip
                    
            if not hostname_encontrado:
                sccm_hostname = fetch_hostname_from_sccm(client_id)
                if sccm_hostname:
                    data['Hostname'] = sccm_hostname
                    hostname_encontrado = sccm_hostname
        else:
            if 'Hostname' not in data:
                data['Hostname'] = hostname_encontrado
        
        return data
    
    except StaleElementReferenceException as e:
        raise e
    except Exception as e:
        logger.error(f"Erro geral linha: {e}")
        return data

def process_all_pages(driver, cache=None):
    all_data, page = [], 1
    while True:
        logger.info(f"Página {page}: extraindo dados...")
        table = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'table.TableSmall'))
        )
        
        rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
        total_linhas = len(rows)
        logger.info(f"Linhas detectadas: {total_linhas}")
        
        for idx, row in enumerate(rows):
            try:
                data = extract_row_data(driver, row, cache=cache)
                all_data.append(data)
                
                chamado_num = data.get('Chamado#', 'N/A')
                logger.info(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")
            
            except StaleElementReferenceException:
                logger.error("Linha obsoleta, tentando novamente")
                table = WebDriverWait(driver, EXPLICIT_WAIT).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'table.TableSmall'))
                )
                rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
                data = extract_row_data(driver, rows[idx], cache=cache)
                all_data.append(data)
                
                chamado_num = data.get('Chamado#', 'N/A')
                logger.info(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")

        try:
            pag = driver.find_element(By.CSS_SELECTOR, 'span.Pagination')
            links = pag.find_elements(By.TAG_NAME, 'a')
            selected = pag.find_element(By.CSS_SELECTOR, 'a.Selected')
            ordered = links
            idx = ordered.index(selected)
            if idx + 1 >= len(ordered):
                break
            next_link = ordered[idx + 1]
            logger.info(f"Indo para página {next_link.text}")
            driver.execute_script("arguments[0].click();", next_link)
            WebDriverWait(driver, EXPLICIT_WAIT).until(EC.staleness_of(table))
            page += 1
        
        except Exception:
            logger.error("Fim da paginação ou erro ao avançar")
            break

    return all_data

def login_page(driver):
    logger.info("Passo 1: Carregando página de login...")
    driver.get(OTRS_URL)

    logger.info("Passo 2: Preenchendo credenciais...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.ID, "User"))
    ).send_keys(USERNAME)
    
    driver.find_element(By.ID, "Password").send_keys(PASSWORD)

    logger.info("Passo 3: Clicando no botão de login...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.element_to_be_clickable((By.ID, "LoginButton"))
    ).click()

def navigation_queue(driver):
    logger.info("Passo 4: Navegando para fila principal...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'Action=AgentTicketQueue')]"))
    )
    
    logger.info("Passo 5: Clicando no link da fila...")
    queue_link = driver.find_element(
        By.XPATH, "//a[contains(@href, 'Action=AgentTicketQueue')]")
    queue_link.click()    

def all_chamados(driver):
    logger.info("Passo 6: Acessando todos os chamados...")
    all_tickets_link = WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@href, 'QueueID=0') and contains(@href, 'Filter=All')]"))
    )
    
    driver.execute_script("arguments[0].scrollIntoView(true);", all_tickets_link)
    driver.execute_script("arguments[0].click();", all_tickets_link)

    logger.info("Passo 7: Validando carregamento...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located(
            (By.XPATH, "//li[@class='Active ']//a[contains(., 'Todos os Chamados')]")
        )
    )    

def pagination_or_not(driver):
    try:
        container = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'span.Pagination'))
        )
        links = container.find_elements(By.TAG_NAME, 'a')
        
        return len(links) > 1
    
    except Exception as e:
        logger.error(f"Erro na verificação de paginação: {e}")
        return False

def data_extract(driver, has_pagination, cache=None):
    logger.info("Passo 9: Extraindo dados da tabela...")
    
    if has_pagination:
        return process_all_pages(driver, cache=cache)
    
    else:
        all_data = []
        table = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'table.TableSmall'))
        )
        rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
        
        total_linhas = len(rows)
        logger.info(f"Linhas detectadas: {total_linhas}")
        
        for idx, row in enumerate(rows):
            try:
                data = extract_row_data(driver, row, cache=cache)
                all_data.append(data)
                
                chamado_num = data.get('Chamado#', 'N/A')
                logger.info(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")
                
            except StaleElementReferenceException:
                logger.error("Linha obsoleta, tentando novamente")
                table = WebDriverWait(driver, EXPLICIT_WAIT).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'table.TableSmall'))
                )
                rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
                data = extract_row_data(driver, rows[idx], cache=cache)
                all_data.append(data)
                
                chamado_num = data.get('Chamado#', 'N/A')
                logger.info(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")
                
        return all_data

def brute_data(data):
    df = pd.DataFrame(data, columns=HEADERS).dropna(subset=['Chamado#'], how='all')
    out_dir = Path("01 - Dados Brutos")
    out_dir.mkdir(exist_ok=True)
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
        'Link': 40,
        'Comentários': 50
    }
    save_df_to_excel_formatted(
        df, file, sheet_name="Sheet1",
        widths=widths, wrap_cols=['Descrição', 'Comentários'], height_col='Descrição'
    )

    logger.info(f"SUCESSO! Total de {len(df)} chamados salvos em: {file}")

def scrape_otrs():
    print_header("SCRAPER OTRS - COLETA DE CHAMADOS", color=CYAN)
    logger.info("🤖 Iniciando processo de raspagem do OTRS...")
    cache = {}
    try:
        out_dir = Path("01 - Dados Brutos")
        existing_files = sorted(out_dir.glob("Chamados_OTRS_*.xlsx"))
        if existing_files:
            latest_file = existing_files[-1]
            logger.info(f"📂 Carregando cache de descrições, IPs e comentários do arquivo mais recente: {latest_file.name}")
            df_old = pd.read_excel(latest_file, dtype=str)
            for _, row_old in df_old.iterrows():
                cid = str(row_old.get('Chamado#', '')).strip()
                desc = row_old.get('Descrição', '')
                ip = row_old.get('IP_Origem', '')
                link = row_old.get('Link', '')
                comments = row_old.get('Comentários', '[]')
                if cid:
                    cache[cid] = {
                        'Descrição': str(desc).strip() if pd.notna(desc) else '',
                        'IP_Origem': str(ip).strip() if pd.notna(ip) else '',
                        'Link': str(link).strip() if pd.notna(link) else '',
                        'Comentários': str(comments).strip() if pd.notna(comments) else '[]'
                    }
            logger.info(f"⚡ Sucesso! {len(cache)} descrições, IPs e comentários carregados no cache de memória do OTRS.")
    except Exception as cache_err:
        logger.warning(f"⚠️ Aviso: Não foi possível carregar cache do OTRS: {cache_err}")

    driver = None
    try:
        driver = get_chrome_driver(headless=HEADLESS, block_media=True)
        driver.implicitly_wait(IMPLICIT_WAIT)

        logger.info("🔑 Realizando login e navegando nas filas...")
        login_page(driver)
        navigation_queue(driver)
        all_chamados(driver)
        
        has_pagination = pagination_or_not(driver)

        logger.info("📥 Extraindo dados da tabela...")
        data = data_extract(driver, has_pagination, cache=cache)
        
        logger.info("💾 Processando dados e salvando planilha...")
        brute_data(data)
        
        cleanup_old_files(INPUT_DIR_BRUTOS, "Chamados_OTRS_*.xlsx", keep_count=10)
        
        logger.info("✅ Coleta do OTRS finalizada com SUCESSO!")
        return True

    except Exception as e:
        timestamp = get_timestamp()
        error_dir = DEBUG_DIR_OTRS / f"erros"
        error_dir.mkdir(exist_ok=True)
        
        logger.error(f"❌ ERRO CRÍTICO no OTRS: {str(e)}")
        print(f"ERRO CRÍTICO: {str(e)}")
        if driver:
            driver.save_screenshot(str(error_dir / f'erro_final_{timestamp}.png'))
            with open(error_dir / f'pagina_final_{timestamp}.html', 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
        return False
        
    finally:
        if driver:
            try:
                driver.quit()
            except Exception as quit_error:
                logger.warning(f"⚠️ Erro ao fechar driver: {str(quit_error)}")

if __name__ == "__main__":
    if scrape_otrs():
        sys.exit(0)
    else:
        sys.exit(1)
