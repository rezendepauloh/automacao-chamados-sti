from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException
)
from config import *
import pandas as pd
import time
import sys
import os
import shutil
from pathlib import Path
from datetime import datetime
from typing import cast
from xlsxwriter.workbook import Workbook as XlsxWorkbook # Alias para não confundir
import logging
from ldap3 import Server, Connection, ALL, SUBTREE, ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES

# Configurações atualizadas de cabeçalhos incluindo a coluna Unidade
HEADERS = ['Chamado#', 'Data Criação', 'Título', 'Cidade - Prédio', 'Unidade', 'Nome do Usuário', 'ID do Cliente', 'Descrição']

# --- Configuração de logging ---
logging.basicConfig(
    filename=str(DEBUG_DIR_OTRS / "otrs_debug.log"),
    filemode="a",
    level=logging.DEBUG,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def get_timestamp():
    return datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

def debug_print(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[DEBUG {ts}] {msg}")
    logging.debug(msg)

# --- Configuração do AD ---
ad_server = Server(DOMINIO, get_info=ALL)
ad_conn = Connection(
    ad_server,
    user=f"MPE\\{USERNAME}",
    password=PASSWORD,
    auto_bind=True
)

# ---------------------------
# AD (Active Directory) - Versão Robusta
# ---------------------------
def fetch_unidade(username):
    """
    Busca a unidade do usuário no AD usando o sAMAccountName.
    Prioridade: department -> physicalDeliveryOfficeName -> Cadastro Incompleto.
    """
    # Proteção básica
    if not username or not ad_conn:
        return ''

    try:
        search_filter = f'(sAMAccountName={username})'
        
        # Agora pedimos os dois atributos na consulta
        target_attrs = ['department', 'physicalDeliveryOfficeName']

        ad_conn.search(
            search_base='DC=in,DC=mpe,DC=ms,DC=gov,DC=br',
            search_filter=search_filter,
            search_scope=SUBTREE,
            attributes=target_attrs
        )

        if not ad_conn.entries:
            return 'Não encontrado no AD'
        
        # Pega a primeira entrada e converte para dicionário (mais seguro)
        entry = ad_conn.entries[0].entry_attributes_as_dict
        
        # 1. TENTATIVA PRINCIPAL: DEPARTAMENTO
        # O LDAP retorna uma lista, pegamos o primeiro item ou string vazia
        dept_list = entry.get('department', [])
        dept = dept_list[0] if dept_list else None
        
        if dept and str(dept).strip():
            return str(dept).strip()
        
        # 2. TENTATIVA SECUNDÁRIA: ESCRITÓRIO (OFFICE)
        office_list = entry.get('physicalDeliveryOfficeName', [])
        office = office_list[0] if office_list else None
        
        if office and str(office).strip():
            return str(office).strip()

        # 3. FALHA
        return 'Cadastro Incompleto (AD)'
    
    except Exception as e:
        # Usa debug_print para manter padrão do arquivo, ou logging se preferir
        msg = f"Erro AD lookup para {username}: {e}"
        # Tenta usar o debug_print definido no arquivo, senão printa normal
        try:
            debug_print(msg)
        except NameError:
            print(msg)
            
        return ''

def backup_master():
    """Create a backup of the existing master file."""
    if INPUT_DIR_BRUTOS.exists():
        shutil.copy2(INPUT_DIR_BRUTOS, BACKUP_PATH_OTRS)
        logging.debug(f"Backup created at {BACKUP_PATH_OTRS}")

def restore_master():
    """Restore from backup if needed."""
    if BACKUP_PATH_OTRS.exists():
        shutil.copy2(BACKUP_PATH_OTRS, INPUT_DIR_BRUTOS)
        logging.debug("Master restored from backup.")

def cleanup_backup():
    """Remove the backup file after successful run."""
    if BACKUP_PATH_OTRS.exists():
        BACKUP_PATH_OTRS.unlink()
        logging.debug("Backup file removed.")

def write_master(df: pd.DataFrame):
    """Write merged DataFrame atomically to master file."""
    df.to_excel(TEMP_PATH_OTRS, index=False)
    # Verify integrity
    tmp = pd.read_excel(TEMP_PATH_OTRS, dtype=str)
    if tmp.shape == df.shape:
        os.replace(TEMP_PATH_OTRS, INPUT_DIR_BRUTOS)
        logging.debug(f"Master updated: {INPUT_DIR_BRUTOS}")
    else:
        raise ValueError("Integrity check failed: row/column mismatch.")

def merge_data(new_df: pd.DataFrame) -> pd.DataFrame:
    """Merge new scraped data with existing master, doing add/update/delete."""
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

    # Remove outdated rows before concatenation
    df_keep = df_keep[~df_keep['Chamado#'].isin(to_update)]

    merged = pd.concat([df_keep, df_update, df_add], ignore_index=True)
    logging.debug(f"Merged data - added: {len(df_add)}, updated: {len(df_update)}, dropped: {len(to_drop)}")
    return merged

def get_ticket_description(driver):
    """
    Extrai apenas o corpo da primeira nota (#1) do container #ArticleItems,
    descartando todas as notas posteriores e cabeçalhos estáticos.
    """
    try:
        # 1) Espera o container principal aparecer
        container = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.ID, "ArticleItems"))
        )

        # 2) Busca todos os iframes dentro dele e pega o primeiro (nota #1)
        iframes = container.find_elements(By.TAG_NAME, "iframe")
        if not iframes:
            raw = container.text
        else:
            driver.switch_to.frame(iframes[0])
            body = WebDriverWait(driver, EXPLICIT_WAIT).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            raw = body.text

        # 3) Remove linhas em branco e faz strip em cada linha
        clean = "\n".join(line.strip() for line in raw.splitlines() if line.strip())
        return clean

    except Exception as e:
        debug_print(f"Erro em get_ticket_description: {type(e).__name__}: {e}")
        return ""

    finally:
        # sempre volta pro contexto principal
        driver.switch_to.default_content()

# Processa um ticket individualmente
def process_ticket(driver, row):
    ticket_id = row.get_attribute('id')
    current_url = driver.current_url

    # abre o chamado
    link = WebDriverWait(row, EXPLICIT_WAIT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "td a.MasterActionLink"))
    )
    driver.execute_script("arguments[0].click();", link)

    WebDriverWait(driver, EXPLICIT_WAIT).until(
        lambda d: d.current_url != current_url and EC.presence_of_element_located((By.ID, "ArticleItems"))(d)
    )

    desc = get_ticket_description(driver)

    # volta para lista
    driver.execute_script("window.history.go(-1);")
    
    try:
        WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.ID, ticket_id))
        )
    
    except:
        WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.TableSmall"))
        )

    return desc

def check_pagination(driver):
    """Verifica se existe paginação de resultados"""
    try:
        pagination_span = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.OverviewActions span.Pagination"))
        )

        # Verificar se existem links de paginação
        page_links = pagination_span.find_elements(By.TAG_NAME, "a")
        
        if len(page_links) > 0:
            debug_print("Passo 8.1: Paginação detectada: Verdadeiro")
            return True
        else:
            debug_print("Passo 8.1: Paginação detectada: Falso (única página)")
            return False
            
    except Exception as e:
        debug_print(f"Passo 8.1: Erro na verificação de paginação: {str(e)}")
        return False

# Extração por linha agora inclui Unidade via AD lookup
def extract_row_data(driver, row):
    data = {h: '' for h in HEADERS}
    try:
        # Garante que a linha está visível
        current = WebDriverWait(driver, EXPLICIT_WAIT).until(EC.visibility_of(row))
        cells = current.find_elements(By.TAG_NAME, 'td')

        # --- Chamado# (Cell 3) ---
        try:
            # O .text retorna string vazia se não tiver nada, então é mais seguro que get_attribute
            data['Chamado#'] = cells[3].find_element(By.TAG_NAME, 'a').text.strip()
        except Exception as e:
            debug_print(f"Erro Chamado#: {e}")

        # --- Data Criação (Cell 4) ---
        try:
            raw_date = cells[4].find_element(By.TAG_NAME, 'div').get_attribute('title')
            data['Data Criação'] = (raw_date or "").strip()
        except Exception as e:
            debug_print(f"Erro Data Criação: {e}")

        # --- Título (Cell 6) ---
        try:
            raw_title = cells[6].find_element(By.TAG_NAME, 'div').get_attribute('title')
            data['Título'] = (raw_title or "").strip()
        except Exception as e:
            debug_print(f"Erro Título: {e}")

        # --- Cidade - Prédio (Cell 8) ---
        try:
            raw_city = cells[8].find_element(By.TAG_NAME, 'div').get_attribute('title')
            data['Cidade - Prédio'] = (raw_city or "").strip()
        except Exception as e:
            debug_print(f"Erro Cidade - Prédio: {e}")

        # --- Nome do Usuário (Cell 9) ---
        try:
            raw_user = cells[9].find_element(By.TAG_NAME, 'div').get_attribute('title')
            client_user = (raw_user or "").strip()
            data['Nome do Usuário'] = client_user
        except Exception as e:
            debug_print(f"Erro Nome do Usuário: {e}")

        # --- ID do Cliente e Unidade (Cell 10) ---
        try:
            raw_client_id = cells[10].find_element(By.TAG_NAME, 'span').get_attribute('title')
            client_id = (raw_client_id or "").strip()
            data['ID do Cliente'] = client_id
            
            # Lookup Unidade no AD
            # (Assumindo que fetch_unidade lida bem com string vazia, senão adicione um if)
            if client_id:
                data['Unidade'] = fetch_unidade(client_id)
            else:
                data['Unidade'] = "N/A"

        except Exception as e:
            debug_print(f"Erro ID do Cliente ou lookup AD: {e}")

        # --- Descrição (Processo separado) ---
        try:
            # Aqui passamos o driver e a linha atual (current)
            data['Descrição'] = process_ticket(driver, current)
        except Exception as e:
            debug_print(f"Erro Descrição: {e}")

        # debug_print(f"Dados extraídos: {data['Chamado#']}")
        
        return data
    
    except Exception as e:
        debug_print(f"Erro geral linha: {e}")
        return data

def process_all_pages(driver):
    all_data, page = [], 1
    while True:
        debug_print(f"Página {page}: extraindo dados...")
        table = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'table.TableSmall'))
        )
        
        rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
        total_linhas = len(rows)
        debug_print(f"Linhas detectadas: {total_linhas}")
        
        for idx, row in enumerate(rows):
            try:
                data = extract_row_data(driver, row)
                all_data.append(data)
                
                # Pega o número do chamado do dicionário retornado
                chamado_num = data.get('Chamado#', 'N/A')
                debug_print(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")
            
            except StaleElementReferenceException:
                debug_print("Linha obsoleta, tentando novamente")
                # Atualiza a lista de elementos da página
                rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
                data = extract_row_data(driver, rows[idx])
                all_data.append(data)
                
                chamado_num = data.get('Chamado#', 'N/A')
                debug_print(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")

        # tentar próxima página via links
        try:
            pag = driver.find_element(By.CSS_SELECTOR, 'span.Pagination')
            links = pag.find_elements(By.TAG_NAME, 'a')
            selected = pag.find_element(By.CSS_SELECTOR, 'a.Selected')
            ordered = links
            idx = ordered.index(selected)
            if idx + 1 >= len(ordered):
                break
            next_link = ordered[idx + 1]
            debug_print(f"Indo para página {next_link.text}")
            driver.execute_script("arguments[0].click();", next_link)
            WebDriverWait(driver, EXPLICIT_WAIT).until(EC.staleness_of(table))
            page += 1
        
        except Exception:
            debug_print("Fim da paginação ou erro ao avançar")
            break

    return all_data

# ========== CONFIGURAÇÃO INICIAL ========== #
def initial_config():
    options = webdriver.ChromeOptions()

    # Desativa logs desnecessários
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"]) # Oculta o "DevTools listening..."

    # --- SILENCIADORES DO CHROME ---
    options.add_argument('--log-level=3')         # Silencia logs internos (mostra apenas erros fatais)
    options.add_argument('--disable-logging')     # Desabilita o motor de log do navegador
    # -------------------------------
    
    # Configurações para evitar pop-ups
    options.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    })
    options.add_argument("--incognito")
    options.add_argument("--disable-infobars")
    options.add_argument("--no-default-browser-check")
    
    if not HEADLESS:
        options.add_argument("--start-maximized")
    else:
        options.add_argument("--headless=new")  # Modo headless moderno
        options.add_argument("--window-size=1920,1080")

    # Bloqueio de imagens/CSS (opcional)
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.fonts": 2,
    }
    options.add_experimental_option("prefs", prefs)

    # Versão alternativa para Chrome antigo
    # options.add_argument("--disable-gpu")
    # options.add_argument("--no-sandbox")

    return options

# ========== ETAPA 1: LOGIN ========== #
def login_page(driver):
    debug_print("Passo 1: Carregando página de login...")
    driver.get(OTRS_URL)
    # time.sleep(2)  # Espera inicial anti-bot

    debug_print("Passo 2: Preenchendo credenciais...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.ID, "User"))
    ).send_keys(USERNAME)
    
    driver.find_element(By.ID, "Password").send_keys(PASSWORD)

    debug_print("Passo 3: Clicando no botão de login...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.element_to_be_clickable((By.ID, "LoginButton"))
    ).click()

# ========== ETAPA 2: NAVEGAÇÃO PARA FILA ========== #
def navigation_queue(driver):
    debug_print("Passo 4: Navegando para fila principal...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'Action=AgentTicketQueue')]"))
    )
    
    debug_print("Passo 5: Clicando no link da fila...")
    queue_link = driver.find_element(
        By.XPATH, "//a[contains(@href, 'Action=AgentTicketQueue')]")
    queue_link.click()    

# ========== ETAPA 3: TODOS OS CHAMADOS ========== #
def all_chamados(driver):
    debug_print("Passo 6: Acessando todos os chamados...")
    all_tickets_link = WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@href, 'QueueID=0') and contains(@href, 'Filter=All')]"))
    )
    
    # Clique robusto com JavaScript
    driver.execute_script("arguments[0].scrollIntoView(true);", all_tickets_link)
    driver.execute_script("arguments[0].click();", all_tickets_link)

    # ========== VERIFICAÇÃO FINAL ========== #
    debug_print("Passo 7: Validando carregamento...")
    WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located(
            (By.XPATH, "//li[@class='Active ']//a[contains(., 'Todos os Chamados')]")
        )
    )    

# ========== ETAPA 4: VERIFICAÇÃO DE PAGINAÇÃO ========== #
def pagination_or_not(driver):
    try:
        container = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'span.Pagination'))
        )
        links = container.find_elements(By.TAG_NAME, 'a')
        
        return len(links) > 1
    
    except Exception as e:
        debug_print(f"Erro na verificação de paginação: {e}")
        return False

# ========== ETAPA 5: EXTRAÇÃO DE DADOS ========== #
def data_extract(driver, has_pagination):
    debug_print("Passo 9: Extraindo dados da tabela...")
    
    # Se tem mais de uma página, vai para a função que arrumamos antes
    if has_pagination:
        return process_all_pages(driver)
    
    # Se tem só UMA página, ele cai aqui:
    else:
        all_data = []
        table = WebDriverWait(driver, EXPLICIT_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'table.TableSmall'))
        )
        rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
        
        total_linhas = len(rows)
        debug_print(f"Linhas detectadas: {total_linhas}")
        
        for idx, row in enumerate(rows):
            try:
                data = extract_row_data(driver, row)
                all_data.append(data)
                
                # Print no formato [1/21] Lido: 46444521
                chamado_num = data.get('Chamado#', 'N/A')
                debug_print(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")
                
            except StaleElementReferenceException:
                debug_print("Linha obsoleta, tentando novamente")
                rows = table.find_elements(By.CSS_SELECTOR, 'tr.MasterAction')
                data = extract_row_data(driver, rows[idx])
                all_data.append(data)
                
                chamado_num = data.get('Chamado#', 'N/A')
                debug_print(f"[{idx + 1}/{total_linhas}] Lido: {chamado_num}")
                
        return all_data

# ========== ETAPA 6: SALVANDO DADOS (DADOS BRUTOS) ========== #
def brute_data(data):
    df = pd.DataFrame(data, columns=HEADERS).dropna(subset=['Chamado#'], how='all')
    out_dir = Path("01 - Dados Brutos")
    out_dir.mkdir(exist_ok=True)
    ts = get_timestamp()
    file = out_dir / f"Chamados_OTRS_{ts}.xlsx"

    with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
        wb = cast(XlsxWorkbook, writer.book)
        wrap = wb.add_format({'text_wrap': True})
        ws = writer.sheets['Sheet1']

        widths = {
            'Chamado#': 15,
            'Data Criação': 20,
            'Título': 40,
            'Cidade - Prédio': 25,
            'Unidade': 40,
            'Nome do Usuário': 25,
            'ID do Cliente': 15,
            'Descrição': 100
        }
        for i, col in enumerate(df.columns):
            fmt = wrap if col=='Descrição' else None
            ws.set_column(i, i, widths.get(col,20), fmt)
        
        for r, desc in enumerate(df['Descrição'], start=1):
            text = '' if pd.isna(desc) else str(desc)
            ws.set_row(r, 15 * (text.count('\n')+1))

    debug_print(f"Passo 10.1: Dados finais salvos em: {out_dir}")
    debug_print(f"Passo 10.2: Total de registros extraídos: {len(df)}")  

def scrape_otrs():
    driver = None
    try:
        # ========== CONFIGURAÇÃO INICIAL ========== #
        options = initial_config()

        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(IMPLICIT_WAIT)

        # Login e navegação
        login_page(driver)
        navigation_queue(driver)
        all_chamados(driver)
        
        # Verifica paginação
        has_pagination = pagination_or_not(driver)

        # Extrai dados
        data = data_extract(driver, has_pagination)
        
        # Salvando em "01 - Dados Brutos"
        brute_data(data)
        
        return True  # Indica sucesso

    except Exception as e:
        timestamp = get_timestamp()
        error_dir = DEBUG_DIR_OTRS / f"erros"
        error_dir.mkdir(exist_ok=True)
        
        print(f"ERRO CRÍTICO: {str(e)}")
        if driver:
            driver.save_screenshot(str(error_dir / f'erro_final_{timestamp}.png'))
            with open(error_dir / f'pagina_final_{timestamp}.html', 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
        return False # Indica falha
        
    finally:
        if driver:
            try:
                driver.quit()
            except Exception as quit_error:
                print(f"AVISO: Erro ao fechar driver - {str(quit_error)}")

if __name__ == "__main__":
    if scrape_otrs():
        sys.exit(0)  # Saída limpa
    else:
        sys.exit(1)  # Saída com erro