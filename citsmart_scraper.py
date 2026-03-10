# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime
import time
import pandas as pd
import re
import logging
import sys
from pathlib import Path
from ldap3 import Server, Connection, ALL, SUBTREE
from typing import cast
from xlsxwriter.workbook import Workbook as XlsxWorkbook # Alias para não confundir
from config import (
    CITSMART_URL, CITSMART_EMAIL, PASSWORD,
    HEADLESS, EXPLICIT_WAIT, DEBUG_DIR_CITSMART,
    DOMINIO, USERNAME
)

# ---------------------------
# Utilitários e Log
# ---------------------------
logging.basicConfig(
    level=logging.DEBUG,
    # Adicionamos os colchetes [] e removemos os milissegundos do datefmt para ficar limpo
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(DEBUG_DIR_CITSMART / "citsmart_scraper.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# --- O SEGREDO ESTÁ AQUI: Cala a boca do Selenium e do urllib3 ---
logging.getLogger('selenium.webdriver.remote.remote_connection').setLevel(logging.WARNING)
logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)
# -----------------------------------------------------------------

def salvar_screenshot(driver, nome_etapa):
    """Tira um print da tela para vermos exatamente o que o Selenium vê."""
    ts = datetime.now().strftime("%H-%M-%S")
    filename = f"debug_citsmart_{ts}_{nome_etapa}.png"
    filepath = DEBUG_DIR_CITSMART / filename
    driver.save_screenshot(str(filepath))
    logger.debug(f"📸 Screenshot salvo: {filename}")

def inspecionar_elemento(driver, seletor, nome_elemento):
    """Extrai informações vitais do elemento antes de tentar clicar."""
    logger.debug(f"🔍 Inspecionando: {nome_elemento} {seletor}")
    try:
        elementos = driver.find_elements(*seletor)
        if not elementos:
            logger.debug(f"❌ O elemento {nome_elemento} não existe no DOM no momento.")
            return None
        
        el = elementos[0]
        html_trecho = el.get_attribute('outerHTML')[:150] # Pega os primeiros 150 caracteres
        
        logger.debug(f"Status de {nome_elemento}:")
        logger.debug(f" - Visível na tela? {el.is_displayed()}")
        logger.debug(f" - Habilitado para clique? {el.is_enabled()}")
        logger.debug(f" - HTML encontrado: {html_trecho}...")
        return el
    except Exception as e:
        logger.error(f"⚠️ Erro ao inspecionar {nome_elemento}: {e}")
        return None

def find_all(ctx, candidates, timeout=5):
    """
    Retorna a primeira lista de elementos encontrada dentre os candidatos.
    """
    for by, sel in candidates:
        try:
            WebDriverWait(ctx, timeout).until(EC.presence_of_element_located((by, sel)))
            els = ctx.find_elements(by, sel)
            if els:
                return els
        except Exception:
            continue
    return []

# ---------------------------
# AD (Active Directory) - Versão Robusta
# ---------------------------
def setup_ad_connection():
    """Tenta conectar no AD. Se falhar, avisa no log mas não quebra o robô."""
    try:
        # Se o SSL estiver dando erro 10054, deixamos use_ssl=False temporariamente
        server = Server(DOMINIO, get_info=ALL) 
        conn = Connection(server, user=f"MPE\\{USERNAME}", password=PASSWORD, auto_bind=True)
        return conn
    except Exception as e:
        logger.error(f"⚠️ Aviso: Não foi possível conectar ao AD. Erro: {e}")
        return None

def fetch_setor_temp(conn, display_name):
    if not display_name or not conn:
        return 'Sem Departamento'
        
    try:
        target_attrs = ['department', 'physicalDeliveryOfficeName']
        
        # Tenta várias formas de encontrar o usuário
        search_filters = [
            f'(displayName={display_name})',
            f'(cn={display_name})',
            f'(name={display_name})',
            f'(displayName=*{display_name}*)'
        ]
        
        entry = None
        for filt in search_filters:
            conn.search(
                search_base='DC=in,DC=mpe,DC=ms,DC=gov,DC=br',
                search_filter=filt,
                search_scope=SUBTREE,
                attributes=target_attrs
            )
            if conn.entries:
                entry = conn.entries[0].entry_attributes_as_dict
                break # Encontrou, para de procurar
        
        if not entry:
            return 'Não encontrado no AD'

        # Tenta Departamento, se vazio tenta Escritório (Office)
        dept = entry.get('department', [''])[0]
        if dept and str(dept).strip():
            return str(dept).strip()
        
        office = entry.get('physicalDeliveryOfficeName', [''])[0]
        if office and str(office).strip():
            return str(office).strip()
        
        return 'Cadastro Incompleto (AD)'
        
    except Exception as e:
        logger.error(f"Erro lookup AD para '{display_name}': {e}")
        return 'Erro na Consulta'

# ---------------------------
# Navegador / Login
# ---------------------------
def initial_config():
    opts = webdriver.ChromeOptions() # type: ignore
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"]) # Oculta o "DevTools listening..."
    opts.add_argument("--disable-blink-features=CSSAnimations,ScrollAnimator")
    opts.add_argument("--incognito")

    # --- SILENCIADORES DO CHROME ---
    opts.add_argument('--log-level=3')         # Silencia logs internos (mostra apenas erros fatais)
    opts.add_argument('--disable-logging')     # Desabilita o motor de log do navegador
    # -------------------------------


    if HEADLESS:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
    else:
        opts.add_argument("--start-maximized")
    opts.page_load_strategy = "eager"
    driver = webdriver.Chrome(options=opts) # type: ignore
    wait = WebDriverWait(driver, timeout=EXPLICIT_WAIT, poll_frequency=0.1)
    return driver, wait

def navigate_to_caixa_entrada(driver, wait):
    logger.info("Acessando CitSmart e fazendo login…")
    driver.get(CITSMART_URL)

    # 1) E-mail
    email = wait.until(EC.element_to_be_clickable((By.NAME, "loginfmt")))
    driver.execute_script("arguments[0].click()", email)
    time.sleep(0.5)
    email.clear(); email.send_keys(CITSMART_EMAIL)
    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    # 2) Senha
    pwd = wait.until(EC.element_to_be_clickable((By.NAME, "passwd")))
    driver.execute_script("arguments[0].click()", pwd)
    time.sleep(0.5)
    pwd.clear(); pwd.send_keys(PASSWORD)
    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    # 3) KMSI
    try:
        wait.until(EC.presence_of_element_located((By.ID, "KmsiCheckboxField")))
        logger.info("Pulando KMSI de manter conectado…")
        wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    except TimeoutException:
        pass

    # 4) Redirecionamento Direto para LowCode
    logger.info("Aguardando carregamento do portal inicial...")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(5) 

    nova_fila = "https://suporte.mpms.mp.br/inbox/lowcode/form/copilot_novo/default"
    logger.info(f"Forçando navegação para: {nova_fila}")
    
    driver.switch_to.default_content()
    driver.execute_script(f"window.location.href = '{nova_fila}';")

    try:
        wait.until(EC.url_contains("copilot_novo"))
        logger.info("URL de destino alcançada.")
        
        logger.info("Aguardando o iframe 'App'...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[title='App']")))
        
        wait.until(EC.presence_of_element_located((By.ID, "pageSize")))
        logger.info("Sucesso! Interface do Copilot detectada via seletor de paginação.")
        
    except Exception as e:
        logger.info(f"Não detectou o elemento interno: {e}")
        driver.save_screenshot(f"{DEBUG_DIR_CITSMART}/erro_iframe_app_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.png")
        raise

# ---------------------------
# Manipulação da Tabela e Paginação
# ---------------------------
def expand_all_records_lowcode(driver, wait):
    """
    Tenta localizar o pager (id='pageSize') e setar para 100 itens.
    Usa o loader específico (.hyper-loading) para sincronizar.
    """
    logger.info("Tentando expandir registros (LowCode)...")
    # salvar_screenshot(driver, "1_inicio_expansao")
    
    # Seletor do GIF de carregamento
    loader_loc = (By.CSS_SELECTOR, "div.hyper-loading")

    # Função auxiliar para esperar o loader sumir
    def wait_loader_vanish(timeout=30):
        try:
            # Espera até que o elemento fique invisível (display: none)
            WebDriverWait(driver, timeout).until(
                EC.invisibility_of_element_located(loader_loc)
            )
        except TimeoutException:
            logger.info("Aviso: O loader demorou muito para sumir ou não apareceu.")

    try:
        time.sleep(3) 
        
        # 1. ANTES DE TUDO: Garante que a página está "quieta"
        wait_loader_vanish()
        # salvar_screenshot(driver, "2_pos_primeiro_loader")

        logger.info("Procurando o dropdown de itens por página...")
        
        # 2. Espera o select específico (novo ID) aparecer na tela
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "pageSize"))
        )
        
        # Usa o Select do Selenium para interagir com ele
        dropdown = Select(dropdown_element)
        
        # Verifica se já está em 100
        try:
            current = dropdown.first_selected_option.text.strip()
        except:
            current = ""

        if "100" in current:
            logger.info("Já está exibindo 100 registros.")
        else:
            # 3. APLICA A MUDANÇA
            dropdown.select_by_visible_text("100")
            logger.info("Sucesso! Paginação alterada para 100 itens.")
            
            # Tira foto exatamente após o clique
            # salvar_screenshot(driver, "3_apos_selecionar_100")
            
            # Dá 1 segundo para o sistema injetar o loader na tela
            time.sleep(1) 
            
            # Agora esperamos ele SUMIR de verdade
            logger.info("Aguardando o loader (.hyper-loading) desaparecer...")
            wait_loader_vanish(timeout=30)
            
            # Tira foto do resultado final da tabela
            # salvar_screenshot(driver, "4_resultado_final")

        # 4. Extrai a contagem final para garantir que atualizou (usando o NOVO HTML)
        try:
            # Busca pela div específica do AngularJS que contém "Mostrando 1–17 de 17"
            pager_info = driver.find_element(By.CSS_SELECTOR, "div[ng-if='totalTickets']")
            text = pager_info.text.strip()  
            logger.info(f"Paginação atualizada: {text}")

            # Usa Regex para capturar o número total que vem depois da palavra "de"
            match = re.search(r"de\s+(\d+)", text)
            if match:
                return int(match.group(1))
        except Exception as e:
            logger.error(f"Aviso: Não consegui ler o texto da paginação. Erro: {e}")
            
        return 0

    except TimeoutException:
        logger.error("Aviso: Dropdown 'pageSize' não encontrado a tempo. Seguindo com a página atual.")
        # salvar_screenshot(driver, "ERRO_timeout")
        return 0
    except Exception as e:
        logger.error(f"Erro ao expandir registros: {e}")
        # salvar_screenshot(driver, "ERRO_excecao")
        return 0

def _list_rows(driver):
    try:
        return driver.find_elements(By.CSS_SELECTOR, "#table tbody tr")
    except:
        return []

def process_page(driver, wait, filtro_grupo=None, ad_conn=None):
    # Nota: Já estamos no iframe correto, não precisa de switch_to_incidents
    rows = _list_rows(driver)
    
    # Se não achou linhas, espera um pouco e tenta de novo (carregamento lento)
    if not rows:
        logger.info("Nenhuma linha encontrada na tabela. Aguardando 3s...")
        time.sleep(3)
        rows = _list_rows(driver)

    logger.info(f"Linhas detectadas: {len(rows)}")
    collected = []

    for idx, row in enumerate(rows):
        try:
            # Função auxiliar para pegar texto de colunas ng-switch
            def get_val(key, is_description=False):
                try:
                    xpath = f".//div[@ng-switch-when='{key}']"
                    element = row.find_element(By.XPATH, xpath)
                    if is_description:
                        # Descrição costuma estar num title de span
                        return element.find_element(By.TAG_NAME, "span").get_attribute("title") or element.text.strip()
                    return element.get_attribute("textContent").strip()
                except:
                    return ""

            # --- Extração ---
            # Chave 1: Número (limpamos ícones com Regex)
            num_bruto = get_val("1")
            num_match = re.search(r'\d+', num_bruto)
            cid = num_match.group(0) if num_match else ""

            if not cid: continue # Pula linhas inválidas

            # Chaves mapeadas do seu HTML
            solicitante_full = get_val("6")
            data_criacao = get_val("9")
            descricao = get_val("10", is_description=True)

            # --- Enriquecimento AD ---
            localidade = "Não encontrada no AD"
            solicitante_nome = solicitante_full
            
            # Tenta limpar "Nome (login)" para buscar só pelo nome ou login
            login_busca = solicitante_full
            if "(" in solicitante_full:
                try:
                    partes = solicitante_full.split("(")
                    solicitante_nome = partes[0].strip()
                    # Pega o que está dentro dos parênteses como login
                    login_busca = partes[1].replace(")", "").strip()
                except:
                    pass
            
            if ad_conn:
                # Tenta buscar pelo login extraído OU pelo nome completo
                # A função fetch_setor_temp agora é inteligente e tenta displayName, cn, etc.
                localidade = fetch_setor_temp(ad_conn, solicitante_nome)

            collected.append({
                "Chamado#": cid,
                "Nome do Usuário": solicitante_nome,
                "Unidade": localidade,
                "Descrição": descricao,
                "Data Criação": data_criacao
            })
            logger.info(f"[{idx+1}/{len(rows)}] Lido: {cid}")

        except Exception as e:
            # Erros de leitura em uma linha não devem parar o script
            continue

    return collected

def ir_para_proxima_pagina(driver, wait):
    try:
        # Busca o botão da seta "Próximo" (›)
        btn_next_container = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "li.pagination-next")
        ))

        # Se tiver a classe 'disabled', acabaram as páginas
        if "disabled" in btn_next_container.get_attribute("class"):
            logger.info("Paginação encerrada: Botão 'Próximo' está desabilitado.")
            return False

        # Clica no link dentro do LI
        link_next = btn_next_container.find_element(By.TAG_NAME, "a")
        driver.execute_script("arguments[0].click();", link_next)
        logger.info("Navegando para a próxima página...")

        # Aguarda tabela atualizar
        time.sleep(3)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#table tbody tr")))
        return True

    except Exception as e:
        logger.error(f"Fim da paginação ou erro: {e}")
        return False

# ---------------------------
# Fluxo principal
# ---------------------------
def scrape_citsmart():
    ad_conn = None
    try:
        ad_conn = setup_ad_connection()
        logger.info("Conexão AD estabelecida.")
    except Exception as e:
        logger.error(f"AD indisponível: {e}")

    driver, wait = initial_config()
    todos_os_dados = []

    try:
        navigate_to_caixa_entrada(driver, wait)
        expand_all_records_lowcode(driver, wait)

        pagina = 1
        while True:
            logger.info(f"--- Processando Página {pagina} ---")
            
            dados_pagina = process_page(driver, wait, filtro_grupo=None, ad_conn=ad_conn)
            if dados_pagina:
                todos_os_dados.extend(dados_pagina)
                logger.info(f"Coletados {len(dados_pagina)} registros nesta página.")
            else:
                logger.info("Aviso: Página retornou 0 registros.")

            # Tenta ir para próxima página
            if not ir_para_proxima_pagina(driver, wait):
                break
            
            pagina += 1

        # Exportação Final
        if todos_os_dados:
            out_dir = Path("01 - Dados Brutos")
            out_dir.mkdir(exist_ok=True)
            ts = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
            file = out_dir / f"Chamados_CitSmart_{ts}.xlsx"

            df = pd.DataFrame(todos_os_dados)
            with pd.ExcelWriter(file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Chamados", index=False)

            with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
                wb = cast(XlsxWorkbook, writer.book)
                wrap = wb.add_format({'text_wrap': True})
                ws = writer.sheets['Sheet1']

                widths = {
                    'Chamado#': 15,
                    'Nome do Usuário': 25,
                    'Unidade': 40,
                    'Descrição': 100,
                    'Data Criação': 20
                }

                for i, col in enumerate(df.columns):
                    fmt = wrap if col=='Descrição' else None
                    ws.set_column(i, i, widths.get(col,20), fmt)
                
                for r, desc in enumerate(df['Descrição'], start=1):
                    text = '' if pd.isna(desc) else str(desc)
                    ws.set_row(r, 15 * (text.count('\n')+1))

            logger.info(f"SUCESSO! Total de {len(todos_os_dados)} chamados salvos em: {file}")
        else:
            logger.info("Nenhum dado foi coletado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_citsmart()