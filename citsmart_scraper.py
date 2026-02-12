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
from pathlib import Path
from ldap3 import Server, Connection, ALL, SUBTREE
from typing import cast
from xlsxwriter.workbook import Workbook as XlsxWorkbook # Alias para não confundir
from config import (
    CITSMART_URL, CITSMART_EMAIL, PASSWORD,
    HEADLESS, EXPLICIT_WAIT,
    DOMINIO, USERNAME
)

# ---------------------------
# Utilitários e Log
# ---------------------------
def debug_print(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[DEBUG {ts}] {msg}")

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
    server = Server(DOMINIO, get_info=ALL)
    conn = Connection(server, user=f"MPE\\{USERNAME}", password=PASSWORD, auto_bind=True)
    return conn

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
        debug_print(f"Erro lookup AD para '{display_name}': {e}")
        return 'Erro na Consulta'

# ---------------------------
# Navegador / Login
# ---------------------------
def initial_config():
    opts = webdriver.ChromeOptions()
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_argument("--disable-blink-features=CSSAnimations,ScrollAnimator")
    opts.add_argument("--incognito")
    if HEADLESS:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
    else:
        opts.add_argument("--start-maximized")
    opts.page_load_strategy = "eager"
    driver = webdriver.Chrome(options=opts)
    wait = WebDriverWait(driver, timeout=EXPLICIT_WAIT, poll_frequency=0.1)
    return driver, wait

def navigate_to_caixa_entrada(driver, wait):
    debug_print("Acessando CitSmart e fazendo login…")
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
        debug_print("Pulando KMSI de manter conectado…")
        wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    except TimeoutException:
        pass

    # 4) Redirecionamento Direto para LowCode
    debug_print("Aguardando carregamento do portal inicial...")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(5) 

    nova_fila = "https://suporte.mpms.mp.br/inbox/lowcode/form/copilot_novo/default"
    debug_print(f"Forçando navegação para: {nova_fila}")
    
    driver.switch_to.default_content()
    driver.execute_script(f"window.location.href = '{nova_fila}';")

    try:
        wait.until(EC.url_contains("copilot_novo"))
        debug_print("URL de destino alcançada.")
        
        debug_print("Aguardando o iframe 'App'...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[title='App']")))
        
        wait.until(EC.presence_of_element_located((By.ID, "pageSize")))
        debug_print("Sucesso! Interface do Copilot detectada via seletor de paginação.")
        
    except Exception as e:
        debug_print(f"Não detectou o elemento interno: {e}")
        driver.save_screenshot("erro_iframe_app.png")
        raise

# ---------------------------
# Manipulação da Tabela e Paginação
# ---------------------------
def expand_all_records_lowcode(driver, wait):
    """
    Tenta localizar o pager e setar para 100 itens.
    Usa o loader específico (.hyper-loading) para sincronizar.
    """
    debug_print("Tentando expandir registros (LowCode)...")
    
    # Seletor do GIF de carregamento que você forneceu
    loader_loc = (By.CSS_SELECTOR, "div.hyper-loading")

    # Função auxiliar para esperar o loader sumir
    def wait_loader_vanish(timeout=30):
        try:
            # Espera até que o elemento fique invisível (display: none)
            WebDriverWait(driver, timeout).until(
                EC.invisibility_of_element_located(loader_loc)
            )
        except TimeoutException:
            debug_print("Aviso: O loader demorou muito para sumir ou não apareceu.")

    try:
        time.sleep(3) 
        
        # 1. ANTES DE TUDO: Garante que a página está "quieta"
        wait_loader_vanish()

        # 2. Localiza o dropdown
        pager_select_loc = (By.CSS_SELECTOR, "span.k-pager-sizes select, .k-pager-sizes select")
        wait.until(EC.element_to_be_clickable(pager_select_loc))
        
        select_elem = driver.find_element(*pager_select_loc)
        sel = Select(select_elem)
        
        # Verifica se já está em 100
        try:
            current = sel.first_selected_option.text.strip()
        except:
            current = ""

        if "100" in current:
            debug_print("Já está exibindo 100 registros.")
        else:
            # 3. APLICA A MUDANÇA
            sel.select_by_visible_text("100")
            debug_print("Selecionado '100' no dropdown.")
            
            # --- O PULO DO GATO ---
            # O sistema demora alguns milissegundos para mudar o display de 'none' para 'block'.
            # Se checarmos a invisibilidade imediatamente, vai dar True (falso positivo).
            # Damos 3 segundos para o loader APARECER.
            time.sleep(3) 
            
            # Agora esperamos ele SUMIR
            debug_print("Aguardando o loader (.hyper-loading) desaparecer...")
            wait_loader_vanish(timeout=30)

        # 4. Extrai a contagem final para garantir que atualizou
        pager_info = driver.find_element(By.CSS_SELECTOR, "span.k-pager-info")
        text = pager_info.text.strip()  # ex: "1 - 100 de 250 itens"
        debug_print(f"Paginação atualizada: {text}")

        match = re.search(r"de\s+(\d+)", text)
        if match:
            return int(match.group(1))
        return 0

    except TimeoutException:
        debug_print("Pager não encontrado ou timeout esperando loader.")
        return 0
    except Exception as e:
        debug_print(f"Erro ao expandir registros: {e}")
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
        debug_print("Nenhuma linha encontrada na tabela. Aguardando 3s...")
        time.sleep(3)
        rows = _list_rows(driver)

    debug_print(f"Linhas detectadas: {len(rows)}")
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
            debug_print(f"[{idx+1}/{len(rows)}] Lido: {cid}")

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
            debug_print("Paginação encerrada: Botão 'Próximo' está desabilitado.")
            return False

        # Clica no link dentro do LI
        link_next = btn_next_container.find_element(By.TAG_NAME, "a")
        driver.execute_script("arguments[0].click();", link_next)
        debug_print("Navegando para a próxima página...")

        # Aguarda tabela atualizar
        time.sleep(3)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#table tbody tr")))
        return True

    except Exception as e:
        debug_print(f"Fim da paginação ou erro: {e}")
        return False

# ---------------------------
# Fluxo principal
# ---------------------------
def scrape_citsmart():
    ad_conn = None
    try:
        ad_conn = setup_ad_connection()
        debug_print("Conexão AD estabelecida.")
    except Exception as e:
        debug_print(f"AD indisponível: {e}")

    driver, wait = initial_config()
    todos_os_dados = []

    try:
        navigate_to_caixa_entrada(driver, wait)
        expand_all_records_lowcode(driver, wait)

        pagina = 1
        while True:
            debug_print(f"--- Processando Página {pagina} ---")
            
            dados_pagina = process_page(driver, wait, filtro_grupo=None, ad_conn=ad_conn)
            if dados_pagina:
                todos_os_dados.extend(dados_pagina)
                debug_print(f"Coletados {len(dados_pagina)} registros nesta página.")
            else:
                debug_print("Aviso: Página retornou 0 registros.")

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

            debug_print(f"SUCESSO! Total de {len(todos_os_dados)} chamados salvos em: {file}")
        else:
            debug_print("Nenhum dado foi coletado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_citsmart()