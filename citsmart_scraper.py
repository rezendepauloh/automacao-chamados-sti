# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import datetime
import time
import pandas as pd
import re
import logging
from pathlib import Path
import json
import html
from ldap3 import SUBTREE
from config import (
    CITSMART_URL, CITSMART_EMAIL, PASSWORD,
    HEADLESS, EXPLICIT_WAIT, DEBUG_DIR_CITSMART,
    setup_logging, save_df_to_excel_formatted,
    setup_ad_connection, get_chrome_driver, fetch_ad_department, cleanup_old_files
)

# ---------------------------
# Utilitários e Log
# ---------------------------
# --- Configuração de logging ---
logger = setup_logging(DEBUG_DIR_CITSMART / "citsmart_scraper.log", __name__)

# (Opcional) Manter o silenciador do Selenium/urllib3 nos scripts de scraping
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
def fetch_setor_temp(conn, query_val, is_username=False):
    """
    Busca o setor do usuário no AD de forma unificada e parametrizável.
    """
    return fetch_ad_department(conn, query_val, is_username=is_username)

# ---------------------------
# Navegador / Login
# ---------------------------
def initial_config():
    driver = get_chrome_driver(headless=HEADLESS, page_load_strategy="eager", disable_gpu=True)
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
        
        # Injeta o interceptador de XHR para capturar o JSON de chamados diretamente do AngularJS/LowCode!
        logger.info("Injetando interceptador XHR no contexto do iframe para captura direta de JSON...")
        driver.execute_script("""
            (function() {
                if (window.__xhr_patched__) return;
                window.__xhr_patched__ = true;
                window.__captured_tickets__ = null;

                var oldOpen = XMLHttpRequest.prototype.open;
                XMLHttpRequest.prototype.open = function(method, url) {
                    this._url = url;
                    return oldOpen.apply(this, arguments);
                };

                var oldSend = XMLHttpRequest.prototype.send;
                XMLHttpRequest.prototype.send = function() {
                    var self = this;
                    var oldOnReadyStateChange = this.onreadystatechange;
                    this.onreadystatechange = function() {
                        if (self.readyState === 4 && self.status === 200) {
                            if (self._url && self._url.indexOf('tb_ticket_queue/list') !== -1) {
                                try {
                                    var data = JSON.parse(self.responseText);
                                    var candidate = null;
                                    if (Array.isArray(data)) {
                                        candidate = data;
                                    } else if (data && Array.isArray(data.list)) {
                                        candidate = data.list;
                                    }
                                    
                                    // Valida se o candidato é de fato um array de chamados (deve conter ticket_id ou id)
                                    if (candidate && candidate.length > 0) {
                                        var has_ticket_id = candidate.some(function(item) {
                                            return item && (item.ticket_id || item.id);
                                        });
                                        if (has_ticket_id) {
                                            // Mantém o maior lote capturado e evita que requests menores de atualização sobrescrevam
                                            if (!window.__captured_tickets__ || candidate.length > window.__captured_tickets__.length) {
                                                window.__captured_tickets__ = candidate;
                                                console.log("🔥 [CAPTURED TICKETS] " + window.__captured_tickets__.length + " chamados gravados!");
                                            }
                                        }
                                    }
                                } catch(e) {
                                    console.error('Error parsing captured tickets:', e);
                                }
                            }
                        }
                        if (oldOnReadyStateChange) {
                            return oldOnReadyStateChange.apply(this, arguments);
                        }
                    };
                    return oldSend.apply(this, arguments);
                };
            })();
        """)

        # Usa o Select do Selenium para interagir com ele
        dropdown = Select(dropdown_element)
        
        # Verifica se já está em 100
        try:
            current = dropdown.first_selected_option.text.strip()
        except:
            current = ""

        if "100" in current:
            logger.info("Já está exibindo 100 registros. Forçando toggle (100 -> 50 -> 100) para registrar captura de rede...")
            dropdown.select_by_visible_text("50")
            time.sleep(1.5)
            # Re-localiza elemento para evitar StaleElementReferenceException
            dropdown_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "pageSize"))
            )
            dropdown = Select(dropdown_element)
            dropdown.select_by_visible_text("100")
            time.sleep(1)
            wait_loader_vanish(timeout=30)
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

def clean_html_comment(html_str):
    if not html_str:
        return ""
    # Remove tags HTML
    txt = re.sub(r'<[^>]*>', ' ', html_str)
    # Decodifica entidades HTML
    txt = html.unescape(txt)
    # Remove linhas vazias redundantes
    lines = [line.strip() for line in txt.splitlines() if line.strip()]
    return "\n".join(lines)

def _list_rows(driver):
    try:
        return driver.find_elements(By.CSS_SELECTOR, "#table tbody tr")
    except:
        return []

def process_page(driver, wait, filtro_grupo=None, ad_conn=None, cache=None):
    # Nota: Já estamos no iframe correto, não precisa de switch_to_incidents
    
    # 1. TENTA EXTRAÇÃO ULTRARRÁPIDA VIA JSON INTERCEPTADO NO XHR
    try:
        captured = driver.execute_script("return window.__captured_tickets__;")
        if captured and isinstance(captured, list) and len(captured) > 0:
            logger.info(f"⚡ [PROCESSO ULTRA-RÁPIDO] Processando {len(captured)} chamados capturados diretamente via JSON de rede!")
            collected = []
            for idx, ticket in enumerate(captured):
                try:
                    cid = str(ticket.get("ticket_id", ""))
                    if not cid:
                        continue
                        
                    solicitante_nome = ticket.get("ticket_requester", "")
                    
                    # Extração do ID do Cliente (username do email_solicitante)
                    id_cliente = ""
                    email_solicitante = ticket.get("email_solicitante", "")
                    if email_solicitante and "@" in email_solicitante:
                        id_cliente = email_solicitante.split("@")[0].strip()
                    
                    # Tratamento inteligente de Data
                    data_criacao = ""
                    iso_str = ticket.get("ticket_creationdate_str", "")
                    if iso_str:
                        try:
                            # "2026-05-07T11:12:07.396Z" -> "07/05/2026 11:12"
                            clean_iso = iso_str.split(".")[0].replace("Z", "")
                            dt = datetime.strptime(clean_iso, "%Y-%m-%dT%H:%M:%S")
                            data_criacao = dt.strftime("%d/%m/%Y %H:%M")
                        except Exception:
                            data_criacao = iso_str

                    # Carrega comentários do cache
                    comments_list = []
                    if cache and cid in cache:
                        cached_comments_str = cache[cid].get('Comentários', '[]')
                        try:
                            comments_list = json.loads(cached_comments_str)
                        except:
                            pass

                    # Tenta extrair o último comentário do JSON atual
                    ticket_ocorrencia = ticket.get("ticket_ocorrencia")
                    if ticket_ocorrencia and ticket_ocorrencia.strip():
                        texto_limpo = clean_html_comment(ticket_ocorrencia)
                        data_reg = ticket.get("ticket_ocorrencia_dataregistro_br") or ticket.get("ticket_ocorrencia_dataregistro")
                        autor = ticket.get("responsible_ocorrencia") or "Sistema"
                        
                        new_comment = {
                            "data": str(data_reg).strip() if data_reg else datetime.now().strftime("%d/%m/%Y %H:%M"),
                            "autor": str(autor).strip(),
                            "texto": texto_limpo
                        }
                        
                        # Evita duplicidade comparando data e conteúdo
                        exists = False
                        for existing in comments_list:
                            if existing.get('data') == new_comment['data'] and existing.get('texto') == new_comment['texto']:
                                exists = True
                                break
                        if not exists:
                            comments_list.append(new_comment)

                    comments_json = json.dumps(comments_list, ensure_ascii=False)

                    # Verificação de Cache
                    if cache and cid in cache:
                        cached_ip = cache[cid].get('IP_Origem') or ""
                        cached_unidade = cache[cid].get('Unidade') or ""
                        cached_desc = cache[cid].get('Descrição') or ""
                        
                        # Se temos o IP no cache, aproveitamos 100%
                        if cached_ip:
                            collected.append({
                                "Chamado#": cid,
                                "ID do Cliente": id_cliente,
                                "Nome do Usuário": solicitante_nome,
                                "Unidade": cached_unidade or "Não encontrada no AD",
                                "Descrição": cached_desc,
                                "Data Criação": data_criacao,
                                "IP_Origem": cached_ip,
                                "Comentários": comments_json
                            })
                            logger.info(f"[{idx+1}/{len(captured)}] ⚡ [CACHE MATCH] Chamado {cid} (com IP: {cached_ip}) recuperado do cache anterior!")
                            continue
                        else:
                            # Se não temos o IP, reaproveitamos os outros dados do cache e apenas buscamos o IP
                            sccm_ip = ""
                            if id_cliente:
                                from config import fetch_ip_from_sccm
                                sccm_ip = fetch_ip_from_sccm(id_cliente)
                            
                            collected.append({
                                "Chamado#": cid,
                                "ID do Cliente": id_cliente,
                                "Nome do Usuário": solicitante_nome,
                                "Unidade": cached_unidade or "Não encontrada no AD",
                                "Descrição": cached_desc,
                                "Data Criação": data_criacao,
                                "IP_Origem": sccm_ip,
                                "Comentários": comments_json
                            })
                            logger.info(f"[{idx+1}/{len(captured)}] ⚡ [CACHE PARCIAL] Chamado {cid} recuperado do cache anterior, consultando IP no SCCM...")
                            continue
                    
                    # Enriquecimento AD se disponível, senão usa unidade nativa do JSON
                    localidade = "Não encontrada no AD"
                    if ad_conn:
                        if id_cliente:
                            localidade = fetch_setor_temp(ad_conn, id_cliente, is_username=True)
                        elif solicitante_nome:
                            localidade = fetch_setor_temp(ad_conn, solicitante_nome, is_username=False)
                    else:
                        localidade = ticket.get("ticket_unit", "") or ticket.get("nome_unidade", "Não encontrada")
                            
                    sccm_ip = ""
                    if id_cliente:
                        from config import fetch_ip_from_sccm
                        sccm_ip = fetch_ip_from_sccm(id_cliente)

                    collected.append({
                        "Chamado#": cid,
                        "ID do Cliente": id_cliente,
                        "Nome do Usuário": solicitante_nome,
                        "Unidade": localidade,
                        "Descrição": ticket.get("ticket_description_long", "") or ticket.get("ticket_description", ""),
                        "Data Criação": data_criacao,
                        "IP_Origem": sccm_ip,
                        "Comentários": comments_json
                    })
                    logger.info(f"[{idx+1}/{len(captured)}] Processado JSON: {cid} (Login: {id_cliente})")
                except Exception as row_err:
                    logger.error(f"Erro ao parsear ticket do JSON: {row_err}")
                    continue
            return collected
    except Exception as e:
        logger.warning(f"Aviso: Não foi possível usar extração ultrarrápida JS ({e}). Usando modo tradicional...")

    # 2. SE NÃO HOUVER JSON, SEGUI COM O MÉTODO TRADICIONAL DE RASPAGEM DO DOM
    rows = _list_rows(driver)
    
    # Se não achou linhas, espera um pouco e tenta de novo (carregamento lento)
    if not rows:
        logger.info("Nenhuma linha encontrada na tabela. Aguardando 3s...")
        time.sleep(3)
        rows = _list_rows(driver)

    logger.info(f"Linhas detectadas no DOM: {len(rows)}")
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
            num_bruto = get_val("1")
            num_match = re.search(r'\d+', num_bruto)
            cid = num_match.group(0) if num_match else ""

            if not cid: continue # Pula linhas inválidas

            solicitante_full = get_val("6")
            data_criacao = get_val("9")

            # --- Tratamento inteligente de login/id_cliente ---
            solicitante_nome = solicitante_full
            id_cliente = ""
            
            if "(" in solicitante_full:
                try:
                    partes = solicitante_full.split("(")
                    solicitante_nome = partes[0].strip()
                    id_cliente = partes[1].replace(")", "").strip()
                except:
                    pass

            # Carrega comentários do cache se existirem no DOM
            comments_json = cache[cid].get('Comentários', '[]') if (cache and cid in cache) else '[]'

            # Verificação de Cache
            if cache and cid in cache:
                cached_ip = cache[cid].get('IP_Origem') or ""
                cached_unidade = cache[cid].get('Unidade') or ""
                cached_desc = cache[cid].get('Descrição') or ""
                
                if cached_ip:
                    collected.append({
                        "Chamado#": cid,
                        "ID do Cliente": id_cliente,
                        "Nome do Usuário": solicitante_nome,
                        "Unidade": cached_unidade or "Não encontrada no AD",
                        "Descrição": cached_desc,
                        "Data Criação": data_criacao,
                        "IP_Origem": cached_ip,
                        "Comentários": comments_json
                    })
                    logger.info(f"[{idx+1}/{len(rows)}] ⚡ [CACHE MATCH] Lido DOM via Cache: {cid} (IP: {cached_ip})")
                    continue
                else:
                    sccm_ip = ""
                    if id_cliente:
                        from config import fetch_ip_from_sccm
                        sccm_ip = fetch_ip_from_sccm(id_cliente)
                        
                    collected.append({
                        "Chamado#": cid,
                        "ID do Cliente": id_cliente,
                        "Nome do Usuário": solicitante_nome,
                        "Unidade": cached_unidade or "Não encontrada no AD",
                        "Descrição": cached_desc,
                        "Data Criação": data_criacao,
                        "IP_Origem": sccm_ip,
                        "Comentários": comments_json
                    })
                    logger.info(f"[{idx+1}/{len(rows)}] ⚡ [CACHE PARCIAL] Lido DOM via Cache: {cid}, consultando IP no SCCM...")
                    continue

            descricao = get_val("10", is_description=True)

            # --- Enriquecimento AD ---
            localidade = "Não encontrada no AD"
            if ad_conn:
                if id_cliente:
                    # Consulta direta pelo sAMAccountName (login) extraído dos parênteses do DOM
                    localidade = fetch_setor_temp(ad_conn, id_cliente, is_username=True)
                else:
                    localidade = fetch_setor_temp(ad_conn, solicitante_nome, is_username=False)

            sccm_ip = ""
            if id_cliente:
                from config import fetch_ip_from_sccm
                sccm_ip = fetch_ip_from_sccm(id_cliente)

            collected.append({
                "Chamado#": cid,
                "ID do Cliente": id_cliente,
                "Nome do Usuário": solicitante_nome,
                "Unidade": localidade,
                "Descrição": descricao,
                "Data Criação": data_criacao,
                "IP_Origem": sccm_ip,
                "Comentários": comments_json
            })
            logger.info(f"[{idx+1}/{len(rows)}] Lido DOM: {cid} (Login: {id_cliente})")

        except Exception as e:
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
    # Carrega cache do último arquivo de CitSmart gerado
    cache = {}
    try:
        out_dir = Path("01 - Dados Brutos")
        existing_files = sorted(out_dir.glob("Chamados_CitSmart_*.xlsx"))
        if existing_files:
            latest_file = existing_files[-1]
            logger.info(f"Carregando cache de descrições e comentários do arquivo mais recente de CitSmart: {latest_file.name}")
            df_old = pd.read_excel(latest_file, dtype=str)
            for _, row_old in df_old.iterrows():
                cid = str(row_old.get('Chamado#', '')).strip()
                desc = row_old.get('Descrição', '')
                ip = row_old.get('IP_Origem', '')
                comments = row_old.get('Comentários', '[]')
                if cid:
                    cache[cid] = {
                        'Descrição': str(desc).strip() if pd.notna(desc) else '',
                        'Unidade': str(row_old.get('Unidade', '')).strip() if pd.notna(row_old.get('Unidade')) else '',
                        'IP_Origem': str(ip).strip() if pd.notna(ip) else '',
                        'Comentários': str(comments).strip() if pd.notna(comments) else '[]'
                    }
            logger.info(f"Sucesso! {len(cache)} chamados carregados no cache do CitSmart.")
    except Exception as cache_err:
        logger.warning(f"Aviso: Não foi possível carregar cache do CitSmart: {cache_err}")

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
            
            dados_pagina = process_page(driver, wait, filtro_grupo=None, ad_conn=ad_conn, cache=cache)
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
            widths = {
                'Chamado#': 15,
                'ID do Cliente': 20,
                'Nome do Usuário': 25,
                'Unidade': 40,
                'Descrição': 100,
                'Data Criação': 20,
                'IP_Origem': 15,
                'Comentários': 50
            }
            save_df_to_excel_formatted(
                df, file, sheet_name="Chamados",
                widths=widths, wrap_cols=['Descrição', 'Comentários'], height_col='Descrição'
            )

            logger.info(f"SUCESSO! Total de {len(todos_os_dados)} chamados salvos em: {file}")
            
            # Limpeza de arquivos antigos (mantém no máximo os 10 últimos do CitSmart)
            cleanup_old_files(out_dir, "Chamados_CitSmart_*.xlsx", keep_count=10)
        else:
            logger.info("Nenhum dado foi coletado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_citsmart()