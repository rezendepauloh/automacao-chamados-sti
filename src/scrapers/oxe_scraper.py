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

import time
import json
import logging
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from src.components.status_banner import check_process_running, read_log_lines
from src.config import (
    OXE_URL, OXE_USER, OXE_PASS,
    HEADLESS, EXPLICIT_WAIT, DEBUG_DIR_OXE,
    setup_logging, save_df_to_excel_formatted,
    get_chrome_driver, cleanup_old_files
)
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE


logger = setup_logging(DEBUG_DIR_OXE / "oxe_scraper.log", __name__)

logging.getLogger('selenium.webdriver.remote.remote_connection').setLevel(logging.WARNING)
logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

def salvar_screenshot(driver, nome_etapa: str):
    try:
        ts = datetime.now().strftime("%H-%M-%S")
        filename = f"debug_oxe_{ts}_{nome_etapa}.png"
        filepath = DEBUG_DIR_OXE / filename
        driver.save_screenshot(str(filepath))
        logger.debug(f"📸 Screenshot salvo em: {filepath.name}")
    except Exception as e:
        logger.warning(f"Não foi possível salvar screenshot '{nome_etapa}': {e}")

def initial_config():
    driver = get_chrome_driver(headless=HEADLESS, disable_gpu=True)
    wait = WebDriverWait(driver, timeout=EXPLICIT_WAIT, poll_frequency=0.2)
    return driver, wait

def realizar_login_oxe(driver, wait):
    target_url = OXE_URL if OXE_URL.endswith("/") else f"{OXE_URL}/"
    login_full_url = f"{target_url}#/login"
    
    logger.info(f"Navegando para a página de login do OXE: {login_full_url}")
    driver.get(login_full_url)
    time.sleep(2)

    logger.info("Injetando interceptador de cabeçalhos de autenticação XHR/Fetch...")
    driver.execute_script("""
        (function() {
            if (window.__xhr_patched__) return;
            window.__xhr_patched__ = true;
            window.__captured_auth_token__ = null;

            var oldSetHeader = XMLHttpRequest.prototype.setRequestHeader;
            XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
                if (header && (header.toLowerCase() === 'authorization' || header.toLowerCase() === 'x-auth-token' || header.toLowerCase() === 'token')) {
                    window.__captured_auth_token__ = value;
                    console.log("🔥 [CAPTURED AUTH TOKEN XHR]", header, value);
                }
                return oldSetHeader.apply(this, arguments);
            };

            var oldFetch = window.fetch;
            window.fetch = function(resource, init) {
                if (init && init.headers) {
                    var h = init.headers;
                    if (h instanceof Headers) {
                        if (h.get('authorization')) window.__captured_auth_token__ = h.get('authorization');
                    } else if (typeof h === 'object') {
                        for (var k in h) {
                            if (k.toLowerCase() === 'authorization') window.__captured_auth_token__ = h[k];
                        }
                    }
                }
                return oldFetch.apply(this, arguments);
            };
        })();
    """)

    try:
        page_src = driver.page_source.lower()
        if "details-button" in page_src or "err_cert" in page_src or "privacidade" in page_src or "particular" in page_src:
            logger.info("Detectada tela de aviso de segurança do Chrome. Ignorando certificado...")
            btn_details = driver.find_element(By.ID, "details-button")
            btn_details.click()
            time.sleep(0.8)
            btn_proceed = driver.find_element(By.ID, "proceed-link")
            btn_proceed.click()
            time.sleep(3)
    except Exception as cert_err:
        logger.debug(f"Bypass de certificado via DOM não foi necessário: {cert_err}")

    logger.info("Aguardando formulário de login...")
    try:
        user_field = wait.until(EC.element_to_be_clickable((By.ID, "username")))
    except TimeoutException:
        try:
            if "details-button" in driver.page_source:
                driver.find_element(By.ID, "details-button").click()
                time.sleep(0.5)
                driver.find_element(By.ID, "proceed-link").click()
                time.sleep(3)
        except Exception:
            pass
        logger.warning("Campo #username não localizado por ID. Tentando por seletor alternativo...")
        user_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text']")))

    user_field.clear()
    user_field.send_keys(OXE_USER)
    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true })); arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", user_field)

    pwd_field = driver.find_element(By.ID, "password")
    pwd_field.clear()
    pwd_field.send_keys(OXE_PASS)
    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true })); arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", pwd_field)

    logger.info(f"Efetuando login no OXE com o usuário '{OXE_USER}'...")
    login_btn = driver.find_element(By.ID, "login-button")
    try:
        login_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", login_btn)

    try:
        form = driver.find_element(By.TAG_NAME, "form")
        driver.execute_script("arguments[0].dispatchEvent(new Event('submit', { bubbles: true }));", form)
    except Exception:
        pass

    time.sleep(3)
    logger.info(f"URL após login: {driver.current_url}")
    
    logger.info("Aguardando gravação de tokens de sessão no storage...")
    token_capturado = None
    for attempt in range(15):
        token_capturado = driver.execute_script(
            "return sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__;"
        )
        if token_capturado:
            logger.info(f"🔑 Token capturado com sucesso! (Tamanho: {len(token_capturado)})")
            break
        time.sleep(0.5)

    if not token_capturado:
        salvar_screenshot(driver, "falha_login_token")
        storage_dump = driver.execute_script("""
            var res = { session: {}, local: {} };
            try {
                for(var i=0; i<sessionStorage.length; i++){
                    var k = sessionStorage.key(i);
                    res.session[k] = sessionStorage.getItem(k);
                }
            } catch(e){}
            try {
                for(var i=0; i<localStorage.length; i++){
                    var k = localStorage.key(i);
                    res.local[k] = localStorage.getItem(k);
                }
            } catch(e){}
            return res;
        """)
        logger.warning(f"⚠️ Token não encontrado. Conteúdo atual dos Storages: {storage_dump}")

    try:
        elements = driver.find_elements(By.XPATH, "//*[contains(translate(text(), 'UTILIZADORES', 'utilizadores'), 'utilizadores') or contains(text(), 'Subscriber')]")
        if elements:
            logger.info(f"Encontrado item de menu de Utilizadores ({len(elements)}). Clicando para inicializar...")
            driver.execute_script("arguments[0].click();", elements[0])
            time.sleep(2)
    except Exception as nav_err:
        logger.debug(f"Interação de menu opcional: {nav_err}")

def fetch_oxe_api_js(driver, api_endpoint: str):
    js_code = f"""
        var callback = arguments[arguments.length - 1];
        
        var idToken = sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__;

        if (!idToken) {{
            for (var s of [sessionStorage, localStorage]) {{
                if (!s) continue;
                for (var i = 0; i < s.length; i++) {{
                    var k = s.key(i);
                    if (k.toLowerCase().includes('token')) {{
                        idToken = s.getItem(k);
                        if (idToken) break;
                    }}
                }}
                if (idToken) break;
            }}
        }}

        if (!idToken) {{
            callback({{ status: 'error', message: 'Nenhum id_token encontrado no sessionStorage/localStorage' }});
            return;
        }}

        var tokenBearer = idToken.startsWith('Bearer ') ? idToken : 'Bearer ' + idToken;
        var tokenRaw = idToken.replace(/^Bearer\\s+/i, '');

        function tryFetch(authHeaderValue, isRetry) {{
            fetch('{api_endpoint}', {{
                method: 'GET',
                headers: {{
                    'Accept': 'application/json, text/plain, */*',
                    'Authorization': authHeaderValue
                }}
            }})
            .then(function(response) {{
                if (response.status === 401 && !isRetry) {{
                    var nextToken = (authHeaderValue === tokenBearer) ? tokenRaw : tokenBearer;
                    tryFetch(nextToken, true);
                }} else if (!response.ok) {{
                    throw new Error('HTTP status ' + response.status);
                }} else {{
                    return response.json().then(function(data) {{
                        callback({{ status: 'success', data: data }});
                    }});
                }}
            }})
            .catch(function(err) {{
                callback({{ status: 'error', message: err.toString() }});
            }});
        }}

        tryFetch(tokenBearer, false);
    """
    res = driver.execute_async_script(js_code)
    if res and res.get("status") == "success":
        return res.get("data")
    else:
        err_msg = res.get("message") if res else "Sem resposta do script JS"
        logger.debug(f"Erro na requisição JS para {api_endpoint}: {err_msg}")
        return None

def fetch_oxe_batch_subscriber_details_js(driver, ramais_list: list) -> dict:
    if not ramais_list:
        return {}

    js_code = """
        var callback = arguments[arguments.length - 1];
        var ramais = arguments[0];
        var idToken = sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__ || '';
        var tokenBearer = idToken.startsWith('Bearer ') ? idToken : (idToken ? 'Bearer ' + idToken : '');

        var results = {};
        var promises = ramais.map(function(num_ramal) {
            var url = '/api/mgt/1.0/Node/1/Subscriber/' + num_ramal;
            return fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json, text/plain, */*',
                    'Authorization': tokenBearer
                }
            })
            .then(function(res) {
                if (res.ok) return res.json();
                return null;
            })
            .then(function(data) {
                if (data) results[num_ramal] = data;
            })
            .catch(function(e) {});
        });

        Promise.all(promises).then(function() {
            callback({ status: 'success', data: results });
        });
    """
    res = driver.execute_async_script(js_code, ramais_list)
    if res and res.get("status") == "success":
        return res.get("data", {})
    return {}

def fetch_oxe_batch_tsc_ip_js(driver, ramais_list: list) -> dict:
    if not ramais_list:
        return {}

    js_code = """
        var callback = arguments[arguments.length - 1];
        var ramais = arguments[0];
        var idToken = sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__ || '';
        var tokenBearer = idToken.startsWith('Bearer ') ? idToken : (idToken ? 'Bearer ' + idToken : '');

        var results = {};
        var promises = ramais.map(function(num_ramal) {
            var url = '/api/mgt/1.0/Node/1/Subscriber/' + num_ramal + '/Tsc_IP_subscriber/' + num_ramal;
            return fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json, text/plain, */*',
                    'Authorization': tokenBearer
                }
            })
            .then(function(res) {
                if (res.ok) return res.json();
                return null;
            })
            .then(function(data) {
                if (data) results[num_ramal] = data;
            })
            .catch(function(e) {});
        });

        Promise.all(promises).then(function() {
            callback({ status: 'success', data: results });
        });
    """
    res = driver.execute_async_script(js_code, ramais_list)
    if res and res.get("status") == "success":
        return res.get("data", {})
    return {}

def extrair_dados_assinantes(driver):
    subscribers_endpoint = (
        "/api/mgt/1.0/Node/1/Subscriber?attributes="
        "Annu_Name,Annu_First_Name,UTF8_Phone_Book_Name,UTF8_Phone_Book_First_Name,"
        "Equipment_Address_Rack,Equipment_Address_Board,Equipment_Address_Terminal,"
        "Station_Type,Opex_License,External_Login,Mail_Address,Station_Sub_Type,"
        "DM_Profile,Entity_Number,Set_Role"
    )
    
    logger.info("Solicitando lista principal de utilizadores/assinantes à API do OXE...")
    subscribers = fetch_oxe_api_js(driver, subscribers_endpoint)
    
    if not subscribers or not isinstance(subscribers, list):
        logger.error("A API do OXE não retornou a lista de assinantes esperada.")
        return []

    logger.info(f"Total de {len(subscribers)} assinantes retornados pela API principal.")

    todos_ramais = [str(s.get("Directory_Number", "")).strip() for s in subscribers if s.get("Directory_Number")]

    logger.info(f"⚡ Disparando consulta paralela em lote de detalhes avançados para {len(todos_ramais)} ramais...")
    subscriber_details_map = {}
    chunk_size = 100
    for i in range(0, len(todos_ramais), chunk_size):
        chunk = todos_ramais[i:i + chunk_size]
        res_chunk = fetch_oxe_batch_subscriber_details_js(driver, chunk)
        if res_chunk:
            subscriber_details_map.update(res_chunk)
        logger.info(f"⚡ Lote Detalhes {i + len(chunk)}/{len(todos_ramais)} processado.")

    logger.info(f"✅ Detalhes avançados (Grupo de Captura, Categoria) obtidos para {len(subscriber_details_map)} ramais.")

    ramais_ip_candidatos = []
    for sub in subscribers:
        num_ramal = str(sub.get("Directory_Number", "")).strip()
        station_type = str(sub.get("Station_Type", "")).upper()
        station_sub_type = str(sub.get("Station_Sub_Type", "")).upper()
        
        if num_ramal and ("IP" in station_type or "SIP" in station_sub_type or "NOE" in station_type):
            ramais_ip_candidatos.append(num_ramal)

    logger.info(f"⚡ Disparando consulta paralela em lote para {len(ramais_ip_candidatos)} ramais IP...")

    tsc_ip_map = {}
    for i in range(0, len(ramais_ip_candidatos), chunk_size):
        chunk = ramais_ip_candidatos[i:i + chunk_size]
        res_chunk = fetch_oxe_batch_tsc_ip_js(driver, chunk)
        if res_chunk:
            tsc_ip_map.update(res_chunk)
        logger.info(f"⚡ Lote IP {i + len(chunk)}/{len(ramais_ip_candidatos)} processado.")

    logger.info(f"✅ Detalhes IP/MAC coletados para {len(tsc_ip_map)} telefones IP.")

    registros_finais = []

    for idx, sub in enumerate(subscribers, start=1):
        try:
            num_ramal = str(sub.get("Directory_Number", "")).strip()
            if not num_ramal:
                continue

            name = str(sub.get("Annu_Name", "")).strip()
            first_name = str(sub.get("Annu_First_Name", "")).strip()
            utf8_name = str(sub.get("UTF8_Phone_Book_Name", "")).strip()
            utf8_first = str(sub.get("UTF8_Phone_Book_First_Name", "")).strip()
            
            nome_completo = f"{name} {first_name}".strip()
            if not nome_completo and utf8_first:
                nome_completo = utf8_first

            station_type = str(sub.get("Station_Type", "")).strip()
            station_sub_type = str(sub.get("Station_Sub_Type", "")).strip()
            mail = str(sub.get("Mail_Address", "")).strip()
            ext_login = str(sub.get("External_Login", "")).strip()
            set_role = str(sub.get("Set_Role", "")).strip()

            rack = sub.get("Equipment_Address_Rack", "")
            board = sub.get("Equipment_Address_Board", "")
            terminal = sub.get("Equipment_Address_Terminal", "")

            registro = {}

            detalhes_ramal = subscriber_details_map.get(num_ramal, {})
            if isinstance(detalhes_ramal, dict):
                for k, v in detalhes_ramal.items():
                    if isinstance(v, (dict, list)):
                        registro[k] = json.dumps(v, ensure_ascii=False)
                    else:
                        registro[k] = v

            registro["Ramal"] = num_ramal
            registro["Nome / Titular"] = name
            registro["Complemento"] = first_name or utf8_first
            registro["Tipo de Estação"] = station_type
            registro["Subtipo"] = station_sub_type
            registro["Função / Role"] = set_role
            registro["Grupo de Captura"] = str(detalhes_ramal.get("Pickup_Group_Name", "")).strip()
            cat_pub = detalhes_ramal.get("Public_Network_Category_Id", "")
            registro["Cat. Rede Pública"] = cat_pub if cat_pub != 255 else "-"
            registro["Login Externo"] = ext_login
            registro["E-mail"] = mail
            registro["Rack"] = rack if rack != 255 else "-"
            registro["Placa"] = board if board != 255 else "-"
            registro["Terminal"] = terminal if terminal != 255 else "-"

            ip_address = ""
            mac_address = ""
            
            tsc_data = tsc_ip_map.get(num_ramal)
            if tsc_data and isinstance(tsc_data, dict):
                ip_address = tsc_data.get("IP_Address") or tsc_data.get("IPv6_address") or ""
                mac_address = str(tsc_data.get("Ethernet_Address", "")).upper()

            registro["Endereço IP"] = ip_address
            registro["MAC Address"] = mac_address

            registros_finais.append(registro)

        except Exception as item_err:
            logger.warning(f"Erro ao processar assinante índice {idx}: {item_err}")
            continue

    return registros_finais

def scrape_oxe():
    print_header("SCRAPER OXE - CENTRAL TELEFÔNICA", color=CYAN)
    logger.info("🤖 Iniciando raspagem de ramais do OXE...")
    
    if not OXE_PASS:
        logger.error("❌ A senha do OXE (OXE_PASS) não foi configurada!")
        return False

    driver = None

    try:
        driver, wait = initial_config()
        realizar_login_oxe(driver, wait)

        dados = extrair_dados_assinantes(driver)

        if not dados:
            logger.error("Nenhum registro foi coletado da Central Telefônica.")
            return False

        out_dir = Path("01 - Dados Brutos")
        out_dir.mkdir(exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        excel_file = out_dir / f"Central_Telefonica_OXE_{ts}.xlsx"

        df = pd.DataFrame(dados)

        widths = {
            "Ramal": 12,
            "Nome / Titular": 25,
            "Complemento": 25,
            "Tipo de Estação": 25,
            "Subtipo": 18,
            "Função / Role": 20,
            "Grupo de Captura": 18,
            "Cat. Rede Pública": 18,
            "Login Externo": 18,
            "Item": 8, "Extension": 12, "Node": 8, "SetType": 18,
            "DirName": 30, "SubNet": 10, "Domain": 10, "AnnuName": 30
        }

        save_df_to_excel_formatted(
            df, excel_file, sheet_name="Central Telefônica",
            widths=widths
        )

        logger.info(f"✅ SUCESSO! Total de {len(dados)} ramais salvos em: {excel_file.name}")

        cleanup_old_files(out_dir, "Central_Telefonica_OXE_*.xlsx", keep_count=10)
        return True

    except Exception as e:
        logger.error(f"❌ Erro crítico no Scraper OXE: {e}", exc_info=True)
        if driver:
            salvar_screenshot(driver, "erro_critico")
        return False

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

def check_oxe_sync_running() -> bool:
    import tempfile
    lock_file = Path(tempfile.gettempdir()) / "oxe_scraper.lock"
    return check_process_running(lock_file)

def read_oxe_last_log_lines(n: int = 15) -> str:
    log_path = Path("debug_logs") / "oxe" / "oxe_scraper.log"
    return read_log_lines(log_path, n)

if __name__ == "__main__":
    import os
    import tempfile

    lock_path = Path(tempfile.gettempdir()) / "oxe_scraper.lock"
    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))

    try:
        success = scrape_oxe()
        if success:
            try:
                from src.preprocess_oxe import preprocess_oxe
                preprocess_oxe()
            except Exception as pe:
                logger.error(f"Erro ao executar pré-processamento do OXE: {pe}")
            sys.exit(0)
        else:
            sys.exit(1)
    finally:
        if lock_path.exists():
            try:
                lock_path.unlink()
            except Exception:
                pass
