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

from src.components.status_banner import check_process_running, read_log_lines
from src.config import (
    OXE_URL, OXE_USER, OXE_PASS,
    HEADLESS, EXPLICIT_WAIT, DEBUG_DIR_OXE,
    setup_logging, save_df_to_excel_formatted,
    cleanup_old_files
)
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE


logger = setup_logging(DEBUG_DIR_OXE / "oxe_scraper.log", __name__)

logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

def salvar_screenshot(page, nome_etapa: str):
    try:
        ts = datetime.now().strftime("%H-%M-%S")
        filename = f"debug_oxe_{ts}_{nome_etapa}.png"
        filepath = DEBUG_DIR_OXE / filename
        page.screenshot(path=str(filepath))
        logger.debug(f"📸 Screenshot salvo em: {filepath.name}")
    except Exception as e:
        logger.warning(f"Não foi possível salvar screenshot '{nome_etapa}': {e}")

def realizar_login_oxe(page):
    target_url = OXE_URL if OXE_URL.endswith("/") else f"{OXE_URL}/"
    login_full_url = f"{target_url}#/login"
    
    logger.info(f"Navegando para a página de login do OXE: {login_full_url}")
    page.goto(login_full_url, wait_until="domcontentloaded", timeout=60000)
    time.sleep(2)

    logger.info("Aguardando formulário de login...")
    try:
        user_input = page.locator("#username, input[type='text']").first
        user_input.wait_for(state="visible", timeout=30000)
        user_input.fill(OXE_USER)
    except Exception as e:
        logger.warning(f"Aviso ao preencher usuário OXE: {e}")

    try:
        pwd_input = page.locator("#password, input[type='password']").first
        pwd_input.fill(OXE_PASS)
    except Exception as e:
        logger.warning(f"Aviso ao preencher senha OXE: {e}")

    logger.info(f"Efetuando login no OXE com o usuário '{OXE_USER}'...")
    login_btn = page.locator("#login-button, button[type='submit']").first
    if login_btn.count() > 0:
        login_btn.click()

    time.sleep(3)
    logger.info(f"URL após login: {page.url}")
    
    logger.info("Aguardando gravação de tokens de sessão no storage...")
    token_capturado = None
    for attempt in range(15):
        token_capturado = page.evaluate(
            "() => sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__;"
        )
        if token_capturado:
            logger.info(f"🔑 Token capturado com sucesso! (Tamanho: {len(token_capturado)})")
            break
        time.sleep(0.5)

    if not token_capturado:
        salvar_screenshot(page, "falha_login_token")
        storage_dump = page.evaluate("""() => {
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
        }""")
        logger.warning(f"⚠️ Token não encontrado. Conteúdo atual dos Storages: {storage_dump}")

    try:
        elements = page.locator("xpath=//*[contains(translate(text(), 'UTILIZADORES', 'utilizadores'), 'utilizadores') or contains(text(), 'Subscriber')]")
        if elements.count() > 0:
            logger.info(f"Encontrado item de menu de Utilizadores ({elements.count()}). Clicando para inicializar...")
            elements.first.click()
            time.sleep(2)
    except Exception as nav_err:
        logger.debug(f"Interação de menu opcional: {nav_err}")

def fetch_oxe_api_js(page, api_endpoint: str):
    js_code = """async (endpoint) => {
        var idToken = sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__;

        if (!idToken) {
            for (var s of [sessionStorage, localStorage]) {
                if (!s) continue;
                for (var i = 0; i < s.length; i++) {
                    var k = s.key(i);
                    if (k.toLowerCase().includes('token')) {
                        idToken = s.getItem(k);
                        if (idToken) break;
                    }
                }
                if (idToken) break;
            }
        }

        if (!idToken) {
            return { status: 'error', message: 'Nenhum id_token encontrado no sessionStorage/localStorage' };
        }

        var tokenBearer = idToken.startsWith('Bearer ') ? idToken : 'Bearer ' + idToken;
        var tokenRaw = idToken.replace(/^Bearer\\s+/i, '');

        try {
            var response = await fetch(endpoint, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json, text/plain, */*',
                    'Authorization': tokenBearer
                }
            });
            if (response.status === 401) {
                response = await fetch(endpoint, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json, text/plain, */*',
                        'Authorization': tokenRaw
                    }
                });
            }
            if (!response.ok) {
                return { status: 'error', message: 'HTTP ' + response.status };
            }
            var data = await response.json();
            return { status: 'success', data: data };
        } catch (err) {
            return { status: 'error', message: err.toString() };
        }
    }"""
    res = page.evaluate(js_code, api_endpoint)
    if res and res.get("status") == "success":
        return res.get("data")
    else:
        err_msg = res.get("message") if res else "Sem resposta do script JS"
        logger.debug(f"Erro na requisição JS para {api_endpoint}: {err_msg}")
        return None

def fetch_oxe_batch_subscriber_details_js(page, ramais_list: list) -> dict:
    if not ramais_list:
        return {}

    js_code = """async (ramais) => {
        var idToken = sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__ || '';
        var tokenBearer = idToken.startsWith('Bearer ') ? idToken : (idToken ? 'Bearer ' + idToken : '');

        var results = {};
        var promises = ramais.map(async function(num_ramal) {
            try {
                var url = '/api/mgt/1.0/Node/1/Subscriber/' + num_ramal;
                var res = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json, text/plain, */*',
                        'Authorization': tokenBearer
                    }
                });
                if (res.ok) {
                    var data = await res.json();
                    if (data) results[num_ramal] = data;
                }
            } catch(e) {}
        });

        await Promise.all(promises);
        return { status: 'success', data: results };
    }"""
    res = page.evaluate(js_code, ramais_list)
    if res and res.get("status") == "success":
        return res.get("data", {})
    return {}

def fetch_oxe_batch_tsc_ip_js(page, ramais_list: list) -> dict:
    if not ramais_list:
        return {}

    js_code = """async (ramais) => {
        var idToken = sessionStorage.getItem('id_token') || localStorage.getItem('id_token') || window.__captured_auth_token__ || '';
        var tokenBearer = idToken.startsWith('Bearer ') ? idToken : (idToken ? 'Bearer ' + idToken : '');

        var results = {};
        var promises = ramais.map(async function(num_ramal) {
            try {
                var url = '/api/mgt/1.0/Node/1/Subscriber/' + num_ramal + '/Tsc_IP_subscriber/' + num_ramal;
                var res = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json, text/plain, */*',
                        'Authorization': tokenBearer
                    }
                });
                if (res.ok) {
                    var data = await res.json();
                    if (data) results[num_ramal] = data;
                }
            } catch(e) {}
        });

        await Promise.all(promises);
        return { status: 'success', data: results };
    }"""
    res = page.evaluate(js_code, ramais_list)
    if res and res.get("status") == "success":
        return res.get("data", {})
    return {}

def extrair_dados_assinantes(page):
    subscribers_endpoint = (
        "/api/mgt/1.0/Node/1/Subscriber?attributes="
        "Annu_Name,Annu_First_Name,UTF8_Phone_Book_Name,UTF8_Phone_Book_First_Name,"
        "Equipment_Address_Rack,Equipment_Address_Board,Equipment_Address_Terminal,"
        "Station_Type,Opex_License,External_Login,Mail_Address,Station_Sub_Type,"
        "DM_Profile,Entity_Number,Set_Role"
    )
    
    logger.info("Solicitando lista principal de utilizadores/assinantes à API do OXE...")
    subscribers = fetch_oxe_api_js(page, subscribers_endpoint)
    
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
        res_chunk = fetch_oxe_batch_subscriber_details_js(page, chunk)
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
        res_chunk = fetch_oxe_batch_tsc_ip_js(page, chunk)
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

    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=HEADLESS,
                args=["--no-sandbox", "--disable-dev-shm-usage", "--ignore-certificate-errors"]
            )
            context = browser.new_context(ignore_https_errors=True)
            page = context.new_page()

            def handle_req(request):
                for h, v in request.headers.items():
                    if h.lower() in ['authorization', 'x-auth-token', 'token']:
                        try:
                            page.evaluate("(val) => { window.__captured_auth_token__ = val; }", v)
                        except Exception:
                            pass
            page.on("request", handle_req)

            realizar_login_oxe(page)
            dados = extrair_dados_assinantes(page)
            browser.close()

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
        return False

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
