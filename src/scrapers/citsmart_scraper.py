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

from datetime import datetime
import time
import pandas as pd
import re
import logging
import json
import html
from ldap3 import SUBTREE
from config import (
    CITSMART_URL, CITSMART_EMAIL, PASSWORD,
    HEADLESS, EXPLICIT_WAIT, DEBUG_DIR_CITSMART,
    setup_logging, save_df_to_excel_formatted,
    setup_ad_connection, fetch_ad_department, cleanup_old_files
)
from src.terminal import log, print_header, CYAN, GREEN, RED, YELLOW, WHITE

# ---------------------------
# Filtros e Constantes
# ---------------------------
GRUPO_MANUTENCAO = "[N1] Manutenção"
USUARIOS_AUTOMATICOS_IGNORAR = {
    "monitoramento adm mpms",
    "adm ticket por email"
}

logger = setup_logging(DEBUG_DIR_CITSMART / "citsmart_scraper.log", __name__)

logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)

def salvar_screenshot(page, nome_etapa):
    ts = datetime.now().strftime("%H-%M-%S")
    filename = f"debug_citsmart_{ts}_{nome_etapa}.png"
    
    if any(k in nome_etapa.lower() for k in ["erro", "falha", "timeout"]):
        error_dir = DEBUG_DIR_CITSMART / "errors"
        error_dir.mkdir(parents=True, exist_ok=True)
        filepath = error_dir / filename
    else:
        filepath = DEBUG_DIR_CITSMART / filename
        
    try:
        page.screenshot(path=str(filepath))
        logger.debug(f"📸 Screenshot salvo: {filepath.name}")
    except Exception as e:
        logger.warning(f"Não foi possível salvar screenshot: {e}")

def fetch_setor_temp(conn, query_val, is_username=False):
    return fetch_ad_department(conn, query_val, is_username=is_username)

def navigate_to_caixa_entrada(page):
    logger.info("Acessando CitSmart e autenticando via Microsoft Entra ID...")
    page.goto(CITSMART_URL, wait_until="domcontentloaded", timeout=60000)

    try:
        login_input = page.locator("input[name='loginfmt'], input[type='email']").first
        login_input.wait_for(state="visible", timeout=30000)
        login_input.fill(CITSMART_EMAIL)
        
        btn_next = page.locator("#idSIButton9, input[type='submit']").first
        btn_next.click()
        
        pwd_input = page.locator("input[name='passwd'], input[type='password']").first
        pwd_input.wait_for(state="visible", timeout=30000)
        time.sleep(0.5)
        pwd_input.fill(PASSWORD or "")
        
        btn_submit = page.locator("#idSIButton9, input[type='submit']").first
        btn_submit.click()

        try:
            kmsi_btn = page.locator("#idSIButton9").first
            kmsi_btn.wait_for(state="visible", timeout=6000)
            logger.info("Pulando tela KMSI de manter conectado...")
            kmsi_btn.click()
        except Exception:
            pass

        logger.info("Aguardando carregamento do portal inicial...")
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        time.sleep(3)

    except Exception as login_err:
        logger.info(f"ℹ️ Fluxo de login concluído ou sessão ativa detectada: {login_err}")

    nova_fila = "https://suporte.mpms.mp.br/inbox/lowcode/form/copilot_novo/default"
    logger.info(f"Forçando navegação para: {nova_fila}")
    page.goto(nova_fila, wait_until="domcontentloaded", timeout=60000)
    time.sleep(3)

def expand_all_records_lowcode(frame):
    logger.info("Tentando expandir registros (LowCode)...")
    try:
        time.sleep(2)
        dropdown = frame.locator("#pageSize").first
        dropdown.wait_for(state="visible", timeout=30000)
        
        try:
            current = dropdown.input_value()
        except Exception:
            current = ""

        if "100" in current:
            logger.info("Já está exibindo 100 registros. Forçando toggle (100 -> 50 -> 100)...")
            dropdown.select_option("50")
            time.sleep(1)
            dropdown.select_option("100")
            time.sleep(1)
        else:
            dropdown.select_option("100")
            logger.info("Sucesso! Paginação alterada para 100 itens.")
            time.sleep(2)

        try:
            pager_info = frame.locator("div[ng-if='totalTickets']").first
            if pager_info.count() > 0:
                text = (pager_info.text_content() or "").strip()
                logger.info(f"Paginação atualizada: {text}")
                match = re.search(r"de\s+(\d+)", text)
                if match:
                    return int(match.group(1))
        except Exception as e:
            logger.error(f"Aviso: Não consegui ler o texto da paginação. Erro: {e}")

        return 0
    except Exception as e:
        logger.error(f"Aviso ao expandir registros: {e}")
        return 0

def clean_html_comment(html_str):
    if not html_str:
        return ""
    txt = re.sub(r'<[^>]*>', ' ', html_str)
    txt = html.unescape(txt)
    lines = [line.strip() for line in txt.splitlines() if line.strip()]
    return "\n".join(lines)

def process_page(frame, captured_tickets, ad_conn=None, cache=None):
    # 1. Obtém os IDs realmente exibidos na tabela da tela para evitar chamados de outras filas/páginas
    dom_cids = set()
    try:
        rows = frame.locator("#table tbody tr")
        row_count = rows.count()
        for idx in range(row_count):
            try:
                el = rows.nth(idx).locator("xpath=.//div[@ng-switch-when='1']").first
                if el.count() > 0:
                    text_c = el.text_content() or ""
                    m = re.search(r'\d+', text_c)
                    if m:
                        dom_cids.add(m.group(0))
            except Exception:
                pass
        if dom_cids:
            logger.info(f"🎯 IDs legítimos detectados na tabela DOM da tela ({len(dom_cids)}): {sorted(list(dom_cids))}")
    except Exception as dom_err:
        logger.debug(f"Aviso ao ler IDs do DOM: {dom_err}")

    # Processamento a partir de tickets interceptados pela rede JSON (Processo Ultra-Rápido)
    if captured_tickets and isinstance(captured_tickets, list) and len(captured_tickets) > 0:
        candidates = captured_tickets
        if dom_cids:
            filtered = [t for t in candidates if str(t.get("ticket_id", "") or t.get("id", "")) in dom_cids]
            if filtered:
                candidates = filtered
                logger.info(f"⚡ [PROCESSO ULTRA-RÁPIDO] Filtrado para {len(candidates)} chamados correspondentes à tabela da fila!")
            else:
                logger.info(f"⚡ [PROCESSO ULTRA-RÁPIDO] Processando {len(candidates)} chamados capturados diretamente via JSON de rede!")
        
        collected = []
        for idx, ticket in enumerate(candidates):
            try:
                cid = str(ticket.get("ticket_id", "") or ticket.get("id", ""))
                if not cid:
                    continue

                # Filtro de Grupo Atual
                grupo_atual = str(
                    ticket.get("grupo_atual") or 
                    ticket.get("current_group") or 
                    ticket.get("ticket_group") or 
                    ticket.get("nome_grupo") or 
                    ticket.get("group_name") or 
                    ticket.get("sigla_grupo") or ""
                ).strip()
                if grupo_atual and "[n1] manutenção" not in grupo_atual.lower() and "manutenção" not in grupo_atual.lower():
                    logger.info(f"⏭️ [IGNORADO GRUPO] Chamado {cid} pertence a outro grupo: '{grupo_atual}'")
                    continue

                solicitante_nome = ticket.get("ticket_requester", "") or ticket.get("solicitante", "")
                if solicitante_nome and solicitante_nome.strip().lower() in USUARIOS_AUTOMATICOS_IGNORAR:
                    logger.info(f"⏭️ [IGNORADO AUTOMÁTICO] Chamado {cid} criado pelo usuário automático '{solicitante_nome}'")
                    continue

                id_cliente = ""
                email_solicitante = ticket.get("email_solicitante", "")
                if email_solicitante and "@" in email_solicitante:
                    id_cliente = email_solicitante.split("@")[0].strip()

                data_criacao = ""
                iso_str = ticket.get("ticket_creationdate_str", "") or ticket.get("data_criacao", "")
                if iso_str:
                    try:
                        clean_iso = iso_str.split(".")[0].replace("Z", "")
                        dt = datetime.strptime(clean_iso, "%Y-%m-%dT%H:%M:%S")
                        data_criacao = dt.strftime("%d/%m/%Y %H:%M")
                    except Exception:
                        data_criacao = iso_str

                comments_list = []
                if cache and cid in cache:
                    cached_comments_str = cache[cid].get('Comentários', '[]')
                    try:
                        comments_list = json.loads(cached_comments_str)
                    except Exception:
                        pass

                raw_comments = ticket.get("comentarios") or ticket.get("comments") or []
                if isinstance(raw_comments, list):
                    for comm in raw_comments:
                        texto_limpo = clean_html_comment(comm.get("texto") or comm.get("comment") or "")
                        if not texto_limpo:
                            continue
                        new_comment = {
                            "data": comm.get("data") or comm.get("date") or "",
                            "autor": comm.get("autor") or comm.get("author") or "",
                            "texto": texto_limpo
                        }
                        if not any(e.get("data") == new_comment["data"] and e.get("texto") == new_comment["texto"] for e in comments_list):
                            comments_list.append(new_comment)

                comments_json = json.dumps(comments_list, ensure_ascii=False)

                if cache and cid in cache:
                    cached_ip = cache[cid].get('IP_Origem') or ""
                    cached_unidade = cache[cid].get('Unidade') or ""
                    cached_desc = cache[cid].get('Descrição') or ""
                    cached_hostname = cache[cid].get('Hostname') or ""
                    if cached_ip:
                        collected.append({
                            "Chamado#": cid,
                            "ID do Cliente": id_cliente,
                            "Nome do Usuário": solicitante_nome,
                            "Unidade": cached_unidade or "Não encontrada no AD",
                            "Descrição": cached_desc,
                            "Data Criação": data_criacao,
                            "IP_Origem": cached_ip,
                            "Hostname": cached_hostname,
                            "Link": f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}",
                            "Comentários": comments_json
                        })
                        logger.info(f"[{idx+1}/{len(candidates)}] ⚡ [CACHE MATCH] Chamado {cid} recuperado do cache!")
                        continue
                    else:
                        sccm_ip = ""
                        sccm_hostname = ""
                        if id_cliente:
                            from config import fetch_ip_from_sccm, fetch_hostname_from_sccm
                            sccm_ip = fetch_ip_from_sccm(id_cliente)
                            sccm_hostname = fetch_hostname_from_sccm(id_cliente)
                        collected.append({
                            "Chamado#": cid,
                            "ID do Cliente": id_cliente,
                            "Nome do Usuário": solicitante_nome,
                            "Unidade": cached_unidade or "Não encontrada no AD",
                            "Descrição": cached_desc,
                            "Data Criação": data_criacao,
                            "IP_Origem": sccm_ip,
                            "Hostname": sccm_hostname,
                            "Link": f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}",
                            "Comentários": comments_json
                        })
                        continue

                localidade = "Não encontrada no AD"
                if ad_conn:
                    if id_cliente:
                        localidade = fetch_setor_temp(ad_conn, id_cliente, is_username=True)
                    elif solicitante_nome:
                        localidade = fetch_setor_temp(ad_conn, solicitante_nome, is_username=False)
                else:
                    localidade = ticket.get("ticket_unit", "") or ticket.get("nome_unidade", "Não encontrada")

                sccm_ip = ""
                sccm_hostname = ""
                if id_cliente:
                    from config import fetch_ip_from_sccm, fetch_hostname_from_sccm
                    sccm_ip = fetch_ip_from_sccm(id_cliente)
                    sccm_hostname = fetch_hostname_from_sccm(id_cliente)

                collected.append({
                    "Chamado#": cid,
                    "ID do Cliente": id_cliente,
                    "Nome do Usuário": solicitante_nome,
                    "Unidade": localidade,
                    "Descrição": ticket.get("ticket_description_long", "") or ticket.get("ticket_description", ""),
                    "Data Criação": data_criacao,
                    "IP_Origem": sccm_ip,
                    "Hostname": sccm_hostname,
                    "Link": f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}",
                    "Comentários": comments_json
                })
                logger.info(f"[{idx+1}/{len(candidates)}] Processado JSON: {cid} (Login: {id_cliente})")
            except Exception as row_err:
                logger.error(f"Erro ao parsear ticket do JSON: {row_err}")
                continue

        if collected:
            return collected

    # 2. Fallback: Extração via DOM
    rows = frame.locator("#table tbody tr")
    row_count = rows.count()
    logger.info(f"Linhas detectadas no DOM: {row_count}")
    collected = []

    for idx in range(row_count):
        try:
            row = rows.nth(idx)
            def get_val(key, is_description=False):
                try:
                    el = row.locator(f"xpath=.//div[@ng-switch-when='{key}']").first
                    if el.count() > 0:
                        if is_description:
                            span = el.locator("span").first
                            if span.count() > 0:
                                return span.get_attribute("title") or (span.text_content() or "").strip()
                        return (el.text_content() or "").strip()
                except Exception:
                    pass
                return ""

            num_bruto = get_val("1")
            num_match = re.search(r'\d+', num_bruto)
            cid = num_match.group(0) if num_match else ""

            solicitante_full = get_val("6")
            data_criacao = get_val("9")

            solicitante_nome = solicitante_full
            id_cliente = ""
            if "(" in solicitante_full:
                try:
                    partes = solicitante_full.split("(")
                    solicitante_nome = partes[0].strip()
                    id_cliente = partes[1].replace(")", "").strip()
                except Exception:
                    pass

            grupo_atual = ""
            for switch_key in ["11", "12", "13", "14", "15", "16", "17", "18"]:
                val_td = get_val(switch_key)
                if val_td and ("[n1]" in val_td.lower() or "manutenção" in val_td.lower()):
                    grupo_atual = val_td
                    break

            if grupo_atual and "[n1] manutenção" not in grupo_atual.lower() and "manutenção" not in grupo_atual.lower():
                continue

            if solicitante_nome and solicitante_nome.strip().lower() in USUARIOS_AUTOMATICOS_IGNORAR:
                continue

            comments_json = cache[cid].get('Comentários', '[]') if (cache and cid in cache) else '[]'

            if cache and cid in cache:
                cached_ip = cache[cid].get('IP_Origem') or ""
                cached_unidade = cache[cid].get('Unidade') or ""
                cached_desc = cache[cid].get('Descrição') or ""
                cached_hostname = cache[cid].get('Hostname') or ""
                if cached_ip:
                    collected.append({
                        "Chamado#": cid,
                        "ID do Cliente": id_cliente,
                        "Nome do Usuário": solicitante_nome,
                        "Unidade": cached_unidade or "Não encontrada no AD",
                        "Descrição": cached_desc,
                        "Data Criação": data_criacao,
                        "IP_Origem": cached_ip,
                        "Hostname": cached_hostname,
                        "Link": f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}",
                        "Comentários": comments_json
                    })
                    continue

            descricao = get_val("10", is_description=True)

            localidade = "Não encontrada no AD"
            if ad_conn:
                if id_cliente:
                    localidade = fetch_setor_temp(ad_conn, id_cliente, is_username=True)
                else:
                    localidade = fetch_setor_temp(ad_conn, solicitante_nome, is_username=False)

            sccm_ip = ""
            sccm_hostname = ""
            if id_cliente:
                from config import fetch_ip_from_sccm, fetch_hostname_from_sccm
                sccm_ip = fetch_ip_from_sccm(id_cliente)
                sccm_hostname = fetch_hostname_from_sccm(id_cliente)

            collected.append({
                "Chamado#": cid,
                "ID do Cliente": id_cliente,
                "Nome do Usuário": solicitante_nome,
                "Unidade": localidade,
                "Descrição": descricao,
                "Data Criação": data_criacao,
                "IP_Origem": sccm_ip,
                "Hostname": sccm_hostname,
                "Link": f"https://suporte.mpms.mp.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load?iframe=true&language=pt-BR#/request?idRequest={cid}",
                "Comentários": comments_json
            })
            logger.info(f"[{idx+1}/{row_count}] Lido DOM: {cid} (Login: {id_cliente})")
        except Exception:
            continue

    return collected

def ir_para_proxima_pagina(frame):
    try:
        btn_next = frame.locator("li.pagination-next").first
        if btn_next.count() == 0:
            return False
            
        cls = btn_next.get_attribute("class") or ""
        if "disabled" in cls:
            logger.info("Paginação encerrada: Botão 'Próximo' está desabilitado.")
            return False

        link_next = btn_next.locator("a").first
        if link_next.count() > 0:
            link_next.click()
            logger.info("Navegando para a próxima página...")
            time.sleep(3)
            return True
        return False
    except Exception as e:
        logger.error(f"Fim da paginação ou erro: {e}")
        return False

def scrape_citsmart():
    print_header("SCRAPER CITSMART - COLETA DE CHAMADOS", color=CYAN)
    logger.info("🤖 Iniciando processo de raspagem do CitSmart...")
    
    cache = {}
    try:
        out_dir = Path("01 - Dados Brutos")
        existing_files = sorted(out_dir.glob("Chamados_CitSmart_*.xlsx"))
        if existing_files:
            latest_file = existing_files[-1]
            logger.info(f"📂 Carregando cache de descrições e comentários do arquivo mais recente: {latest_file.name}")
            df_old = pd.read_excel(latest_file, dtype=str)
            for _, row_old in df_old.iterrows():
                cid = str(row_old.get('Chamado#', '')).strip()
                desc = row_old.get('Descrição', '')
                ip = row_old.get('IP_Origem', '')
                link = row_old.get('Link', '')
                comments = row_old.get('Comentários', '[]')
                hostname = row_old.get('Hostname', '')
                if cid:
                    clean_ip = str(ip).strip() if pd.notna(ip) and str(ip).strip().lower() not in ["nan", "none", "null"] else ''
                    clean_host = str(hostname).strip() if pd.notna(hostname) and str(hostname).strip().lower() not in ["nan", "none", "null"] else ''
                    clean_desc = str(desc).strip() if pd.notna(desc) and str(desc).strip().lower() not in ["nan", "none", "null"] else ''
                    clean_unid = str(row_old.get('Unidade', '')).strip() if pd.notna(row_old.get('Unidade')) and str(row_old.get('Unidade')).strip().lower() not in ["nan", "none", "null"] else ''
                    clean_link = str(link).strip() if pd.notna(link) and str(link).strip().lower() not in ["nan", "none", "null"] else ''
                    clean_comm = str(comments).strip() if pd.notna(comments) and str(comments).strip().lower() not in ["nan", "none", "null"] else '[]'
                    
                    cache[cid] = {
                        'Descrição': clean_desc,
                        'Unidade': clean_unid,
                        'IP_Origem': clean_ip,
                        'Hostname': clean_host,
                        'Link': clean_link,
                        'Comentários': clean_comm
                    }
            logger.info(f"⚡ Sucesso! {len(cache)} chamados carregados no cache do CitSmart.")
    except Exception as cache_err:
        logger.warning(f"⚠️ Aviso: Não foi possível carregar cache do CitSmart: {cache_err}")

    ad_conn = None
    try:
        ad_conn = setup_ad_connection()
        logger.info("Conexão AD estabelecida.")
    except Exception as e:
        logger.error(f"AD indisponível: {e}")

    todos_os_dados = []

    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=HEADLESS,
                args=["--no-sandbox", "--disable-dev-shm-usage", "--ignore-certificate-errors"]
            )
            context = browser.new_context(ignore_https_errors=True)
            page = context.new_page()

            captured_network_tickets = []
            def on_response(response):
                try:
                    url = response.url.lower()
                    if "tb_ticket_queue/list" in url:
                        ct = response.headers.get("content-type", "")
                        if "json" in ct or "text" in ct:
                            data = response.json()
                            candidate = data if isinstance(data, list) else (data.get("list", []) if isinstance(data, dict) else [])
                            if candidate:
                                captured_network_tickets.extend(candidate)
                                logger.info(f"🔥 [PLAYWRIGHT NETWORK] {len(candidate)} chamados interceptados da API CitSmart!")
                except Exception:
                    pass

            page.on("response", on_response)

            logger.info("🔑 Navegando até a caixa de entrada do CitSmart via Playwright...")
            navigate_to_caixa_entrada(page)

            frame = page.frame_locator("iframe[title='App']").first
            expand_all_records_lowcode(frame)

            pagina = 1
            while True:
                logger.info(f"📄 --- Processando Página {pagina} do CitSmart ---")
                
                dados_pagina = process_page(frame, captured_network_tickets, ad_conn=ad_conn, cache=cache)
                if dados_pagina:
                    todos_os_dados.extend(dados_pagina)
                    logger.info(f"Coletados {len(dados_pagina)} registros nesta página.")
                else:
                    logger.info("Aviso: Página retornou 0 registros.")

                if not ir_para_proxima_pagina(frame):
                    break
                
                pagina += 1

            browser.close()
            logger.info("Navegador Playwright encerrado com sucesso.")

        if todos_os_dados:
            out_dir = Path("01 - Dados Brutos")
            out_dir.mkdir(exist_ok=True)
            ts = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
            file = out_dir / f"Chamados_CitSmart_{ts}.xlsx"

            df = pd.DataFrame(todos_os_dados)
            df = df.fillna("")
            widths = {
                'Chamado#': 15,
                'ID do Cliente': 20,
                'Nome do Usuário': 25,
                'Unidade': 40,
                'Descrição': 100,
                'Data Criação': 20,
                'IP_Origem': 15,
                'Hostname': 20,
                'Link': 40,
                'Comentários': 50
            }
            save_df_to_excel_formatted(
                df, file, sheet_name="Chamados",
                widths=widths, wrap_cols=['Descrição', 'Comentários'], height_col='Descrição'
            )

            logger.info(f"✅ SUCESSO! Total de {len(todos_os_dados)} chamados salvos em: {file.name}")
            cleanup_old_files(out_dir, "Chamados_CitSmart_*.xlsx", keep_count=10)
            return True
        else:
            logger.warning("⚠️ Nenhum dado foi coletado do CitSmart.")
            return False

    except Exception as e:
        logger.error(f"❌ Erro crítico no Scraper CitSmart: {e}", exc_info=True)
        return False

if __name__ == "__main__":
    if scrape_citsmart():
        sys.exit(0)
    else:
        sys.exit(1)
