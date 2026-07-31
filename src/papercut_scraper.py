import os
import sys
import time
import shutil
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Garante importações dos módulos do projeto
root_dir = Path(__file__).parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))
if str(root_dir / "src") not in sys.path:
    sys.path.insert(0, str(root_dir / "src"))

from src.config import (
    PAPERCUT_URL, PAPERCUT_PRINTER_LIST_URL, PAPERCUT_DEVICE_LIST_URL,
    PAPERCUT_USER, PAPERCUT_PASS, DEBUG_DIR_PAPERCUT, HEADLESS,
    setup_logging, get_chrome_driver, USER_HOME, INPUT_DIR_BRUTOS,
    cleanup_old_files
)
from src.database import save_impressoras_to_db

logger = setup_logging(DEBUG_DIR_PAPERCUT / "papercut_scraper.log", __name__)

DOWNLOAD_DIR = USER_HOME / "Downloads"
PAPERCUT_BRUTOS_DIR = INPUT_DIR_BRUTOS / "papercut"
PAPERCUT_BRUTOS_DIR.mkdir(parents=True, exist_ok=True)



def cleanup_old_papercut_downloads():
    """Remove arquivos antigos do tipo printer_list*.csv e device_list*.csv na pasta Downloads."""
    logger.info("🧹 Limpando arquivos CSV anteriores do PaperCut na pasta Downloads...")
    for pattern in ["printer_list*.csv", "device_list*.csv"]:
        for f in DOWNLOAD_DIR.glob(pattern):
            try:
                f.unlink()
                logger.info(f"  └─ Arquivo limpo: {f.name}")
            except Exception as e:
                logger.warning(f"  └─ Não foi possível apagar {f.name}: {e}")


def clean_encoding_text(val):
    """Corrige encodings corrompidos (Latin1 / UTF-8 misturados ou bytes mal decodificados)."""
    if not isinstance(val, str):
        return val
    val = val.strip()
    if not val or val.lower() == 'nan':
        return ''
    
    try:
        if any(c in val for c in ['Ã', 'Â', 'ï¿½', 'Ã§', 'Ã£', 'Ã©', 'Ã³', 'Ã¡', 'Ãº']):
            val = val.encode('latin1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    
    return val


def load_csv_safe(file_path: Path) -> pd.DataFrame:
    """Carrega um arquivo CSV do PaperCut (usando ';' como separador e '#' como comentário)."""
    logger.info(f"--- Tentando carregar arquivo CSV: {file_path} ---")
    if not file_path.exists():
        logger.warning(f"❌ Arquivo não encontrado no disco: {file_path}")
        return pd.DataFrame()

    file_size = file_path.stat().st_size
    logger.info(f"Tamanho do arquivo '{file_path.name}': {file_size} bytes")
    if file_size == 0:
        logger.warning(f"⚠️ O arquivo '{file_path.name}' está vazio (0 bytes).")
        return pd.DataFrame()

    encodings = ['latin1', 'utf-8-sig', 'cp1252', 'iso-8859-1', 'utf-16']
    df = None
    
    # 1. Tenta carregar com separador ';' e ignorando linhas de comentário '#'
    for enc in encodings:
        try:
            temp_df = pd.read_csv(file_path, encoding=enc, sep=';', comment='#', engine='python')
            if len(temp_df.columns) >= 2:
                df = temp_df
                logger.info(f"✅ CSV '{file_path.name}' decodificado com sep=';' e comment='#' em '{enc}'. Linhas: {len(df)}, Colunas: {len(df.columns)}")
                break
        except Exception as e:
            logger.debug(f"Falha ao tentar sep=';' com encoding '{enc}' em '{file_path.name}': {e}")
            continue

    # 2. Fallback com auto-detecção de separador se o primeiro falhar
    if df is None:
        for enc in encodings:
            try:
                temp_df = pd.read_csv(file_path, encoding=enc, sep=None, comment='#', engine='python')
                if len(temp_df.columns) >= 2:
                    df = temp_df
                    logger.info(f"✅ Fallback sep=None executado para '{file_path.name}' em '{enc}'. Linhas: {len(df)}")
                    break
            except Exception:
                continue

    if df is None or df.empty:
        logger.error(f"💥 Erro fatal: Não foi possível estruturar o CSV '{file_path.name}'.")
        return pd.DataFrame()

    # Tratamento dos nomes de colunas e dados
    df.columns = [clean_encoding_text(str(col)).strip() for col in df.columns]
    for col in df.columns:
        df[col] = df[col].apply(clean_encoding_text)

    logger.info(f"📋 Colunas identificadas em '{file_path.name}': {list(df.columns)}")
    if not df.empty:
        logger.info(f"🔹 Exemplo da 1ª linha de '{file_path.name}': {df.iloc[0].to_dict()}")

    return df


def get_flexible_value(row_dict: dict, candidates: list, default=""):
    """Busca um valor no dicionário da linha comparando chaves de forma flexível."""
    for cand in candidates:
        cand_clean = cand.lower().strip()
        for k, v in row_dict.items():
            k_clean = str(k).lower().strip()
            if cand_clean == k_clean or cand_clean in k_clean or k_clean in cand_clean:
                if pd.notna(v) and str(v).strip() != '':
                    return str(v).strip()
    return default


def is_valid_printer_name(nome: str) -> bool:
    if not nome or len(nome) <= 1:
        return False
    nome_lower = nome.lower()
    if ';' in nome or 'dispositivo;' in nome_lower or 'tipo de dispositivo' in nome_lower or 'atividade;' in nome_lower:
        return False
    return True


def merge_and_normalize_papercut_data(df_printers: pd.DataFrame, df_devices: pd.DataFrame) -> pd.DataFrame:
    """
    Unifica e normaliza os dados de Impressoras (filas) e Dispositivos Físicos do PaperCut.
    """
    logger.info("--- Iniciando Normalização e Fusão dos DataFrames ---")
    records = []

    # Processa Lista de Impressoras (PrinterList)
    if df_printers is not None and not df_printers.empty:
        logger.info(f"Processando {len(df_printers)} linhas do DataFrame de Impressoras (PrinterList)...")
        for idx, row in df_printers.iterrows():
            row_dict = row.to_dict()
            
            nome = get_flexible_value(row_dict, ['impressora', 'printer name', 'nome da impressora', 'name', 'nome', 'printer'])
            servidor = get_flexible_value(row_dict, ['servidor', 'servidores', 'server', 'server name', 'host'])
            localizacao = get_flexible_value(row_dict, ['localização', 'localizacao', 'location', 'local', 'prédio', 'sala'])
            status = get_flexible_value(row_dict, ['status', 'estado', 'situação'], default='OK')
            paginas_raw = get_flexible_value(row_dict, ['total páginas', 'total paginas', 'total printed', 'pages', 'páginas', 'paginas', 'total trabalhos'], default='0')
            modelo = get_flexible_value(row_dict, ['modelo', 'model', 'tipo/modelo', 'atributos', 'fabricante'])
            ip_host = get_flexible_value(row_dict, ['ip/host', 'nome físico', 'physical name', 'ip address', 'endereço ip', 'hostname', 'ip'], default=servidor)

            # Tenta extrair servidor do nome se vier no formato "servidor\impressora"
            if not servidor and '\\' in nome:
                parts = nome.split('\\')
                servidor = parts[0]

            try:
                paginas = int(float(str(paginas_raw).replace('.', '').replace(',', '')))
            except (ValueError, TypeError):
                paginas = 0

            if is_valid_printer_name(nome):
                records.append({
                    'nome': nome,
                    'servidor': servidor if servidor else 'PaperCut',
                    'tipo': 'Fila de Impressão',
                    'modelo': modelo,
                    'localizacao': localizacao,
                    'ip_host': ip_host if ip_host else servidor,
                    'status': status,
                    'total_paginas': paginas,
                    'filas_relacionadas': '',
                    'detalhes_extra': 'Origem: PrinterList'
                })
            else:
                logger.debug(f"Linha {idx} de PrinterList ignorada por nome inválido/curto: {row_dict}")

        logger.info(f"✅ Extração de PrinterList concluída: {len(records)} registros válidos obtidos.")
    else:
        logger.warning("⚠️ DataFrame de Impressoras (PrinterList) está VAZIO.")

    printers_count = len(records)

    # Processa Lista de Dispositivos (DeviceList)
    if df_devices is not None and not df_devices.empty:
        logger.info(f"Processando {len(df_devices)} linhas do DataFrame de Dispositivos (DeviceList)...")
        for idx, row in df_devices.iterrows():
            row_dict = row.to_dict()
            
            nome = get_flexible_value(row_dict, ['nome do dispositivo', 'device name', 'dispositivo', 'device', 'name', 'nome'])
            servidor = get_flexible_value(row_dict, ['alojado em', 'servidor', 'server', 'host'])
            ip_host = get_flexible_value(row_dict, ['nome do host', 'ip address', 'endereço ip', 'endereco ip', 'ip/host', 'host', 'ip'])
            localizacao = get_flexible_value(row_dict, ['localização', 'localizacao', 'location', 'local', 'prédio'])
            status = get_flexible_value(row_dict, ['status', 'estado', 'situação'], default='OK')
            modelo = get_flexible_value(row_dict, ['tipo', 'função', 'funcao', 'tipo/modelo', 'model', 'modelo'])
            paginas_raw = get_flexible_value(row_dict, ['total páginas impressas', 'total páginas', 'total printed', 'pages', 'contador'], default='0')
            
            try:
                paginas = int(float(str(paginas_raw).replace('.', '').replace(',', '')))
            except (ValueError, TypeError):
                paginas = 0

            if is_valid_printer_name(nome):
                records.append({
                    'nome': nome,
                    'servidor': servidor if servidor else ip_host,
                    'tipo': 'Dispositivo Físico (MFD)',
                    'modelo': modelo,
                    'localizacao': localizacao,
                    'ip_host': ip_host if ip_host else servidor,
                    'status': status,
                    'total_paginas': paginas,
                    'filas_relacionadas': '',
                    'detalhes_extra': 'Origem: DeviceList'
                })
            else:
                logger.debug(f"Linha {idx} de DeviceList ignorada por nome inválido/curto: {row_dict}")

        logger.info(f"✅ Extração de DeviceList concluída: {len(records) - printers_count} registros válidos obtidos.")
    else:
        logger.warning("⚠️ DataFrame de Dispositivos (DeviceList) está VAZIO.")

    if not records:
        logger.warning("❌ Nenhum registro pôde ser montado a partir dos DataFrames informados.")
        return pd.DataFrame()

    df_merged = pd.DataFrame(records)
    logger.info(f"Total de registros unificados antes de desduplicar: {len(df_merged)}")
    
    # Remove duplicatas baseadas no nome, mantendo o primeiro registro válido
    df_merged.drop_duplicates(subset=['nome'], keep='first', inplace=True)
    logger.info(f"Total de registros finais após remover duplicatas: {len(df_merged)}")

    if not df_merged.empty:
        logger.info(f"🔹 Exemplo do 1º registro final: {df_merged.iloc[0].to_dict()}")

    return df_merged


def run_papercut_scraper():
    """
    Executa a raspagem dos dados do PaperCut via Selenium baixando e tratando os arquivos CSVs atualizados.
    """
    logger.info("============================================================")
    logger.info("INICIANDO PROCESSO DE RASPAGEM E SINCRONIZAÇÃO DO PAPERCUT")
    logger.info("============================================================")
    logger.info(f"URL de Login: '{PAPERCUT_URL}'")
    logger.info(f"URL PrinterList: '{PAPERCUT_PRINTER_LIST_URL}'")
    logger.info(f"URL DeviceList: '{PAPERCUT_DEVICE_LIST_URL}'")
    logger.info(f"Usuário PaperCut: '{PAPERCUT_USER}' | Senha preenchida: {'SIM' if PAPERCUT_PASS else 'NÃO'}")

    # Limpa arquivos antigos da pasta Downloads antes de iniciar o novo download
    cleanup_old_papercut_downloads()

    driver = None
    scraped_success = False

    try:
        if PAPERCUT_USER and PAPERCUT_PASS and PAPERCUT_URL:
            logger.info("🚀 Iniciando navegador Chrome com Selenium...")
            driver = get_chrome_driver(headless=HEADLESS)
            logger.info("✅ Navegador Chrome iniciado com sucesso.")

            logger.info(f"Navegando para página de login: {PAPERCUT_URL}")
            driver.get(PAPERCUT_URL)
            time.sleep(2)
            
            logger.info(f"Página de Login -> URL Atual: '{driver.current_url}' | Título: '{driver.title}'")

            # Preenche o formulário de login
            username_fields = driver.find_elements(By.ID, "inputUsername") or driver.find_elements(By.NAME, "inputUsername")
            password_fields = driver.find_elements(By.ID, "inputPassword") or driver.find_elements(By.NAME, "inputPassword")
            
            if username_fields and password_fields:
                username_fields[0].clear()
                username_fields[0].send_keys(PAPERCUT_USER)
                password_fields[0].clear()
                password_fields[0].send_keys(PAPERCUT_PASS)
                logger.info(f"Credenciais preenchidas para usuário '{PAPERCUT_USER}'. Submetendo...")

                submit_btn = (driver.find_elements(By.NAME, "$Submit$0") or 
                              driver.find_elements(By.CSS_SELECTOR, "input.loginSubmit") or 
                              driver.find_elements(By.CSS_SELECTOR, "input[type='submit']"))
                
                if submit_btn:
                    submit_btn[0].click()
                    time.sleep(4)
                    logger.info(f"Pós-Login -> URL Atual: '{driver.current_url}' | Título: '{driver.title}'")

            # -----------------------------------------------------------------------------
            # NAVEGAÇÃO E DOWNLOAD: PRINTER LIST
            # -----------------------------------------------------------------------------
            logger.info(f"Navegando para PrinterList: {PAPERCUT_PRINTER_LIST_URL}")
            driver.get(PAPERCUT_PRINTER_LIST_URL)
            time.sleep(3)
            logger.info(f"PrinterList -> URL Atual: '{driver.current_url}' | Título: '{driver.title}'")
            
            # Localiza o link exato do CSV apontado no HTML do PaperCut:
            # <a href="/app?service=direct/1/PrinterList/$ReportLink.csv&sp=SCSV&sp=F"...>
            csv_links_printer = (
                driver.find_elements(By.XPATH, "//a[contains(@href, 'PrinterList/$ReportLink.csv') or contains(@href, 'sp=SCSV')]") or
                driver.find_elements(By.XPATH, "//img[@alt='CSV' or @title='CSV']/parent::a")
            )
            
            logger.info(f"Links de exportação CSV encontrados em PrinterList: {len(csv_links_printer)}")
            if csv_links_printer:
                target_url = csv_links_printer[0].get_attribute('href')
                logger.info(f"📥 Clicando no link de download CSV do PrinterList: {target_url}")
                csv_links_printer[0].click()
                time.sleep(5)
            else:
                logger.warning("⚠️ Nenhum link de download CSV encontrado em PrinterList.")

            # -----------------------------------------------------------------------------
            # NAVEGAÇÃO E DOWNLOAD: DEVICE LIST
            # -----------------------------------------------------------------------------
            logger.info(f"Navegando para DeviceList: {PAPERCUT_DEVICE_LIST_URL}")
            driver.get(PAPERCUT_DEVICE_LIST_URL)
            time.sleep(3)
            logger.info(f"DeviceList -> URL Atual: '{driver.current_url}' | Título: '{driver.title}'")
            
            # Localiza o link exato do CSV apontado no HTML do PaperCut:
            # <a href="/app?service=direct/1/DeviceList/$ReportLink.csv&sp=SCSV&sp=F"...>
            csv_links_device = (
                driver.find_elements(By.XPATH, "//a[contains(@href, 'DeviceList/$ReportLink.csv') or contains(@href, 'sp=SCSV')]") or
                driver.find_elements(By.XPATH, "//img[@alt='CSV' or @title='CSV']/parent::a")
            )
            
            logger.info(f"Links de exportação CSV encontrados em DeviceList: {len(csv_links_device)}")
            if csv_links_device:
                target_url = csv_links_device[0].get_attribute('href')
                logger.info(f"📥 Clicando no link de download CSV do DeviceList: {target_url}")
                csv_links_device[0].click()
                time.sleep(5)
            else:
                logger.warning("⚠️ Nenhum link de download CSV encontrado em DeviceList.")

            scraped_success = True
            logger.info("✅ Fluxo Selenium finalizado.")

        else:
            logger.warning("⚠️ Credenciais ou URL do PaperCut não configuradas no .env. Saltando etapa Selenium.")

    except Exception as e:
        logger.error(f"💥 Falha durante execução do Selenium Scraper: {e}", exc_info=True)
    finally:
        if driver:
            try:
                driver.quit()
                logger.info("Navegador Selenium encerrado.")
            except Exception:
                pass

    # -----------------------------------------------------------------------------
    # IDENTIFICAÇÃO DOS ARQUIVOS CSV BAIXADOS
    # -----------------------------------------------------------------------------
    logger.info("--- Procurando arquivos CSV recém-baixados na pasta Downloads ---")
    
    downloaded_printers = list(DOWNLOAD_DIR.glob("printer_list*.csv"))
    downloaded_devices = list(DOWNLOAD_DIR.glob("device_list*.csv"))

    printer_csv_file = downloaded_printers[0] if downloaded_printers else PAPERCUT_BRUTOS_DIR / "printer_list.csv"
    device_csv_file = downloaded_devices[0] if downloaded_devices else PAPERCUT_BRUTOS_DIR / "device_list.csv"

    logger.info(f"Arquivo selecionado para PrinterList: {printer_csv_file}")
    logger.info(f"Arquivo selecionado para DeviceList: {device_csv_file}")

    # Carrega os DataFrames
    df_printers = load_csv_safe(printer_csv_file)
    df_devices = load_csv_safe(device_csv_file)

    # Copia backups com timestamp para a pasta de dados brutos do projeto
    ts = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    if printer_csv_file.exists() and printer_csv_file.parent == DOWNLOAD_DIR:
        try:
            shutil.copy2(printer_csv_file, PAPERCUT_BRUTOS_DIR / f"printer_list_{ts}.csv")
            shutil.copy2(printer_csv_file, PAPERCUT_BRUTOS_DIR / "printer_list.csv")
            logger.info(f"Cópia de backup salva em: {PAPERCUT_BRUTOS_DIR / f'printer_list_{ts}.csv'}")
        except Exception as e:
            logger.warning(f"Não foi possível salvar backup de printer_list: {e}")

    if device_csv_file.exists() and device_csv_file.parent == DOWNLOAD_DIR:
        try:
            shutil.copy2(device_csv_file, PAPERCUT_BRUTOS_DIR / f"device_list_{ts}.csv")
            shutil.copy2(device_csv_file, PAPERCUT_BRUTOS_DIR / "device_list.csv")
            logger.info(f"Cópia de backup salva em: {PAPERCUT_BRUTOS_DIR / f'device_list_{ts}.csv'}")
        except Exception as e:
            logger.warning(f"Não foi possível salvar backup de device_list: {e}")

    # Mantém até 10 backups rotativos históricos do PaperCut
    cleanup_old_files(PAPERCUT_BRUTOS_DIR, "printer_list_*.csv", keep_count=10)
    cleanup_old_files(PAPERCUT_BRUTOS_DIR, "device_list_*.csv", keep_count=10)


    # Unifica e trata os dados
    df_merged = merge_and_normalize_papercut_data(df_printers, df_devices)

    if not df_merged.empty:
        logger.info(f"Gravando {len(df_merged)} impressoras/dispositivos no banco de dados SQLite...")
        save_impressoras_to_db(df_merged)
        logger.info(f"🎉 SUCESSO COMPLETO: {len(df_merged)} ativos gravados com sucesso no SQLite!")
        
        # Limpa os arquivos temporários baixados em Downloads para não acumular
        for f in [printer_csv_file, device_csv_file]:
            if f and f.exists() and f.parent == DOWNLOAD_DIR:
                try:
                    f.unlink()
                    logger.info(f"🗑️ Arquivo de download temporário removido: {f.name}")
                except Exception as e:
                    logger.warning(f"Não foi possível remover temporário {f.name}: {e}")
    else:
        logger.warning("⚠️ Nenhum dado de impressora pôde ser extraído ou carregado dos CSVs.")

    logger.info("============================================================\n")


if __name__ == "__main__":
    run_papercut_scraper()
