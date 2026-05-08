import os
import keyring
from pathlib import Path
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env para a memória do script
load_dotenv()

# -----------------------------------------------------------------------------
# Instalações antes de rodar
# -----------------------------------------------------------------------------

# Credenciais
CITSMART_URL = os.getenv("CITSMART_LINK", "")
CITSMART_NOVA_FILA = os.getenv("CITSMART_LINK_NOVO", "")
OTRS_URL = os.getenv("OTRS_LINK", "")
PROMOTORIAS_URL = "https://www.mpms.mp.br/promotorias"
PROCURADORIAS_URL = "https://www.mpms.mp.br/procuradorias"

USERNAME = os.getlogin()
CITSMART_EMAIL = f"{USERNAME}@{os.getenv('AD_EMAIL', '')}"

# Tenta pegar senha do keyring, ou deixa vazia se falhar (para não quebrar no PC de outros)
try:
    PASSWORD = keyring.get_password("otrs", USERNAME)
except:
    PASSWORD = None

# Domínio
DOMINIO = os.getenv("AD_DOMAIN", "")
DOMINIO_CURTO = os.getenv("AD_SHORT", "")
DOMINIO_MMC = os.getenv("AD_MMC", "")

# Configurações do WebDriver
DRIVER_PATH = "./chromedriver.exe"  # Baixe a versão correspondente ao seu Chrome
HEADLESS = True  # Mude para True após testes

# Adicione estes novos parâmetros
IMPLICIT_WAIT = 10  # Espera implícita global
MAX_WAIT_DESCRIPTION = 15  # Aumente se necessário
EXPLICIT_WAIT = 30  # Espera explícita para elementos críticos
MAX_RETRIES = 5     # Número de tentativas por página

# -----------------------------------------------------------------------------
# Diretórios
# -----------------------------------------------------------------------------

# Pega automaticamente a pasta do usuário atual
USER_HOME = Path.home()

BASE_DIR              = Path(__file__).parent
INPUT_DIR_BRUTOS      = BASE_DIR / "01 - Dados Brutos"
INPUT_DIR_BRUTOS.mkdir(exist_ok=True)
OUTPUT_DIR_TRATADOS   = BASE_DIR / "02 - Dados tratados"
OUTPUT_DIR_TRATADOS.mkdir(exist_ok=True)
OUTPUT_DIR_PRONTO     = BASE_DIR / "03 - Dados prontos"
OUTPUT_DIR_PRONTO.mkdir(exist_ok=True)
MODEL_DIR             = BASE_DIR / "models"
MODEL_DIR.mkdir(exist_ok=True)
MASTER_FILE_PATH = USER_HOME / os.getenv("SHAREPOINT_RELATIVE_PATH", "")

# OTRS
DEBUG_DIR_OTRS = BASE_DIR / "debug_logs" / "otrs"
DEBUG_DIR_OTRS.mkdir(parents=True, exist_ok=True)
BACKUP_CSV_OTRS = BASE_DIR / "debug_logs" / "otrs" / "backup_stream.csv"

# Master spreadsheet path
BACKUP_PATH_OTRS = INPUT_DIR_BRUTOS.with_suffix('.backup.xlsx')
TEMP_PATH_OTRS = INPUT_DIR_BRUTOS.with_suffix('.tmp.xlsx')

# CITSMART
DEBUG_DIR_CITSMART = BASE_DIR / "debug_logs" / "citsmart"
DEBUG_DIR_CITSMART.mkdir(parents=True, exist_ok=True)
BACKUP_CSV_CITSMART = BASE_DIR /"debug_logs" / "citsmart" / "backup_stream.csv"

# PREPROCESS
DEBUG_DIR_PREPROCESS = BASE_DIR / "debug_logs" / "preprocess"
DEBUG_DIR_PREPROCESS.mkdir(parents=True, exist_ok=True)
BACKUP_CSV_PREPROCESS = BASE_DIR /"debug_logs" / "preprocess" / "backup_stream.csv"

# Tag Classfier
DEBUG_DIR_TAG = BASE_DIR / "debug_logs" / "tag"
DEBUG_DIR_TAG.mkdir(parents=True, exist_ok=True)
TREINO_PATH = OUTPUT_DIR_TRATADOS / "Chamados_Treino.xlsx"
MODEL_PATH  = MODEL_DIR / "tag_classifier.joblib"

# Sync Master
DEBUG_DIR_SYNC = BASE_DIR / "debug_logs" / "sync"
DEBUG_DIR_SYNC.mkdir(parents=True, exist_ok=True)

# Orquestrador
DEBUG_DIR_ORQUESTRADOR = BASE_DIR / "debug_logs" / "orquestrador"
DEBUG_DIR_ORQUESTRADOR.mkdir(parents=True, exist_ok=True)
LOG_FILE_ORQUESTRADOR = DEBUG_DIR_ORQUESTRADOR / "orquestrador.log"

# -----------------------------------------------------------------------------
# LOGGING CENTRALIZADO
# -----------------------------------------------------------------------------
import sys
import logging
import pandas as pd
from logging.handlers import RotatingFileHandler

class SafeStreamWrapper:
    """Wrapper para streams que previne travamentos catastróficos por UnicodeEncodeError no Windows."""
    def __init__(self, stream):
        self.stream = stream

    def write(self, data):
        try:
            self.stream.write(data)
        except UnicodeEncodeError:
            try:
                # Tenta codificar substituindo caracteres não suportados por '?'
                encoding = getattr(self.stream, "encoding", None) or "ascii"
                safe_data = data.encode(encoding, errors="replace").decode(encoding)
                self.stream.write(safe_data)
            except Exception:
                # Fallback final em ASCII absoluto
                safe_data = data.encode("ascii", errors="replace").decode("ascii")
                self.stream.write(safe_data)

    def flush(self):
        if hasattr(self.stream, "flush"):
            self.stream.flush()

    def __getattr__(self, name):
        return getattr(self.stream, name)

def setup_logging(log_file: Path, name: str = __name__) -> logging.Logger:
    """Configura o logging rotativo e para terminal de forma unificada e centralizada com proteção Unicode."""
    log_file.parent.mkdir(parents=True, exist_ok=True)
    
    file_handler = RotatingFileHandler(
        filename=log_file,
        maxBytes=5 * 1024 * 1024,  # 5 MB em bytes
        backupCount=3,             # Mantém apenas 3 arquivos de histórico
        encoding='utf-8'
    )
    
    # Envelopa sys.stdout com nosso wrapper de proteção unicode
    safe_stdout = SafeStreamWrapper(sys.stdout)
    stream_handler = logging.StreamHandler(safe_stdout)
    
    # Configura o basicConfig. force=True reinicia handlers anteriores
    logging.basicConfig(
        level=logging.DEBUG,
        format='[%(asctime)s] [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[file_handler, stream_handler],
        force=True
    )
    return logging.getLogger(name)


def save_df_to_excel_formatted(
    df: pd.DataFrame,
    output_path: Path,
    sheet_name: str = "Sheet1",
    widths: dict = None,
    wrap_cols: list = None,
    height_col: str = None
):
    """
    Salva um DataFrame no Excel usando xlsxwriter e aplica formatação de 
    larguras, quebra de texto (wrap text) e ajuste dinâmico de altura de linhas.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    if widths is None:
        widths = {}
    if wrap_cols is None:
        wrap_cols = []
        
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        wb = writer.book
        ws = writer.sheets[sheet_name]
        wrap_format = wb.add_format({'text_wrap': True})
        
        # 1. Aplica larguras e quebra de texto nas colunas
        for i, col in enumerate(df.columns):
            fmt = wrap_format if col in wrap_cols else None
            width = widths.get(col, 20)
            ws.set_column(i, i, width, fmt)
            
        # 2. Ajusta dinamicamente a altura das linhas com base em quebras de linha (\n)
        if height_col and height_col in df.columns:
            for r, cell in enumerate(df[height_col], start=1):
                cell_str = str(cell) if pd.notna(cell) else ""
                lines = cell_str.count("\n") + 1
                ws.set_row(r, 15 * lines)


def cleanup_old_files(directory: Path, pattern: str, keep_count: int = 10) -> None:
    """
    Remove arquivos antigos correspondentes ao padrão especificado no diretório,
    mantendo apenas a quantidade dos mais recentes definidos por `keep_count`.
    """
    try:
        import os
        files = sorted(directory.glob(pattern), key=os.path.getmtime)
        if len(files) > keep_count:
            to_delete = files[:-keep_count]
            for file in to_delete:
                try:
                    file.unlink()
                    logging.getLogger().info(f"[LIMPEZA] Arquivo antigo removido: {file.name}")
                except Exception as e:
                    logging.getLogger().warning(f"Não foi possível remover arquivo {file.name}: {e}")
    except Exception as e:
        logging.getLogger().error(f"Erro na rotina de limpeza para o padrão '{pattern}': {e}")


def setup_ad_connection():
    """Tenta conectar no AD (Active Directory) de forma unificada e centralizada."""
    try:
        from ldap3 import Server, Connection, ALL
        server = Server(DOMINIO, get_info=ALL)
        conn = Connection(server, user=f"{DOMINIO_CURTO}\\{USERNAME}", password=PASSWORD, auto_bind=True)
        return conn
    except Exception as e:
        import logging
        logging.getLogger().debug(f"⚠️ Aviso: Não foi possível conectar ao AD. Erro: {e}")
        return None


def fetch_ad_department(conn, query_val: str, is_username: bool = True) -> str:
    """
    Busca o departamento/unidade de um usuário no AD de forma robusta e unificada.
    Se is_username=True, busca por sAMAccountName (usado no OTRS).
    Se is_username=False, busca por displayName/cn/name (usado no CitSmart).
    Retorna o departamento, ou escritório, ou mensagens padronizadas de falha.
    """
    if not conn or not query_val:
        return ""
        
    try:
        from ldap3 import SUBTREE
        target_attrs = ['department', 'physicalDeliveryOfficeName']
        
        if is_username:
            search_filters = [f'(sAMAccountName={query_val})']
        else:
            search_filters = [
                f'(displayName={query_val})',
                f'(cn={query_val})',
                f'(name={query_val})',
                f'(displayName=*{query_val}*)'
            ]
            
        entry = None
        for filt in search_filters:
            conn.search(
                search_base=f'{DOMINIO_MMC}',
                search_filter=filt,
                search_scope=SUBTREE,
                attributes=target_attrs
            )
            if conn.entries:
                entry = conn.entries[0].entry_attributes_as_dict
                break
                
        if not entry:
            return 'Não encontrado no AD'
            
        # 1. Tenta Departamento (department)
        dept_list = entry.get('department', [])
        dept = dept_list[0] if dept_list else None
        if dept and str(dept).strip():
            return str(dept).strip()
            
        # 2. Tenta Escritório/Prédio (physicalDeliveryOfficeName)
        office_list = entry.get('physicalDeliveryOfficeName', [])
        office = office_list[0] if office_list else None
        if office and str(office).strip():
            return str(office).strip()
            
        return 'Cadastro Incompleto (AD)'
        
    except Exception as e:
        import logging
        logging.getLogger().error(f"Erro AD lookup para {query_val}: {e}")
        return 'Erro na Consulta'


def get_chrome_driver(
    headless: bool = True,
    page_load_strategy: str = None,
    block_media: bool = False,
    disable_gpu: bool = False
):
    """
    Inicializa o Selenium Chrome Driver de forma unificada e profissional com silenciadores de log,
    modo incógnito, modo headless e tratamento de tamanho de tela.
    """
    from selenium import webdriver
    
    opts = webdriver.ChromeOptions()
    
    # Desativa logs desnecessários do Chrome
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_argument('--log-level=3')
    opts.add_argument('--disable-logging')
    opts.add_argument("--incognito")
    opts.add_argument("--disable-infobars")
    opts.add_argument("--no-default-browser-check")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-blink-features=CSSAnimations,ScrollAnimator")
    
    # Previne pop-ups de senhas e credenciais
    opts.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    })
    
    # Estratégia de carregamento de página
    if page_load_strategy:
        opts.page_load_strategy = page_load_strategy
        
    # Bloqueio opcional de imagens e CSS (para economia de banda e CPU no OTRS)
    if block_media:
        prefs = {
            "profile.managed_default_content_settings.images": 2,
            "profile.managed_default_content_settings.stylesheets": 2,
            "profile.managed_default_content_settings.fonts": 2,
        }
        opts.add_experimental_option("prefs", prefs)
        
    # Configuração headless
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
        if disable_gpu:
            opts.add_argument("--disable-gpu")
    else:
        opts.add_argument("--start-maximized")
        
    driver = webdriver.Chrome(options=opts)
    return driver