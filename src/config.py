import os
import sys
import asyncio
import keyring
from pathlib import Path
from dotenv import load_dotenv

# Silencia exceção WinError 10054 no asyncio/Proactor ao fechar abas/conexões no Windows
if sys.platform == 'win32':
    try:
        from asyncio.proactor_events import _ProactorBasePipeTransport
        _orig_call_connection_lost = _ProactorBasePipeTransport._call_connection_lost

        def _silenced_call_connection_lost(self, exc):
            try:
                _orig_call_connection_lost(self, exc)
            except (ConnectionResetError, ConnectionAbortedError, OSError):
                pass

        _ProactorBasePipeTransport._call_connection_lost = _silenced_call_connection_lost
    except Exception:
        pass

# Carrega as variáveis do arquivo .env para a memória do script
load_dotenv()

# Função auxiliar para ler do banco de dados com fallback para os.getenv
def _cfg(key: str, default: str = "") -> str:
    try:
        from src.database.settings_db import get_setting
        val = get_setting(key)
        if val is not None and str(val).strip() != "":
            return str(val).strip()
    except Exception:
        pass
    return os.getenv(key, default)

# -----------------------------------------------------------------------------
# Instalações antes de rodar
# -----------------------------------------------------------------------------

# Credenciais
CITSMART_URL = _cfg("CITSMART_LINK", "")
CITSMART_NOVA_FILA = _cfg("CITSMART_LINK_NOVO", "")
OTRS_URL = _cfg("OTRS_LINK", "")
PROMOTORIAS_URL = "https://www.mpms.mp.br/promotorias"
PROCURADORIAS_URL = "https://www.mpms.mp.br/procuradorias"

PAPERCUT_URL = _cfg("PAPERCUT_URL", "")
PAPERCUT_PRINTER_LIST_URL = _cfg("PAPERCUT_PRINTER_LIST_URL", "")
PAPERCUT_DEVICE_LIST_URL = _cfg("PAPERCUT_DEVICE_LIST_URL", "")

OXE_URL = _cfg("OXE_URL", "")

# Pega automaticamente a pasta do usuário e a raiz do projeto
USER_HOME = Path.home()
BASE_DIR = Path(__file__).parent.parent

# Scripts de Automação PowerShell (Padrão: src/scripts_powershell/)
PS_SCRIPTS_DIR = BASE_DIR / "src" / "scripts_powershell"
PS_SCRIPT_ANALISADOR = PS_SCRIPTS_DIR / "analisador" / "Analisador.ps1"
PS_SCRIPT_MANUTENCAO = PS_SCRIPTS_DIR / "manutencao" / "Manutencao.ps1"
PS_SCRIPT_REMOVER_USUARIOS = PS_SCRIPTS_DIR / "perfis" / "RemoverUsuarios.ps1"


def _get_username() -> str:
    """Obtém o nome de usuário do sistema de forma segura para Banco, Docker, Linux e Windows."""
    db_user = _cfg("AD_USER")
    if db_user and db_user.strip() and db_user.strip() != "root":
        return db_user.strip()

    env_user = os.getenv("AD_USER") or os.getenv("CITSMART_USER")
    if env_user and env_user.strip() and env_user.strip() != "root":
        return env_user.strip()
    
    host_user = os.getenv("USER") or os.getenv("USERNAME")
    if host_user and host_user.strip() and host_user.strip() != "root":
        return host_user.strip()

    try:
        import getpass
        user = getpass.getuser()
        if user and user.strip() and user.strip() != "root":
            return user.strip()
    except Exception:
        pass
        
    return "paulogoncalves"

USERNAME = _get_username()
CITSMART_EMAIL = f"{USERNAME}@{_cfg('AD_EMAIL', os.getenv('AD_EMAIL', ''))}"

# Tenta pegar senha do banco de dados, keyring ou env AD_PASSWORD
try:
    PASSWORD = _cfg("AD_PASSWORD") or os.getenv("AD_PASSWORD") or keyring.get_password("otrs", USERNAME)
except Exception:
    PASSWORD = _cfg("AD_PASSWORD") or os.getenv("AD_PASSWORD")

try:
    PAPERCUT_USER = _cfg("PAPERCUT_USER", os.getenv("PAPERCUT_USER", keyring.get_password("papercut_user", "papercut") or "admin"))
    PAPERCUT_PASS = _cfg("PAPERCUT_PASS", os.getenv("PAPERCUT_PASS", keyring.get_password("papercut", PAPERCUT_USER) or ""))
except:
    PAPERCUT_USER = _cfg("PAPERCUT_USER", os.getenv("PAPERCUT_USER", "admin"))
    PAPERCUT_PASS = _cfg("PAPERCUT_PASS", os.getenv("PAPERCUT_PASS", ""))

try:
    OXE_USER = _cfg("OXE_USER", os.getenv("OXE_USER", keyring.get_password("oxe_user", "oxe") or "mtcl"))
    OXE_PASS = _cfg("OXE_PASS", os.getenv("OXE_PASS", keyring.get_password("oxe", OXE_USER) or ""))
except:
    OXE_USER = _cfg("OXE_USER", os.getenv("OXE_USER", "mtcl"))
    OXE_PASS = _cfg("OXE_PASS", os.getenv("OXE_PASS", ""))


# Domínio
DOMINIO = _cfg("AD_DOMAIN", os.getenv("AD_DOMAIN", ""))
DOMINIO_CURTO = _cfg("AD_SHORT", os.getenv("AD_SHORT", ""))
DOMINIO_MMC = _cfg("AD_MMC", os.getenv("AD_MMC", ""))

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

INPUT_DIR_BRUTOS      = BASE_DIR / "01 - Dados Brutos"
INPUT_DIR_BRUTOS.mkdir(exist_ok=True)
OUTPUT_DIR_TRATADOS   = BASE_DIR / "02 - Dados tratados"
OUTPUT_DIR_TRATADOS.mkdir(exist_ok=True)
OUTPUT_DIR_PRONTO     = BASE_DIR / "03 - Dados prontos"
OUTPUT_DIR_PRONTO.mkdir(exist_ok=True)
MODEL_DIR             = BASE_DIR / "models"
MODEL_DIR.mkdir(exist_ok=True)
MASTER_FILE_PATH = USER_HOME / _cfg("SHAREPOINT_RELATIVE_PATH", os.getenv("SHAREPOINT_RELATIVE_PATH", ""))
DONATIONS_FILE_PATH = USER_HOME / _cfg("DONATIONS_EXCEL_RELATIVE_PATH", os.getenv("DONATIONS_EXCEL_RELATIVE_PATH", ""))
WARRANTY_FILE_PATH = USER_HOME / _cfg("WARRANTY_EXCEL_RELATIVE_PATH", os.getenv("WARRANTY_EXCEL_RELATIVE_PATH", ""))
VIAGENS_FILE_PATH = USER_HOME / _cfg("VIAGENS_EXCEL_RELATIVE_PATH", os.getenv("VIAGENS_EXCEL_RELATIVE_PATH", ""))
SHAREPOINT_MATUTINO_URL = _cfg("SHAREPOINT_MATUTINO_URL", os.getenv("SHAREPOINT_MATUTINO_URL", ""))

VIDEO_FAQ_ENV = os.getenv("VIDEO_FAQ_PATH", "")
IMAGE_FAQ_ENV = os.getenv("IMAGE_FAQ_PATH", "")

if VIDEO_FAQ_ENV.startswith("http://") or VIDEO_FAQ_ENV.startswith("https://"):
    VIDEO_FAQ_URL = VIDEO_FAQ_ENV
else:
    VIDEO_FAQ_URL = ""

if IMAGE_FAQ_ENV.startswith("http://") or IMAGE_FAQ_ENV.startswith("https://"):
    IMAGE_FAQ_URL = IMAGE_FAQ_ENV
else:
    IMAGE_FAQ_URL = ""

# Configuração de Vídeos e Imagens FAQ
# Padrão Cloud-First (Linux/RedHat/Docker): BASE_DIR / "uploads" / "faq"
VIDEO_FAQ_DIR = BASE_DIR / "uploads" / "faq" / "videos"
IMAGE_FAQ_DIR = BASE_DIR / "uploads" / "faq" / "imagens"

# Fallback opcional: só pesquisa pasta sincronizada do OneDrive se a variável de ambiente USE_ONEDRIVE_FALLBACK for ativada
if os.getenv("USE_ONEDRIVE_FALLBACK", "false").lower() == "true":
    _onedrive_base_candidates = [
        Path("/mnt/c/Users/paulogoncalves/OneDrive - Ministerio Público do Estado de Mato Grosso do Sul/Documentos SharePoint DIT-Manutenção/Tutoriais-FAQs"),
        Path("/mnt/c/Users/paulogoncalves/OneDrive - Ministério Público do Estado de Mato Grosso do Sul/Documentos SharePoint DIT-Manutenção/Tutoriais-FAQs")
    ]
    for _cand in _onedrive_base_candidates:
        if _cand.exists():
            for _v_name in ["Vídeos FAQ", "Videos FAQ"]:
                if (_cand / _v_name).exists():
                    VIDEO_FAQ_DIR = _cand / _v_name
                    break
            for _i_name in ["Imagens FAQ", "imagens"]:
                if (_cand / _i_name).exists():
                    IMAGE_FAQ_DIR = _cand / _i_name
                    break
            break

VIDEO_FAQ_DIR.mkdir(parents=True, exist_ok=True)
IMAGE_FAQ_DIR.mkdir(parents=True, exist_ok=True)

# Garantia Logs
DEBUG_DIR_GARANTIA = BASE_DIR / "debug_logs" / "garantia"
DEBUG_DIR_GARANTIA.mkdir(parents=True, exist_ok=True)

# Atos e Normas
ATOS_NORMAS_API_URL = os.getenv("ATOS_NORMAS_API_URL", "")
ATOS_NORMAS_DOWNLOAD_URL = os.getenv("ATOS_NORMAS_DOWNLOAD_URL", "")




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

# Donations
DEBUG_DIR_DONATIONS = BASE_DIR / "debug_logs" / "donations"
DEBUG_DIR_DONATIONS.mkdir(parents=True, exist_ok=True)


# Orquestrador
DEBUG_DIR_ORQUESTRADOR = BASE_DIR / "debug_logs" / "orquestrador"
DEBUG_DIR_ORQUESTRADOR.mkdir(parents=True, exist_ok=True)

# Unidades & Ramais
DEBUG_DIR_UNIDADES = BASE_DIR / "debug_logs" / "unidades"
DEBUG_DIR_UNIDADES.mkdir(parents=True, exist_ok=True)

DEBUG_DIR_RAMAIS = BASE_DIR / "debug_logs" / "ramais"
DEBUG_DIR_RAMAIS.mkdir(parents=True, exist_ok=True)

# Leaflet Map Logs
DEBUG_DIR_LEAFLET = BASE_DIR / "debug_logs" / "leaflet"
DEBUG_DIR_LEAFLET.mkdir(parents=True, exist_ok=True)

# FAQ Logs
DEBUG_DIR_FAQ = BASE_DIR / "debug_logs" / "faq"
DEBUG_DIR_FAQ.mkdir(parents=True, exist_ok=True)

# Plantao Logs
DEBUG_DIR_PLANTOES = BASE_DIR / "debug_logs" / "plantoes"
DEBUG_DIR_PLANTOES.mkdir(parents=True, exist_ok=True)

# PaperCut Logs
DEBUG_DIR_PAPERCUT = BASE_DIR / "debug_logs" / "papercut"
DEBUG_DIR_PAPERCUT.mkdir(parents=True, exist_ok=True)

# OXE Central Telefonica Logs
DEBUG_DIR_OXE = BASE_DIR / "debug_logs" / "oxe"
DEBUG_DIR_OXE.mkdir(parents=True, exist_ok=True)

# Scripts Logs
DEBUG_DIR_SCRIPTS = BASE_DIR / "debug_logs" / "scripts"
DEBUG_DIR_SCRIPTS.mkdir(parents=True, exist_ok=True)

# Fiscalizacao Logs
DEBUG_DIR_FISCALIZACAO = BASE_DIR / "debug_logs" / "fiscalizacao"
DEBUG_DIR_FISCALIZACAO.mkdir(parents=True, exist_ok=True)

DEBUG_DIR_VIAGENS = BASE_DIR / "debug_logs" / "viagens"
DEBUG_DIR_VIAGENS.mkdir(parents=True, exist_ok=True)



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

class ANSIColoredFormatter(logging.Formatter):
    """Formatador de logging com cores ANSI automáticas baseadas no nível da mensagem para o terminal."""
    RESET = "\033[0m"
    BOLD = "\033[1m"
    
    COLORS = {
        logging.DEBUG: "\033[90m",                      # Cinza Escuro
        logging.INFO: "\033[36m",                       # Ciano
        logging.WARNING: "\033[33m",                    # Amarelo
        logging.ERROR: "\033[31m\033[1m",               # Vermelho com Negrito
        logging.CRITICAL: "\033[41m\033[37m\033[1m",    # Fundo Vermelho / Texto Branco Negrito
    }

    def format(self, record):
        color = self.COLORS.get(record.levelno, self.RESET)
        log_fmt = f"{color}[%(asctime)s] [%(levelname)s]{self.RESET} %(message)s"
        formatter = logging.Formatter(log_fmt, datefmt='%Y-%m-%d %H:%M:%S')
        return formatter.format(record)

def setup_logging(log_file: Path, name: str = __name__) -> logging.Logger:
    """Configura o logging rotativo e para terminal de forma unificada e centralizada com proteção Unicode e cores ANSI."""
    log_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Handler para Arquivo (Log Puro em disco sem códigos ANSI)
    file_handler = RotatingFileHandler(
        filename=log_file,
        maxBytes=5 * 1024 * 1024,  # 5 MB em bytes
        backupCount=3,             # Mantém apenas 3 arquivos de histórico
        encoding='utf-8'
    )
    plain_formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setFormatter(plain_formatter)
    
    # Handler para Terminal (Console com Cores ANSI e proteção Unicode)
    safe_stdout = SafeStreamWrapper(sys.stdout)
    stream_handler = logging.StreamHandler(safe_stdout)
    color_formatter = ANSIColoredFormatter()
    stream_handler.setFormatter(color_formatter)
    
    # Configura o logger raiz/específico com handlers separados
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    logger.handlers = []  # Limpa handlers anteriores para evitar duplicidades
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.propagate = False
    
    return logger


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


def _detect_powershell_executable() -> tuple:
    """
    Detecta o melhor executável do PowerShell disponível no ambiente:
    1. 'pwsh' / 'pwsh.exe' (PowerShell 7 Core - nativo no Linux/Container ou Windows)
    2. 'pwsh.exe' via WSL Interop (WindowsApps / Program Files)
    3. 'powershell.exe' (Windows PowerShell 5.1 via WSL Interop)
    4. 'powershell' (Windows PowerShell 5.1 nativo no Windows)
    Retorna uma tupla: (caminho_ou_comando, tipo: 'pwsh' | 'windows_ps' | None)
    """
    import shutil

    # 1. Tenta PowerShell 7+ (pwsh) no PATH (Containers Linux, Red Hat ou Windows)
    pwsh_bin = shutil.which("pwsh") or shutil.which("pwsh.exe")
    if pwsh_bin:
        return pwsh_bin, "pwsh"

    # 2. Se estiver no Linux/WSL, tenta buscar o PowerShell 7 (pwsh.exe) do Windows host
    if sys.platform != "win32":
        ps7_wsl_candidates = [
            "/mnt/c/Program Files/PowerShell/7/pwsh.exe",
            "/mnt/c/Program Files (x86)/PowerShell/7/pwsh.exe",
            f"/mnt/c/Users/{USERNAME}/AppData/Local/Microsoft/WindowsApps/pwsh.exe" if 'USERNAME' in globals() and USERNAME else None,
            "/mnt/c/Users/paulogoncalves/AppData/Local/Microsoft/WindowsApps/pwsh.exe",
        ]
        for candidate in ps7_wsl_candidates:
            if candidate and shutil.which(candidate):
                return candidate, "pwsh"

        # Fallback para Windows PowerShell 5.1 no WSL
        ps5_wsl_candidates = [
            shutil.which("powershell.exe"),
            "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe",
            "/mnt/c/WINDOWS/System32/WindowsPowerShell/v1.0/powershell.exe",
        ]
        for candidate in ps5_wsl_candidates:
            if candidate and shutil.which(candidate):
                return candidate, "windows_ps"

    # 3. Se estiver no Windows nativo
    if sys.platform == "win32":
        ps_bin = shutil.which("powershell") or shutil.which("powershell.exe")
        if ps_bin:
            return ps_bin, "windows_ps"

    return None, None


_sccm_cache = {}


def fetch_sccm_data(username: str) -> dict:
    """
    Consulta o SCCM via PowerShell (CIM/WMI) para obter os dados do último computador
    onde o usuário esteve logado, retornando um dicionário {"ip": "...", "hostname": "..."}.
    Suporta PowerShell 7 (pwsh) em containers Linux (Red Hat/Debian) e PowerShell no Windows/WSL.
    Utiliza um cache em memória para evitar consultas duplicadas na mesma execução.
    """
    if not username:
        return {"ip": "", "hostname": ""}
        
    username_lower = username.lower().strip()
    if username_lower in _sccm_cache:
        return _sccm_cache[username_lower]
        
    import subprocess
    import logging
    import json
    import re
    import time
    import keyring
    
    logger = logging.getLogger(__name__)
    res_data = {"ip": "N/A", "hostname": "N/A"}
    
    site_server = _cfg("SCCM_SERVER", os.getenv("SCCM_SERVER", ""))
    site_code = _cfg("SCCM_SITE_CODE", os.getenv("SCCM_SITE_CODE", ""))
    
    if not site_server or not site_code:
        logger.warning("⚠️ Variáveis 'SCCM_SERVER' ou 'SCCM_SITE_CODE' não configuradas. A consulta no SCCM será ignorada.")
        _sccm_cache[username_lower] = res_data
        return res_data

    # 1. Recupera as credenciais do administrador do SCCM (Banco de Dados / Keyring / Env)
    admin_user = _cfg("SCCM_ADMIN_USER", os.getenv("SCCM_ADMIN_USER", ""))
    admin_password = None
    if admin_user:
        # Tenta primeiro a senha salva/criptografada no banco
        admin_password = _cfg("SCCM_ADMIN_PASSWORD")
        if not admin_password:
            try:
                admin_password = keyring.get_password("sccm_admin", admin_user) or keyring.get_password("sccm", admin_user)
            except Exception as ke:
                logger.debug(f"[SCCM] Keyring indisponível ou inacessível no ambiente: {ke}")

        # Fallback essencial para Containers (Docker / Red Hat) onde não há GUI/Keyring interativo
        if not admin_password:
            admin_password = os.getenv("SCCM_ADMIN_PASSWORD") or os.getenv("SCCM_PASSWORD")
            if admin_password:
                logger.debug(f"[SCCM] Usando credencial de '{admin_user}' obtida via variável de ambiente (.env).")
    else:
        logger.warning("⚠️ Variável 'SCCM_ADMIN_USER' não configurada. A consulta no SCCM prosseguirá sem credenciais administrativas dedicadas.")

    # 2. Detecta o executável do PowerShell disponível (pwsh no Linux/Container ou powershell.exe no WSL/Windows)
    ps_executable, ps_flavor = _detect_powershell_executable()
    logger.debug(f"[SCCM] SO: {sys.platform} | Executável PS: '{ps_executable}' | Tipo: '{ps_flavor}'")

    if not ps_executable:
        logger.warning(
            f"[SCCM] Nenhum executável PowerShell ('pwsh' ou 'powershell.exe') encontrado no ambiente para consultar SCCM para '{username}'."
        )
        _sccm_cache[username_lower] = res_data
        return res_data

    # 3. Monta a consulta WMI/CIM com correspondência EXATA no LastLogonUserName
    query = f"SELECT * FROM SMS_R_System WHERE LastLogonUserName = '{username}'"
    
    # 4. Constrói o comando do PowerShell de acordo com a versão e credenciais
    if admin_password:
        logger.info(f"[SCCM] Consultando SCCM para '{username}' via {ps_flavor} (usando credenciais de '{admin_user}')")
        domain_user = admin_user
        if "\\" not in domain_user and "@" not in domain_user:
            short_domain = DOMINIO_CURTO or "MPE"
            domain_user = f"{short_domain}\\{admin_user}"
            
        escaped_password = admin_password.replace('"', '`"').replace('$', '`$')
        
        if ps_flavor == "pwsh":
            # No PowerShell 7 (Linux/Windows), Get-WmiObject foi removido; utiliza-se Get-CimInstance
            ps_command = (
                f'$secpasswd = ConvertTo-SecureString "{escaped_password}" -AsPlainText -Force; '
                f'$mycreds = New-Object System.Management.Automation.PSCredential ("{domain_user}", $secpasswd); '
                f'Get-CimInstance -ComputerName {site_server} -Namespace \'root\\sms\\site_{site_code}\' '
                f'-Query "{query}" -Credential $mycreds | Select-Object IPAddresses, Name | ConvertTo-Json'
            )
        else:
            # No Windows PowerShell 5.1, Get-WmiObject com PacketPrivacy garante compatibilidade DCOM completa
            ps_command = (
                f'$secpasswd = ConvertTo-SecureString "{escaped_password}" -AsPlainText -Force; '
                f'$mycreds = New-Object System.Management.Automation.PSCredential ("{domain_user}", $secpasswd); '
                f'Get-WmiObject -ComputerName {site_server} -Namespace \'root\\sms\\site_{site_code}\' '
                f'-Query "{query}" -Credential $mycreds -Authentication PacketPrivacy | Select-Object IPAddresses, Name | ConvertTo-Json'
            )
    else:
        logger.info(f"[SCCM] Consultando SCCM para '{username}' via {ps_flavor} (sem credenciais administrativas adicionais)")
        ps_command = f"Get-CimInstance -ComputerName {site_server} -Namespace 'root\\sms\\site_{site_code}' -Query \"{query}\" | Select-Object IPAddresses, Name | ConvertTo-Json"
    
    t0 = time.time()
    try:
        run_kwargs = {
            "capture_output": True,
            "text": True,
            "timeout": 10,
            "encoding": "utf-8",
            "errors": "replace",
        }
        if sys.platform == "win32":
            run_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            [ps_executable, "-NoProfile", "-NonInteractive", "-Command", ps_command],
            **run_kwargs
        )
        duration = time.time() - t0
        logger.debug(f"[SCCM] Consulta finalizada em {duration:.2f}s com código de retorno: {result.returncode}")
        
        if result.returncode != 0:
            error_output = result.stderr.strip()
            logger.debug(f"[SCCM] Stderr da execução: {error_output}")
            if any(term in error_output for term in ["Acesso negado", "Access denied", "Acesso Negado", "UnauthorizedAccessException"]):
                logger.warning(f"[SCCM] Acesso negado ao consultar SCCM para '{username}'. Requer privilégios elevados de rede.")
                res_data = {"ip": "Acesso Negado", "hostname": "Acesso Negado"}
                _sccm_cache[username_lower] = res_data
                return res_data
            logger.error(f"[SCCM] Erro ao consultar SCCM para '{username}': {error_output}")
            _sccm_cache[username_lower] = res_data
            return res_data
            
        output = result.stdout.strip()
        logger.debug(f"[SCCM] Stdout da execução (primeiros 200 chars): {output[:200]}")
        
        if not output:
            logger.info(f"[SCCM] Nenhum registro encontrado no SCCM para '{username}'.")
            _sccm_cache[username_lower] = res_data
            return res_data
            
        # Tenta decodificar o JSON retornado pelo PowerShell
        parsed = None
        try:
            parsed = json.loads(output)
        except Exception as je:
            logger.debug(f"[SCCM] Não foi possível decodificar saída JSON direta: {je}. Tentando extração por regex.")
            
        if parsed:
            items = parsed if isinstance(parsed, list) else [parsed]
            best_ip = ""
            best_hostname = ""
            
            for item in items:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("Name", "")).strip()
                ips = item.get("IPAddresses", [])
                if isinstance(ips, str):
                    ips = [ips]
                elif not isinstance(ips, list):
                    ips = []
                    
                # Prioriza IPs de rede interna que iniciam com '10.'
                ip_10 = next((ip for ip in ips if str(ip).startswith("10.")), None)
                if ip_10:
                    best_ip = str(ip_10)
                    best_hostname = name
                    break
                
                # Fallback para qualquer IPv4 válido encontrado
                if not best_ip:
                    ipv4_regex = re.compile(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$')
                    any_ipv4 = next((ip for ip in ips if ipv4_regex.match(str(ip))), None)
                    if any_ipv4:
                        best_ip = str(any_ipv4)
                        best_hostname = name
            
            if best_ip:
                res_data = {"ip": best_ip, "hostname": best_hostname}
            else:
                first_item = items[0] if items else {}
                name = str(first_item.get("Name", "")).strip()
                ips = first_item.get("IPAddresses", [])
                if isinstance(ips, str):
                    first_ip = ips
                elif isinstance(ips, list) and ips:
                    first_ip = str(ips[0])
                else:
                    first_ip = ""
                res_data = {"ip": first_ip, "hostname": name}
        else:
            # Fallback robusto por regex caso a saída não seja JSON puro
            ips = re.findall(r'10\.\d{1,3}\.\d{1,3}\.\d{1,3}', output)
            ip_val = ""
            if ips:
                ip_val = ips[0]
            else:
                ipv4s = re.findall(r'\d{1,3}\.\d{1,3}\.\d{1,3}', output)
                if ipv4s:
                    ip_val = ipv4s[0]
            name_match = re.search(r'"?Name"?\s*:\s*"?([^"\r\n\s]+)"?', output, re.IGNORECASE)
            name_val = name_match.group(1).strip() if name_match else ""
            res_data = {"ip": ip_val, "hostname": name_val}

        logger.info(f"[SCCM] Resultado para '{username}': IP={res_data.get('ip')} | Hostname={res_data.get('hostname')}")
        
    except subprocess.TimeoutExpired:
        logger.error(f"[SCCM] Timeout ({run_kwargs.get('timeout', 10)}s) ao consultar SCCM para '{username}'.")
        res_data = {"ip": "Timeout", "hostname": "Timeout"}
    except Exception as e:
        logger.error(f"[SCCM] Exceção ao consultar SCCM para '{username}': {e}", exc_info=True)
        res_data = {"ip": "Erro", "hostname": "Erro"}
        
    _sccm_cache[username_lower] = res_data
    return res_data


def fetch_ip_from_sccm(username: str) -> str:
    """
    Consulta o SCCM para obter o IP do último computador
    onde o usuário esteve logado.
    """
    return fetch_sccm_data(username)["ip"]


def fetch_hostname_from_sccm(username: str) -> str:
    """
    Consulta o SCCM para obter o Hostname do último computador
    onde o usuário esteve logado.
    """
    return fetch_sccm_data(username)["hostname"]


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
    
    # Aceita certificados SSL autoassinados / inseguros
    opts.accept_insecure_certs = True
    opts.add_argument('--ignore-certificate-errors')
    opts.add_argument('--ignore-ssl-errors=yes')
    opts.add_argument('--allow-insecure-localhost')
    
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
    else:
        opts.add_argument("--start-maximized")
    # Configurações essenciais para execução em containers Docker/Linux e headless
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-software-rasterizer")

    # Localiza binário do Chromium no Linux (incluindo o cache do Playwright)
    if sys.platform != "win32":
        possible_binaries = [
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
        ]
        playwright_cache = Path.home() / ".cache" / "ms-playwright"
        if playwright_cache.exists():
            for p in sorted(playwright_cache.glob("chromium-*/chrome-linux/chrome"), reverse=True):
                possible_binaries.insert(0, str(p))

        for binary in possible_binaries:
            if Path(binary).exists():
                opts.binary_location = str(binary)
                break

    from selenium.webdriver.chrome.service import Service
    if sys.platform == "win32":
        service = Service(creationflags=0x08000000)
    else:
        service = Service()
    driver = webdriver.Chrome(service=service, options=opts)
    return driver


def clean_otrs_description(desc: str) -> str:
    """
    Limpa de forma centralizada e altamente robusta a estrutura de formulários e campos do OTRS.
    Extrai apenas o conteúdo real escrito pelo usuário após campos como 'Descrição do Pedido:' ou 'Descrição:'.
    Remove saudações, termos de cortesia e assinaturas de e-mail.
    """
    import re
    import pandas as pd
    
    if pd.isna(desc):
        return ""
    text = str(desc).strip()
    
    # 1. Extração do conteúdo real após o cabeçalho estruturado do OTRS
    # Procura padrões comuns como "Descrição do Pedido:", "Descrição do pedido:", "Descrição:" (com ou sem acentos)
    match_desc = re.search(
        r'(?si)(?:descrição\s+(?:do\s+pedido|do\s+chamado)?|descricao\s+(?:do\s+pedido|do\s+chamado)?):\s*(.*)$',
        text
    )
    if match_desc:
        text = match_desc.group(1).strip()
        
    # 2. Remoção de rodapés específicos do OTRS
    text = re.sub(r'(?si)[\r\n]+(?:Para acompanhamento.*|É possível acompanhar.*)$', '', text)
    text = re.sub(r'(?m)^\.\.\.\s*$', '', text)
    text = re.sub(r'(?m)^Prazo:.*$', '', text)

    # 3. Remoção de saudações no início do texto
    text = re.sub(r'(?si)^(?:bom\s+dia|boa\s+tarde|boa\s+noite|prezados?|prezadas?|caros?|caras?|olá|ola)[,\-\s]*[\r\n]*', '', text)

    # 4. Remoção de despedidas e assinaturas
    padrao_despedida = r"(?si)\b(atenciosamente|att\.?|at\.te|grato|grata|obrigada?|obrigados|cordialmente|respeitosamente|saudações)\b[\s\S]*"
    text = re.sub(padrao_despedida, '', text)
    
    # Fallback para assinaturas no estilo e-mail ("--")
    text = re.sub(r"(?si)[\r\n]+--[\s\S]*", "", text)
    
    # 5. Descarta histórico de réplicas se houver
    parts = re.split(r'(?m)^#2\b', text, maxsplit=1)
    block1 = parts[0]
    
    # 6. Limpeza linha a linha de lixo residual do OTRS
    cleaned = []
    for line in block1.splitlines():
        l = line.strip()
        if not l:
            continue
        if re.fullmatch(r'[A-Z]{1,2}', l):
            continue
        if l.lower().startswith('responder a nota') or l.lower() in ('imprimir', 'dividir'):
            continue
        cleaned.append(l)
        
    return '\n'.join(cleaned).strip()


def clean_otrs_comments(comments_val) -> list:
    """
    Filtra e limpa a lista de comentários do OTRS.
    Ignora completamente comentários gerados automaticamente pela central de atendimento do suporte técnico
    (cujo autor contenha 'suporte@mpms.mp.br' ou 'Central de Atendimento').
    Retorna uma lista limpa de dicionários de comentários.
    """
    import json
    import pandas as pd
    
    if comments_val is None:
        return []
        
    if not isinstance(comments_val, (list, str)):
        try:
            if pd.isna(comments_val):
                return []
        except ValueError:
            pass
        
    if isinstance(comments_val, list):
        comments_list = comments_val
    elif isinstance(comments_val, str):
        val_stripped = comments_val.strip()
        if not val_stripped or val_stripped == '[]':
            return []
        try:
            comments_list = json.loads(val_stripped)
        except Exception:
            return []
    else:
        return []
        
    if not isinstance(comments_list, list):
        return []
        
    cleaned_comments = []
    for c in comments_list:
        if not isinstance(c, dict):
            continue
            
        autor = str(c.get('autor', '')).strip()
        
        # Ignora comentários gerados pela Central de Atendimento ao Usuário de TI
        if "suporte@mpms.mp.br" in autor or "Central de Atendimento ao Usuário" in autor:
            continue
            
        cleaned_comments.append(c)
        
    return cleaned_comments