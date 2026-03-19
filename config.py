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