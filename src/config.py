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

BASE_DIR              = Path(__file__).parent.parent
INPUT_DIR_BRUTOS      = BASE_DIR / "01 - Dados Brutos"
INPUT_DIR_BRUTOS.mkdir(exist_ok=True)
OUTPUT_DIR_TRATADOS   = BASE_DIR / "02 - Dados tratados"
OUTPUT_DIR_TRATADOS.mkdir(exist_ok=True)
OUTPUT_DIR_PRONTO     = BASE_DIR / "03 - Dados prontos"
OUTPUT_DIR_PRONTO.mkdir(exist_ok=True)
MODEL_DIR             = BASE_DIR / "models"
MODEL_DIR.mkdir(exist_ok=True)
MASTER_FILE_PATH = USER_HOME / os.getenv("SHAREPOINT_RELATIVE_PATH", "")
DONATIONS_FILE_PATH = USER_HOME / os.getenv("DONATIONS_EXCEL_RELATIVE_PATH", "")
SHAREPOINT_MATUTINO_URL = os.getenv("SHAREPOINT_MATUTINO_URL", "")
VIDEO_FAQ_DIR = USER_HOME / os.getenv("VIDEO_FAQ_PATH", "")

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

# Leaflet Map Logs
DEBUG_DIR_LEAFLET = BASE_DIR / "debug_logs" / "leaflet"
DEBUG_DIR_LEAFLET.mkdir(parents=True, exist_ok=True)

# FAQ Logs
DEBUG_DIR_FAQ = BASE_DIR / "debug_logs" / "faq"
DEBUG_DIR_FAQ.mkdir(parents=True, exist_ok=True)

# Plantao Logs
DEBUG_DIR_PLANTOES = BASE_DIR / "debug_logs" / "plantoes"
DEBUG_DIR_PLANTOES.mkdir(parents=True, exist_ok=True)

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


_sccm_cache = {}

def fetch_sccm_data(username: str) -> dict:
    """
    Consulta o SCCM via PowerShell (CIM/WMI) para obter os dados do último computador
    onde o usuário esteve logado, retornando um dicionário {"ip": "...", "hostname": "..."}.
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
    import keyring
    
    logger = logging.getLogger(__name__)
    
    site_server = os.getenv("SCCM_SERVER")
    site_code = os.getenv("SCCM_SITE_CODE")
    
    if not site_server or not site_code:
        logger.warning("⚠️ Variáveis 'SCCM_SERVER' ou 'SCCM_SITE_CODE' não configuradas no .env. A consulta no SCCM será ignorada.")
        return res_data

    
    # 1. Recupera as credenciais do administrador do SCCM no cofre de senhas do Windows
    admin_user = os.getenv("SCCM_ADMIN_USER")
    admin_password = None
    if admin_user:
        admin_password = keyring.get_password("sccm_admin", admin_user)
    else:
        logger.warning("⚠️ Variável 'SCCM_ADMIN_USER' não configurada no .env. A consulta no SCCM prosseguirá sem credenciais administrativas dedicadas.")


    
    # 2. Monta a consulta WMI com correspondência EXATA no LastLogonUserName
    query = f"SELECT * FROM SMS_R_System WHERE LastLogonUserName = '{username}'"
    
    # 3. Constrói o comando do PowerShell dependendo de termos ou não a senha do admin
    # Selecionamos IPAddresses e Name (que é o Hostname/NetbiosName no SCCM)
    if admin_password:
        logger.info(f"Consultando SCCM para o usuário: {username} (Usando credenciais de {admin_user})")
        domain_user = admin_user
        if "\\" not in domain_user and "@" not in domain_user:
            short_domain = DOMINIO_CURTO or "MPE"
            domain_user = f"{short_domain}\\{admin_user}"
            
        escaped_password = admin_password.replace('"', '`"').replace('$', '`$')
        
        ps_command = (
            f'$secpasswd = ConvertTo-SecureString "{escaped_password}" -AsPlainText -Force; '
            f'$mycreds = New-Object System.Management.Automation.PSCredential ("{domain_user}", $secpasswd); '
            f'Get-WmiObject -ComputerName {site_server} -Namespace \'root\\sms\\site_{site_code}\' '
            f'-Query "{query}" -Credential $mycreds -Authentication PacketPrivacy | Select-Object IPAddresses, Name | ConvertTo-Json'
        )
    else:
        logger.info(f"Consultando SCCM para o usuário: {username} (Sem credenciais adicionais)")
        ps_command = f"Get-CimInstance -ComputerName {site_server} -Namespace 'root\\sms\\site_{site_code}' -Query \"{query}\" | Select-Object IPAddresses, Name | ConvertTo-Json"
    
    res_data = {"ip": "", "hostname": ""}
    
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_command],
            capture_output=True,
            text=True,
            timeout=15,
            encoding='cp1252',
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        if result.returncode != 0:
            error_output = result.stderr.strip()
            if "Acesso negado" in error_output or "Access denied" in error_output or "Acesso Negado" in error_output:
                logger.warning(f"Acesso negado ao consultar SCCM para {username}. Requer privilégios elevados de rede.")
                res_data = {"ip": "Acesso Negado", "hostname": "Acesso Negado"}
                _sccm_cache[username_lower] = res_data
                return res_data
            logger.error(f"Erro ao consultar SCCM para {username}: {error_output}")
            _sccm_cache[username_lower] = res_data
            return res_data
            
        output = result.stdout.strip()
        if not output:
            logger.info(f"Nenhum registro encontrado no SCCM para {username}.")
            _sccm_cache[username_lower] = res_data
            return res_data
            
        # Tenta decodificar o JSON
        parsed = None
        try:
            parsed = json.loads(output)
        except Exception as je:
            logger.debug(f"Erro ao decodificar JSON do SCCM: {je}")
            
        if parsed:
            # Pode ser um dicionário ou uma lista de dicionários
            items = parsed if isinstance(parsed, list) else [parsed]
            
            # Vamos procurar um item com IP válido
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
                    
                # Procura IP que começa com 10.
                ip_10 = next((ip for ip in ips if str(ip).startswith("10.")), None)
                if ip_10:
                    best_ip = str(ip_10)
                    best_hostname = name
                    break
                
                # Se não achou 10., mas achou qualquer IPv4
                if not best_ip:
                    ipv4_regex = re.compile(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$')
                    any_ipv4 = next((ip for ip in ips if ipv4_regex.match(str(ip))), None)
                    if any_ipv4:
                        best_ip = str(any_ipv4)
                        best_hostname = name
            
            if best_ip:
                res_data = {"ip": best_ip, "hostname": best_hostname}
            else:
                # Se não achou IP pelas regras normais, pega o primeiro Name/IP
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
            # Fallback robusto por regex se o JSON falhar
            # Procura por IPs que começam com 10.
            ips = re.findall(r'10\.\d{1,3}\.\d{1,3}\.\d{1,3}', output)
            ip_val = ""
            if ips:
                ip_val = ips[0]
            else:
                ipv4s = re.findall(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', output)
                if ipv4s:
                    ip_val = ipv4s[0]
            # Procura pelo Name com ou sem aspas de forma robusta e flexível
            name_match = re.search(r'"?Name"?\s*:\s*"?([^"\r\n\s]+)"?', output, re.IGNORECASE)
            name_val = name_match.group(1).strip() if name_match else ""
            res_data = {"ip": ip_val, "hostname": name_val}

            
        logger.info(f"Dados do SCCM para {username}: {res_data}")
        
    except subprocess.TimeoutExpired:
        logger.error(f"Timeout ao consultar SCCM para {username}.")
        res_data = {"ip": "Timeout", "hostname": "Timeout"}
    except Exception as e:
        logger.error(f"Exceção ao consultar SCCM para {username}: {e}")
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