import os
import keyring
from datetime import datetime
from src.database.connection import get_connection, DB_TYPE
from src.crypto_utils import encrypt_value, decrypt_value

def setup_settings_table():
    """Cria a tabela de configurações do sistema se ela não existir."""
    conn = get_connection()
    cursor = conn.cursor()

    if DB_TYPE in ["postgres", "postgresql"]:
        create_sql = """
        CREATE TABLE IF NOT EXISTS system_settings (
            key VARCHAR(100) PRIMARY KEY,
            value TEXT,
            is_secret BOOLEAN DEFAULT FALSE,
            category VARCHAR(50),
            description TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """
    else:
        create_sql = """
        CREATE TABLE IF NOT EXISTS system_settings (
            key TEXT PRIMARY KEY,
            value TEXT,
            is_secret INTEGER DEFAULT 0,
            category TEXT,
            description TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """

    cursor.execute(create_sql)
    conn.commit()
    cursor.close()
    conn.close()

def get_setting(key: str, default: str = None) -> str:
    """
    Obtém uma configuração do banco de dados.
    Se for segredo/senha, decriptografa automaticamente.
    Se não existir no banco, retorna o default informado (ou fallback de env).
    """
    try:
        conn = get_connection()
        cursor = conn.cursor()
        if DB_TYPE in ["postgres", "postgresql"]:
            cursor.execute("SELECT value, is_secret FROM system_settings WHERE key = %s", (key,))
        else:
            cursor.execute("SELECT value, is_secret FROM system_settings WHERE key = ?", (key,))

        row = cursor.fetchone()
        cursor.close()
        conn.close()

        if row:
            val, is_secret = row[0], bool(row[1])
            if val is None:
                return default
            if is_secret:
                return decrypt_value(val)
            return val
    except Exception:
        pass

    return default

def set_setting(key: str, value: str, is_secret: bool = False, category: str = "geral", description: str = "") -> bool:
    """Salva ou atualiza uma configuração no banco (criptografando caso is_secret=True)."""
    try:
        setup_settings_table()
        val_to_save = encrypt_value(value) if is_secret and value else (value if value is not None else "")

        conn = get_connection()
        cursor = conn.cursor()

        if DB_TYPE in ["postgres", "postgresql"]:
            sql = """
            INSERT INTO system_settings (key, value, is_secret, category, description, updated_at)
            VALUES (%s, %s, %s, %s, %s, NOW())
            ON CONFLICT (key) DO UPDATE SET
                value = EXCLUDED.value,
                is_secret = EXCLUDED.is_secret,
                category = EXCLUDED.category,
                description = EXCLUDED.description,
                updated_at = NOW();
            """
            cursor.execute(sql, (key, val_to_save, is_secret, category, description))
        else:
            sql = """
            INSERT INTO system_settings (key, value, is_secret, category, description, updated_at)
            VALUES (?, ?, ?, ?, ?, datetime('now', 'localtime'))
            ON CONFLICT(key) DO UPDATE SET
                value = excluded.value,
                is_secret = excluded.is_secret,
                category = excluded.category,
                description = excluded.description,
                updated_at = datetime('now', 'localtime');
            """
            cursor.execute(sql, (key, val_to_save, 1 if is_secret else 0, category, description))

        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        print(f"Erro ao salvar configuração '{key}': {e}")
        return False

def get_all_settings(decrypt: bool = True) -> dict:
    """Retorna todas as configurações como um dicionário {chave: {'value': ..., 'is_secret': ..., 'category': ..., 'description': ...}}."""
    setup_settings_table()
    result = {}
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT key, value, is_secret, category, description, updated_at FROM system_settings")
        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        for r in rows:
            k, v, is_sec, cat, desc, up_at = r[0], r[1], bool(r[2]), r[3], r[4], r[5]
            val = decrypt_value(v) if (is_sec and decrypt and v) else v
            result[k] = {
                "value": val or "",
                "is_secret": is_sec,
                "category": cat or "geral",
                "description": desc or "",
                "updated_at": up_at
            }
    except Exception as e:
        print(f"Erro ao ler todas as configurações: {e}")
    return result

def seed_settings_from_env_if_empty(force: bool = False):
    """
    Se a tabela de configurações estiver vazia (ou se force=True),
    importa automaticamente as variáveis já existentes no .env e no keyring,
    garantindo que o usuário comece com tudo 100% preenchido.
    """
    setup_settings_table()
    current = get_all_settings(decrypt=False)

    # Mapeamento com categorias e flags de segredo
    # Obtém senhas atuais do keyring se existirem
    ad_user = os.getenv("AD_USER", "paulogoncalves")
    try:
        ad_pass = os.getenv("AD_PASSWORD") or keyring.get_password("otrs", ad_user) or ""
    except Exception:
        ad_pass = os.getenv("AD_PASSWORD", "")

    sccm_user = os.getenv("SCCM_ADMIN_USER", "paulo_admin")
    try:
        sccm_pass = os.getenv("SCCM_ADMIN_PASSWORD") or keyring.get_password("sccm_admin", sccm_user) or keyring.get_password("sccm", sccm_user) or ""
    except Exception:
        sccm_pass = os.getenv("SCCM_ADMIN_PASSWORD", "")

    pc_user = os.getenv("PAPERCUT_USER", "admin")
    try:
        pc_pass = os.getenv("PAPERCUT_PASS") or keyring.get_password("papercut", pc_user) or ""
    except Exception:
        pc_pass = os.getenv("PAPERCUT_PASS", "")

    oxe_user = os.getenv("OXE_USER", "mtcl")
    try:
        oxe_pass = os.getenv("OXE_PASS") or keyring.get_password("oxe", oxe_user) or ""
    except Exception:
        oxe_pass = os.getenv("OXE_PASS", "")

    defaults = [
        # Active Directory & Rede
        ("AD_USER", ad_user, False, "rede", "Usuário de Rede / Active Directory"),
        ("AD_PASSWORD", ad_pass, True, "rede", "Senha de Rede / Active Directory (OTRS & CitSmart)"),
        ("AD_DOMAIN", os.getenv("AD_DOMAIN", "in.mpe.ms.gov.br"), False, "rede", "Domínio FQDN do Active Directory"),
        ("AD_MMC", os.getenv("AD_MMC", "DC=in,DC=mpe,DC=ms,DC=gov,DC=br"), False, "rede", "Base DN / MMC do Active Directory"),
        ("AD_SHORT", os.getenv("AD_SHORT", "MPE"), False, "rede", "Nome NetBIOS curto do Domínio"),
        ("AD_EMAIL", os.getenv("AD_EMAIL", "mpms.mp.br"), False, "rede", "Sufixo de e-mail institucional"),

        # SCCM
        ("SCCM_ADMIN_USER", sccm_user, False, "sccm", "Conta de Administrador para consultas SCCM"),
        ("SCCM_ADMIN_PASSWORD", sccm_pass, True, "sccm", "Senha da conta Administradora do SCCM"),
        ("SCCM_SERVER", os.getenv("SCCM_SERVER", "srv-1046.in.mpe.ms.gov.br"), False, "sccm", "Servidor FQDN do SCCM"),
        ("SCCM_SITE_CODE", os.getenv("SCCM_SITE_CODE", "PGJ"), False, "sccm", "Código do Site do SCCM"),

        # PaperCut
        ("PAPERCUT_USER", pc_user, False, "papercut", "Usuário Administrador do PaperCut"),
        ("PAPERCUT_PASS", pc_pass, True, "papercut", "Senha do Administrador do PaperCut"),
        ("PAPERCUT_URL", os.getenv("PAPERCUT_URL", "http://impressora.mpms.mp.br:9191/admin"), False, "papercut", "URL do painel administrativo do PaperCut"),
        ("PAPERCUT_PRINTER_LIST_URL", os.getenv("PAPERCUT_PRINTER_LIST_URL", "http://impressora.mpms.mp.br:9191/app?service=page/PrinterList"), False, "papercut", "URL da listagem de impressoras"),
        ("PAPERCUT_DEVICE_LIST_URL", os.getenv("PAPERCUT_DEVICE_LIST_URL", "http://impressora.mpms.mp.br:9191/app?service=page/DeviceList"), False, "papercut", "URL da listagem de dispositivos multifuncionais"),

        # Telefonia (OXE)
        ("OXE_USER", oxe_user, False, "oxe", "Usuário de acesso à Central Telefônica OXE"),
        ("OXE_PASS", oxe_pass, True, "oxe", "Senha de acesso à Central Telefônica OXE"),
        ("OXE_URL", os.getenv("OXE_URL", "https://10.12.32.30"), False, "oxe", "URL da Central Telefônica OXE"),

        # IA Gemini
        ("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""), True, "ia", "Chave da API do Google Gemini"),

        # Portais e URLs
        ("CITSMART_LINK", os.getenv("CITSMART_LINK", "https://suporte.mpms.mp.br"), False, "urls", "URL principal do CitSmart"),
        ("CITSMART_LINK_NOVO", os.getenv("CITSMART_LINK_NOVO", "https://suporte.mpms.mp.br/inbox/lowcode/form/copilot_novo/default"), False, "urls", "URL do formulário de abertura de chamados do CitSmart"),
        ("OTRS_LINK", os.getenv("OTRS_LINK", "https://central.mpms.mp.br"), False, "urls", "URL do OTRS"),
        ("ATOS_NORMAS_API_URL", os.getenv("ATOS_NORMAS_API_URL", "https://www.mpms.mp.br/atos-e-normas/listAll"), False, "urls", "Endpoint da API de Atos e Normas MPMS"),
        ("ATOS_NORMAS_DOWNLOAD_URL", os.getenv("ATOS_NORMAS_DOWNLOAD_URL", "https://www.mpms.mp.br/atos-e-normas/download/"), False, "urls", "URL base para download de Atos e Normas"),

        # Planilhas e Nuvem
        ("SHAREPOINT_RELATIVE_PATH", os.getenv("SHAREPOINT_RELATIVE_PATH", r"OneDrive - Ministerio Público do Estado de Mato Grosso do Sul\Documentos SharePoint DIT-Manutenção\Chamados\Chamados_Unificados_Final.xlsx"), False, "sharepoint", "Caminho relativo da Planilha de Chamados no OneDrive/SharePoint"),
        ("DONATIONS_EXCEL_RELATIVE_PATH", os.getenv("DONATIONS_EXCEL_RELATIVE_PATH", ""), False, "sharepoint", "URL/Caminho da Planilha de Doações e Baixas"),
        ("FISCAL_EXCEL_RELATIVE_PATH", os.getenv("FISCAL_EXCEL_RELATIVE_PATH", ""), False, "sharepoint", "URL/Caminho da Planilha de Fiscalização de Contratos"),
        ("WARRANTY_EXCEL_RELATIVE_PATH", os.getenv("WARRANTY_EXCEL_RELATIVE_PATH", ""), False, "sharepoint", "URL/Caminho da Planilha de Garantia"),
        ("VIAGENS_EXCEL_RELATIVE_PATH", os.getenv("VIAGENS_EXCEL_RELATIVE_PATH", ""), False, "sharepoint", "URL/Caminho da Planilha de Viagens da Bancada"),
        ("SHAREPOINT_MATUTINO_URL", os.getenv("SHAREPOINT_MATUTINO_URL", ""), False, "sharepoint", "URL da Planilha de Escala Matutina"),
        ("VIDEO_FAQ_PATH", os.getenv("VIDEO_FAQ_PATH", ""), False, "sharepoint", "URL da pasta de Vídeos FAQ no SharePoint"),
        ("IMAGE_FAQ_PATH", os.getenv("IMAGE_FAQ_PATH", ""), False, "sharepoint", "URL da pasta de Imagens FAQ no SharePoint"),
    ]

    for key, val, is_sec, cat, desc in defaults:
        if force or key not in current:
            set_setting(key, val, is_secret=is_sec, category=cat, description=desc)
