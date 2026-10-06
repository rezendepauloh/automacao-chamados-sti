import os
import time
from typing import Optional, Dict, Any, Tuple
import itsdangerous
import streamlit as st

from src.crypto_utils import _get_or_create_key
from src.services.ad_ldap_service import authenticate_user_credentials
from src.terminal import log, GREEN, YELLOW, RED, CYAN

# Cookie e Token constantes
COOKIE_NAME = "bancada_auth_token"
AUTH_SALT = "bancada-auth-token-salt"
SESSION_DURATION_DAYS = 30
SESSION_DURATION_SECONDS = SESSION_DURATION_DAYS * 24 * 60 * 60

# Mapeamento oficial dos Técnicos da Bancada (Administradores com Acesso Total)
# Conforme especificado na Seção 5.1 do CONTEXTO_GERAL.md
ADMIN_USERS: Dict[str, Dict[str, Any]] = {
    "paulogoncalves": {
        "display_name": "Paulo Henrique Gonçalves Rezende",
        "role": "admin",
        "admin_sys": "paulo_admin",
        "admin_ad": "paulo_admin_ad",
        "email": "paulogoncalves@mpms.mp.br",
    },
    "reginaldosb": {
        "display_name": "Reginaldo da Silva Bandeira",
        "role": "admin",
        "admin_sys": "reginaldo_admin",
        "admin_ad": "reginaldo_admin_ad",
        "email": "reginaldosb@mpms.mp.br",
    },
    "luizvillalba": {
        "display_name": "Luiz Leonardo Villalba",
        "role": "admin",
        "admin_sys": "villalba_admin",
        "admin_ad": "villalba_admin_ad",
        "email": "luizvillalba@mpms.mp.br",
    },
}

def get_auth_secret() -> str:
    """Retorna a chave secreta HMAC para assinatura e validação do token."""
    secret = os.getenv("JWT_SECRET_KEY") or os.getenv("APP_SECRET_KEY")
    if secret and secret.strip():
        return secret.strip()
    return _get_or_create_key().decode()

def _get_serializer() -> itsdangerous.URLSafeTimedSerializer:
    """Retorna o serializador seguro assinado com timestamp embutido."""
    return itsdangerous.URLSafeTimedSerializer(
        secret_key=get_auth_secret(),
        salt=AUTH_SALT
    )

def create_auth_token(username: str) -> str:
    """Gera um token seguro assinado digitalmente com dados de perfil do usuário e validade longa."""
    clean_user = username.strip().lower()
    user_info = ADMIN_USERS.get(clean_user, {
        "display_name": clean_user,
        "role": "viewer",
        "admin_sys": None,
        "admin_ad": None,
        "email": f"{clean_user}@mpms.mp.br",
    })

    now = int(time.time())
    payload = {
        "sub": clean_user,
        "username": clean_user,
        "display_name": user_info["display_name"],
        "role": user_info["role"],
        "admin_sys": user_info.get("admin_sys"),
        "admin_ad": user_info.get("admin_ad"),
        "iat": now,
    }
    s = _get_serializer()
    return s.dumps(payload)

def decode_auth_token(token: str) -> Optional[Dict[str, Any]]:
    """Decodifica e valida a assinatura criptográfica e tempo de expiração do token."""
    if not token or not str(token).strip():
        return None
    try:
        s = _get_serializer()
        payload = s.loads(str(token).strip(), max_age=SESSION_DURATION_SECONDS)
        return payload
    except itsdangerous.SignatureExpired:
        return None
    except itsdangerous.BadSignature:
        return None
    except Exception:
        return None

def get_current_user() -> Optional[Dict[str, Any]]:
    """Retorna os dados do usuário autenticado no session_state atual."""
    return st.session_state.get("authenticated_user")

def is_authenticated() -> bool:
    """Verifica se há um usuário válido na sessão."""
    user = get_current_user()
    return bool(user and user.get("username"))

def is_admin() -> bool:
    """Verifica se o usuário logado tem privilégios de Administrador."""
    user = get_current_user()
    return bool(user and user.get("role") == "admin")

def can_edit() -> bool:
    """Alias semântico para verificar se o usuário pode realizar alterações."""
    return is_admin()

def inject_cookie_setter(token: str):
    """Injeta JavaScript para gravar o cookie seguro com validade de 30 dias no navegador."""
    js_code = f"""
    <script>
        const cookieName = "{COOKIE_NAME}";
        const cookieVal = "{token}";
        const maxAge = {SESSION_DURATION_SECONDS};
        document.cookie = `${{cookieName}}=${{cookieVal}}; max-age=${{maxAge}}; path=/; SameSite=Lax`;
        try {{
            localStorage.setItem(cookieName, cookieVal);
        }} catch(e) {{}}
    </script>
    """
    st.components.v1.html(js_code, height=0, width=0)

def inject_cookie_remover():
    """Injeta JavaScript para apagar o cookie e localStorage no logout."""
    js_code = f"""
    <script>
        document.cookie = "{COOKIE_NAME}=; max-age=0; path=/; SameSite=Lax";
        try {{
            localStorage.removeItem("{COOKIE_NAME}");
        }} catch(e) {{}}
    </script>
    """
    st.components.v1.html(js_code, height=0, width=0)

def restore_session_from_cookie() -> bool:
    """
    Tenta restaurar a sessão a partir do cookie presente no st.context.cookies ou query params.
    Retorna True se a sessão foi restaurada com sucesso.
    """
    if is_authenticated():
        return True

    # 1. Tenta recuperar do st.context.cookies (Streamlit 1.57+)
    token = None
    try:
        if hasattr(st, "context") and hasattr(st.context, "cookies"):
            token = st.context.cookies.get(COOKIE_NAME)
    except Exception:
        pass

    # 2. Fallback: suporte a token transmitido em parâmetro ou session state
    if not token:
        token = st.session_state.get("pending_auth_token")

    if not token:
        return False

    payload = decode_auth_token(token)
    if payload:
        st.session_state["authenticated_user"] = payload
        st.session_state["auth_token"] = token
        return True

    return False

def login_user(username: str, password: str) -> Tuple[bool, str]:
    """
    Valida as credenciais corporativas no AD/LDAP e autentica o usuário,
    gerando o token assinado e registrando no session_state.
    """
    sucesso, msg = authenticate_user_credentials(username, password)
    if not sucesso:
        return False, msg

    clean_user = username.strip().lower()
    token = create_auth_token(clean_user)
    payload = decode_auth_token(token)

    st.session_state["authenticated_user"] = payload
    st.session_state["auth_token"] = token
    st.session_state["just_logged_in"] = True

    return True, "Login realizado com sucesso!"

def logout_user():
    """Encerra a sessão atual, limpa st.session_state e programa remoção do cookie no navegador."""
    st.session_state["authenticated_user"] = None
    st.session_state["auth_token"] = None
    st.session_state["do_logout_cleanup"] = True
    st.rerun()

def render_login_page():
    """Renderiza a tela de login moderna com identidade visual da Bancada STI / MPMS."""
    # Se uma limpeza de logout foi solicitada, aciona script JS
    if st.session_state.get("do_logout_cleanup"):
        inject_cookie_remover()
        st.session_state["do_logout_cleanup"] = False

    st.markdown("""
    <style>
    .login-container {
        max-width: 440px;
        margin: 40px auto 20px auto;
        padding: 30px;
        background: rgba(30, 41, 59, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 14px;
        backdrop-filter: blur(12px);
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
        text-align: center;
    }
    .login-logo {
        font-size: 3rem;
        margin-bottom: 10px;
    }
    .login-title {
        font-size: 1.6rem;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 6px;
    }
    .login-subtitle {
        font-size: 0.95rem;
        color: #94a3b8;
        margin-bottom: 25px;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div class="login-container">
            <div class="login-logo">🛡️</div>
            <div class="login-title">Bancada STI</div>
            <div class="login-subtitle">Autenticação com Credenciais de Rede (Active Directory)</div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("form_bancada_login", clear_on_submit=False):
            st.markdown("##### 🔑 Entrar com sua conta de rede")
            user_input = st.text_input(
                "Usuário do Domínio (sAmAccountName)",
                placeholder="Ex: paulogoncalves ou reginaldosb",
                key="login_user_field"
            )
            pass_input = st.text_input(
                "Senha de Rede",
                type="password",
                placeholder="••••••••••••",
                key="login_pass_field"
            )

            btn_entrar = st.form_submit_button("🚀 Entrar no Sistema", use_container_width=True, type="primary")

            if btn_entrar:
                if not user_input or not pass_input:
                    st.error("Por favor, preencha o usuário e a senha de rede.")
                else:
                    with st.spinner("Autenticando contra o Active Directory..."):
                        sucesso, msg = login_user(user_input, pass_input)
                        if sucesso:
                            st.success("Autenticado com sucesso! Carregando painel...")
                            st.rerun()
                        else:
                            st.error(f"Falha de autenticação: {msg}")

        st.caption("🔒 Sessão persistente por até 30 dias. Para encerrar, utilize o botão de Sair no menu.")
