import keyring
import os
import sys
import getpass

def _get_username() -> str:
    # 1. Verifica variável de ambiente explícita
    env_user = os.getenv("AD_USER") or os.getenv("CITSMART_USER")
    if env_user and env_user.strip() and env_user.strip() != "root":
        return env_user.strip()
    
    # 2. Tenta pegar do ambiente Linux host ou USER
    host_user = os.getenv("USER") or os.getenv("USERNAME")
    if host_user and host_user.strip() and host_user.strip() != "root":
        return host_user.strip()
        
    try:
        user = getpass.getuser()
        if user and user.strip() and user.strip() != "root":
            return user.strip()
    except Exception:
        pass
        
    # Padrão fallback para paulogoncalves
    return "paulogoncalves"

def prompt_senha(service_name, username_label, default_user=None):
    """
    Verifica se a senha já existe no keyring para o serviço/usuário especificado.
    Se existir, mostra o status e pergunta se deseja atualizar ou manter (pular).
    """
    existing_user = keyring.get_password(f"{service_name}_user", service_name) if default_user is None else default_user
    if existing_user is None:
        user_to_check = username_label
    else:
        user_to_check = existing_user

    senha_atual = None
    try:
        senha_atual = keyring.get_password(service_name, user_to_check)
    except Exception:
        pass

    if senha_atual:
        print(f"🔒 [STATUS] Já existe uma senha salva para '{service_name}' (usuário: {user_to_check}).")
        resp = input("   Deseja ATUALIZAR esta senha? (S/N) [Padrão: N - Pular]: ").strip().upper()
        if resp != 'S':
            print("   ⏩ Mantendo a senha atual existente (pulado).\n")
            return False, user_to_check, senha_atual

    return True, user_to_check, None


# -----------------------------------------------------------------------------
# 1. SENHA DE REDE / AD (OTRS & CitSmart)
# -----------------------------------------------------------------------------
print("="*60)
print("🔑 CONFIGURAÇÃO DE SENHA DA REDE / AD (OTRS & CitSmart)")
print("="*60)

usuario_windows = _get_username()
deve_pedir, target_user, _ = prompt_senha("otrs", usuario_windows, default_user=usuario_windows)

if deve_pedir:
    senha_real = getpass.getpass(f"Digite a sua senha da rede/AD para o usuário '{usuario_windows}': ").strip()
    if senha_real:
        try:
            keyring.set_password("otrs", usuario_windows, senha_real)
            keyring.set_password("citSmart", usuario_windows, senha_real)
            print(f"✅ Senha de rede salva com sucesso para o usuário: {usuario_windows}\n")
        except Exception as e:
            print(f"⚠️ Não foi possível salvar no keyring: {e}\n")
    else:
        print("⚠️ Nenhuma senha digitada. Etapa pulada.\n")


# -----------------------------------------------------------------------------
# 2. USUÁRIO ADMINISTRADOR DO SCCM
# -----------------------------------------------------------------------------
print("="*60)
print("🔑 CONFIGURAÇÃO OPCIONAL: USUÁRIO ADMINISTRADOR DO SCCM")
print("="*60)
print("Se você deseja rodar a consulta de IPs no SCCM com privilégios de Administrador,")
print("podemos salvar a senha da conta administradora (ex: paulo_admin) no seu cofre.")

admin_user = "paulo_admin"
deve_pedir_sccm, target_admin, _ = prompt_senha("sccm_admin", admin_user, default_user=admin_user)

if deve_pedir_sccm:
    quer_salvar = input("Deseja salvar/atualizar a senha da conta de Admin para SCCM? (S/N): ").strip().upper()
    if quer_salvar == 'S':
        print(f"--> Configurando credenciais para: {admin_user}")
        senha_admin = getpass.getpass(f"Digite a senha de rede do usuário administrador ({admin_user}): ").strip()
        if senha_admin:
            try:
                keyring.set_password("sccm_admin", admin_user, senha_admin)
                print(f"✅ Senha de administrador salva com sucesso para a conta: {admin_user}!\n")
            except Exception as e:
                print(f"⚠️ Erro ao salvar senha no keyring: {e}\n")
        else:
            print("⚠️ Nenhuma senha digitada. Configuração de Admin ignorada.\n")
    else:
        print("ℹ️ Configuração do SCCM ignorada.\n")


# -----------------------------------------------------------------------------
# 3. USUÁRIO ADMINISTRADOR DO PAPERCUT
# -----------------------------------------------------------------------------
print("="*60)
print("🖨️ CONFIGURAÇÃO OPCIONAL: USUÁRIO ADMINISTRADOR DO PAPERCUT")
print("="*60)

current_pc_user = keyring.get_password("papercut_user", "papercut") or "admin"
deve_pedir_pc, _, _ = prompt_senha("papercut", current_pc_user, default_user=current_pc_user)

if deve_pedir_pc:
    quer_salvar_papercut = input("Deseja salvar/atualizar as credenciais de Admin do PaperCut? (S/N): ").strip().upper()
    if quer_salvar_papercut == 'S':
        pc_user = input(f"Digite o usuário do PaperCut (padrão: {current_pc_user}): ").strip() or current_pc_user
        pc_pass = getpass.getpass("Digite a senha do PaperCut: ").strip()
        if pc_pass:
            try:
                keyring.set_password("papercut_user", "papercut", pc_user)
                keyring.set_password("papercut", pc_user, pc_pass)
                print(f"✅ Credenciais do PaperCut salvas para o usuário: {pc_user}\n")
            except Exception as e:
                print(f"⚠️ Erro ao salvar senha no keyring: {e}\n")
        else:
            print("⚠️ Nenhuma senha digitada. Configuração do PaperCut ignorada.\n")
    else:
        print("ℹ️ Configuração do PaperCut ignorada.\n")


# -----------------------------------------------------------------------------
# 4. CENTRAL TELEFÔNICA (OXE)
# -----------------------------------------------------------------------------
print("="*60)
print("📞 CONFIGURAÇÃO OPCIONAL: CENTRAL TELEFÔNICA (OXE)")
print("="*60)

current_oxe_user = keyring.get_password("oxe_user", "oxe") or "mtcl"
deve_pedir_oxe, _, _ = prompt_senha("oxe", current_oxe_user, default_user=current_oxe_user)

if deve_pedir_oxe:
    oxe_input = input("Digite a senha do OXE (ou 'S' para alterar usuário / N para ignorar): ").strip()
    if oxe_input and oxe_input.upper() != 'N':
        if oxe_input.upper() == 'S':
            oxe_user = input(f"Digite o usuário do OXE (padrão: {current_oxe_user}): ").strip() or current_oxe_user
            oxe_pass = getpass.getpass(f"Digite a senha do OXE para o usuário '{oxe_user}': ").strip()
        else:
            oxe_user = current_oxe_user
            oxe_pass = oxe_input

        if oxe_pass:
            try:
                keyring.set_password("oxe_user", "oxe", oxe_user)
                keyring.set_password("oxe", oxe_user, oxe_pass)
                print(f"✅ Credenciais do OXE salvas com sucesso para o usuário: {oxe_user}\n")
            except Exception as e:
                print(f"⚠️ Erro ao salvar senha no keyring: {e}\n")
        else:
            print("⚠️ Nenhuma senha digitada. Configuração do OXE ignorada.\n")
    else:
        print("ℹ️ Configuração do OXE ignorada.\n")