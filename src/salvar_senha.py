import keyring
import os

def _get_username() -> str:
    env_user = os.getenv("AD_USER") or os.getenv("CITSMART_USER") or os.getenv("USER") or os.getenv("USERNAME")
    if env_user and env_user.strip():
        return env_user.strip()
    try:
        import getpass
        user = getpass.getuser()
        if user and user.strip():
            return user.strip()
    except Exception:
        pass
    try:
        return os.getlogin()
    except Exception:
        pass
    return "usuario"

# Pega o mesmo usuário que o config.py usa (username) para o contexto diário
usuario_windows = _get_username() 
senha_real = input("Digite a sua senha da rede/AD (para OTRS/CitSmart): ")

# Salva no cofre do Windows para o usuário logado
keyring.set_password("otrs", usuario_windows, senha_real)
keyring.set_password("citSmart", usuario_windows, senha_real)
print(f"✅ Senha salva com sucesso no cofre do Windows para o usuário: {usuario_windows}")

print("\n" + "="*60)
print("🔑 CONFIGURAÇÃO OPCIONAL: USUÁRIO ADMINISTRADOR DO SCCM")
print("="*60)
print("Se você deseja rodar a consulta de IPs no SCCM com privilégios de Administrador,")
print("podemos salvar a senha da conta administradora (ex: paulo_admin) no seu cofre.")
print("Assim, o script diário continuará rodando com o seu usuário comum, mas fará")
print("as consultas do SCCM fingindo ser o administrador com segurança!")

quer_salvar = input("\nDeseja salvar a senha da conta de Admin para SCCM? (S/N): ").strip().upper()
if quer_salvar == 'S':
    admin_user = "paulo_admin"
    print(f"\n--> Configurando credenciais para: {admin_user}")
    senha_admin = input("Digite a senha de rede do usuário administrador (paulo_admin): ").strip()
    if senha_admin:
        keyring.set_password("sccm_admin", admin_user, senha_admin)
        print(f"\n✅ Senha de administrador salva com sucesso para a conta: {admin_user}!")
    else:
        print("\n⚠️ Nenhuma senha digitada. Configuração de Admin ignorada.")
print("\n" + "="*60)
print("🖨️ CONFIGURAÇÃO OPCIONAL: USUÁRIO ADMINISTRADOR DO PAPERCUT")
print("="*60)
quer_salvar_papercut = input("Deseja salvar as credenciais de Admin do PaperCut? (S/N): ").strip().upper()
if quer_salvar_papercut == 'S':
    pc_user = input("Digite o usuário do PaperCut (padrão: admin): ").strip() or "admin"
    pc_pass = input("Digite a senha do PaperCut: ").strip()
    if pc_pass:
        keyring.set_password("papercut_user", "papercut", pc_user)
        keyring.set_password("papercut", pc_user, pc_pass)
        print(f"✅ Credenciais do PaperCut salvas para o usuário: {pc_user}")
    else:
        print("⚠️ Nenhuma senha digitada. Configuração do PaperCut ignorada.")
else:
    print("ℹ️ Configuração do PaperCut ignorada.")

print("\n" + "="*60)
print("📞 CONFIGURAÇÃO OPCIONAL: CENTRAL TELEFÔNICA (OXE)")
print("="*60)
oxe_input = input("Digite a senha do OXE (ou 'S' para alterar usuário / ENTER para ignorar): ").strip()
if oxe_input and oxe_input.upper() != 'N':
    if oxe_input.upper() == 'S':
        oxe_user = input("Digite o usuário do OXE (padrão: mtcl): ").strip() or "mtcl"
        oxe_pass = input(f"Digite a senha do OXE para o usuário '{oxe_user}': ").strip()
    else:
        oxe_user = "mtcl"
        oxe_pass = oxe_input

    if oxe_pass:
        keyring.set_password("oxe_user", "oxe", oxe_user)
        keyring.set_password("oxe", oxe_user, oxe_pass)
        print(f"✅ Credenciais da Central Telefônica (OXE) salvas com sucesso no cofre do Windows para o usuário: {oxe_user}")
    else:
        print("⚠️ Nenhuma senha digitada. Configuração do OXE ignorada.")
else:
    print("ℹ️ Configuração do OXE ignorada.")