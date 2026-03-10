import keyring
import os

# Pega o mesmo usuário que o config.py usa (username)
usuario_windows = os.getlogin() 
senha_real = input("Digite a sua senha da rede/AD: ")

# Salva no cofre do Windows
keyring.set_password("otrs", usuario_windows, senha_real)
keyring.set_password("citSmart", usuario_windows, senha_real)
print(f"✅ Senha salva com sucesso no cofre do Windows para o usuário: {usuario_windows}")