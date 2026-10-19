# Configuração de Acesso Externo e Firewall (WSL2 / Windows)

Este documento orienta como liberar e persistir o acesso externo ao Sistema Bancada (Streamlit) na rede local do MPMS (ex: `http://mpe-80703:8501/` ou `http://<IP_DO_WINDOWS>:8501/`).

---

## 📌 Por que essa configuração é necessária?

No ambiente WSL2:
1. O Docker e a aplicação rodam em uma máquina virtual interna com uma faixa de IP própria (ex: `172.21.x.x`).
2. O Windows possui seu próprio IP na rede corporativa (ex: `10.111.144.90`).
3. Toda vez que o Windows é reiniciado, **o IP interno do WSL2 pode mudar**.
4. Para que outras máquinas da rede acessem a aplicação:
   - Uma regra de **Inbound (Entrada) no Firewall do Windows** precisa estar liberada para a porta `8501`.
   - Um redirecionamento de porta (**PortProxy**) precisa encaminhar o tráfego da porta `8501` do Windows para a porta `8501` do IP atual do WSL2.

---

## 🚀 Como Aplicar / Persistir Automaticamente

Criamos o script [`configurar_portproxy_firewall.ps1`](util/rede_firewall/configurar_portproxy_firewall.ps1) na pasta `util/rede_firewall/`.

### Opção 1: Executar Manualmente quando reiniciar (PowerShell Admin)
Abra o **PowerShell como Administrador** no Windows e execute:
```powershell
powershell -ExecutionPolicy Bypass -File "\\wsl$\Ubuntu\home\paulogoncalves\PythonProjects\automated-OTRS-and-CitSmart\util\rede_firewall\configurar_portproxy_firewall.ps1"
```
*(ou navegue até a pasta pelo Windows Explorer e execute com o botão direito `Executar com o PowerShell` em modo Admin).*

---

### Opção 2: Persistir Automaticamente no Boot / Logon (Agendador de Tarefas)

Para que essa regra nunca mais quebre mesmo que a máquina reinicie ou o IP do WSL mude, crie uma Tarefa Agendada no Windows:

1. Abra o **PowerShell como Administrador** no Windows.
2. Execute o comando abaixo uma única vez para registrar a tarefa agendada:

```powershell
$ScriptPath = "\\wsl$\Ubuntu\home\paulogoncalves\PythonProjects\automated-OTRS-and-CitSmart\util\rede_firewall\configurar_portproxy_firewall.ps1"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
$Trigger = New-ScheduledTaskTrigger -AtLogOn
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

Register-ScheduledTask -TaskName "WSL2_PortProxy_Bancada_8501" -Action $Action -Trigger $Trigger -Principal $Principal -Force
```

Dessa forma, sempre que o Windows iniciar ou você fizer logon, o script detectará o IP dinâmico atual do WSL2 e atualizará o PortProxy automaticamente.

---

## 🛠️ Comandos Manuais (Referência Rápida)

Se preferir rodar manualmente no PowerShell (Admin):

### 1. Descobrir o IP atual do WSL:
```powershell
$wslIp = (wsl -d Ubuntu hostname -I).Trim().Split(" ")[0]
Write-Host "IP do WSL: $wslIp"
```

### 2. Configurar o PortProxy:
```powershell
netsh interface portproxy delete v4tov4 listenport=8501 listenaddress=0.0.0.0
netsh interface portproxy add v4tov4 listenport=8501 listenaddress=0.0.0.0 connectport=8501 connectaddress=$wslIp
```

### 3. Criar / Garantir a Regra de Firewall:
```powershell
New-NetFirewallRule -DisplayName "Streamlit Bancada 8501" -Direction Inbound -LocalPort 8501 -Protocol TCP -Action Allow -Profile Domain,Private,Public -ErrorAction SilentlyContinue
```

### 4. Verificar se está ativo:
```powershell
netsh interface portproxy show v4tov4
Get-NetFirewallRule -DisplayName "*8501*"
```
