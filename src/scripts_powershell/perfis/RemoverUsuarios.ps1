<#
.SYNOPSIS
  Script para habilitar WinRM (via DCOM/WS-Man) e remover perfis de usuário locais inativos remotamente em uma estação Windows.

.DESCRIPTION
  Este script realiza as seguintes etapas em ordem:
    1. Carrega credenciais criptografadas (DPAPI) de um arquivo XML.
    2. Testa conectividade ICMP com o host remoto especificado.
    3. Tenta criar uma CimSession via DCOM (RPC) para o host remoto e, se necessário:
       a) Habilita o serviço WinRM no host remoto via DCOM.
       b) Abre a exceção de firewall para permitir o tráfego WinRM no host remoto.
       c) Verifica se o WinRM está respondendo. Caso falhe, tenta criar CimSession via WS-Man.
    4. Para cada usuário especificado em UsersToPurge:
       a) Tenta remover a conta local (net user <Usuario> /delete) na estação remota.
       b) Busca perfis locais cujo caminho termine em “\<UserName>” e, se encontrado:
          i. Se o perfil estiver carregado (em uso ativo), emite um aviso e pula a remoção, assegurando a integridade do usuário ativo na estação.
          ii. Tenta Win32_UserProfile.Delete() para remover a pasta e a chave de registro de forma limpa.
          iii. Se Win32_UserProfile.Delete() falhar, executa um fallback manual:
               - Remove a pasta “C:\Users\<UserName>” via rmdir /S /Q.
               - Remove a chave de registro em “HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\<SID>” usando StdRegProv.
    5. Fecha a CimSession ao final do processo e gera um log de execução na pasta raiz do script.

.PARAMETER ComputerName
  Nome ou IP da estação Windows onde os perfis serão removidos (exemplo: "PGJ-58099").

.PARAMETER UsersToPurge
  Conjunto de strings indicando os nomes “básicos” das pastas de usuário que se deseja remover.
  Exemplo: @("pedrofonseca","amandabarbosa").

.EXAMPLE
  PS> .\Remove-RemoteUserProfiles.ps1 -ComputerName "MPE-58063" -UsersToPurge @("muriloananias","patrickoliveira","lohanlima")
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [Parameter(Mandatory = $true)]
    [string[]]$UsersToPurge
)

# Normaliza o array de usuários caso algum elemento contenha múltiplos logins separados por vírgula (comum em chamadas externas CLI/GUI)
$cleanUsers = @()
foreach ($item in $UsersToPurge) {
    if ($item -match '[,; ]') {
        $cleanUsers += $item.Split(@(',', ';', ' '), [System.StringSplitOptions]::RemoveEmptyEntries)
    } else {
        $cleanUsers += $item
    }
}
$UsersToPurge = $cleanUsers

#==================================================================================
# 1) CONFIGURAÇÃO DE LOGGING
#==================================================================================

$LogPath = Join-Path -Path $PSScriptRoot -ChildPath "Remove-RemoteUserProfiles.log"

# Rotação automática de logs (Limite de 3 MB, mantendo no máximo 3 arquivos de histórico)
try {
    if (Test-Path $LogPath) {
        $logItem = Get-Item $LogPath -ErrorAction SilentlyContinue
        if ($logItem -and $logItem.Length -gt 3MB) {
            for ($i = 2; $i -ge 1; $i--) {
                $oldLog = "$LogPath.$i"
                $newLog = "$LogPath." + ($i + 1)
                if (Test-Path $oldLog) {
                    Move-Item -Path $oldLog -Destination $newLog -Force -ErrorAction SilentlyContinue
                }
            }
            Move-Item -Path $LogPath -Destination "$LogPath.1" -Force -ErrorAction SilentlyContinue
        }
    }
} catch {
    # Ignora falhas silenciosas na rotação de logs para não interromper a execução do fluxo principal
}

function Write-Log {
    <#
    .SYNOPSIS
      Função auxiliar padronizada para registro em console e arquivo de log.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"

    # Saída no console
    switch ($Level) {
        "INFO"    { Write-Host $logLine -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $logLine -ForegroundColor Green }
        "WARNING" { Write-Warning $Message }
        "ERROR"   { Write-Error $Message }
    }

    # Gravar no arquivo de log
    Add-Content -Path $LogPath -Value $logLine
}

Write-Log "=== Iniciando processo de limpeza de perfis remotos em '$ComputerName' ==="
Write-Log "Arquivo de log inicializado em: $LogPath"

#==================================================================================
# 2) CARREGAR CREDENCIAIS CRIPTOGRAFADAS VIA DPAPI
#==================================================================================

$CredXmlPath = Join-Path $PSScriptRoot "cred_admin.xml"

if (-not (Test-Path $CredXmlPath)) {
    Write-Log "Arquivo de credencial não encontrado em '$CredXmlPath'." -Level "ERROR"
    exit 1
}

try {
    $script:CredentialAdmin = Import-Clixml -Path $CredXmlPath
    Write-Log "Credenciais DPAPI carregadas com sucesso." -Level "SUCCESS"
}
catch {
    Write-Log "Falha ao importar credenciais de '$CredXmlPath': $_" -Level "ERROR"
    exit 1
}

#==================================================================================
# 3) TESTE DE CONECTIVIDADE (PING ÚNICO)
#==================================================================================

Write-Log "Testando conectividade ICMP com '$ComputerName'..."
if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
    Write-Log "Máquina '$ComputerName' está inacessível via ICMP. Abortando." -Level "ERROR"
    exit 2
}
Write-Log "Host '$ComputerName' está acessível via ICMP." -Level "SUCCESS"

#==================================================================================
# 4) FUNÇÕES AUXILIARES PARA CONFIGURAR WinRM VIA DCOM
#==================================================================================

function Enable-WinRM-viaDCOM {
    param([Parameter(Mandatory = $true)][CimSession]$CimSession)
    try {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'" -CimSession $CimSession -ErrorAction Stop
        if ($svc.StartMode -ne "Auto") {
            $null = Invoke-CimMethod -InputObject $svc -MethodName ChangeStartMode -Arguments @{ StartMode = "Automatic" } -CimSession $CimSession -ErrorAction Stop
            Write-Log "Serviço WinRM configurado para Inicialização Automática." -Level "SUCCESS"
        }
        if ($svc.State -ne "Running") {
            $null = Invoke-CimMethod -InputObject $svc -MethodName StartService -CimSession $CimSession -ErrorAction Stop
            Write-Log "Serviço WinRM iniciado com sucesso." -Level "SUCCESS"
        }
    }
    catch {
        Write-Log "Falha ao habilitar/iniciar WinRM via DCOM: $_" -Level "WARNING"
    }
}

function Open-Firewall-WinRM {
    param([Parameter(Mandatory = $true)][CimSession]$CimSession)
    $cmd = 'netsh advfirewall firewall set rule group="Windows Remote Management" new enable=yes'
    try {
        $proc = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $cmd } -CimSession $CimSession -ErrorAction Stop
        if ($proc.ReturnValue -eq 0) {
            Write-Log "Regra de Firewall para WinRM habilitada com sucesso via DCOM." -Level "SUCCESS"
        } else {
            Write-Log "Falha ao habilitar regra de Firewall (código $($proc.ReturnValue))." -Level "WARNING"
        }
    }
    catch {
        Write-Log "Erro ao executar netsh via DCOM: $_" -Level "WARNING"
    }
}

function Set-TrustedHosts-viaDCOM {
    param(
        [Parameter(Mandatory = $true)][CimSession]$CimSession,
        [Parameter(Mandatory = $true)][string]$Hosts
    )
    $HKLM      = [uint32]2147483650
    $subKey    = "SOFTWARE\Policies\Microsoft\Windows\WinRM\Client"
    $valueName = "TrustedHosts"

    $argsReg = @{
        hDefKey     = $HKLM
        sSubKeyName = $subKey
        sValueName  = $valueName
        sValue      = $Hosts
    }

    try {
        Invoke-CimMethod -Namespace root\default -ClassName StdRegProv -MethodName SetStringValue -Arguments $argsReg -CimSession $CimSession -ErrorAction Stop | Out-Null
        Write-Log "TrustedHosts em HKLM:\$subKey\$valueName definido para '$Hosts'." -Level "SUCCESS"
    }
    catch {
        Write-Log "Falha ao ajustar TrustedHosts via StdRegProv: $_" -Level "WARNING"
    }
}

#==================================================================================
# 5) CONECTIVIDADE CIM: DCOM PRIMEIRO, COM FALLBACK AUTOMÁTICO PARA WS-MAN
#==================================================================================

Write-Log "Estabelecendo sessão CIM com '$ComputerName'..."
try {
    Write-Log "Tentando conexão DCOM (RPC) com '$ComputerName'..."
    $optDcom = New-CimSessionOption -Protocol Dcom
    $script:CimSession = New-CimSession -ComputerName $ComputerName -Credential $script:CredentialAdmin -SessionOption $optDcom -ErrorAction Stop

    Enable-WinRM-viaDCOM -CimSession $script:CimSession
    Open-Firewall-WinRM -CimSession $script:CimSession
    Set-TrustedHosts-viaDCOM -CimSession $script:CimSession -Hosts $env:COMPUTERNAME

    Test-WSMan -ComputerName $ComputerName -ErrorAction SilentlyContinue | Out-Null
    if ($?) {
        Write-Log "WinRM operacional e respondendo em '$ComputerName'." -Level "SUCCESS"
    }
    Write-Log "Sessão CIM DCOM estabelecida com sucesso." -Level "SUCCESS"
}
catch {
    Write-Log "Falha ao conectar via DCOM. Tentando WS-MAN em '$ComputerName'..." -Level "WARNING"
    try {
        $script:CimSession = New-CimSession -ComputerName $ComputerName -Credential $script:CredentialAdmin -ErrorAction Stop
        Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null
        Write-Log "Sessão CIM WS-MAN estabelecida com sucesso." -Level "SUCCESS"
    }
    catch {
        Write-Log "Falha definitiva ao conectar via DCOM e WS-MAN em '$ComputerName': $_" -Level "ERROR"
        exit 3
    }
}

#==================================================================================
# 6) REMOÇÃO REMOTA DE PERFIS DE USUÁRIO
#==================================================================================

$HKLM = [uint32]2147483650
$profileListKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

Write-Log "Iniciando processamento de remoção para os usuários: $($UsersToPurge -join ', ')"

foreach ($user in $UsersToPurge) {
    Write-Log "-------------------------------------------------------------"
    Write-Log "Processando conta de usuário: '$user'" -Level "INFO"

    # 6.1) Tenta remover a conta de usuário local, caso exista
    try {
        $delUserCmd = "net user `"$user`" /delete"
        $delUserRes = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $delUserCmd } -CimSession $script:CimSession -ErrorAction Stop
        if ($delUserRes.ReturnValue -eq 0) {
            Write-Log "Conta local '$user' removida com sucesso (net user)." -Level "SUCCESS"
        } else {
            Write-Log "Conta local '$user' não encontrada ou não removível via net user (código $($delUserRes.ReturnValue)). Prosseguindo para exclusão de perfis..." -Level "INFO"
        }
    }
    catch {
        Write-Log "Aviso ao tentar executar net user para '$user': $_" -Level "WARNING"
    }

    # 6.2) Busca perfis do usuário no WMI (Win32_UserProfile)
    try {
        $allProfiles = Get-CimInstance -ClassName Win32_UserProfile -CimSession $script:CimSession -ErrorAction Stop
        $targetProfiles = $allProfiles | Where-Object {
            $_.LocalPath -and ($_.LocalPath.TrimEnd('\') -like "*\$user")
        }

        if (-not $targetProfiles) {
            Write-Log "Nenhum perfil Win32_UserProfile encontrado correspondente a '*\$user'." -Level "WARNING"
            continue
        }

        foreach ($p in $targetProfiles) {
            $path = $p.LocalPath
            $sid  = $p.SID

            Write-Log "Perfil detectado: Path='$path', SID='$sid'"

            # 6.3) Verificação de segurança: perfil carregado (ativo)?
            if ($p.Loaded) {
                Write-Log "SEGURANÇA: O perfil '$path' (SID: $sid) está CARREGADO (em uso ativo). A remoção foi cancelada para proteger a integridade do usuário logado." -Level "WARNING"
                continue
            }

            # 6.4) Tenta exclusão nativa pelo método Delete() do Win32_UserProfile
            $deletedViaWmi = $false
            try {
                Write-Log "Invocando Win32_UserProfile.Delete() para '$path'..."
                $null = Invoke-CimMethod -InputObject $p -MethodName Delete -CimSession $script:CimSession -ErrorAction Stop
                Write-Log "Perfil '$path' excluído com sucesso via Win32_UserProfile.Delete()." -Level "SUCCESS"
                $deletedViaWmi = $true
            }
            catch {
                Write-Log "Falha ao excluir via Win32_UserProfile.Delete() ($path): $_. Tentando fallback manual de exclusão..." -Level "WARNING"
            }

            # 6.5) Fallback manual caso a exclusão WMI tenha falhado
            if (-not $deletedViaWmi) {
                # A) Remove a pasta física em C:\Users
                try {
                    Write-Log "Removendo pasta física '$path' via rmdir /S /Q..."
                    $rmdirCmd = "cmd.exe /c rmdir /S /Q `"$path`""
                    $rmdirRes = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $rmdirCmd } -CimSession $script:CimSession -ErrorAction Stop
                    if ($rmdirRes.ReturnValue -eq 0) {
                        Write-Log "Pasta física '$path' removida com sucesso." -Level "SUCCESS"
                    } else {
                        Write-Log "Falha ao remover a pasta física '$path' (código $($rmdirRes.ReturnValue))." -Level "WARNING"
                    }
                }
                catch {
                    Write-Log "Erro ao tentar remover pasta física '$path': $_" -Level "ERROR"
                }

                # B) Remove a chave no registro ProfileList via StdRegProv
                try {
                    Write-Log "Removendo chave de registro ProfileList\$sid via StdRegProv..."
                    $delRegKey = "$profileListKey\$sid"
                    $delRegRes = Invoke-CimMethod -Namespace root\default -ClassName StdRegProv -MethodName DeleteKey -Arguments @{ hDefKey = $HKLM; sSubKeyName = $delRegKey } -CimSession $script:CimSession -ErrorAction Stop
                    if ($delRegRes.ReturnValue -eq 0) {
                        Write-Log "Chave de registro '$delRegKey' removida com sucesso." -Level "SUCCESS"
                    } else {
                        Write-Log "Falha ao remover chave de registro '$delRegKey' (código $($delRegRes.ReturnValue))." -Level "WARNING"
                    }
                }
                catch {
                    Write-Log "Erro ao tentar remover chave de registro para SID '$sid': $_" -Level "ERROR"
                }
            }
        }
    }
    catch {
        Write-Log "Erro geral ao processar perfis para o usuário '$user': $_" -Level "ERROR"
    }
}

#==================================================================================
# 7) FINALIZAÇÃO E LIMPEZA DE RECURSOS
#==================================================================================

if ($script:CimSession) {
    Write-Log "Encerrando sessão CIM..."
    Remove-CimSession -CimSession $script:CimSession -ErrorAction SilentlyContinue
}

Write-Log "=== Processo de remoção de perfis finalizado para '$ComputerName' ===" -Level "SUCCESS"
