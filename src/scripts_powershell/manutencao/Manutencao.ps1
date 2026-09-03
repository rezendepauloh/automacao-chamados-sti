<#
.SYNOPSIS
  Script unificado para manutenção e limpeza de estação Windows local ou remota (via PowerShell/WinRM).

.DESCRIPTION
  Este script realiza uma série de tarefas de limpeza e manutenção em uma estação Windows, seja localmente ou remotamente.
  Se o parâmetro -ComputerName for fornecido, ele tentará estabelecer conexão remota (DCOM → WS-Man) para habilitar/configurar WinRM e, em seguida, criará uma PSSession (WinRM) para executar os comandos remotos.  
  Caso não seja especificado -ComputerName, todas as ações serão executadas localmente na máquina onde o script está rodando.

  Etapas principais (em ordem):
    1. Definição e carregamento de credenciais criptografadas (DPAPI) de um arquivo XML, se remoto.
    2. (Remoto) Teste de conectividade ICMP e configuração de WinRM via CimSession/DCOM (Enable-WinRM-viaDCOM, Open-Firewall-WinRM, Set-TrustedHosts-viaDCOM).
    3. Criação de PSSession (WinRM) para execução dos comandos remotos.
    4. Coleta do espaço livre em C: antes das operações (local ou remoto).
    5. Execução das rotinas de limpeza:
       a. Limpeza de arquivos temporários (Temp do usuário e sistema).
       b. Limpeza da pasta Prefetch.
       c. Limpeza da Lixeira.
       d. Limpeza de logs de eventos (Event Logs).
       e. Limpeza do cache do Windows Update (SoftwareDistribution\Download).
       f. Limpeza de caches adicionais (ex.: DNS).
    6. Execução das rotinas de manutenção:
       a. Repair-Volume (se Windows 10+), ou chkdsk em versões anteriores.
       b. Otimização/Defrag do volume C: (Optimize-Volume ou defrag).
       c. SFC /scannow.
    7. Coleta do espaço livre em C: após as operações.
    8. Exibição de relatório de “Free Space Before” e “Free Space After”.
    9. (Remoto) Fechamento da PSSession.

.PARAMETER ComputerName
  (Opcional) Nome ou IP da estação Windows para manutenção remota. Ex.: “PGJ-58099”.
  Se fornecido, o script assume operação remota via WinRM/DCOM. Se omitido, roda localmente.

.PARAMETER CleanWindowsOld
  (Opcional) Switch para remover instalações e atualizações anteriores do Windows (C:\Windows.old).

.PARAMETER CleanDeliveryOptimization
  (Opcional) Switch para limpar o cache do Delivery Optimization.

.PARAMETER CleanCrashDumps
  (Opcional) Switch para limpar relatórios de erros do Windows (WER) e dumps de memória.

.EXAMPLE
  # 1) Limpeza e manutenção local padrão:
  PS> .\MaintenanceAndCleanup.ps1

  # 2) Limpeza e manutenção remota padrão (credenciais pré-criadas):
  PS> .\MaintenanceAndCleanup.ps1 -ComputerName "MPE-58063" -Verbose

  # 3) Limpeza remota profunda (incluindo Windows.old, Delivery Optimization e Crash Dumps):
  PS> .\MaintenanceAndCleanup.ps1 -ComputerName "MPE-58063" -CleanWindowsOld -CleanDeliveryOptimization -CleanCrashDumps -Verbose

.NOTES
  - O caminho de $CredXmlPath está embutido no script. Ajuste conforme local do seu XML de credenciais.
  - Para criar o arquivo de credenciais (DPAPI):
      PS> $cred = Get-Credential
      PS> $cred | Export-Clixml -Path "C:\Scripts\cred_admin.xml"
    Esse XML ficará criptografado para o seu usuário atual do Windows, e só poderá ser lido por ele.
  - O usuário que rodar o script remoto precisa ser Administrador local na máquina-alvo.
  - Testado em Windows 10/Server 2016 e superiores.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ComputerName,

    [Parameter(Mandatory = $false)]
    [switch]$CleanWindowsOld,

    [Parameter(Mandatory = $false)]
    [switch]$CleanDeliveryOptimization,

    [Parameter(Mandatory = $false)]
    [switch]$CleanCrashDumps
)

#==================================================================================
# 0) CONFIGURAÇÃO DE LOGGING (TRANSCRIPT NATIVO)
#==================================================================================

$LogPath = Join-Path -Path $PSScriptRoot -ChildPath "MaintenanceAndCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
try { Start-Transcript -Path $LogPath -Append -Force -ErrorAction SilentlyContinue | Out-Null } catch {}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "INFO"    { Write-Host $logLine -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $logLine -ForegroundColor Green }
        "WARNING" { Write-Warning $Message }
        "ERROR"   { Write-Error $Message }
    }
}

Write-Log "Arquivo de log inicializado em: $LogPath" -Level "SUCCESS"

#==================================================================================
# 1) CAMINHO DO ARQUIVO XML QUE CONTÉM AS CREDENCIAIS CRIPTOGRAFADAS VIA DPAPI (REMOTO)
#==================================================================================

$CredXmlPath = Join-Path $PSScriptRoot "cred_admin.xml"

# Se vamos operar remotamente, precisamos importar as credenciais
if ($ComputerName -and $ComputerName -ne ".") {
    if (-not (Test-Path $CredXmlPath)) {
        Write-Log "Arquivo de credencial não encontrado em '$CredXmlPath'. Abortando." -Level "ERROR"
        exit 1
    }
    try {
        $Global:CredentialAdmin = Import-Clixml -Path $CredXmlPath
        Write-Log "Credenciais importadas com sucesso de '$CredXmlPath'." -Level "SUCCESS"
    }
    catch {
        Write-Log "Falha ao importar credenciais de '$CredXmlPath': $_" -Level "ERROR"
        exit 1
    }
}

#==================================================================================
# 2) FUNÇÕES AUXILIARES PARA CONFIGURAR WinRM VIA DCOM (USADAS APENAS EM MODO REMOTO)
#==================================================================================

# 2.1) Verifica se está online via ping
if ($ComputerName -and $ComputerName -ne "." -and -not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
    Write-Log "Máquina '$ComputerName' está inacessível via ICMP. Abortando." -Level "ERROR"
    exit 2
}

function Enable-WinRM-viaDCOM {
    <#
    .SYNOPSIS
      Habilita/configura o serviço WinRM no host remoto usando CimSession via DCOM.
    .PARAMETER CimSession
      Sessão CIM já aberta (via DCOM) para o host remoto.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [CimSession]$CimSession
    )

    try {
        $svc = Get-CimInstance -ClassName Win32_Service `
              -Filter "Name='WinRM'" `
              -CimSession $CimSession `
              -ErrorAction Stop

        if ($svc.StartMode -ne "Auto") {
            Invoke-CimMethod -InputObject $svc `
                             -MethodName ChangeStartMode `
                             -Arguments @{ StartMode = "Automatic" } `
                             -CimSession $CimSession `
                             -ErrorAction Stop | Out-Null
            Write-Log "WinRM StartMode ajustado para 'Automatic'." -Level "SUCCESS"
        }

        if ($svc.State -ne "Running") {
            Invoke-CimMethod -InputObject $svc `
                             -MethodName StartService `
                             -CimSession $CimSession `
                             -ErrorAction Stop | Out-Null
            Write-Log "Serviço WinRM iniciado no host remoto." -Level "SUCCESS"
        }
    }
    catch {
        Write-Log "Falha ao habilitar/configurar WinRM via DCOM: $_" -Level "WARNING"
    }
}

function Open-Firewall-WinRM {
    <#
    .SYNOPSIS
      Abre exceção de firewall para o grupo “Windows Remote Management” no host remoto via DCOM.
    .PARAMETER CimSession
      Sessão CIM já aberta (via DCOM) para o host remoto.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [CimSession]$CimSession
    )

    $cmd = 'netsh advfirewall firewall set rule group="Windows Remote Management" new enable=yes'

    try {
        $proc = Invoke-CimMethod -ClassName Win32_Process `
                                 -MethodName Create `
                                 -Arguments @{ CommandLine = $cmd } `
                                 -CimSession $CimSession `
                                 -ErrorAction Stop

        if ($proc.ReturnValue -eq 0) {
            Write-Log "Exceção de firewall para WinRM habilitada via DCOM." -Level "SUCCESS"
        }
        else {
            Write-Log "Falha ao habilitar regra de firewall (código $($proc.ReturnValue))." -Level "WARNING"
        }
    }
    catch {
        Write-Log "Erro ao executar netsh via DCOM: $_" -Level "WARNING"
    }
}

function Set-TrustedHosts-viaDCOM {
    <#
    .SYNOPSIS
      Ajusta TrustedHosts no registro remoto (WinRM\Client) usando StdRegProv via DCOM.
    .PARAMETER CimSession
      Sessão CIM já aberta (via DCOM) para o host remoto.
    .PARAMETER Hosts
      String com hosts permitidos (por exemplo, "*" ou "PC1,PC2").
    #>
    param(
        [Parameter(Mandatory = $true)]
        [CimSession]$CimSession,

        [Parameter(Mandatory = $true)]
        [string]$Hosts
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
        Invoke-CimMethod -Namespace root\default `
                         -ClassName StdRegProv `
                         -MethodName SetStringValue `
                         -Arguments $argsReg `
                         -CimSession $CimSession `
                         -ErrorAction Stop | Out-Null

        Write-Log "TrustedHosts em HKLM:\$subKey\$valueName definido para '$Hosts'." -Level "SUCCESS"
    }
    catch {
        Write-Log "Falha ao ajustar TrustedHosts via StdRegProv: $_" -Level "WARNING"
    }
}

#==================================================================================
# 3) FUNÇÃO PARA CRIAR PSSession REMOTO (APÓS CONFIGURAR WinRM)
#==================================================================================

function New-RemotePSSession {
    <#
    .SYNOPSIS
      Tenta criar uma PSSession remota via WinRM. Retorna o objeto PSSession ou $null.
    .PARAMETER ComputerName
      Nome ou IP da máquina remota.
    .PARAMETER Credential
      Objeto PSCredential para autenticação.
    .OUTPUTS
      Retorna PSSession ou $null em caso de falha.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [PSCredential]$Credential
    )

    try {
        Write-Log "Tentando criar PSSession via WinRM em '$ComputerName'..." -Level "INFO"
        $ps = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
        Write-Log "PSSession (WinRM) estabelecida com sucesso." -Level "SUCCESS"
        return $ps
    }
    catch {
        Write-Log "Falha ao criar PSSession via WinRM em '$ComputerName': $_" -Level "ERROR"
        return $null
    }
}

#==================================================================================
# 4) FUNÇÕES DE LIMPEZA (LOCAL OU REMOTO)
#==================================================================================

function Clear-TempFiles {
    <#
    .SYNOPSIS
      Limpa os arquivos de TEMP de TODOS os perfis em C:\Users\<usuário>\AppData\Local\Temp
      e também C:\Windows\Temp. Funciona local ou via PSSession.
    .PARAMETER Session
      PSSession para execução remota. Se $null, executa local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando arquivos temporários de TODOS os perfis em C:\Users e C:\Windows\Temp..."

        # 1) Varre todas as pastas em C:\Users
        Get-ChildItem -Path "C:\Users" -Directory |
            ForEach-Object {
                $userFolder = $_.FullName
                $tempPath   = Join-Path $userFolder "AppData\Local\Temp"
                if (Test-Path $tempPath) {
                    try {
                        Remove-Item -Path "$tempPath\*" -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Host "   • Limpo: $tempPath"
                    }
                    catch {
                        Write-Warning "     Falha ao limpar $($tempPath): $_"
                    }
                }
            }

        # 2) Limpa C:\Windows\Temp
        $systemTemp = "C:\Windows\Temp"
        if (Test-Path $systemTemp) {
            try {
                Remove-Item -Path "$systemTemp\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "   • Limpo: $systemTemp"
            }
            catch {
                Write-Warning "     Falha ao limpar $($systemTemp): $_"
            }
        }

        Write-Host "   • Todas as pastas Temp limpas."
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-Prefetch {
    <#
    .SYNOPSIS
      Limpa a pasta C:\Windows\Prefetch.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando Prefetch..."
        try { Remove-Item -Path "C:\Windows\Prefetch\*" -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        Write-Host "   • Prefetch limpo."
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-RecycleBin {
    <#
    .SYNOPSIS
      Esvazia a Lixeira.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Esvaziando Lixeira..."
        try {
            $shell = New-Object -ComObject Shell.Application
            $bin   = $shell.Namespace(0x0a)
            $items = $bin.Items()
            if ($items.Count -gt 0) {
                $bin.InvokeVerb("Esvaziar Lixeira")
                Write-Host "   • Lixeira esvaziada."
            }
            else {
                Write-Host "   • Lixeira já está vazia."
            }
        }
        catch {
            Write-Warning "Falha ao tentar esvaziar Lixeira: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-EventLogs {
    <#
    .SYNOPSIS
      Limpa apenas logs dos tipos “Admin” e “Operational”, evitando erros de acesso negado.
    .PARAMETER Session
      PSSession para execução remota. Se $null, executa local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando Event Logs (somente Admin/Operational)..."

        try {
            # Tentar obter a lista de logs; se falhar em algum, pulamos silenciosamente
            $allLogs = @()
            try {
                $allLogs = Get-WinEvent -ListLog * -ErrorAction Stop
            }
            catch {
                # Se houve erro ao listar, tentamos listar canal por canal para não abandonar todo o processo
                Write-Verbose "   • Erro genérico ao listar logs, tentando por canal individualmente..."
                $names = (wevtutil el 2>$null) -split "`r?`n"
                foreach ($name in $names) {
                    if ($name.Trim()) {
                        try {
                            $logInfo = Get-WinEvent -ListLog $name -ErrorAction Stop
                            $allLogs += $logInfo
                        }
                        catch {
                            # Pula este canal
                            Write-Verbose "     • Não foi possível listar o canal '$name'. Pulando."
                        }
                    }
                }
            }

            # Filtrar apenas Admin ou Operational
            $logsToClear = $allLogs |
                Where-Object { $_.LogType -in @("Admin", "Operational") } |
                Select-Object -ExpandProperty LogName

            foreach ($logName in $logsToClear) {
                try {
                    wevtutil cl $logName 2>$null
                    Write-Host "   • Limpo: $logName"
                }
                catch {
                    Write-Verbose "     • Falha ao limpar '$logName'. Pulando."
                }
            }

            Write-Host "   • Event Logs (Admin/Operational) limpos."
        }
        catch {
            Write-Warning "   • Falha geral ao listar/limpar Event Logs: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-WindowsUpdateCache {
    <#
    .SYNOPSIS
      Limpa o cache de Windows Update em C:\Windows\SoftwareDistribution\Download.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando cache do Windows Update..."
        try {
            Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
            Start-Service -Name wuauserv -ErrorAction SilentlyContinue
            Write-Host "   • Cache do Windows Update limpo."
        }
        catch {
            Write-Warning "Falha ao limpar cache do Windows Update: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-AdditionalCaches {
    <#
    .SYNOPSIS
      Exemplo de outras limpezas de cache (opcional).
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando caches adicionais (DNS, etc.)..."
        try {
            ipconfig /flushdns | Out-Null
            Write-Host "   • Caches adicionais limpos."
        }
        catch {
            Write-Warning "Falha ao limpar caches adicionais: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-WinSxS {
    <#
    .SYNOPSIS
      Executa DISM /Online /Cleanup-Image /StartComponentCleanup para limpar componentes obsoletos do WinSxS.
    .PARAMETER Session
      PSSession para execução remota. Se $null, roda local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando WinSxS (Component Store) com DISM..."
        try {
            # Remove componentes obsoletos e libera espaço
            dism.exe /Online /Cleanup-Image /StartComponentCleanup | Out-Null
            Write-Host "   • WinSxS limpo (StartComponentCleanup concluído)."
        }
        catch {
            Write-Warning "   • Falha ao executar DISM /StartComponentCleanup: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-ThumbnailCache {
    <#
    .SYNOPSIS
      Remove arquivos de cache de miniaturas (thumbcache_*) de todos os usuários e do sistema.
    .PARAMETER Session
      PSSession para execução remota. Se $null, roda local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando cache de miniaturas (thumbcache) de todos os perfis..."

        # 1) Cache global (Windows Explorer)
        $systemThumbCache = "C:\Users\Public\Libraries\*Thumb*.*"
        try { if (Test-Path (Split-Path $systemThumbCache)) { Remove-Item -Path $systemThumbCache -Force -ErrorAction SilentlyContinue } } catch {}
        # (Em geral, existe também em C:\Users\<user>\AppData\Local\Microsoft\Windows\Explorer\thumbcache_*.db)

        foreach ($user in Get-ChildItem -Path "C:\Users" -Directory) {
            $thumbFolder = Join-Path $user.FullName "AppData\Local\Microsoft\Windows\Explorer"
            if (Test-Path $thumbFolder) {
                Get-ChildItem -Path "$thumbFolder\thumbcache_*" -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
                        Write-Host "   • Removido: $($_.FullName)"
                    }
                    catch {
                        Write-Warning "     Falha ao remover $($_.FullName): $_"
                    }
                }
            }
        }

        Write-Host "   • Cache de miniaturas limpo."
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-StoreCache {
    <#
    .SYNOPSIS
      Executa wsreset.exe para limpar o cache da Microsoft Store.
    .PARAMETER Session
      PSSession para execução remota. Se $null, roda local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando cache da Microsoft Store..."
        try {
            # Executa wsreset para resetar o cache da Store
            Start-Process -FilePath "wsreset.exe" -ArgumentList "-quiet" -NoNewWindow -Wait
            Write-Host "   • Cache da Store limpo."
        }
        catch {
            Write-Warning "   • Falha ao executar wsreset: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-CBSLogs {
    <#
    .SYNOPSIS
      Trunca arquivos de log em C:\Windows\Logs\CBS e subpastas, pulando os que estiverem em uso.
    .PARAMETER Session
      PSSession para execução remota. Se $null, roda local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando logs do CBS (C:\Windows\Logs\CBS)..."
        $cbsFolder = "C:\Windows\Logs\CBS"

        if (-not (Test-Path $cbsFolder)) {
            Write-Host "   • Pasta CBS não existe. Nada a fazer."
            return
        }

        # Pegar todos os arquivos dentro de CBS
        $files = Get-ChildItem -Path "$cbsFolder\*" -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                # Se o arquivo estiver em uso, Get-Content pode falhar; capturamos e pulamos
                Set-Content -Path $file.FullName -Value $null -ErrorAction Stop
                Write-Host "   • Truncado: $($file.Name)"
            }
            catch {
                Write-Verbose "     • Não foi possível truncar '$($file.Name)' (talvez em uso). Pulando."
            }
        }

        Write-Host "   • Logs do CBS limpos (arquivos em uso foram ignorados)."
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-WindowsOld {
    <#
    .SYNOPSIS
      Remove diretórios de instalações e atualizações anteriores do Windows.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando instalações anteriores do Windows (C:\Windows.old)..."
        $paths = @("C:\Windows.old", "C:\`$Windows.~BT", "C:\`$Windows.~WS")
        foreach ($p in $paths) {
            if (Test-Path $p) {
                try {
                    Remove-Item -Path "$p\*" -Recurse -Force -ErrorAction SilentlyContinue
                    cmd.exe /c "rd /s /q `"$p`"" 2>$null
                    Write-Host "   • Removido: $p"
                } catch {
                    Write-Verbose "     • Falha ao remover $($p): $_"
                }
            }
        }
        Write-Host "   • Instalações anteriores limpas."
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-DeliveryOptimization {
    <#
    .SYNOPSIS
      Limpa o cache de Otimização de Entrega do Windows.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando cache de Otimização de Entrega (Delivery Optimization)..."
        try {
            if (Get-Command Remove-DeliveryOptimizationCache -ErrorAction SilentlyContinue) {
                Remove-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue | Out-Null
            }
            $doPath = "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache\*"
            if (Test-Path "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache") {
                Remove-Item -Path $doPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-Host "   • Cache de Otimização de Entrega limpo."
        } catch {
            Write-Verbose "     • Falha ao limpar Delivery Optimization: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Clear-CrashDumps {
    <#
    .SYNOPSIS
      Limpa filas e relatórios de erro do WER e Dumps de memória.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Limpando relatórios de erros (WER) e Crash Dumps..."
        $werPaths = @(
            "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*",
            "C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*",
            "C:\Windows\Minidump\*",
            "C:\Windows\MEMORY.DMP"
        )

        foreach ($p in $werPaths) {
            try {
                if (Test-Path (Split-Path $p)) {
                    Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }

        # Limpa CrashDumps de todos os usuários
        foreach ($user in Get-ChildItem -Path "C:\Users" -Directory) {
            $cdPath = Join-Path $user.FullName "AppData\Local\CrashDumps\*"
            if (Test-Path (Split-Path $cdPath)) {
                try { Remove-Item -Path $cdPath -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            }
        }

        Write-Host "   • Relatórios de erros e Dumps limpos."
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

#==================================================================================
# 5) FUNÇÕES DE MANUTENÇÃO (REPAIR, DEFRAG, SFC)
#==================================================================================

function Start-Repair-Device {
    <#
    .SYNOPSIS
      Executa Repair-Volume (Windows 10+/Server 2016+) ou chkdsk em versões anteriores.
    .PARAMETER Session
      PSSession para execução remota. Se $null, roda local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Reparando volume C: (Repair-Volume ou chkdsk)..."
        try {
            $osversion = (Get-CimInstance Win32_OperatingSystem).Version
            $major = [int]($osversion.Split('.')[0])
            if ($major -ge 10) {
                Repair-Volume -DriveLetter C -OfflineScanAndFix -ErrorAction SilentlyContinue
                Write-Host "   • Repair-Volume (Windows 10+) executado."
            }
            else {
                chkdsk C: /F /X | Out-Null
                Write-Host "   • chkdsk C: /F executado (versão antiga do Windows)."
            }
        }
        catch {
            Write-Warning "Falha ao executar Repair/Chkdsk: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Start-Defrag-Device {
    <#
    .SYNOPSIS
      Realiza otimização/desfragmentação do volume C: somente se houver HDD no sistema.
      Se todos os discos forem SSD, pula a etapa.
    .PARAMETER Session
      PSSession para execução remota. Se $null, executa local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Verificando tipo de disco para decidir sobre desfragmentação/otimização de C:..."

        try {
            # 1) Coleta todos os discos físicos e seus MediaTypes
            $physicalDisks = Get-PhysicalDisk | Select-Object DeviceId, MediaType

            # Se não conseguir identificar discos (por versão do SO), assume HDD
            if (-not $physicalDisks) {
                Write-Host "   • Não foi possível ler Get-PhysicalDisk. Assumindo HDD para desfragmentação."
                $needsDefrag = $true
            }
            else {
                # Checa se existe pelo menos um HDD
                $hasHDD = $physicalDisks | Where-Object { $_.MediaType -eq 'HDD' }
                if ($hasHDD) {
                    $needsDefrag = $true
                } else {
                    $needsDefrag = $false
                }
            }

            if (-not $needsDefrag) {
                Write-Host "   • Todos os discos são SSD. Pulando desfragmentação (não recomendado em SSD)."
                return
            }

            # 2) Se precisa desfragmentar/otimizar, detecta versão do Windows
            $osVersion = (Get-CimInstance Win32_OperatingSystem).Version
            $major = [int]($osVersion.Split('.')[0])

            if ($major -ge 10) {
                Write-Host "   • HDD detectado. Executando Optimize-Volume -DriveLetter C -Defrag..."
                Optimize-Volume -DriveLetter C -Defrag -ErrorAction SilentlyContinue
                Write-Host "   • Optimize-Volume (Defrag) concluído."
            }
            else {
                Write-Host "   • HDD detectado. Executando defrag C: /T /U /V..."
                defrag C: /T /U /V | Out-Null
                Write-Host "   • defrag.exe concluído."
            }
        }
        catch {
            Write-Warning "   • Falha ao executar Defrag/Optimize: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

function Start-Sfc-Device {
    <#
    .SYNOPSIS
      Verifica se há corrupção na imagem do sistema via DISM CheckHealth e executa SFC /scannow apenas se necessário.
    .PARAMETER Session
      PSSession para execução remota. Se $null, roda local.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    $scriptBlock = {
        Write-Host "→ Verificando saúde da imagem do sistema com DISM CheckHealth..."
        try {
            $health = dism.exe /Online /Cleanup-Image /CheckHealth 2>&1
            $output = $health | Out-String

            if ($output -match "No component store corruption detected|Nenhuma corrupção") {
                Write-Host "   • Imagem saudável (nenhuma corrupção detectada). Pulando SFC /scannow."
            }
            else {
                Write-Host "→ Alerta de corrupção detectado. Executando SFC /scannow (aguarde a conclusão)..." -ForegroundColor Yellow
                sfc /scannow
                Write-Host "   • SFC /scannow concluído."
            }
        }
        catch {
            Write-Warning "Falha ao verificar/executar SFC /scannow: $_"
        }
    }

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $scriptBlock
    } else {
        & $scriptBlock
    }
}

#==================================================================================
# 6) FUNÇÃO PARA OBTER FREE SPACE EM C:
#==================================================================================

function Get-FreeSpaceGB {
    <#
    .SYNOPSIS
      Retorna o espaço livre em GB na unidade C:. Se PSSession for fornecida, faz remotamente.
    .PARAMETER Session
      PSSession para execução remota. Se $null, faz local.
    .OUTPUTS
      [double] – Espaço livre em GB.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [object]$Session
    )

    if ($Session) {
        $obj = Invoke-Command -Session $Session -ScriptBlock {
            Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
        }
    }
    else {
        $obj = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
    }

    if ($obj -and $obj.FreeSpace) {
        return [math]::Round($obj.FreeSpace / 1GB, 2)
    }
    else {
        return 0
    }
}

#==================================================================================
# 7) FUNÇÃO PRINCIPAL: INVOCAR LIMPEZA E MANUTENÇÃO
#==================================================================================

function Invoke-Maintenance {
    <#
    .SYNOPSIS
      Roteia execução para modo local ou remoto (se ComputerName estiver definido).
    .PARAMETER ComputerName
      Nome ou IP da máquina remota (se fornecido, rodarão operações remotamente).
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$ComputerName,

        [Parameter(Mandatory = $false)]
        [switch]$CleanWindowsOld,

        [Parameter(Mandatory = $false)]
        [switch]$CleanDeliveryOptimization,

        [Parameter(Mandatory = $false)]
        [switch]$CleanCrashDumps
    )

    # 7.1) Se ComputerName NÃO for especificado → execução local
    if (-not $ComputerName -or $ComputerName -eq ".") {
        Write-Log "=== Executando manutenção LOCALMENTE em $env:COMPUTERNAME ===" -Level "INFO"

        $fsBefore = Get-FreeSpaceGB
        Write-Log "Espaço livre antes         : $fsBefore GB" -Level "INFO"

        # Rotinas de limpeza local
        Clear-TempFiles
        Clear-Prefetch
        Clear-RecycleBin
        Clear-EventLogs
        Clear-WindowsUpdateCache
        Clear-AdditionalCaches
        Clear-WinSxS
        Clear-ThumbnailCache
        Clear-StoreCache
        Clear-CBSLogs
        if ($CleanWindowsOld) { Clear-WindowsOld }
        if ($CleanDeliveryOptimization) { Clear-DeliveryOptimization }
        if ($CleanCrashDumps) { Clear-CrashDumps }

        # Rotinas de manutenção local
        Start-Repair-Device
        Start-Defrag-Device
        Start-Sfc-Device

        $fsAfter = Get-FreeSpaceGB
        Write-Log "Espaço livre depois        : $fsAfter GB" -Level "INFO"
        $freed = [math]::Round($fsAfter - $fsBefore, 2)
        Write-Log "Espaço recuperado          : $freed GB" -Level "SUCCESS"

        Write-Log "→ Manutenção local CONCLUÍDA com êxito." -Level "SUCCESS"
        return
    }

    # 7.2) Execução remota (ComputerName foi fornecido)
    Write-Log "=== Iniciando manutenção REMOTA em '$ComputerName' ===" -Level "INFO"

    # 7.2.1) Testa conectividade ICMP
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
        Write-Log "Máquina '$ComputerName' inacessível via ICMP. Abortando." -Level "ERROR"
        return
    }

    # 7.2.2) Cria CimSession para configuração WinRM via DCOM
    Write-Log "Tentando criar CimSession via DCOM para configurar WinRM..." -Level "INFO"
    try {
        $optDcom = New-CimSessionOption -Protocol Dcom
        $cim = New-CimSession -ComputerName $ComputerName -Credential $Global:CredentialAdmin -SessionOption $optDcom -ErrorAction Stop
        Write-Log "CimSession DCOM estabelecida com sucesso." -Level "SUCCESS"

        # Configura WinRM no host remoto
        Enable-WinRM-viaDCOM    -CimSession $cim
        Open-Firewall-WinRM     -CimSession $cim
        Set-TrustedHosts-viaDCOM -CimSession $cim -Hosts $env:COMPUTERNAME
    }
    catch {
        Write-Log "Falha ao configurar WinRM via CimSession/DCOM: $_" -Level "WARNING"
        # Continua, pois o WinRM pode já estar funcional
    }
    finally {
        if ($cim) { Remove-CimSession -CimSession $cim -ErrorAction SilentlyContinue }
    }

    # 7.2.3) Cria PSSession (WinRM) para operações remotas
    $psSession = New-RemotePSSession -ComputerName $ComputerName -Credential $Global:CredentialAdmin
    if (-not $psSession) {
        Write-Log "Não foi possível estabelecer PSSession (WinRM) em '$ComputerName'. Abortando." -Level "ERROR"
        return
    }

    # 7.2.4) Obter espaço livre antes
    $fsBefore = Get-FreeSpaceGB -Session $psSession
    Write-Log "Espaço livre antes (remoto)   : $fsBefore GB" -Level "INFO"

    # 7.2.5) Rotinas de limpeza remota
    Clear-TempFiles            -Session $psSession
    Clear-Prefetch             -Session $psSession
    Clear-RecycleBin           -Session $psSession
    Clear-EventLogs            -Session $psSession
    Clear-WindowsUpdateCache   -Session $psSession
    Clear-AdditionalCaches     -Session $psSession
    Clear-WinSxS               -Session $psSession
    Clear-ThumbnailCache       -Session $psSession
    Clear-StoreCache           -Session $psSession
    Clear-CBSLogs              -Session $psSession
    if ($CleanWindowsOld) { Clear-WindowsOld -Session $psSession }
    if ($CleanDeliveryOptimization) { Clear-DeliveryOptimization -Session $psSession }
    if ($CleanCrashDumps) { Clear-CrashDumps -Session $psSession }

    # 7.2.6) Rotinas de manutenção remota
    Start-Repair-Device        -Session $psSession
    Start-Defrag-Device        -Session $psSession
    Start-Sfc-Device           -Session $psSession

    # 7.2.7) Obter espaço livre depois
    $fsAfter = Get-FreeSpaceGB -Session $psSession
    Write-Log "Espaço livre depois (remoto)  : $fsAfter GB" -Level "INFO"
    $freedRem = [math]::Round($fsAfter - $fsBefore, 2)
    Write-Log "Espaço recuperado (remoto)    : $freedRem GB" -Level "SUCCESS"

    Write-Log "→ Manutenção remota em '$ComputerName' CONCLUÍDA com êxito." -Level "SUCCESS"

    # 7.2.8) Fechar PSSession
    Remove-PSSession -Session $psSession
    Write-Log "PSSession fechada com sucesso." -Level "SUCCESS"
}

#==================================================================================
# 8) EXECUTAR INVOKE-MAINTENANCE COM O PARÂMETRO ComputerName (SE HOUVER)
#==================================================================================
Invoke-Maintenance -ComputerName $ComputerName -CleanWindowsOld:$CleanWindowsOld -CleanDeliveryOptimization:$CleanDeliveryOptimization -CleanCrashDumps:$CleanCrashDumps
try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch {}
