# ==============================================================================
# Script: bancada-launcher.ps1
# Função: Executor local do Protocol Handler 'bancada://' para estações Windows.
# ==============================================================================

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$UriString
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Remove o prefixo do protocolo e normaliza a query string
$cleanUri = $UriString -replace "^bancada:/*", ""
$cleanUri = $cleanUri -replace "^run/*\?", ""
if ($cleanUri -match "\?(.*)$") {
    $cleanUri = $matches[1]
}

# Parse dos parâmetros da query string
$params = @{}
$cleanUri.Split('&') | ForEach-Object {
    if ($_ -match "^(?<key>[^=]+)=(?<val>.*)$") {
        $key = [System.Uri]::UnescapeDataString($matches['key'])
        $val = [System.Uri]::UnescapeDataString($matches['val'])
        $params[$key] = $val
    }
}

$tool       = $params['tool']
$targetHost = $params['host']
$serverUrl  = $params['server']
$skipMajor  = $params['skip_major'] -eq 'true'
$timeoutSec = if ($params['timeout']) { [int]$params['timeout'] } else { 30 }
$usersPurge = $params['users']
$psEngine   = $params['ps_engine']

# Detecta e escolhe o executável do PowerShell desejado
$pwsh7Path = "C:\Program Files\PowerShell\7\pwsh.exe"
$hasPwsh7 = Test-Path $pwsh7Path
if (-not $hasPwsh7) {
    $cmdTest = Get-Command "pwsh.exe" -ErrorAction SilentlyContinue
    if ($cmdTest) {
        $pwsh7Path = $cmdTest.Source
        $hasPwsh7 = $true
    }
}

$desiredEngine = "powershell.exe"
if ($psEngine -eq "pwsh") {
    if ($hasPwsh7) { $desiredEngine = $pwsh7Path }
} elseif ($psEngine -eq "powershell") {
    $desiredEngine = "powershell.exe"
} else {
    # Detectar Automaticamente (Padrão): usa pwsh se existir, senão powershell 5.1
    if ($hasPwsh7) {
        $desiredEngine = $pwsh7Path
    } else {
        $desiredEngine = "powershell.exe"
    }
}

# Se estamos rodando no Windows PowerShell 5.1 e o usuário/auto pediu pwsh.exe, repassa a execução para o pwsh.exe em nova janela visível e fecha a atual silenciosamente
if ($PSVersionTable.PSVersion.Major -lt 7 -and $desiredEngine -ne "powershell.exe" -and $hasPwsh7) {
    # Minimiza / esconde a janela atual do PowerShell 5.1
    $null = Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    ' -ErrorAction SilentlyContinue
    $consolePtr = [Console.Window]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) {
        [Console.Window]::ShowWindow($consolePtr, 0) # 0 = SW_HIDE
    }

    Start-Process -FilePath $desiredEngine -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath, "`"$UriString`"")
    exit 0
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "       SISTEMA BANCADA — DISPARADOR LOCAL WINDOWS" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$currentEngineName = if ($PSVersionTable.PSVersion.Major -ge 7) { "⚡ PowerShell 7+ ($($PSVersionTable.PSVersion))" } else { "💻 Windows PowerShell 5.1" }
Write-Host " [INFO] Interpretador em uso  : $currentEngineName" -ForegroundColor Cyan
Write-Host " [INFO] Ferramenta solicitada : $tool" -ForegroundColor Green
Write-Host " [INFO] Máquina alvo          : $targetHost" -ForegroundColor Green
if ($serverUrl) {
    Write-Host " [INFO] Servidor de Origem    : $serverUrl" -ForegroundColor Gray
}

if (-not $tool -or -not $targetHost) {
    Write-Error "Parâmetros insuficientes na chamada: tool e host são obrigatórios."
    Write-Host "Pressione ENTER para fechar..."
    Read-Host
    exit 1
}

# Cria diretório temporário isolado para a execução
$tempFolder = Join-Path $env:TEMP ("bancada_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
Write-Host " [INFO] Pasta temporária criada : $tempFolder" -ForegroundColor Gray

try {
    # Lista de arquivos a obter
    $scriptFiles = @()
    if ($tool -eq "analisador") {
        $scriptFiles = @("Analisador.ps1", "GeradorHtml.ps1", "Mapeamentos.ps1", "cred_admin.xml")
        $mainScript = Join-Path $tempFolder "Analisador.ps1"
    } elseif ($tool -eq "manutencao") {
        $scriptFiles = @("Manutencao.ps1", "cred_admin.xml")
        $mainScript = Join-Path $tempFolder "Manutencao.ps1"
    } elseif ($tool -eq "perfis") {
        $scriptFiles = @("RemoverUsuarios.ps1", "cred_admin.xml")
        $mainScript = Join-Path $tempFolder "RemoverUsuarios.ps1"
    } else {
        throw "Ferramenta desconhecida: '$tool'"
    }

    # Baixa ou copia os arquivos do script
    Write-Host " [INFO] Obtendo arquivos necessarios para '$tool'..." -ForegroundColor Yellow
    foreach ($file in $scriptFiles) {
        $destPath = Join-Path $tempFolder $file
        $copied = $false

        # 1) Tenta caminhos do WSL / rede local
        $candidates = @(
            "\\wsl.localhost\Ubuntu-26.04\home\paulogoncalves\PythonProjects\automated-OTRS-and-CitSmart\src\scripts_powershell\$tool\$file",
            "\\wsl$\Ubuntu-26.04\home\paulogoncalves\PythonProjects\automated-OTRS-and-CitSmart\src\scripts_powershell\$tool\$file",
            "\\wsl.localhost\Ubuntu\home\paulogoncalves\PythonProjects\automated-OTRS-and-CitSmart\src\scripts_powershell\$tool\$file",
            "\\wsl$\Ubuntu\home\paulogoncalves\PythonProjects\automated-OTRS-and-CitSmart\src\scripts_powershell\$tool\$file"
        )
        foreach ($cand in $candidates) {
            if (Test-Path $cand) {
                Copy-Item $cand $destPath -Force
                Write-Host "  -> [WSL OK] $file obtido de $cand" -ForegroundColor Gray
                $copied = $true
                break
            }
        }

        # 2) Se nao achou no WSL, tenta download via HTTP se serverUrl estiver definido
        if (-not $copied -and $serverUrl) {
            $downloadUrl = "$serverUrl/api/scripts/$tool/$file"
            try {
                Invoke-RestMethod -Uri $downloadUrl -OutFile $destPath -TimeoutSec 15
                Write-Host "  -> [HTTP OK] $file baixado de $downloadUrl" -ForegroundColor Gray
                $copied = $true
            } catch {
                Write-Warning "  -> [FALHA] Nao foi possivel baixar $file via $downloadUrl"
            }
        }

        if (-not (Test-Path $destPath)) {
            Write-Warning "  -> [AVISO] Arquivo $file nao foi localizado."
        }
    }

    if (-not (Test-Path $mainScript)) {
        throw "Não foi possível obter o script principal '$mainScript'."
    }

    Write-Host " [INFO] Disparando execução local no Windows..." -ForegroundColor Yellow

    # Monta os argumentos
    $outDir = Join-Path $env:USERPROFILE "DeviceReports"
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    # Monta os argumentos via hashtable (splatting) e executa de forma segura
    if ($tool -eq "analisador") {
        $splat = @{
            ComputerName = $targetHost
            OutputFolder = $outDir
            TimeoutSec   = $timeoutSec
        }
        if ($skipMajor) { $splat['SkipMajorData'] = $true }
        & $mainScript @splat
    } elseif ($tool -eq "manutencao") {
        & $mainScript -ComputerName $targetHost -Verbose
    } elseif ($tool -eq "perfis") {
        & $mainScript -ComputerName $targetHost -UsersToPurge $usersPurge
    }

    Write-Host ""
    Write-Host " [OK] Execução concluída com sucesso!" -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host " [ERRO] Ocorreu uma falha durante a execução:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
} finally {
    # Remove a pasta temporária dos scripts baixados para não deixar resquícios
    if (Test-Path $tempFolder) {
        Write-Host " [INFO] Limpando pasta temporária de execução..." -ForegroundColor Gray
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " Execução finalizada! Pressione ENTER para fechar a janela..." -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
    Read-Host
}
