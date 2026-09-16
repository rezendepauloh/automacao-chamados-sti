# ==============================================================================
# Script: bancada-launcher.ps1
# Função: Executor local do Protocol Handler 'bancada://' para estações Windows.
# ==============================================================================

param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$UriString = ""
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Garante que NENHUMA falha ou encerramento feche o PowerShell sem o usuário ver
try {
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "       SISTEMA BANCADA — DISPARADOR LOCAL WINDOWS" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan

    # Se chamado sem argumentos (ex: teste manual), solicita o link
    if ([string]::IsNullOrWhiteSpace($UriString)) {
        Write-Host " [AVISO] Nenhuma URL recebida por argumento de linha de comando." -ForegroundColor Yellow
        $UriString = Read-Host " Insira ou cole a URL bancada:// (ou pressione Enter para sair)"
        if ([string]::IsNullOrWhiteSpace($UriString)) {
            Write-Host " Nenhuma URL informada. Encerrando." -ForegroundColor Gray
            return
        }
    }

    Write-Host " [DEBUG] Argumento bruto recebido: $UriString" -ForegroundColor Gray

    # Limpeza profunda da URI (aspas, barras extras que browsers às vezes injetam no final)
    $cleanUri = $UriString.Trim().Trim('"').Trim("'").Trim('/')
    
    # Remove prefixo do protocolo: bancada://run? ou bancada:/run? ou bancada:? ou bancada://
    $cleanUri = $cleanUri -replace '^bancada:/*(run/?)?\??', ''
    if ($cleanUri.Contains('?')) {
        $cleanUri = $cleanUri.Substring($cleanUri.IndexOf('?') + 1)
    }

    # Extrai os parâmetros chave=valor
    $params = @{}
    $pairs = $cleanUri.Split('&')
    foreach ($pair in $pairs) {
        if ($pair -match "^(?<key>[^=]+)=(?<val>.*)$") {
            try {
                $k = [System.Uri]::UnescapeDataString($matches['key'])
                $v = [System.Uri]::UnescapeDataString($matches['val'])
                $params[$k] = $v
            } catch {
                $params[$matches['key']] = $matches['val']
            }
        }
    }

    $tool       = if ($params['tool']) { $params['tool'].ToLower().Trim() } else { "" }
    $targetHost = if ($params['host']) { $params['host'].Trim() } else { "" }
    $serverUrl  = $params['server']
    $skipMajor  = $params['skip_major'] -eq 'true'
    $timeoutSec = if ($params['timeout']) { [int]$params['timeout'] } else { 30 }
    $usersPurge = $params['users']
    $psEngine   = $params['ps_engine']

    $currentEngineName = if ($PSVersionTable.PSVersion.Major -ge 7) { "⚡ PowerShell 7+ ($($PSVersionTable.PSVersion))" } else { "💻 Windows PowerShell 5.1 ($($PSVersionTable.PSVersion))" }
    Write-Host " [INFO] Interpretador em uso  : $currentEngineName" -ForegroundColor Cyan
    Write-Host " [INFO] Ferramenta solicitada : $tool" -ForegroundColor Green
    Write-Host " [INFO] Máquina alvo          : $targetHost" -ForegroundColor Green
    if ($serverUrl) {
        Write-Host " [INFO] Servidor de Origem    : $serverUrl" -ForegroundColor Gray
    }

    if (-not $tool -or -not $targetHost) {
        Write-Host ""
        Write-Host " [ERRO] Parâmetros insuficientes na chamada:" -ForegroundColor Red
        Write-Host "        tool='$tool', host='$targetHost'" -ForegroundColor Red
        Write-Host "        URL original: $UriString" -ForegroundColor Yellow
        return
    }

    # Ferramentas leves locais (rdp, explorer, ping) rodam IMEDIATAMENTE no processo atual sem troca de engine
    if ($tool -eq "rdp") {
        Write-Host " [INFO] Iniciando Conexão de Área de Trabalho Remota (MSTSC) para: $targetHost" -ForegroundColor Green
        Start-Process "mstsc.exe" -ArgumentList "/v:$targetHost"
        Write-Host " [OK] Conexão RDP disparada com sucesso!" -ForegroundColor Green
        return
    }
    
    if ($tool -eq "explorer") {
        $sharePath = if ($params['path']) { "\\$targetHost\$($params['path'])" } else { "\\$targetHost\c$" }
        Write-Host " [INFO] Abrindo compartilhamento de rede no Explorer: $sharePath" -ForegroundColor Green
        Start-Process "explorer.exe" -ArgumentList $sharePath
        Write-Host " [OK] Compartilhamento aberto no Windows Explorer com sucesso!" -ForegroundColor Green
        return
    }
    
    if ($tool -eq "ping") {
        Write-Host " [INFO] Disparando teste de conectividade ICMP contínuo para $targetHost (Pressione Ctrl+C para parar)..." -ForegroundColor Yellow
        Write-Host ""
        ping.exe $targetHost -t
        Write-Host ""
        Write-Host " [OK] Teste de conectividade finalizado." -ForegroundColor Green
        return
    }

    # Para scripts pesados de manutenção/análise, checa se deve migrar para pwsh 7
    $pwsh7Path = "C:\Program Files\PowerShell\7\pwsh.exe"
    $hasPwsh7 = Test-Path $pwsh7Path
    if (-not $hasPwsh7) {
        $cmdTest = Get-Command "pwsh.exe" -ErrorAction SilentlyContinue
        if ($cmdTest) {
            $pwsh7Path = $cmdTest.Source
            $hasPwsh7 = $true
        }
    }

    $desiredEngine = if ($hasPwsh7 -and $psEngine -ne "powershell") { $pwsh7Path } else { "powershell.exe" }

    if ($PSVersionTable.PSVersion.Major -lt 7 -and $desiredEngine -ne "powershell.exe" -and $hasPwsh7) {
        Write-Host " [INFO] Migrando execução para PowerShell 7..." -ForegroundColor Cyan
        Start-Process -FilePath $desiredEngine -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath, "`"$UriString`"")
        return
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
            Write-Host " [ERRO] Ferramenta desconhecida: '$tool'" -ForegroundColor Red
            return
        }

        # Baixa os scripts do servidor bancada
        if (-not $serverUrl) {
            $serverUrl = "http://localhost:8502"
        }
        $serverUrl = $serverUrl.TrimEnd('/')

        Write-Host " [INFO] Baixando scripts auxiliares de $serverUrl..." -ForegroundColor Gray
        foreach ($file in $scriptFiles) {
            $url = "$serverUrl/static/scripts/$file"
            $dest = Join-Path $tempFolder $file
            try {
                Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing -TimeoutSec 15
                Write-Host "   -> Baixado: $file" -ForegroundColor Gray
            } catch {
                Write-Host "   [!] Aviso ao baixar $file via $url : $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        if (-not (Test-Path $mainScript)) {
            Write-Host " [ERRO] O script principal não foi encontrado em: $mainScript" -ForegroundColor Red
            return
        }

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
        Write-Host " [OK] Execução do script concluída com sucesso!" -ForegroundColor Green

    } finally {
        # Remove a pasta temporária dos scripts baixados para não deixar resquícios
        if ($tempFolder -and (Test-Path $tempFolder)) {
            Write-Host " [INFO] Limpando pasta temporária de execução..." -ForegroundColor Gray
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

} catch {
    Write-Host ""
    Write-Host " [ERRO FATAL] Ocorreu uma falha no launcher:" -ForegroundColor Red
    Write-Host $_.Exception.ToString() -ForegroundColor Red
} finally {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " Execução finalizada. Pressione ENTER para fechar esta janela..." -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
    Read-Host
}
