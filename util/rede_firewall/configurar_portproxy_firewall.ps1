<#
.SYNOPSIS
    Configura o PortProxy e Firewall do Windows para redirecionar as portas do WSL2 (Sistema Bancada).
.DESCRIPTION
    Obtem dinamicamente o endereco IP atual da instancia WSL e mapeia a porta 8501 no Windows,
    alem de garantir a liberacao no Firewall do Windows.
#>

# Parametros das Portas a mapear
$ports = @(8501)

Write-Host "Obtendo IP do WSL2..." -ForegroundColor Cyan

# Tenta obter o IP do WSL
try {
    $wslIp = (wsl hostname -I).Trim().Split(" ")[0]
} catch {
    Write-Error "Nao foi possivel obter o IP do WSL. Verifique se o WSL esta em execucao."
    exit 1
}

if (-not $wslIp) {
    Write-Error "Endereco IP do WSL retornou vazio."
    exit 1
}

Write-Host "IP do WSL identificado: $wslIp" -ForegroundColor Green

foreach ($port in $ports) {
    Write-Host "Configurando redirecionamento para porta $port..." -ForegroundColor Yellow
    
    # Remove mapeamento anterior para evitar duplicidade
    netsh interface portproxy delete v4tov4 listenport=$port listenaddress=0.0.0.0 2>$null | Out-Null
    
    # Adiciona mapeamento atualizado com o IP corrente do WSL
    netsh interface portproxy add v4tov4 listenport=$port listenaddress=0.0.0.0 connectport=$port connectaddress=$wslIp | Out-Null
    
    # Garante regra no Firewall do Windows
    $ruleName = "WSL2 PortProxy - Porta $port"
    $existingRule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
    if (-not $existingRule) {
        Write-Host "Criando regra no Firewall para a porta $port..." -ForegroundColor Yellow
        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -LocalPort $port -Protocol TCP -Action Allow -Profile Domain,Private,Public | Out-Null
    }
}

Write-Host "`nConfiguracao concluida com sucesso!" -ForegroundColor Green
Write-Host "Status do PortProxy:" -ForegroundColor Cyan
netsh interface portproxy show v4tov4
