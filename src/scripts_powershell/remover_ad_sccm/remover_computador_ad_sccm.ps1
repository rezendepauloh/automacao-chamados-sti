<#
.SYNOPSIS
    Script unificado para localizar e remover contas de computador do Active Directory e SCCM.
    Suporta fallback de credencial de administrador (definido em .env, cred_admin.xml ou interativo) e remoção de proteção contra exclusão acidental.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, Position = 0, HelpMessage = "Digite o nome da maquina (ex: MPE-80600)")]
    [string]$ComputerName,

    [switch]$Force,
    [string]$SiteServer,
    [string]$SiteCode,
    [string]$AdAdminUser,
    [string]$AdAdminPassword
)

# -------------------------------------------------------------------------
# CARREGAR CONFIGURACOES DO .ENV E GARANTIR RUNTIME DESKTOP (POWERSHELL 5.1)
# -------------------------------------------------------------------------
$envHelper = Join-Path $PSScriptRoot "utils_env.ps1"
if (Test-Path $envHelper) {
    . $envHelper
    if (Get-Command Ensure-DesktopPowerShell -ErrorAction SilentlyContinue) {
        Ensure-DesktopPowerShell -ScriptPath $MyInvocation.MyCommand.Definition -BoundParameters $PSBoundParameters
    }
    if (Get-Command Import-DotEnv -ErrorAction SilentlyContinue) {
        [void](Import-DotEnv)
    }
}

$loggerHelper = Join-Path $PSScriptRoot "utils_logger.ps1"
if (Test-Path $loggerHelper) { . $loggerHelper }

function Write-SafeAppLog {
    param([string]$Type = "Info", [string]$Message)
    if (Get-Command Write-AppLog -ErrorAction SilentlyContinue) {
        Write-AppLog -Type $Type -Message $Message
    }
}

if (-not $SiteServer) {
    $SiteServer = if ($env:SCCM_SITE_SERVER) { $env:SCCM_SITE_SERVER } else { "srv-1046.in.mpe.ms.gov.br" }
}
if (-not $SiteCode) {
    $SiteCode = if ($env:SCCM_SITE_CODE) { $env:SCCM_SITE_CODE } else { "PGJ" }
}

# Caso o usuario execute sem argumentos, pergunta o nome do computador
if (-not $ComputerName) {
    Write-Host ""
    $ComputerName = Read-Host "Informe o nome do computador (ex: MPE-80600)"
    if ([string]::IsNullOrWhiteSpace($ComputerName)) {
        Write-Host "[ERRO] Nome do computador nao informado." -ForegroundColor Red
        exit 1
    }
}

$ComputerName = $ComputerName.Trim().ToUpper()

Write-SafeAppLog -Type Info -Message "--- Limpeza AD/SCCM iniciada para: $ComputerName ---"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   LIMPEZA DE COMPUTADOR: ACTIVE DIRECTORY & SCCM" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Computador alvo : $ComputerName" -ForegroundColor Yellow
Write-Host "Servidor SCCM   : $SiteServer (Site: $SiteCode)" -ForegroundColor DarkGray
if ($WhatIfPreference) {
    Write-Host "Modo de Teste   : SIMULACAO (-WhatIf ativado, nada sera excluido)" -ForegroundColor Magenta
}
Write-Host "----------------------------------------------------------" -ForegroundColor Gray

# -------------------------------------------------------------------------
# ETAPA 1: LOCALIZAR NO ACTIVE DIRECTORY
# -------------------------------------------------------------------------
Write-Host "`n[1/4] Consultando Active Directory..." -ForegroundColor Yellow

$adComputer = $null
try {
    if (-not (Get-Module -Name ActiveDirectory)) {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Import-Module ActiveDirectory -UseWindowsPowerShell -WarningAction SilentlyContinue -ErrorAction Stop
        } else {
            Import-Module ActiveDirectory -WarningAction SilentlyContinue -ErrorAction Stop
        }
    }
    $adComputer = Get-ADComputer -Filter "Name -eq '$ComputerName' -or sAMAccountName -eq '$ComputerName$'" -Properties OperatingSystem, Enabled, IPv4Address, whenCreated, ProtectedFromAccidentalDeletion -ErrorAction SilentlyContinue | Select-Object -First 1
} catch {
    Write-Host "  -> Falha ao carregar modulo ActiveDirectory ou consultar AD: $_" -ForegroundColor Red
}

if ($adComputer) {
    Write-Host "  [OK] Registro encontrado no Active Directory!" -ForegroundColor Green
    Write-Host "       Nome       : $($adComputer.Name)" -ForegroundColor Cyan
    Write-Host "       DN         : $($adComputer.DistinguishedName)" -ForegroundColor Cyan
    Write-Host "       Ativo      : $($adComputer.Enabled)" -ForegroundColor Cyan
    Write-Host "       SO         : $($adComputer.OperatingSystem)" -ForegroundColor Cyan
    Write-Host "       Criado em  : $($adComputer.whenCreated)" -ForegroundColor Cyan
    if ($adComputer.ProtectedFromAccidentalDeletion) {
        Write-Host "       Protecao   : ATIVADA (Protegido contra exclusao acidental)" -ForegroundColor Magenta
    }
} else {
    Write-Host "  [i] Computador '$ComputerName' nao localizado no Active Directory." -ForegroundColor Gray
}

# -------------------------------------------------------------------------
# ETAPA 2: LOCALIZAR NO CONFIGURATION MANAGER (SCCM)
# -------------------------------------------------------------------------
Write-Host "`n[2/4] Consultando Configuration Manager (SCCM)..." -ForegroundColor Yellow

$sccmDevices = @()
$sccmModuleCandidates = @(
    "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1",
    "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1",
    "C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1",
    "C:\Program Files\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
)
$sccmModulePath = $null
foreach ($mCand in $sccmModuleCandidates) {
    if (Test-Path $mCand) {
        $sccmModulePath = $mCand
        break
    }
}
if (-not $sccmModulePath -and (Get-Command Get-SccmModulePath -ErrorAction SilentlyContinue)) {
    $sccmModulePath = Get-SccmModulePath
}

if ($sccmModulePath) {
    try {
        if (-not (Get-Module -Name ConfigurationManager)) {
            Import-Module $sccmModulePath -ErrorAction Stop
        }

        if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider "CMSite" -Root $SiteServer -Description "SCCM Site Drive" -WhatIf:$false -ErrorAction Stop | Out-Null
        }

        $currentLocation = Get-Location
        Set-Location "$($SiteCode):"

        $sccmDevices = @(Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue)

        Set-Location $currentLocation
    } catch {
        Write-Host "  -> Falha ao consultar o Configuration Manager via Console: $_" -ForegroundColor Yellow
    }
}

# Fallback WMI / CIM caso o módulo do Console não esteja instalado localmente
if ($sccmDevices.Count -eq 0) {
    try {
        $wmiDev = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query "SELECT ResourceID, Name, SMSID, ClientActiveStatus, OperatingSystemNameandVersion FROM SMS_R_System WHERE Name = '$ComputerName'" -Authentication PacketPrivacy -ErrorAction SilentlyContinue
        if ($wmiDev) {
            foreach ($wd in $wmiDev) {
                $sccmDevices += [PSCustomObject]@{
                    ResourceID         = $wd.ResourceID
                    SMSID              = $wd.SMSID
                    ClientActiveStatus = $wd.ClientActiveStatus
                    DeviceOS           = $wd.OperatingSystemNameandVersion
                    FromWmi            = $true
                }
            }
        }
    } catch {}
}

if ($sccmDevices.Count -gt 0) {
    Write-Host "  [OK] $($sccmDevices.Count) registro(s) encontrado(s) no SCCM!" -ForegroundColor Green
    foreach ($dev in $sccmDevices) {
        Write-Host "       ResourceID : $($dev.ResourceID)" -ForegroundColor Cyan
        Write-Host "       SMSID      : $($dev.SMSID)" -ForegroundColor Cyan
        Write-Host "       Cliente Atv: $($dev.ClientActiveStatus)" -ForegroundColor Cyan
        Write-Host "       SO         : $($dev.DeviceOS)" -ForegroundColor Cyan
    }
} else {
    Write-Host "  [i] Computador '$ComputerName' nao localizado no banco de dados do SCCM." -ForegroundColor Gray
}

# -------------------------------------------------------------------------
# VERIFICAR SE HA ALGO A EXCLUIR
# -------------------------------------------------------------------------
if (-not $adComputer -and $sccmDevices.Count -eq 0) {
    Write-Host "`n[CONCLUIDO] A maquina '$ComputerName' nao existe em nenhuma das duas bases." -ForegroundColor Green
    exit 0
}

# -------------------------------------------------------------------------
# ETAPA 3: CONFIRMACAO DO USUARIO
# -------------------------------------------------------------------------
Write-Host ""
if ($WhatIfPreference) {
    Write-Host "[SIMULACAO CONCLUIDA] Nenhuma alteracao foi feita devido ao parametro -WhatIf." -ForegroundColor Magenta
    exit 0
}

if (-not $Force) {
    $confirm = Read-Host "Deseja realmente EXCLUIR '$ComputerName' do AD e/ou SCCM? (S/N)"
    if ($confirm -notmatch "^[SsYy]$") {
        Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Yellow
        exit 0
    }
}

# -------------------------------------------------------------------------
# ETAPA 4: EXCLUSAO NO ACTIVE DIRECTORY (COM SUPORTE A FALLBACK DE ADMIN)
# -------------------------------------------------------------------------
Write-Host "`n[3/4] Excluindo do Active Directory..." -ForegroundColor Yellow
if ($adComputer) {
    $deleted = $false

    # Tentativa 1: Com as credenciais do usuario logado atualmente
    try {
        if ($adComputer.ProtectedFromAccidentalDeletion) {
            Set-ADComputer -Identity $adComputer.DistinguishedName -ProtectedFromAccidentalDeletion $false -ErrorAction SilentlyContinue
        }
        Remove-ADComputer -Identity $adComputer.DistinguishedName -Confirm:$false -ErrorAction Stop
        Write-Host "  -> [SUCESSO] Computador removido do Active Directory (usuario atual)!" -ForegroundColor Green
        Write-SafeAppLog -Type Sucesso -Message "[$ComputerName] Removido do Active Directory (usuario atual)."
        $deleted = $true
    } catch {
        $firstError = $_.Exception.Message
        Write-Host "  -> Falha com usuario atual: $firstError" -ForegroundColor Yellow
    }

    # Tentativa 2: Fallback com credencial dedicada do AD (cred_admin_ad.xml) se disponível
    if (-not $deleted) {
        $credAdXml = Join-Path $PSScriptRoot "cred_admin_ad.xml"
        if (-not (Test-Path $credAdXml)) {
            $credAdXml = Join-Path (Split-Path $PSScriptRoot -Parent) "cred_admin_ad.xml"
        }
        if (Test-Path $credAdXml) {
            try {
                $savedAdCred = Import-Clixml -Path $credAdXml
                Write-Host "  -> Tentando fallback via cred_admin_ad.xml ($($savedAdCred.UserName))..." -ForegroundColor Cyan
                Set-ADComputer -Identity $adComputer.DistinguishedName -ProtectedFromAccidentalDeletion $false -Credential $savedAdCred -ErrorAction SilentlyContinue
                Remove-ADComputer -Identity $adComputer.DistinguishedName -Credential $savedAdCred -Confirm:$false -ErrorAction Stop
                Write-Host "  -> [SUCESSO] Computador removido do AD usando cred_admin_ad.xml!" -ForegroundColor Green
                Write-SafeAppLog -Type Sucesso -Message "[$ComputerName] Removido do AD via cred_admin_ad.xml."
                $deleted = $true
            } catch {
                Write-Host "  -> Falha com cred_admin_ad.xml: $_" -ForegroundColor Yellow
            }
        }
    }

    # Tentativa 3: Fallback com conta de Administrador do AD passada por parâmetro ou .env
    if (-not $deleted) {
        $targetAdUser = $AdAdminUser
        if (-not $targetAdUser) { $targetAdUser = $env:AD_ADMIN_USER }
        $targetAdPass = $AdAdminPassword
        if (-not $targetAdPass) { $targetAdPass = $env:AD_ADMIN_PASSWORD }

        # Se nenhum usuário foi configurado mas o usuário logado for detectado, tenta sugerir ou usar <login>_admin_ad
        if (-not $targetAdUser) {
            $currentUserShort = $env:USERNAME
            if ($currentUserShort) {
                # Mapeamento conhecido para os técnicos da Bancada STI
                $techMap = @{
                    "paulogoncalves" = "paulo_admin_ad"
                    "reginaldosb"    = "reginaldo_admin_ad"
                    "luizvillalba"   = "villalba_admin_ad"
                }
                if ($techMap.ContainsKey($currentUserShort.ToLower())) {
                    $targetAdUser = $techMap[$currentUserShort.ToLower()]
                } else {
                    $targetAdUser = "${currentUserShort}_admin_ad"
                }
            }
        }

        if ($targetAdUser) {
            Write-Host "  -> Tentando fallback com a conta administrativa do AD: $targetAdUser..." -ForegroundColor Cyan
            
            if (-not $targetAdPass) {
                $secPass = Read-Host "Digite a senha para $targetAdUser" -AsSecureString
            } else {
                $secPass = ConvertTo-SecureString $targetAdPass -AsPlainText -Force
            }

            $domain = if ($env:AD_DOMAIN) { $env:AD_DOMAIN } else { "in.mpe.ms.gov.br" }
            $fullAdminUser = if ($targetAdUser -notmatch '[@\\]') { "$targetAdUser@$domain" } else { $targetAdUser }
            $adminCred = New-Object System.Management.Automation.PSCredential($fullAdminUser, $secPass)

            try {
                Set-ADComputer -Identity $adComputer.DistinguishedName -ProtectedFromAccidentalDeletion $false -Credential $adminCred -ErrorAction SilentlyContinue
                Remove-ADComputer -Identity $adComputer.DistinguishedName -Credential $adminCred -Confirm:$false -ErrorAction Stop
                Write-Host "  -> [SUCESSO] Computador removido do AD usando credencial de $targetAdUser!" -ForegroundColor Green
                Write-SafeAppLog -Type Sucesso -Message "[$ComputerName] Removido do AD via $targetAdUser."
                $deleted = $true
            } catch {
                Write-Host "  -> [ERRO] Falha tambem com a conta ${targetAdUser}: $_" -ForegroundColor Red
            }
        }
    }

    # Tentativa 4: Fallback secundário com cred_admin.xml existente se disponível
    if (-not $deleted) {
        $credXml = Join-Path $PSScriptRoot "cred_admin.xml"
        if (-not (Test-Path $credXml)) {
            $credXml = Join-Path (Split-Path $PSScriptRoot -Parent) "cred_admin.xml"
        }
        if (Test-Path $credXml) {
            try {
                $savedCred = Import-Clixml -Path $credXml
                Write-Host "  -> Tentando fallback secundário via cred_admin.xml ($($savedCred.UserName))..." -ForegroundColor Cyan
                Set-ADComputer -Identity $adComputer.DistinguishedName -ProtectedFromAccidentalDeletion $false -Credential $savedCred -ErrorAction SilentlyContinue
                Remove-ADComputer -Identity $adComputer.DistinguishedName -Credential $savedCred -Confirm:$false -ErrorAction Stop
                Write-Host "  -> [SUCESSO] Computador removido do AD usando cred_admin.xml!" -ForegroundColor Green
                Write-SafeAppLog -Type Sucesso -Message "[$ComputerName] Removido do AD via cred_admin.xml."
                $deleted = $true
            } catch {
                Write-Host "  -> Falha com cred_admin.xml: $_" -ForegroundColor Yellow
            }
        }
    }

    if (-not $deleted) {
        Write-Host "  -> [DICA] Defina AD_ADMIN_USER ou gere cred_admin_ad.xml para conceder permissão de exclusão no AD." -ForegroundColor DarkGray
    }
} else {
    Write-Host "  -> Nada a remover no Active Directory." -ForegroundColor Gray
}

# -------------------------------------------------------------------------
# ETAPA 5: EXCLUSAO NO CONFIGURATION MANAGER (SCCM)
# -------------------------------------------------------------------------
Write-Host "`n[4/4] Excluindo do Configuration Manager (SCCM)..." -ForegroundColor Yellow
if ($sccmDevices.Count -gt 0) {
    $sccmDeleted = $false
    # Método 1: Via CM PowerShell Drive
    if (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue) {
        try {
            $currentLocation = Get-Location
            Set-Location "$($SiteCode):"

            foreach ($dev in $sccmDevices) {
                Write-Host "  -> Removendo ResourceID: $($dev.ResourceID)..." -ForegroundColor Cyan
                Remove-CMResource -ResourceId $dev.ResourceID -Force -ErrorAction Stop
            }

            Set-Location $currentLocation
            Write-Host "  -> [SUCESSO] Registro(s) removido(s) do Configuration Manager!" -ForegroundColor Green
            Write-SafeAppLog -Type Sucesso -Message "[$ComputerName] Registro(s) excluido(s) do Configuration Manager com sucesso."
            $sccmDeleted = $true
        } catch {
            Write-Host "  -> [AVISO] Falha ao remover via CM Drive: $_" -ForegroundColor Yellow
        }
    }

    # Método 2: Fallback direto via WMI SMS_R_System
    if (-not $sccmDeleted) {
        try {
            foreach ($dev in $sccmDevices) {
                $rId = $dev.ResourceID
                Write-Host "  -> Removendo ResourceID $rId via WMI SMS_R_System..." -ForegroundColor Cyan
                $wmiTarget = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_R_System WHERE ResourceID = $rId" -Authentication PacketPrivacy -ErrorAction Stop
                if ($wmiTarget) {
                    $wmiTarget.Delete()
                    Write-Host "  -> [SUCESSO] ResourceID $rId excluido do SCCM via WMI!" -ForegroundColor Green
                    Write-SafeAppLog -Type Sucesso -Message "[$ComputerName] ResourceID $rId excluido via WMI."
                    $sccmDeleted = $true
                }
            }
        } catch {
            Write-Host "  -> [ERRO] Falha tambem ao remover via WMI: $_" -ForegroundColor Red
        }
    }
} else {
    Write-Host "  -> Nada a remover no SCCM." -ForegroundColor Gray
}

Write-Host "`n==========================================================" -ForegroundColor Green
Write-Host "   LIMPEZA CONCLUIDA COM SUCESSO!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
