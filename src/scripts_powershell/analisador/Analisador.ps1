<#
.SYNOPSIS
  Analisador de Dispositivos de Máquina – Versão Refatorada Compatível com PowerShell 5.1+
.DESCRIPTION
  Coleta informações principais de uma máquina remota via CIM/WSMan,
  sem expor senhas em texto claro. Utiliza um arquivo XML criptografado
  pelo DPAPI para armazenar credenciais (Export-Clixml/Import-Clixml), 
  cria apenas uma CimSession única e usa Get-CimInstance -CimSession em
  vez de Invoke-Command com CimSession para otimizar conexões. Por fim,
  gera arquivos HTML, PDF e Excel com todos os resultados formatados.

  Atenção: ajustado para PowerShell 5.1 no remoto (WS-Man/WinRM). No host local
  você pode executar com PowerShell 7+, mas as chamadas CIM/AD rodam em 5.1.

.PARAMETERS
  -ComputerName       (obrigatório): nome ou IP da máquina remota a ser escaneada.
  -OutputFolder       (opcional)   : pasta onde o relatório será salvo (HTML, PDF, XLSX).
                          Default: "$env:USERPROFILE\DeviceReports"
  -TimeoutSec         (opcional)   : timeout em segundos para operações CIM. Default: 30.
  -SkipMajorData      (opcional) switch: se presente, pula a coleta de Drivers, Programas,
                          Serviços, Processos e Usuários logados.

.EXAMPLE
  .\Analisador.ps1 -ComputerName "PJCHA-54491" -OutputFolder "C:\get-computer-info"

  .\Analisador.ps1 -ComputerName "PJCHA-54491" -SkipMajorData -OutputFolder "C:\get-computer-info"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [Parameter()]
    [string]$OutputFolder = "$env:USERPROFILE\DeviceReports",

    [Parameter()]
    [int]$TimeoutSec = 30,

    [switch]$SkipMajorData
)

# ----------------------------------------
# 0) WINRM e FIREWALL
# ----------------------------------------

function Enable-WinRM-viaDCOM {
    <#
    .SYNOPSIS
      Usa CimSession DCOM para habilitar/configurar o serviço WinRM no host remoto.
    .PARAMETER CimSession
      CimSession já aberto (DCOM) para o computador-alvo.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [CimSession]$CimSession
    )

    try {
        # 1) Recupera o serviço “WinRM” no host remoto
        $svc = Get-CimInstance -ClassName Win32_Service `
              -Filter "Name='WinRM'" `
              -CimSession $CimSession `
              -ErrorAction Stop

        # 2) Se o StartMode não for “Auto”, altera para “Auto”
        if ($svc.StartMode -ne "Auto") {
            $argsMode = @{ StartMode = "Automatic" }
            $null = Invoke-CimMethod -InputObject $svc `
                             -MethodName ChangeStartMode `
                             -Arguments $argsMode `
                             -CimSession $CimSession `
                             -ErrorAction Stop
        }

        # 3) Se o serviço não estiver rodando, inicia-o
        if ($svc.State -ne "Running") {
            $null = Invoke-CimMethod -InputObject $svc `
                             -MethodName StartService `
                             -CimSession $CimSession `
                             -ErrorAction Stop
        }

        Write-Verbose "WinRM configurado para Automatic e iniciado no host remoto."
    }
    catch {
        Write-Warning "Falha ao habilitar/configurar WinRM via DCOM: $_"
    }
}


function Open-Firewall-WinRM {
    <#
    .SYNOPSIS
      Usa Win32_Process.Create (via CIM/DCOM) para adicionar a regra de firewall
      que libera o grupo “Windows Remote Management” no host remoto.
    .PARAMETER CimSession
      CimSession já aberto (DCOM) para o computador-alvo.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [CimSession]$CimSession
    )

    # Comando NETSH que habilita a regra de firewall para WinRM
    $cmd = 'netsh advfirewall firewall set rule group="Windows Remote Management" new enable=yes'

    try {
        $proc = Invoke-CimMethod -ClassName Win32_Process `
                                 -MethodName Create `
                                 -Arguments @{ CommandLine = $cmd } `
                                 -CimSession $CimSession `
                                 -ErrorAction Stop

        if ($proc.ReturnValue -eq 0) {
            Write-Verbose "Exceção de firewall para WinRM habilitada via DCOM."
        } else {
            Write-Warning "Falha ao habilitar regra de firewall (código $($proc.ReturnValue))."
        }
    }
    catch {
        Write-Warning "Não foi possível executar netsh via DCOM: $_"
    }
}


function Set-TrustedHosts-viaDCOM {
    <#
    .SYNOPSIS
      Ajusta a política de TrustedHosts no registro remoto (WinRM\Client)
      usando StdRegProv (DCOM). Precisa de hDefKey correto como [uint32].
    .PARAMETER CimSession
      CimSession já ativo (DCOM) para o computador-alvo.
    .PARAMETER Hosts
      String com os hosts permitidos (por exemplo, “*” ou “HOST1,HOST2”).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [CimSession]$CimSession,

        [Parameter(Mandatory = $true)]
        [string]$Hosts
    )

    # HKEY_LOCAL_MACHINE em StdRegProv (uint32!)
    $HKLM = [uint32]2147483650

    # Caminho de registro WinRM\Client no Policies
    $subKey = "SOFTWARE\Policies\Microsoft\Windows\WinRM\Client"
    $valueName = "TrustedHosts"

    # Prepara argumentos para SetStringValue
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
                         -ErrorAction Stop

        Write-Verbose "Registry: HKLM:\$subKey\$valueName definido para '$Hosts'."
    }
    catch {
        Write-Warning "Falha ao ajustar TrustedHosts via StdRegProv: $_"
    }
}

# ----------------------------------------
# 1) CONFIGURAÇÕES INICIAIS
# ----------------------------------------

# Garantir resolução de variáveis de ambiente no OutputFolder ($env:USERPROFILE, %USERPROFILE%, etc.)
$OutputFolder = [System.Environment]::ExpandEnvironmentVariables($OutputFolder)
if ($OutputFolder -match '^\$env:(\w+)(.*)$') {
    $envVar = $Matches[1]
    $rest = $Matches[2]
    $envVal = [System.Environment]::GetEnvironmentVariable($envVar)
    if ($envVal) {
        $OutputFolder = Join-Path $envVal $rest.TrimStart('\', '/')
    }
}
$OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)

# Garantir que a pasta de saída exista
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}
Write-Host "📁 Diretório de saída dos relatórios: $OutputFolder" -ForegroundColor Cyan

# Caminho do arquivo XML com credenciais criptografadas pelo DPAPI
$CredXmlPath = Join-Path $PSScriptRoot "cred_admin.xml"

# Detecta se a máquina está no domínio do MPMS
$domain = (Get-CimInstance Win32_ComputerSystem).Domain
$isMPMSDomain = $domain -match "(?i)mpe"

if (-not (Test-Path $CredXmlPath)) {
    if ($isMPMSDomain) {
        Write-Warning "Arquivo de credencial nao encontrado em '$CredXmlPath'. Solicitando novas credenciais..."
        $newCred = Get-Credential -Message "Digite seu login administrativo (ex: paulo_admin) e senha para acesso remoto"
        $username = $newCred.UserName
        
        if ($username -match "\\") {
            if ($username -notmatch "^mpe\\") {
                $username = "mpe\" + ($username -split "\\")[-1]
            }
        } else {
            $username = "mpe\$username"
        }
        
        $Global:CredentialAdmin = New-Object System.Management.Automation.PSCredential($username, $newCred.Password)
        $Global:CredentialAdmin | Export-Clixml -Path $CredXmlPath
    } else {
        Write-Verbose "Arquivo de credencial nao encontrado e nao estamos no dominio MPMS. Usando credenciais atuais."
        $Global:CredentialAdmin = $null
    }
} else {
    try {
        $Global:CredentialAdmin = Import-Clixml -Path $CredXmlPath
    } catch {
        Write-Warning "Falha ao descriptografar '$CredXmlPath' (gerado por outro usuario/DPAPI). Solicitando credenciais..."
        
        if ($isMPMSDomain) {
            $newCred = Get-Credential -Message "Digite seu login administrativo (ex: paulo_admin) e senha para acesso remoto"
            $username = $newCred.UserName
            
            if ($username -match "\\") {
                if ($username -notmatch "^mpe\\") {
                    $username = "mpe\" + ($username -split "\\")[-1]
                }
            } else {
                $username = "mpe\$username"
            }
            
            $Global:CredentialAdmin = New-Object System.Management.Automation.PSCredential($username, $newCred.Password)
            $Global:CredentialAdmin | Export-Clixml -Path $CredXmlPath
        } else {
            Write-Warning "Falha ao descriptografar credenciais e nao estamos no dominio MPMS. Usando credenciais atuais."
            $Global:CredentialAdmin = $null
        }
    }
}

# Extrair usuário e senha (SecureString) do objeto PSCredential
$Global:UserNameCred = $Global:CredentialAdmin.UserName
$Global:PWord        = $Global:CredentialAdmin.Password    # SecureString

# ----------------------------------------
# 2) TESTAR CONECTIVIDADE E CRIAR CIMSESSION (PRIMEIRO TENTA DCOM, DEPOIS WS-MAN)
# ----------------------------------------

# 2.1) Verifica se está online via ping
if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
    Write-Error "Máquina '$ComputerName' está inacessível via ICMP. Abortando."
    exit 2
}

# 2.2) Tenta CimSession usando DCOM (RPC)
try {
    # 1) Criar sessão CIM via DCOM
    Write-Verbose "Tentando criar CimSession via DCOM (RPC) em '$ComputerName'..."
    $optDcom = New-CimSessionOption -Protocol Dcom
    $cimParams = @{
        ComputerName = $ComputerName
        SessionOption = $optDcom
    }
    if ($Global:CredentialAdmin) {
        $cimParams["Credential"] = $Global:CredentialAdmin
    }
    $Global:CimSession = New-CimSession @cimParams

    # 2) Habilitar o serviço WinRM no host remoto (via DCOM)
    Enable-WinRM-viaDCOM -CimSession $Global:CimSession

    # 3) Abrir a exceção de firewall para WinRM (via DCOM)
    Open-Firewall-WinRM -CimSession $Global:CimSession

    $meuHost = $env:COMPUTERNAME
    Set-TrustedHosts-viaDCOM -CimSession $Global:CimSession -Hosts $meuHost   

    # (c) (Opcional) Verifica se agora o WinRM está ativo, tentando um Test-WSMan
    #     Isso pode falhar se o firewall do remoto continuar bloqueando WS-Man,
    #     mas ao menos o serviço WinRM estará em “Running” e “Automatic”.
    Test-WSMan -ComputerName $ComputerName -ErrorAction SilentlyContinue | Out-Null
    if ($?) {
        Write-Verbose "WinRM configurado e respondendo no host remoto."
    } else {
        Write-Verbose "WinRM iniciado no remoto, mas WS-Man pode estar bloqueado pelo firewall."
    }

    # Teste rápido (pega 1 linha de Win32_OperatingSystem) para validar a sessão
    Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $Global:CimSession -ErrorAction Stop | Select-Object -First 1 | Out-Null
    Write-Verbose "CimSession DCOM bem-sucedida."
}
catch {
    Write-Warning "Falha ao criar CimSession via DCOM em '$ComputerName'. Tentando WS-MAN..."

    # 2.3) Se DCOM falhar, tenta pelo método original (WS-MAN)
    try {
        $cimParams = @{
            ComputerName = $ComputerName
        }
        if ($Global:CredentialAdmin) {
            $cimParams["Credential"] = $Global:CredentialAdmin
        }
        $Global:CimSession = New-CimSession @cimParams

        # Confirma que WinRM responde
        Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null
        Write-Verbose "CimSession WS-MAN bem-sucedida."
    }
    catch {
        Write-Error "Não foi possível conectar nem via DCOM nem via WS-MAN em '$ComputerName'."
        Write-Error "Se WS-MAN for necessário, certifique-se de executar no host remoto: 'winrm quickconfig' e 'Enable-PSRemoting -Force'."
        exit 3
    }
}

# ----------------------------------------
# 3) TABELAS DE MAPEAMENTO (Carregamento Externo)
# ----------------------------------------

# Carrega os mapeamentos de computadores e monitores de um arquivo externo para facilitar a manutenção
$MapeamentosPath = Join-Path $PSScriptRoot "Mapeamentos.ps1"
if (Test-Path $MapeamentosPath) {
    . $MapeamentosPath
} else {
    Write-Warning "Arquivo de mapeamentos não encontrado em: $MapeamentosPath. Usando tabelas de fallback vazias."
    $ModelMap = @{}
    $MonitorMap = @{}
    $jedecMap = @{}
    $videoTechMap = @{}
    $chassisTypeMap = @{}
    $GeneralManufacturerMap = @{}
}

# Função auxiliar para padronizar nomes de fabricantes (ex: LENOVO -> Lenovo)
function Get-FriendlyManufacturer {
    param([string]$Manufacturer)
    if ([string]::IsNullOrWhiteSpace($Manufacturer)) {
        return 'N/D'
    }
    $trimmed = $Manufacturer.Trim()
    $upper = $trimmed.ToUpper()
    if ($GeneralManufacturerMap.ContainsKey($upper)) {
        return $GeneralManufacturerMap[$upper]
    }
    return $trimmed
}

# Carrega o gerador de HTML externo para modularizar o código e facilitar a manutenção estética
$GeradorHtmlPath = Join-Path $PSScriptRoot "GeradorHtml.ps1"
if (Test-Path $GeradorHtmlPath) {
    . $GeradorHtmlPath
} else {
    Write-Error "Arquivo gerador de HTML essencial não encontrado em: $GeradorHtmlPath"
    exit 1
}

#region "Função para obter impressoras mapeadas por usuário"
function Get-UserPrinter {
    <#
    .SYNOPSIS
      Recupera as impressoras mapeadas (e o padrão) de cada usuário que já logou no host remoto,
      lendo em HKEY_USERS\<SID>\Printers\Connections e no valor “Device” em
      HKEY_USERS\<SID>\Software\Microsoft\Windows NT\CurrentVersion\Windows, via Invoke-Command (WS-Man).

    .PARAMETER ComputerName
      Nome (ou IP) da máquina remota.

    .PARAMETER Credential
      PSCredential com usuário/dominio que tenha permissão para ler registro no host remoto.

    .OUTPUTS
      Uma lista de PSCustomObject com as propriedades:
        - Usuário              : “DOMÍNIO\UserSamAccountName”
        - Impressoras Mapeadas : string com lista de impressoras (ex.: "\\PrintSrv\Laser1, \\PrintSrv\Inkjet" ou "Nenhuma")
        - Impressora Padrão    : nome da impressora padrão (ex.: "\\PrintSrv\Laser1" ou "N/D")
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]   $ComputerName,

        [Parameter(Mandatory = $true)]
        [PSCredential] $Credential
    )

    $scriptBlock = {
        # ======== Bloco executado NO HOST REMOTO ========

        # 1) Enumerar SIDs de perfis que existem em HKEY_USERS (exceto .DEFAULT e *Classes)
        $sids = Get-ChildItem Registry::HKEY_USERS |
                Where-Object { $_.Name -notlike "*\.DEFAULT" -and $_.Name -notlike "*\*\Classes" } |
                Select-Object -ExpandProperty Name

        $results = @()

        foreach ($fullKey in $sids) {
            # “fullKey” é algo como “HKEY_USERS\S-1-5-21-xxx-yyy-zzz-1001”
            $sid = $fullKey.Substring($fullKey.LastIndexOf('\') + 1)

            # 2) Monta caminho de “Printers\Connections” para este SID
            $connKey = "Registry::HKEY_USERS\$sid\Printers\Connections"
            if (Test-Path $connKey) {
                # 2a) Se existir, pega todas as subchaves (cada subchave corresponde a algo como "\\SERVER,Impressora,Ne02")
                $rawSub = Get-ChildItem $connKey -ErrorAction SilentlyContinue |
                          Select-Object -ExpandProperty Name

                if ($rawSub) {
                    # 2b) Converte cada entrada para nome legível (parte após última “\” e troca vírgula por “\”)
                    $printerNames = foreach ($entry in $rawSub) {
                        # Exemplo: "\\SERVER,Laser1,Ne02" -> "Laser1" -> "\\SERVER\Laser1"
                        $justName = if ($entry.Contains('\')) {
                            $entry.Substring($entry.LastIndexOf('\') + 1)
                        } else {
                            $entry
                        }
                        $justName -replace ',', '\'
                    }

                    # 3) Para buscar “Device” (impressora padrão), olha em HKU\<SID>\Software\Microsoft\Windows NT\CurrentVersion\Windows
                    $defKey = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows NT\CurrentVersion\Windows"
                    $default = "N/D"
                    if (Test-Path $defKey) {
                        try {
                            $valDevice = Get-ItemProperty -Path $defKey -Name Device -ErrorAction Stop |
                                         Select-Object -ExpandProperty Device
                            if ($valDevice -and $valDevice.Contains(',')) {
                                $default = $valDevice.Substring(0, $valDevice.IndexOf(','))
                            } elseif ($valDevice) {
                                $default = $valDevice
                            }
                        } catch {
                            # Se não encontrar “Device”, mantemos $default como N/D
                        }
                    }

                    # 4) Traduz SID → NTAccount (“DOMÍNIO\Usuario”)
                    try {
                        $ntAccount = (New-Object System.Security.Principal.SecurityIdentifier($sid)).
                                     Translate([System.Security.Principal.NTAccount]).Value
                    } catch {
                        $ntAccount = $sid
                    }

                    $printersFormatted = if ($printerNames) { $printerNames -join ', ' } else { "Nenhuma" }

                    # 5) Cria PSCustomObject para este usuário
                    $results += [PSCustomObject]@{
                        'Usuário'              = $ntAccount
                        'Impressoras Mapeadas' = $printersFormatted
                        'Impressora Padrão'    = $default
                    }
                }
            }
        }

        return $results
    }

    # ======== Fim do bloco remoto ========

    try {
        $mapped = Invoke-Command -ComputerName $ComputerName `
                                 -Credential $Credential `
                                 -ScriptBlock $scriptBlock `
                                 -ErrorAction Stop

        return $mapped | Select-Object 'Usuário', 'Impressoras Mapeadas', 'Impressora Padrão'
    }
    catch {
        Write-Warning "Falha ao coletar impressoras mapeadas via Invoke-Command: $_"
        return @()
    }
}
#endregion

#region "Função para obter usuários logados / perfis já acessados na máquina"
function Get-Users-Data {
    <#
    .SYNOPSIS
      Coleta informações dos usuários de domínio que já logaram naquela máquina,
      incluindo o “último uso” baseado no LastWriteTime da pasta de perfil remota,
      e calcula o espaço usado (em GB) localmente no host remoto.

    .PARAMETER ComputerName
      Nome (ou IP) da máquina remota.

    .PARAMETER Credential
      PSCredential válido para autenticação remota e leitura de registro/AD.

    .OUTPUTS
      Array de PSCustomObject com colunas:
        Nome Usuário,
        Usuário (samAccountName),
        Descrição (Cargo),
        Local de trabalho (Office),
        LocalPath (caminho físico “C:\Users\…”),
        Última vez usado (LastWriteTime do diretório),
        SID,
        Loaded (Logado / Não logado),
        Special (Sim / Não),
        Espaço Usado no Disco (GB, calculado remotamente).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]      $ComputerName,

        [Parameter(Mandatory = $true)]
        [PSCredential] $Credential
    )

    Write-Host "  ✓ Coletando usuários de domínio de '$ComputerName'..." -ForegroundColor Cyan

    try {
        # (1) Buscamos as informações brutas dos perfis via WMI e calculamos a data real
        #     de modificação do perfil usando o mesmo algoritmo nativo que o Windows usa na tela
        #     de "Perfis de Usuário" (painel de controle avançado).
        #     Ele combina as chaves de registro LocalProfileLoadTime e LocalProfileUnLoadTime,
        #     selecionando a mais recente, e usa o NTUSER.DAT como fallback.
        #
        $profilesData = Invoke-Command -ComputerName $ComputerName `
                                       -Credential $Credential `
                                       -ScriptBlock {
                                           $profiles = Get-CimInstance -ClassName Win32_UserProfile |
                                                       Where-Object { $_.LocalPath -and $_.SID } |
                                                       Select-Object SID, LocalPath, Loaded, Special

                                           # Cria um mapa do SID para a data real de logon ("Modificado em")
                                           $timeMap = @{}
                                           foreach ($p in $profiles) {
                                               $sid = $p.SID
                                               $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
                                               $bestTime = $null

                                               if (Test-Path $regPath) {
                                                   $props = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                                                   
                                                   $loadTime = $null
                                                   if ($null -ne $props -and $null -ne $props.LocalProfileLoadTimeHigh -and $null -ne $props.LocalProfileLoadTimeLow) {
                                                       $ft64 = ([UInt64]$props.LocalProfileLoadTimeHigh -shl 32) -bor $props.LocalProfileLoadTimeLow
                                                       if ($ft64 -gt 0) {
                                                           $loadTime = [datetime]::FromFileTime($ft64)
                                                       }
                                                   }
                                                   
                                                   $unloadTime = $null
                                                   if ($null -ne $props -and $null -ne $props.LocalProfileUnLoadTimeHigh -and $null -ne $props.LocalProfileUnLoadTimeLow) {
                                                       $ft64 = ([UInt64]$props.LocalProfileUnLoadTimeHigh -shl 32) -bor $props.LocalProfileUnLoadTimeLow
                                                       if ($ft64 -gt 0) {
                                                           $unloadTime = [datetime]::FromFileTime($ft64)
                                                       }
                                                   }
                                                   
                                                   # Pega a mais recente entre carregamento e descarregamento do perfil
                                                   if ($null -ne $loadTime -and $null -ne $unloadTime) {
                                                       if ($loadTime -gt $unloadTime) {
                                                           $bestTime = $loadTime
                                                       } else {
                                                           $bestTime = $unloadTime
                                                       }
                                                   } elseif ($null -ne $loadTime) {
                                                       $bestTime = $loadTime
                                                   } elseif ($null -ne $unloadTime) {
                                                       $bestTime = $unloadTime
                                                   }
                                               }

                                               # Fallback para o ntuser.dat caso as chaves de registro não existam ou estejam zeradas
                                               if ($null -eq $bestTime) {
                                                   $ntuserPath = Join-Path $p.LocalPath "ntuser.dat"
                                                   if (Test-Path $ntuserPath) {
                                                       $bestTime = (Get-Item -Path $ntuserPath -Force -ErrorAction SilentlyContinue).LastWriteTime
                                                   }
                                               }

                                               if ($null -ne $bestTime) {
                                                   $timeMap[$sid] = $bestTime
                                               }
                                           }

                                           [PSCustomObject]@{
                                               Profiles = $profiles
                                               TimeMap  = $timeMap
                                           }
                                       } -ErrorAction Stop

        $profiles = $profilesData.Profiles
        $timeMap  = $profilesData.TimeMap
        $results = @()

        foreach ($uProfile in $profiles) {
            $sid       = $uProfile.SID
            $localPath = $uProfile.LocalPath

            #
            # (2) Pular contas que não existem no AD (ex.: perfis de serviço, System, NetworkService etc.)
            #
            $adObj = $null
            try {
                $adObj = Get-ADUser -Identity $sid -Properties Name, SamAccountName, Description, Office -ErrorAction Stop
            } catch {
                $adObj = $null
            }
            if (-not $adObj) {
                continue
            }

            #
            # (3) Extrai dados do AD para uso no relatório
            #
            $simpleUser  = $adObj.SamAccountName
            $displayName = $adObj.Name
            $description = $adObj.Description
            $office      = $adObj.Office

            #
            # (4) Traduz “Loaded” e “Special”
            #
            $loadedText = "Não logado"
            if ($uProfile.Loaded) {
                $loadedText = "Logado"
            }
            $specialText = "Não"
            if ($uProfile.Special) {
                $specialText = "Sim"
            }

            #
            # (5) Obter “Última vez usado” de forma extremamente precisa.
            #     Buscamos o LastWriteTime do arquivo NTUSER.DAT mapeado na nossa tabela em memória,
            #     o que reflete o logon real e autêntico de um ser humano na máquina (evitando falsos
            #     positivos gerados por processos de antivírus e atualizações do Windows).
            #
            $realLoginTime = $null
            if ($timeMap.ContainsKey($sid)) {
                $realLoginTime = $timeMap[$sid]
            }

            if ($null -ne $realLoginTime) {
                if ($realLoginTime -is [System.DateTime]) {
                    $lastUsedText = $realLoginTime.ToString('dd/MM/yyyy HH:mm:ss')
                } else {
                    $lastUsedText = $realLoginTime
                }
            } else {
                $lastUsedText = "N/D"
            }

            # (6) Calcular “Espaço Usado no Disco” localmente no host remoto, para cada perfil
            # Usando Robocopy em modo de listagem (/L) com exclusão de junções (/XJD), que é imbatível em velocidade e precisão de permissões.
            #
            $sizeGB = Invoke-Command -ComputerName $ComputerName `
                                     -Credential $Credential `
                                     -ScriptBlock {
                                         param($path)
                                         if (-not (Test-Path $path)) { return 0 }
                                         try {
                                             # Executa robocopy silencioso de listagem para contar bytes sem loops de junções
                                             $raw = robocopy $path NULL /L /S /NJH /NFL /NDL /R:0 /W:0 /BYTES /XJD
                                             $bytesLine = $raw | Where-Object { $_ -match '^\s*Bytes\s*:\s*(\d+)' }
                                             if ($bytesLine -and $matches[1]) {
                                                 $bytes = [double]$matches[1]
                                                 [math]::Round($bytes / 1GB, 2)
                                             } else {
                                                 return 0
                                             }
                                         } catch {
                                             return 0
                                         }
                                     } -ArgumentList $localPath `
                                       -ErrorAction SilentlyContinue

            if ($null -eq $sizeGB) {
                $espacoTexto = "N/D"
            } else {
                $espacoTexto = "{0:N2} GB" -f $sizeGB
            }

            #
            # (7) Montar o objeto final para este usuário de domínio
            #
            $results += [PSCustomObject]@{
                'Nome Usuário'          = $displayName
                'Usuário'               = $simpleUser
                'Descrição (Cargo)'     = $description
                'Local de trabalho'     = $office
                'LocalPath'             = $localPath
                'Última vez usado'      = $lastUsedText
                'SID'                   = $sid
                'Loaded'                = $loadedText
                'Special'               = $specialText
                'Espaço Usado no Disco' = $espacoTexto
            }
        }

        return $results
    }
    catch {
        Write-Warning "Erro ao coletar dados de usuários via Invoke-Command: $_"
        return @()
    }
}
#endregion

# ----------------------------------------
# 4) FUNÇÕES DE COLETA DE DADOS
# ----------------------------------------

<#
.SYNOPSIS
  Coleta várias informações de WMI/CIM (HW, usuário atual, OS, BIOS, CPU, MB, Chassi,
  Memória, Disco, Rede, Drivers, Monitores, Programas, Serviços, Processos).
.PARAMETER CimSession
  Objeto CimSession já autenticado.
.OUTPUTS
  PSCustomObject contendo propriedades:
    ComputerName,
    HardwareInfo,
    CurrentUser,
    CurrentUserDisplayName,
    CurrentUserDomain,
    OS,
    BIOS,
    Processor,
    Motherboard,
    Chassis,
    MemoryModules,
    DiskDrives,
    LogicalDisks,
    NetworkConfigs,
    Monitors,
    VideoController,
    SoundDevices,
    PrintersLocal,
    PrintersMapped,
    InstalledDrivers,
    InstalledPrograms,
    RunningServices,
    ActiveProcesses,
    LastBootTime.
#>
function Get-AllSystemData {
    param(
        [Parameter(Mandatory = $true)]
        [string]      $ComputerName,
    
        [Parameter(Mandatory = $true)]
        [CimSession]$CimSession,

        [switch]$SkipMajorData
    )

    try {
        #########################
        # 4.1) INFORMAÇÕES DO COMPUTADOR (HW e USUÁRIO ATUAL)
        #########################
        Write-Host "  ✓ Informações do computador '$($ComputerName)'..." -ForegroundColor Cyan

        # Função para detectar portas de vídeo ativas consultando o WMI
        function Get-ComputerVideoPorts {
            param($CimSession)
            
            try {
                # Busca as conexões de monitores ativas em tempo real e parâmetros de display
                $connections = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorConnectionParams -CimSession $CimSession -ErrorAction Stop
                $displayParams = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -CimSession $CimSession -ErrorAction Stop
                
                $paramMap = @{}
                $displayParams | ForEach-Object { $paramMap[$_.InstanceName.ToLower()] = $_ }

                # Filtra conexões ativas e traduz para nomes legíveis com conversão segura de tipo para Int32 usando o mapeamento global $videoTechMap
                $activePorts = $connections | Where-Object { $_.Active } | ForEach-Object {
                    $inst = $_.InstanceName.ToLower()
                    $p = $paramMap[$inst]
                    if ($p -and $p.VideoInputType -eq 0) {
                        'VGA'
                    } else {
                        $code = $_.VideoOutputTechnology
                        if ($null -ne $code) {
                            $codeInt = [int]$code
                            if ($videoTechMap.ContainsKey($codeInt)) { $videoTechMap[$codeInt] } else { "Outra ($code)" }
                        }
                    }
                } | Sort-Object -Unique

                if ($activePorts) {
                    return $activePorts
                }
            }
            catch {
                # Falha silenciosa em VMs ou hosts sem monitor ativo
            }
            return @()
        }

        try {
            # Coleta de dados principal
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $CimSession -ErrorAction Stop
            $bios = Get-CimInstance -ClassName Win32_BIOS -CimSession $CimSession -ErrorAction Stop

            # Detectar portas de vídeo ativas
            $activePorts = Get-ComputerVideoPorts -CimSession $CimSession 2>$null
            $portasUsadas = ($activePorts | Sort-Object -Unique) -join ', '

            # Mapeamento de modelos
            $rawModel = $cs.Model.Trim()
            $friendlyComputer = $rawModel
            $entradasVideo = 'Não mapeado'

            if ($ModelMap.ContainsKey($rawModel)) {
                $friendlyComputer = $ModelMap[$rawModel].FriendlyName
                $entradasVideo = $ModelMap[$rawModel].VideoPorts
            }

            # Montar objeto de hardware (Computador como primeira coluna, Modelo como segunda)
            $hwObject = [PSCustomObject]@{
                'Computador'               = $friendlyComputer
                'Modelo'                   = $rawModel
                'Fabricante'               = Get-FriendlyManufacturer $cs.Manufacturer
                'Memória Física (GB)'      = "$([math]::Round($cs.TotalPhysicalMemory / 1GB, 2)) GB"
                'Processadores Lógicos'    = $cs.NumberOfLogicalProcessors
                'Processadores Físicos'    = $cs.NumberOfProcessors
                'Serial Number'            = $bios.SerialNumber
                'Portas de Vídeo Ativas'   = if ($portasUsadas) { $portasUsadas } else { 'Nenhuma detectada' }
                'Entradas de vídeo'        = $entradasVideo
            }
        }
        catch {
            Write-Warning "Falha na coleta de informações do computador: $($_.Exception.Message)"
            $hwObject = [PSCustomObject]@{
                Erro            = "Falha na coleta de dados"
                Detalhes        = $_.Exception.Message
            }
        }

        #########################
        # 4.2) INFORMAÇÕES DO USUÁRIO
        #########################
        # Usuário atual / domínio
        $loggedUserRaw = $cs.UserName      # pode vir “DOMÍNIO\usuario” ou apenas “usuario”
        $loggedDomain  = $cs.Domain

        if ([string]::IsNullOrWhiteSpace($loggedUserRaw)) {
            $simpleUser  = ''
            $displayName = ''
        } else {
            # Se vier “DOMÍNIO\usuario”, extrair só o nome de usuário
            if ($loggedUserRaw -like '*\*') {
                $parts      = $loggedUserRaw.Split('\')
                $simpleUser = $parts[1]
            } else {
                $simpleUser = $loggedUserRaw
            }

            # Tentar consultar AD (no remoto roda 5.1) para obter DisplayName
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                $adUser      = Get-ADUser -Identity $simpleUser -Properties DisplayName -ErrorAction Stop
                $displayName = $adUser.DisplayName
            } catch {
                # Se falhar (usuário não existir, permissão, etc.), deixar vazio
                $displayName = ''
            }
        }

        #########################
        # 4.3) INFORMAÇÕES DO SISTEMA OPERACIONAL (OS)
        #########################
        Write-Host "  ✓ Informações do Sistema Operacional '$($ComputerName)'..." -ForegroundColor Cyan

        $wmiOS = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $CimSession |
                 Select-Object Caption, Version, BuildNumber, CSName, LastBootUpTime

        # Coleta "Versão do Windows" via StdRegProv (DCOM) – não precisa de WinRM
        try {
            $regPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
            $subKey  = $regPath
            $value1  = "DisplayVersion"
            $value2  = "ReleaseId"
            # Constante para HKLM no StdRegProv
            $HKLM = [uint32]2147483650

            # 1) Tenta “DisplayVersion”
            $argsDisplay = @{
                hDefKey     = $HKLM
                sSubKeyName = $subKey
                sValueName  = $value1
            }
            $getDisp = Invoke-CimMethod -Namespace root\default -ClassName StdRegProv `
                         -MethodName GetStringValue `
                         -Arguments $argsDisplay `
                         -CimSession $CimSession `
                         -ErrorAction SilentlyContinue

            if ($getDisp -and $getDisp.sValue) {
                $winRelease = $getDisp.sValue
            } else {
                # 2) Se não houver “DisplayVersion”, tenta “ReleaseId”
                $argsRel = @{
                    hDefKey     = $HKLM
                    sSubKeyName = $subKey
                    sValueName  = $value2
                }
                $getRel = Invoke-CimMethod -Namespace root\default -ClassName StdRegProv `
                             -MethodName GetStringValue `
                             -Arguments $argsRel `
                             -CimSession $CimSession `
                             -ErrorAction SilentlyContinue

                if ($getRel -and $getRel.sValue) {
                    $winRelease = $getRel.sValue
                } else {
                    # 3) Se nenhum dos dois existir (acesso negado ou chave não existe),
                    # usamos o “Version” do WMI como fallback (que sempre vai existir)
                    $winRelease = $wmiOS.Version
                }
            }
        } catch {
            # Qualquer exceção no StdRegProv, cai aqui e pega o WMI.Version
            $winRelease = $wmiOS.Version
        }       

        $osObject = [PSCustomObject]@{
            'SO'                       = $wmiOS.Caption
            'Versão WMI (Version)'     = $wmiOS.Version
            'Versão do Windows (HnH)'  = $winRelease
            'Número da Build'          = $wmiOS.BuildNumber
            'Nome da Máquina'          = $wmiOS.CSName
            'Último Boot'              = $wmiOS.LastBootUpTime.ToString('dd/MM/yyyy HH:mm:ss')
        }

        #########################
        # 4.4) BIOS
        #########################
        Write-Host "  ✓ Informações da Bios '$($ComputerName)'..." -ForegroundColor Cyan
        
        $biosInfo = Get-CimInstance -ClassName Win32_BIOS -CimSession $CimSession |
                    Select-Object Manufacturer, SMBIOSBIOSVersion, SerialNumber, ReleaseDate

        $biosInfo = [PSCustomObject]@{
            'Fabricante'               = Get-FriendlyManufacturer $biosInfo.Manufacturer
            'Versão da BIOS'           = $biosInfo.SMBIOSBIOSVersion
            'Número de Serial'         = $biosInfo.SerialNumber
            'Data de lançamento'       = $biosInfo.ReleaseDate
        }
        
        #########################
        # 4.5) Processador
        #########################
        Write-Host "  ✓ Informações da Processador '$($ComputerName)'..." -ForegroundColor Cyan

        $cpuInfo = Get-CimInstance -ClassName Win32_Processor -CimSession $CimSession |
                   Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, CurrentVoltage, DataWidth, L2CacheSize, L3CacheSize

        $cpuInfo = [PSCustomObject]@{
            'Nome'                             = $cpuInfo.Name
            'Número de Cores'                  = $cpuInfo.NumberOfCores
            'Número de processadores lógicos'  = $cpuInfo.NumberOfLogicalProcessors
            'Clock máximo'                     = if ($null -ne $cpuInfo.MaxClockSpeed) { "$($cpuInfo.MaxClockSpeed) Mhz" } else { 'N/D' }
            'Voltagem'                         = if ($null -ne $cpuInfo.CurrentVoltage) { "$($cpuInfo.CurrentVoltage) V" } else { 'N/D' }
            'L1 Cache'                         = if ($null -ne $cpuInfo.DataWidth) { "$($cpuInfo.DataWidth) KB" } else { 'N/D' }
            'L2 Cache'                         = if ($null -ne $cpuInfo.L2CacheSize) { "$($cpuInfo.L2CacheSize) KB" } else { 'N/D' }
            'L3 Cache'                         = if ($null -ne $cpuInfo.L3CacheSize) { "$($cpuInfo.L3CacheSize) KB" } else { 'N/D' }
        }

        #########################
        # 4.6) Placa-Mãe
        #########################
        Write-Host "  ✓ Informações da Placa Mãe '$($ComputerName)'..." -ForegroundColor Cyan

        $motherboardInfo= Get-CimInstance -ClassName Win32_BaseBoard -CimSession $CimSession | Select-Object Manufacturer, Product, Version
        
        $motherboardInfo = [PSCustomObject]@{
            'Fabricante'               = Get-FriendlyManufacturer $motherboardInfo.Manufacturer
            'Produto'                  = $motherboardInfo.Product
            'Versão'                   = $motherboardInfo.Version
        }   

        #########################
        # 4.7) Chassi
        #########################
        Write-Host "  ✓ Informações da Chassi '$($ComputerName)'..." -ForegroundColor Cyan

        # 1) Coleta bruta do chassi
        $rawChassis = Get-CimInstance -ClassName Win32_SystemEnclosure -CimSession $CimSession |
            Select-Object ChassisTypes, Manufacturer

        # O mapa $chassisTypeMap foi carregado externamente do arquivo Mapeamentos.ps1

        # 3) Traduzir cada valor de ChassisTypes para o texto correspondente,
        #    convertendo $_ (UInt16) para [int] antes de consultar o hashtable.
        $chassisInfo = [PSCustomObject]@{
            'Tipo de Chassi' = ($rawChassis.ChassisTypes | ForEach-Object {
                $codeInt = [int]$_
                if ($chassisTypeMap.ContainsKey($codeInt)) {
                    $chassisTypeMap[$codeInt]
                } else {
                    "Unknown($codeInt)"
                }
            }) -join ", "
            'Fabricante' = Get-FriendlyManufacturer $rawChassis.Manufacturer
        } 

        #########################
        # 4.8) MEMÓRIA RAM
        #########################
        Write-Host "  ✓ Informações da Memória Ram '$($ComputerName)'..." -ForegroundColor Cyan
        
        # Obtém os módulos de memória física utilizando a classe robusta Win32_PhysicalMemory diretamente
        try {
            $rawMem = Get-CimInstance -ClassName Win32_PhysicalMemory -CimSession $CimSession -ErrorAction Stop |
                    Select-Object Manufacturer, Capacity, Speed, PartNumber, SerialNumber

            # O dicionário $jedecMap foi carregado externamente do arquivo Mapeamentos.ps1

            # Traduz os dados usando o dicionário JEDEC ou mantendo o formato bruto
            $memoryInfo = $rawMem | ForEach-Object {
                $rawMan = $_.Manufacturer
                $friendlyMan = 'N/D'

                if ($null -ne $rawMan) {
                    $rawManTrim = $rawMan.Trim()
                    if ($rawManTrim -match '^(?:0x)?([A-Fa-f0-9]+)') {
                        try {
                            $code = [Convert]::ToInt32($Matches[1], 16)
                            if ($jedecMap.ContainsKey($code)) {
                                $friendlyMan = $jedecMap[$code]
                            } else {
                                $friendlyMan = "Desconhecido (0x{0:X})" -f $code
                            }
                        } catch {
                            $friendlyMan = $rawManTrim
                        }
                    } elseif ($rawManTrim.Length -gt 0) {
                        $friendlyMan = $rawManTrim
                    }
                }

                [PSCustomObject]@{
                    'Fabricante'       = $friendlyMan
                    'Capacidade (GB)'  = "{0} GB" -f ([math]::Round($_.Capacity / 1GB, 2))
                    'Velocidade (MHz)' = if ($_.Speed) { "$($_.Speed) MHz" } else { 'N/D' }
                    'Part Number'      = if ($_.PartNumber) { $_.PartNumber.Trim() } else { 'N/D' }
                    'Número de Série'  = if ($_.SerialNumber) { $_.SerialNumber.Trim() } else { 'N/D' }
                }
            }
        }
        catch {
            Write-Warning "Falha na coleta de informações de memória ram: $($_.Exception.Message)"
            $memoryInfo = @()
        }

        ################################
        # 4.9) MONITOR INFO (via EDID)
        ################################
        Write-Host "  ✓ Informações do Monitor '$($ComputerName)'..." -ForegroundColor Cyan

        # 1) array final de PSCustomObject
        $monitors = @()

        # O dicionário $videoTechMap foi carregado externamente do arquivo Mapeamentos.ps1

        function Convert-VideoOutputTech {
            param([int]$Code)
            if ($videoTechMap.ContainsKey($Code)) { $videoTechMap[$Code] } else { "Desconhecido ($Code)" }
        }

        function Convert-MonitorManufacturer {
            param([string]$Code)
            if ($null -ne $ManufacturerMap -and $ManufacturerMap.ContainsKey($Code)) { $ManufacturerMap[$Code] } else { $Code }
        }

        # 3) Primeiro, tentar com CimSession
        try {
            # a) Coleta EDID melhorada
            $rawMonitors = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -CimSession $CimSession -ErrorAction Stop | ForEach-Object {
                $monitor = $_
                
                # Conversão correta de dados binários
                $manufacturer = ($monitor.ManufacturerName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''
                $serial = ($monitor.SerialNumberID | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''
                $name = ($monitor.UserFriendlyName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''

                [PSCustomObject]@{
                    InstanceId       = $monitor.InstanceName.ToLower()
                    RawName          = $monitor.InstanceName.Split('\')[1]
                    ManufacturerName  = $manufacturer
                    SerialNumberID    = $serial
                    UserFriendlyName  = $name
                    Active           = $monitor.Active
                }
            }

            # b) Coleta de conexões aprimorada
            $rawConnections = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorConnectionParams -CimSession $CimSession -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{
                    InstanceId            = $_.InstanceName.ToLower()
                    VideoOutputTechnology = $_.VideoOutputTechnology
                    Active               = $_.Active
                }
            }

            # c) Coleta de parâmetros físicos
            $physicalParams = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -CimSession $CimSession | ForEach-Object {
                [PSCustomObject]@{
                    InstanceId     = $_.InstanceName.ToLower()
                    Horizontal     = $_.MaxHorizontalImageSize
                    Vertical       = $_.MaxVerticalImageSize
                    VideoInputType = $_.VideoInputType
                }
            }

            # Criar dicionários para junção usando InstanceId único
            $connMap = @{}; $paramMap = @{}
            $rawConnections | ForEach-Object { $connMap[$_.InstanceId] = $_ }
            $physicalParams | ForEach-Object { $paramMap[$_.InstanceId] = $_ }

            # d) Montagem do objeto final
            foreach ($m in $rawMonitors) {
                $conn = $connMap[$m.InstanceId]
                $params = $paramMap[$m.InstanceId]
                
                $portaAtiva = if ($params -and $params.VideoInputType -eq 0) {
                    'VGA'
                } elseif ($conn) {
                    Convert-VideoOutputTech $conn.VideoOutputTechnology
                } else {
                    'N/A'
                }
                
                $obj = [PSCustomObject]@{
                    Nome            = $m.RawName
                    Polegada        = if ($params) { [math]::Round(($params.Horizontal/2.54 * $params.Vertical/2.54)/100,1) } else { 'N/A' }
                    'Resolução'     = if ($params) { "$($params.Horizontal)x$($params.Vertical)" } else { 'N/A' }
                    Tamanho         = if ($params) { "$($params.Horizontal)mm x $($params.Vertical)mm" } else { 'N/A' }
                    Ativo           = if ($m.Active) { "Sim" } else { "Não" }
                    'Serial Number' = $m.SerialNumberID
                    'Nome amigável' = $m.UserFriendlyName
                    Fabricante      = Convert-MonitorManufacturer $m.ManufacturerName
                    'Entradas de vídeo' = 'N/A'  # Será substituído pelo MonitorMap
                    'Porta Ativa'   = $portaAtiva
                }

                # Aplicar mapeamento manual
                if ($MonitorMap.ContainsKey($obj.Nome)) {
                    $cfg = $MonitorMap[$obj.Nome]
                    $obj.Nome               = $cfg.Nome
                    $obj.Polegada           = $cfg.Polegada
                    $obj.'Resolução'        = $cfg.'Resolução'
                    $obj.Tamanho            = $cfg.Tamanho
                    $obj.'Entradas de vídeo' = $cfg.VideoPorts
                }

                $monitors += $obj
            }
        }
        catch {
            # 4) Fallback via Invoke-Command
            try {
                $remoteData = Invoke-Command -ComputerName $ComputerName -Credential $Global:CredentialAdmin -ScriptBlock {
                    # Coletar todos dados localmente
                    $data = @{}
                    
                    # Monitores
                    $data.monitors = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
                        [PSCustomObject]@{
                            InstanceId       = $_.InstanceName.ToLower()
                            RawName          = $_.InstanceName.Split('\')[1]
                            ManufacturerName = ($_.ManufacturerName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''
                            SerialNumberID   = ($_.SerialNumberID | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''
                            UserFriendlyName = ($_.UserFriendlyName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''
                            Active          = $_.Active
                        }
                    }
                    
                    # Conexões
                    $data.connections = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorConnectionParams | ForEach-Object {
                        [PSCustomObject]@{
                            InstanceId = $_.InstanceName.ToLower()
                            VideoOutputTechnology = $_.VideoOutputTechnology
                        }
                    }
                    
                    # Parâmetros físicos
                    $data.physical = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams | ForEach-Object {
                        [PSCustomObject]@{
                            InstanceId     = $_.InstanceName.ToLower()
                            Horizontal     = $_.MaxHorizontalImageSize
                            Vertical       = $_.MaxVerticalImageSize
                            VideoInputType = $_.VideoInputType
                        }
                    }
                    
                    return $data
                }

                # Processar dados remotos usando InstanceId único
                $connMap = @{}; $paramMap = @{}
                $remoteData.connections | ForEach-Object { $connMap[$_.InstanceId] = $_ }
                $remoteData.physical | ForEach-Object { $paramMap[$_.InstanceId] = $_ }

                foreach ($m in $remoteData.monitors) {
                    $conn = $connMap[$m.InstanceId]
                    $params = $paramMap[$m.InstanceId]
                    
                    $portaAtiva = if ($params -and $params.VideoInputType -eq 0) {
                        'VGA'
                    } elseif ($conn) {
                        Convert-VideoOutputTech $conn.VideoOutputTechnology
                    } else {
                        'N/A'
                    }
                    
                    $obj = [PSCustomObject]@{
                        Nome            = $m.RawName
                        Polegada        = if ($params) { [math]::Round(($params.Horizontal/2.54 * $params.Vertical/2.54)/100,1) } else { 'N/A' }
                        'Resolução'     = if ($params) { "$($params.Horizontal)x$($params.Vertical)" } else { 'N/A' }
                        Tamanho         = if ($params) { "$($params.Horizontal)mm x $($params.Vertical)mm" } else { 'N/A' }
                        Ativo           = if ($m.Active) { "Sim" } else { "Não" }
                        'Serial Number' = $m.SerialNumberID
                        'Nome amigável' = $m.UserFriendlyName
                        Fabricante      = Convert-MonitorManufacturer $m.ManufacturerName
                        'Entradas de vídeo' = 'N/A'
                        'Porta Ativa'   = $portaAtiva
                    }

                    if ($MonitorMap.ContainsKey($obj.Nome)) {
                        $cfg = $MonitorMap[$obj.Nome]
                        $obj.Nome               = $cfg.Nome
                        $obj.Polegada           = $cfg.Polegada
                        $obj.'Resolução'        = $cfg.'Resolução'
                        $obj.Tamanho            = $cfg.Tamanho
                        $obj.'Entradas de vídeo' = $cfg.VideoPorts
                    }

                    $monitors += $obj
                }
            }
            catch {
                Write-Warning "Falha na coleta remota: $($_.Exception.Message)"
                $monitors = [PSCustomObject]@{
                    Nome             = 'Erro na coleta'
                    Polegada         = 'N/A'
                    'Resolução'      = 'N/A'
                    Tamanho          = 'N/A'
                    Ativo            = 'N/A'
                    'Serial Number'  = 'N/A'
                    'Nome amigável'  = 'N/A'
                    Fabricante       = 'N/A'
                    'Entradas de vídeo' = 'N/A'
                    'Porta Ativa'    = 'N/A'
                }
            }
        }
        
        ################################
        # 4.10) CONTROLADOR DE VÍDEO
        ################################
        Write-Host "  ✓ Informações da Placa de vídeo '$($ComputerName)'..." -ForegroundColor Cyan

        $videoRaw        = Get-CimInstance -ClassName Win32_VideoController -CimSession $CimSession |
                           Select-Object Name, VideoProcessor, AdapterCompatibility, AdapterRAM, DeviceID, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentNumberOfColors, DriverVersion, DriverDate
        $videoInfo       = $videoRaw | ForEach-Object {
            [PSCustomObject]@{
                'Nome'                           = $_.Name
                'Processador de Vídeo'           = $_.VideoProcessor
                'Adaptador de compatibilidade'   = $_.AdapterCompatibility
                'Capacidade (GB)'                = ("{0} GB" -f ([math]::Round($_.AdapterRAM / 1GB, 2)))
                'DeviceID'                       = $_.DeviceID
                'Resolução Horizontal'           = $_.CurrentHorizontalResolution
                'Resolução Vertical'             = $_.CurrentVerticalResolution
                'Número de Cores'                = $_.CurrentNumberOfColors
                'Versão do Driver'               = $_.DriverVersion
                'Data do Driver Instalado'       = ($_.DriverDate.ToString('dd/MM/yyyy HH:mm:ss'))
            }
        }        
        
        ################################
        # 4.11) Disco Rígido físico
        ################################
        Write-Host "  ✓ Informações do Disco Rígido físico '$($ComputerName)'..." -ForegroundColor Cyan

        # 1) Obter Win32_DiskDrive via CIM
        $diskDrives = Get-CimInstance -ClassName Win32_DiskDrive -CimSession $CimSession |
            Select-Object DeviceID, Model, InterfaceType, SerialNumber, @{
                Name       = "Size(GB)"
                Expression = {[math]::Round($_.Size / 1GB, 2)}
            }

        # 2) Obter Get-PhysicalDisk no remoto para pegar DeviceId (int) e MediaType (string)
        $physicalDisks = Invoke-Command -ComputerName $ComputerName `
            -Credential $Global:CredentialAdmin `
            -ScriptBlock {
                # Este comando roda no host remoto (PS 5.1) e retorna algo como:
                # DeviceId  MediaType
                # --------  ---------
                # 0         SSD
                # 1         HDD
                Get-PhysicalDisk | Select-Object DeviceId, MediaType
            }      

        # Função auxiliar inteligente para traduzir fabricante e modelo do disco
        function Get-ParsedDiskInfo {
            param([string]$RawModel)
            
            if ($null -eq $RawModel -or $RawModel.Trim().Length -eq 0) {
                return @{ Fabricante = 'N/D'; Modelo = 'N/D' }
            }

            $modelTrimmed = $RawModel.Trim()

            if ($null -ne $DiskModelMap -and $DiskModelMap.ContainsKey($modelTrimmed)) {
                return $DiskModelMap[$modelTrimmed]
            }

            # Lógica inteligente para extrair fabricante a partir de prefixos conhecidos
            $m = $modelTrimmed
            $fab = 'Desconhecido'

            if ($m -imatch '^SAMSUNG\s+(.*)') {
                $fab = 'Samsung'; $m = $Matches[1]
            } elseif ($m -imatch '^SKHynix[_\s]+(.*)') {
                $fab = 'SK Hynix'; $m = $Matches[1]
            } elseif ($m -imatch '^KINGSTON\s+(.*)') {
                $fab = 'Kingston'; $m = $Matches[1]
            } elseif ($m -imatch '^WDC\s+(.*)') {
                $fab = 'Western Digital'; $m = $Matches[1]
            } elseif ($m -imatch '^ST\d+(.*)') {
                $fab = 'Seagate';
            } elseif ($m -imatch '^CT\d+(.*)') {
                $fab = 'Crucial';
            } elseif ($m -imatch '^ADATA\s+(.*)') {
                $fab = 'ADATA'; $m = $Matches[1]
            } elseif ($m -imatch '^INTEL\s+(.*)') {
                $fab = 'Intel'; $m = $Matches[1]
            } elseif ($m -imatch '^KIOXIA\s+(.*)') {
                $fab = 'Kioxia'; $m = $Matches[1]
            } elseif ($m -imatch '^TOSHIBA\s+(.*)') {
                $fab = 'Toshiba'; $m = $Matches[1]
            } elseif ($m -imatch '^SanDisk\s+(.*)') {
                $fab = 'SanDisk'; $m = $Matches[1]
            } elseif ($m -imatch '^CORSAIR\s+(.*)') {
                $fab = 'Corsair'; $m = $Matches[1]
            } elseif ($m -imatch '^CRUCIAL\s+(.*)') {
                $fab = 'Crucial'; $m = $Matches[1]
            } elseif ($m -imatch '^PATRIOT\s+(.*)') {
                $fab = 'Patriot'; $m = $Matches[1]
            }

            return @{ Fabricante = $fab; Modelo = $m }
        }

        # 3) Montar $diskInfo usando regex para extrair o número do disco
        $diskInfo = foreach ($d in $diskDrives) {
            # Tenta extrair um número após 'PHYSICALDRIVE', ex.: "\\.\PHYSICALDRIVE0" -> "0"
            if ($d.DeviceID -match 'PHYSICALDRIVE(\d+)$') {
                $driveNumber = [int]$Matches[1]
                # Procurar na coleção de Get-PhysicalDisk
                $match = $physicalDisks | Where-Object { $_.DeviceId -eq $driveNumber }
                $mediaTypeText = if ($match) { $match.MediaType } else { "Unspecified" }
            } else {
                # Se não conseguir extrair número, marcamos como Unspecified
                $driveNumber    = $null
                $mediaTypeText  = "Unspecified"
            }

            $parsedDisk = Get-ParsedDiskInfo $d.Model

            [PSCustomObject]@{
                'Fabricante'        = $parsedDisk.Fabricante
                'Modelo'            = $parsedDisk.Modelo
                'Tipo de interface' = $d.InterfaceType
                'Tamanho em GB'     = if ($null -ne $d.'Size(GB)') { "$($d.'Size(GB)') GB" } else { 'N/D' }
                'Tipo de mídia'     = $mediaTypeText
                'Número de Série'   = if ($d.SerialNumber) { $d.SerialNumber.Trim() } else { 'N/D' }
            }
        }

        ################################
        # 4.12) Disco Rígido lógico
        ################################
        Write-Host "  ✓ Informações do Disco Rígido lógico '$($ComputerName)'..." -ForegroundColor Cyan

        # DISCO LÓGICO (Win32_LogicalDisk)
        $logicalRaw      = Get-CimInstance -ClassName Win32_LogicalDisk -CimSession $CimSession |
                           Select-Object DeviceID, FreeSpace, Size, VolumeName, FileSystem
        $logicalInfo     = $logicalRaw | ForEach-Object {
            $freeGB     = [math]::Round($_.FreeSpace / 1GB, 2)
            $totalGB    = [math]::Round($_.Size / 1GB, 2)
            $usedGB     = [math]::Round($totalGB - $freeGB, 2)
            $volName    = if ([string]::IsNullOrWhiteSpace($_.VolumeName)) { "Windows" } else { $_.VolumeName }
            [PSCustomObject]@{
                'DeviceID'         = $_.DeviceID
                'Espaço Livre'     = "$( $freeGB ) GB"
                'Espaço Usado'     = "$( $usedGB ) GB"
                'Espaço Total'     = "$( $totalGB ) GB"
                'Nome do Volume'   = $volName
                'Sistema de Arquivos' = $_.FileSystem
            }
        }        

        ################################
        # 4.13) Rede
        ################################
        Write-Host "  ✓ Informações da Rede '$($ComputerName)'..." -ForegroundColor Cyan

        $networkInfo    = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -CimSession $CimSession |
                          Select-Object Description, MACAddress, @{Name="IPAddresses";Expression={$_.IPAddress -join ", "}}, @{Name="Gateways";Expression={$_.DefaultIPGateway -join ", "}}

        $networkInfo = [PSCustomObject]@{
            'Nome da placa'            = $networkInfo.Description
            'MAC Address'              = $networkInfo.MACAddress
            'Endereços IPs'            = $networkInfo.IPAddresses
            'Gateway'                  = $networkInfo.Gateways
        }

        ################################
        # 4.14) Som
        ################################
        Write-Host "  ✓ Informações do Som '$($ComputerName)'..." -ForegroundColor Cyan

        $soundRaw        = Get-CimInstance -ClassName Win32_SoundDevice -CimSession $CimSession |
                           Select-Object Name, Manufacturer, Status, StatusInfo
        $soundInfo       = $soundRaw | ForEach-Object {
            switch ($_.StatusInfo) {
                1 { $stat = "Outro" }
                2 { $stat = "Desconhecido" }
                3 { $stat = "Habilitado" }
                4 { $stat = "Desabilitado" }
                5 { $stat = "Não aplicável" }
                default { $stat = $_.StatusInfo }
            }
            [PSCustomObject]@{
                'Nome'                 = $_.Name
                'Fabricante'           = Get-FriendlyManufacturer $_.Manufacturer
                'Status'               = $_.Status
                'Informação do Status' = $stat
            }
        }

        #########################
        # 4.15) IMPRESSORAS LOCAIS
        #########################
        Write-Host "  ✓ Informações das Impressoras locais '$($ComputerName)'..." -ForegroundColor Cyan

        $printersLocalRaw = Get-CimInstance -ClassName Win32_Printer -CimSession $CimSession |
                            Select-Object Name, PrinterState, PrinterStatus, ShareName, PortName
        $printersInfo     = $printersLocalRaw | ForEach-Object {
            $statusPrint = switch ($_.PrinterStatus) {
                1  { "Outro Status" }
                2  { "Desconhecido" }
                3  { "Parada" }
                4  { "Imprimindo" }
                5  { "Está em aquecimento" }
                6  { "Impressão interrompida" }
                7  { "Offline" }
                default { $_.PrinterStatus }
            }
            $statePrint = switch ($_.PrinterState) {
                0  { "Parada" }
                1  { "Pausada" }
                2  { "Erro" }
                3  { "Pendente de deletar" }
                4  { "Atolamento de papel" }
                5  { "Sem papel" }
                6  { "Alimentação manual" }
                7  { "Problema de papel" }
                8  { "Offline" }
                9  { "Entrada/Saída ativa" }
                10 { "Ocupada" }
                11 { "Imprimindo" }
                12 { "" }
                default { $_.PrinterState }
            }
            [PSCustomObject]@{
                'Nome'                     = $_.Name
                'Estado da impressão'      = $statePrint
                'Status da impressão'      = $statusPrint
                'Nome de compartilhamento' = $_.ShareName
                'Porta'                    = $_.PortName
            }
        }

        #########################
        # 4.16) IMPRESSORAS MAPEADAS (por usuário)
        #########################
        Write-Host "  ✓ Informações das Impressoras mapeadas '$($ComputerName)'..." -ForegroundColor Cyan

        $mappedPrinters = Get-UserPrinter -ComputerName $ComputerName -Credential $Global:CredentialAdmin

        #########################
        # 4.17) INSTALLED DRIVERS (ou array vazio se SkipMajorData)
        #########################
        
        if (-not $SkipMajorData.IsPresent) {
            Write-Host "  ✓ Informações dos drivers instalados '$($ComputerName)'..." -ForegroundColor Cyan

            $rawDrivers = Get-CimInstance -ClassName Win32_PnPSignedDriver -CimSession $CimSession
            $driversInfo = foreach ($drv in $rawDrivers) {
                if ($null -eq $drv) { continue }
                
                # Trata a data do driver se existir
                $drvDate = if ($drv.DriverDate) { $drv.DriverDate.ToString('dd/MM/yyyy') } else { "N/D" }
                $signedStr = if ($drv.IsSigned) { "Sim" } else { "Não" }

                [PSCustomObject]@{
                    "Nome do dispositivo" = $drv.DeviceName
                    "Fabricante"          = $drv.Manufacturer
                    "Versão do driver"    = $drv.DriverVersion
                    "Data do driver"      = $drvDate
                    "Arquivo INF"         = $drv.InfName
                    "Assinado"            = $signedStr
                }
            }
        } else {
            $driversInfo = @()
        }

        #########################
        # 4.18) PROGRAMAS INSTALADOS (ou array vazio se SkipMajorData)
        #########################

        if (-not $SkipMajorData.IsPresent) {
            Write-Host "  ✓ Informações dos programas instalados '$($ComputerName)'..." -ForegroundColor Cyan
            
            $regPaths = @(
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
            )
            $rawPrograms = foreach ($path in $regPaths) {
                Invoke-Command -ComputerName $ComputerName -Credential $Global:CredentialAdmin -ScriptBlock {
                    param($p)
                    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
                      Where-Object { $_.DisplayName -and $_.DisplayName.Trim().Length -gt 0 } |
                      Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
                } -ArgumentList $path
            }

            $installedPrograms = foreach ($prog in $rawPrograms) {
                if ($null -eq $prog) { continue }
                
                # Trata e formata a data no padrão brasileiro dd/MM/yyyy
                $rawDate = ($prog.InstallDate -as [string]).Trim()
                $formattedDate = ""
                if ($rawDate -match '^\d{8}$') {
                    # Formato YYYYMMDD (ex: 20260313 -> 13/03/2026)
                    $formattedDate = "$($rawDate.Substring(6,2))/$($rawDate.Substring(4,2))/$($rawDate.Substring(0,4))"
                } elseif ($rawDate -match '^(\d{4})-(\d{2})-(\d{2})') {
                    # Formato YYYY-MM-DD
                    $formattedDate = "$($Matches[3])/$($Matches[2])/$($Matches[1])"
                } else {
                    $formattedDate = $rawDate
                }

                [PSCustomObject]@{
                    "Nome do programa"   = $prog.DisplayName
                    "Versão"             = $prog.DisplayVersion
                    "Fabricante"         = $prog.Publisher
                    "Data da instalação" = $formattedDate
                }
            }
        } else {
            $installedPrograms = @()
        }

        #########################
        # 4.19) SERVIÇOS EM EXECUÇÃO (ou array vazio se SkipMajorData)
        #########################

        if (-not $SkipMajorData.IsPresent) {
            Write-Host "  ✓ Informações dos serviços executados '$($ComputerName)'..." -ForegroundColor Cyan
            
            $rawServices = Get-CimInstance -ClassName Win32_Service -Filter "State='Running'" -CimSession $CimSession
            $runningServices = foreach ($srv in $rawServices) {
                if ($null -eq $srv) { continue }
                
                # Traduz Status e Tipo de Inicialização
                $statusTranslated = switch ($srv.State) {
                    'Running' { 'Em execução' }
                    'Paused'  { 'Pausado' }
                    'Stopped' { 'Parado' }
                    default   { $srv.State }
                }
                $startModeTranslated = switch ($srv.StartMode) {
                    'Auto'      { 'Automático' }
                    'Manual'    { 'Manual' }
                    'Disabled'  { 'Desativado' }
                    'System'    { 'Sistema' }
                    'Boot'      { 'Boot' }
                    default     { $srv.StartMode }
                }

                [PSCustomObject]@{
                    "Nome técnico"          = $srv.Name
                    "Nome de exibição"      = $srv.DisplayName
                    "Status"                = $statusTranslated
                    "Tipo de inicialização" = $startModeTranslated
                    "Conta de logon"        = $srv.StartName
                    "Caminho do executável" = $srv.PathName
                }
            }
        } else {
            $runningServices = @()
        }

        #########################
        # 4.20) PROCESSOS ATIVOS (ou array vazio se SkipMajorData)
        #########################

        if (-not $SkipMajorData.IsPresent) {
            Write-Host "  ✓ Informações dos processos ativos '$($ComputerName)'..." -ForegroundColor Cyan
            
            $rawProcesses = Invoke-Command -ComputerName $ComputerName -Credential $Global:CredentialAdmin -ScriptBlock {
                Get-CimInstance -ClassName Win32_Process | ForEach-Object {
                    $ownerRes = $_ | Invoke-CimMethod -MethodName GetOwner -ErrorAction SilentlyContinue
                    $owner = if ($ownerRes.User) { "$($ownerRes.Domain)\$($ownerRes.User)" } else { "Sistema" }
                    [PSCustomObject]@{
                        ProcessId      = $_.ProcessId
                        Name           = $_.Name
                        WorkingSetSize = $_.WorkingSetSize
                        ThreadCount    = $_.ThreadCount
                        HandleCount    = $_.HandleCount
                        ExecutablePath = $_.ExecutablePath
                        CommandLine    = $_.CommandLine
                        Owner          = $owner
                    }
                }
            }

            $activeProcesses = foreach ($proc in $rawProcesses) {
                if ($null -eq $proc) { continue }
                
                # Calcula o Uso de Memória em MB
                $memMB = if ($proc.WorkingSetSize) { [math]::Round($proc.WorkingSetSize / 1MB, 2) } else { 0.00 }

                [PSCustomObject]@{
                    "PID"                = $proc.ProcessId
                    "Nome do executável" = $proc.Name
                    "Usuário"            = $proc.Owner
                    "Uso de Memória (MB)"= $memMB
                    "Threads"            = $proc.ThreadCount
                    "Handles"            = $proc.HandleCount
                    "Caminho do arquivo" = $proc.ExecutablePath
                    "Linha de comando"   = $proc.CommandLine
                }
            }
        } else {
            $activeProcesses = @()
        }

        #########################
        # 4.21) LAST BOOTTIME
        #########################
        Write-Host "  ✓ Informações do último boot feito '$($ComputerName)'..." -ForegroundColor Cyan

        $lastBoot = $wmiOS.LastBootUpTime.ToString('dd/MM/yyyy HH:mm:ss')

        # Construir o objeto final que agrega tudo
        return [PSCustomObject]@{
            ComputerName            = $ComputerName
            HardwareInfo            = $hwObject
            CurrentUser             = $simpleUser
            CurrentUserDisplayName  = $displayName
            CurrentUserDomain       = $loggedDomain
            OS                      = $osObject
            BIOS                    = $biosInfo
            Processor               = $cpuInfo
            Motherboard             = $motherboardInfo
            Chassis                 = $chassisInfo
            MemoryModules           = $memoryInfo
            DiskDrives              = $diskInfo
            LogicalDisks            = $logicalInfo
            NetworkConfigs          = $networkInfo
            Monitors                = $monitors
            VideoController         = $videoInfo
            SoundDevices            = $soundInfo
            PrintersLocal           = $printersInfo
            PrintersMapped          = $mappedPrinters            
            InstalledDrivers        = $driversInfo
            InstalledPrograms       = $installedPrograms
            RunningServices         = $runningServices
            ActiveProcesses         = $activeProcesses
            LastBootTime            = $lastBoot
        }
    } catch {
        Write-Error "Falha ao coletar dados principais via CIM/Invoke-Command: $_"
        return $null
    }
}

# ----------------------------------------
# 5) COLETA DE DADOS NO CLIENTE
# ----------------------------------------

Write-Host "Iniciando coleta para '$ComputerName'..." -ForegroundColor Cyan

$systemData = Get-AllSystemData -CimSession $Global:CimSession -ComputerName $ComputerName -SkipMajorData:$SkipMajorData.IsPresent
if (-not $systemData) {
    Write-Error "Não foi possível coletar dados principais. Abortando."
    Remove-CimSession -CimSession $Global:CimSession
    exit 4
}

#########################
# 6) USUÁRIOS
#########################
Write-Host "  ✓ Informações dos usuários '$($ComputerName)'..." -ForegroundColor Cyan
$usersData = Get-Users-Data -ComputerName $ComputerName -Credential $Global:CredentialAdmin

# Agora sim removemos a sessão (DCOM ou WS-MAN)
Remove-CimSession -CimSession $Global:CimSession

# ----------------------------------------
# 7) GERAÇÃO DO RELATÓRIO HTML
# ----------------------------------------

# Função auxiliar: converte objeto ou coleção em fragmento HTML (tabela)
function Convert-ToHtmlFragment {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Object[]]$InputObject
    )
    begin {
        $allObjects = @()
    }
    process {
        $allObjects += $InputObject
    }
    end {
        $fragment = $allObjects | ConvertTo-Html -Fragment
        return "<div class='table-container'>$($fragment -join '')</div>"
    }
}

# Para montar o nome do arquivo de saída, primeiro sanitizamos quaisquer caracteres inválidos:
$rawName = "$($systemData.HardwareInfo.Computador)_$($systemData.OS.'Nome da Máquina')"
$safeName = $rawName -replace '[\\/:\*\?"<>|\s]', '_'
$outFile = Join-Path $OutputFolder ("Report_{0}_{1:yyyyMMdd_HHmmss}.html" -f $safeName, (Get-Date))

# -----------------------------------------------
# Montagem do HTML via módulo externo GeradorHtml.ps1
# -----------------------------------------------
$htmlContent = Get-DeviceReportHtml -systemData $systemData -usersData $usersData -SkipMajorData:$SkipMajorData

# Gravar o HTML no arquivo
try {
    $htmlContent | Out-File -FilePath $outFile -Encoding UTF8
    Write-Host "RELATORIO_HTML_PATH: $outFile" -ForegroundColor Green
    Write-Host "Relatório salvo em: $outFile" -ForegroundColor Green
} catch {
    Write-Error "Falha ao escrever arquivo HTML em '$outFile': $_"
    exit 5
}

# Abre automaticamente o relatório no navegador padrão
try {
    Start-Process "$outFile"
} catch {
    try {
        Start-Process explorer.exe "`"$outFile`""
    } catch {
        Write-Warning "Não foi possível abrir o navegador automaticamente: $_"
    }
}

# ----------------------------------------
# 8) GERAÇÃO DO RELATÓRIO PDF
# ----------------------------------------

# 2) Gera PDF via Edge headless
$pdfFile = [System.IO.Path]::ChangeExtension($outFile, '.pdf')
$path64 = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
$path86 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

if (Test-Path $path64) {
    $edgeExe = $path64
} elseif (Test-Path $path86) {
    $edgeExe = $path86
} else {
    $edgeExe = $null
}

if ($edgeExe) {
    # Executa o Edge headless redirecionando erros internos de telemetria/renderers do Chromium para nul
    Start-Process -FilePath $edgeExe -ArgumentList @(
        "--headless=new"
        "--disable-gpu"
        "--no-first-run"
        "--no-default-browser-check"
        "--disable-extensions"
        "--disable-logging"
        "--log-level=3"
        "--print-to-pdf=`"$pdfFile`""
        "`"$outFile`""
    ) -NoNewWindow -Wait -RedirectStandardError "$env:TEMP\edge_headless_err.log" -ErrorAction SilentlyContinue

    Write-Host "RELATORIO_PDF_PATH: $pdfFile" -ForegroundColor Green
    Write-Host "PDF gerado em: $pdfFile" -ForegroundColor Green
} else {
    Write-Warning "msedge.exe não encontrado nos caminhos padrão; pulando geração de PDF."
}

# ----------------------------------------
# 9) GERAÇÃO DO EXCEL em UMA ÚNICA ABA via ImportExcel
# ----------------------------------------
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Warning "Módulo ImportExcel não está instalado. Para gerar Excel, execute: Install-Module ImportExcel"
} else {
    $excelFile = Join-Path $OutputFolder ("Report_{0}_{1:yyyyMMdd_HHmmss}.xlsx" -f $safeName, (Get-Date))
    $sheetName = 'Relatorio'

    # Começa em row 1
    $row = 1

    # 1) HardwareInfo (objeto único)
    $systemData.HardwareInfo | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize -ClearSheet
    # Avança linhas: número de propriedades + 2 (uma para cabeçalho, uma para espaço)
    $row += ($systemData.HardwareInfo | Get-Member -MemberType NoteProperty).Count + 2

    # 2) UsuárioAtual (objeto único)
    [PSCustomObject]@{
        'Usuário Logado agora' = $systemData.CurrentUser
        'DisplayName AD'       = $systemData.CurrentUserDisplayName
        'Domínio'              = $systemData.CurrentUserDomain
        'Nome do computador'   = $systemData.HardwareInfo.Computador
    } | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += 5  # 4 colunas + espaço

    # 3) Sistema Operacional
    $systemData.OS | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.OS | Get-Member -MemberType NoteProperty).Count + 2

    # 4) BIOS
    $systemData.BIOS | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.BIOS | Measure-Object).Count + 2

    # 5) Processador
    $systemData.Processor | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.Processor | Get-Member -MemberType NoteProperty).Count + 2

    # 6) Placa-Mãe
    $systemData.Motherboard | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.Motherboard | Get-Member -MemberType NoteProperty).Count + 2

    # 7) Chassi
    $systemData.Chassis | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.Chassis | Get-Member -MemberType NoteProperty).Count + 2

    # 8) Memória
    $systemData.MemoryModules | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.MemoryModules.Count + 1) + 2

    # 9) Discos Físicos
    $systemData.DiskDrives | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.DiskDrives.Count + 1) + 2

    # 10) Discos Lógicos
    $systemData.LogicalDisks | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.LogicalDisks.Count + 1) + 2

    # 11) Rede
    $systemData.NetworkConfigs | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.NetworkConfigs.Count + 1) + 2

    # 12) Controlador de Vídeo
    $systemData.VideoController | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.VideoController.Count + 1) + 2

    # 13) Informações de Som
    $systemData.SoundDevices | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.SoundDevices.Count + 1) + 2

    # 14) Impressoras Locais
    $systemData.PrintersLocal | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.PrintersLocal.Count + 1) + 2

    # 15) Impressoras Mapeadas
    $systemData.PrintersMapped | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.PrintersMapped.Count + 1) + 2

    # 16) Monitores
    $systemData.Monitors | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
    $row += ($systemData.Monitors.Count + 1) + 2

    # 17) Usuários na Máquina (se não pular)
    if (-not $SkipMajorData) {
        $usersData | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
        $row += ($usersData.Count + 1) + 2
    }

    # 18) Se não pular, adiciona Programas, Serviços, Processos e Drivers
    if (-not $SkipMajorData) {
        # Programas
        $systemData.InstalledPrograms | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
        $row += ($systemData.InstalledPrograms.Count + 1) + 2
        # Serviços
        $systemData.RunningServices | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
        $row += ($systemData.RunningServices.Count + 1) + 2
        # Processos
        $systemData.ActiveProcesses | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
        $row += ($systemData.ActiveProcesses.Count + 1) + 2
        # Drivers
        $systemData.InstalledDrivers | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow $row -AutoSize
        $row += ($systemData.InstalledDrivers.Count + 1) + 2
    }

    Write-Host "RELATORIO_EXCEL_PATH: $excelFile" -ForegroundColor Green
    Write-Host "Arquivo Excel gerado em: $excelFile" -ForegroundColor Green
}

Write-Host "Coleta concluída com sucesso!" -ForegroundColor Cyan
