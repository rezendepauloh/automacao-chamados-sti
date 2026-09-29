param(
    [string]$SiteServer = "srv-1046.in.mpe.ms.gov.br",
    [string]$SiteCode = "PGJ",
    [string]$Username = "MPE\paulo_admin",
    [string]$Password = "Abcd4268#",
    [string]$OutputFile = ""
)

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   SISTEMA BANCADA — SINCRONIZADOR DE INVENTARIO SCCM" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$sec = ConvertTo-SecureString $Password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($Username, $sec)

Write-Host " [1/4] Consultando Colecoes de Dispositivos e Usuarios..." -ForegroundColor Yellow
$colQuery = "SELECT CollectionID, Name, CollectionType, MemberCount, Comment, LastRefreshTime FROM SMS_Collection"
$collections = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $colQuery -Credential $cred -Authentication PacketPrivacy
Write-Host "       -> $($collections.Count) colecoes encontradas." -ForegroundColor Green

Write-Host " [2/4] Consultando Dispositivos do SCCM (SMS_R_System)..." -ForegroundColor Yellow
$devQuery = "SELECT ResourceID, Name, LastLogonUserName, IPAddresses, MACAddresses, OperatingSystemNameandVersion, Build, ClientVersion, Active, ADSiteName, DistinguishedName FROM SMS_R_System"
$devices = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $devQuery -Credential $cred -Authentication PacketPrivacy
Write-Host "       -> $($devices.Count) estacoes encontradas." -ForegroundColor Green

Write-Host " [3/4] Consultando Inventario de Hardware (Computador, CPU, RAM, Discos)..." -ForegroundColor Yellow
$compSys = @{}
try {
    $csQuery = "SELECT ResourceID, Manufacturer, Model, TotalPhysicalMemory FROM SMS_G_System_COMPUTER_SYSTEM"
    Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $csQuery -Credential $cred -Authentication PacketPrivacy -ErrorAction SilentlyContinue | ForEach-Object {
        $compSys[$_.ResourceID] = $_
    }
} catch {}

$procs = @{}
try {
    $procQuery = "SELECT ResourceID, Name FROM SMS_G_System_PROCESSOR"
    Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $procQuery -Credential $cred -Authentication PacketPrivacy -ErrorAction SilentlyContinue | ForEach-Object {
        if (-not $procs.ContainsKey($_.ResourceID)) {
            $procs[$_.ResourceID] = $_.Name
        }
    }
} catch {}

$disks = @{}
try {
    $diskQuery = "SELECT ResourceID, DeviceID, Size, FreeSpace FROM SMS_G_System_LOGICAL_DISK WHERE DriveType = 3"
    Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $diskQuery -Credential $cred -Authentication PacketPrivacy -ErrorAction SilentlyContinue | ForEach-Object {
        $freeGB = [math]::Round($_.FreeSpace / 1024, 1)
        $totalGB = [math]::Round($_.Size / 1024, 1)
        $dInfo = "$($_.DeviceID) ${freeGB}GB livres de ${totalGB}GB"
        if ($disks.ContainsKey($_.ResourceID)) {
            $disks[$_.ResourceID] += "; $dInfo"
        } else {
            $disks[$_.ResourceID] = $dInfo
        }
    }
} catch {}
Write-Host "       -> Inventario de componentes de hardware correlacionado." -ForegroundColor Green

Write-Host " [4/4] Consultando Usuarios do SCCM (SMS_R_User)..." -ForegroundColor Yellow
$userQuery = "SELECT ResourceID, UserName, FullUserName, WindowsNTDomain, DistinguishedName FROM SMS_R_User"
$users = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $userQuery -Credential $cred -Authentication PacketPrivacy
Write-Host "       -> $($users.Count) usuarios encontrados." -ForegroundColor Green

$enrichedDevices = @($devices | ForEach-Object {
    $rId = $_.ResourceID
    $cs = $compSys[$rId]
    $cpu = $procs[$rId]
    $dsk = $disks[$rId]

    $mfg = if ($cs -and $cs.Manufacturer) { $cs.Manufacturer.Trim() } else { "" }
    $mod = if ($cs -and $cs.Model) { $cs.Model.Trim() } else { "" }
    $ramStr = if ($cs -and $cs.TotalPhysicalMemory) {
        $gb = [math]::Round($cs.TotalPhysicalMemory / 1048576, 0)
        "${gb} GB"
    } else { "" }

    [PSCustomObject]@{
        ResourceID = $_.ResourceID
        Name = $_.Name
        LastLogonUserName = $_.LastLogonUserName
        IPAddresses = $_.IPAddresses
        MACAddresses = $_.MACAddresses
        Manufacturer = $mfg
        Model = $mod
        MemoryRAM = $ramStr
        Processor = if ($cpu) { $cpu.Trim() } else { "" }
        DiskDrives = if ($dsk) { $dsk.Trim() } else { "" }
        OperatingSystemNameandVersion = $_.OperatingSystemNameandVersion
        Build = $_.Build
        ClientVersion = $_.ClientVersion
        Active = $_.Active
        ADSiteName = $_.ADSiteName
        DistinguishedName = $_.DistinguishedName
    }
})

$exportData = @{
    generated_at = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    collections = @($collections | Select-Object CollectionID, Name, CollectionType, MemberCount, Comment, LastRefreshTime)
    devices = $enrichedDevices
    users = @($users | Select-Object ResourceID, UserName, FullUserName, WindowsNTDomain, DistinguishedName)
}

$json = $exportData | ConvertTo-Json -Depth 4 -Compress

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $OutputFile = Join-Path $env:USERPROFILE "sccm_inventory.json"
}

[System.IO.File]::WriteAllText($OutputFile, $json, [System.Text.Encoding]::UTF8)
Write-Host ""
Write-Host " [OK] Inventario SCCM exportado com sucesso para: $OutputFile" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
