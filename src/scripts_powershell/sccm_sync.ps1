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

Write-Host " [1/3] Consultando Colecoes de Dispositivos e Usuarios..." -ForegroundColor Yellow
$colQuery = "SELECT CollectionID, Name, CollectionType, MemberCount, Comment, LastRefreshTime FROM SMS_Collection"
$collections = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $colQuery -Credential $cred -Authentication PacketPrivacy
Write-Host "       -> $($collections.Count) colecoes encontradas." -ForegroundColor Green

Write-Host " [2/3] Consultando Dispositivos do SCCM (SMS_R_System)..." -ForegroundColor Yellow
$devQuery = "SELECT ResourceID, Name, LastLogonUserName, IPAddresses, MACAddresses, OperatingSystemNameandVersion, Build, ClientVersion, Active, ADSiteName, DistinguishedName FROM SMS_R_System"
$devices = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $devQuery -Credential $cred -Authentication PacketPrivacy
Write-Host "       -> $($devices.Count) estacoes encontradas." -ForegroundColor Green

Write-Host " [3/3] Consultando Usuarios do SCCM (SMS_R_User)..." -ForegroundColor Yellow
$userQuery = "SELECT ResourceID, UserName, FullUserName, WindowsNTDomain, DistinguishedName FROM SMS_R_User"
$users = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Query $userQuery -Credential $cred -Authentication PacketPrivacy
Write-Host "       -> $($users.Count) usuarios encontrados." -ForegroundColor Green

$exportData = @{
    generated_at = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    collections = @($collections | Select-Object CollectionID, Name, CollectionType, MemberCount, Comment, LastRefreshTime)
    devices = @($devices | Select-Object ResourceID, Name, LastLogonUserName, IPAddresses, MACAddresses, OperatingSystemNameandVersion, Build, ClientVersion, Active, ADSiteName, DistinguishedName)
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
