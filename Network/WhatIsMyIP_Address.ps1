<# 
.SYNOPSIS
    Outil de diagnostic Helpdesk permettant d'afficher rapidement les principales informations d'un poste Windows.
.DESCRIPTION
    Ce script PowerShell collecte et affiche les informations essentielles d'un poste de travail afin de faciliter les opérations de support à distance, de diagnostic et d'inventaire
#>

function Format-GB($bytes) {
    if (-not $bytes) { return "n/a" }
    "{0:N0}GB" -f ($bytes / 1GB)
}
function Write-Separator { Write-Host ("=" * 62) -ForegroundColor Gray }
function Write-Section($title) { Write-Host ("{0,-13}: " -f $title) -NoNewline -ForegroundColor Gray }

# --- IP ---
$ipv4 = Get-NetIPAddress -AddressFamily IPv4 -PrefixOrigin Dhcp,Manual -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notmatch '^169\.254\.' -and $_.IPAddress -ne '127.0.0.1' } |
        Sort-Object InterfaceMetric, SkipAsSource |
        Select-Object -ExpandProperty IPAddress
$ip1 = $ipv4[0]; if (-not $ip1) { $ip1 = 'n/a' }
$ip2 = $ipv4[1]; if (-not $ip2) { $ip2 = 'n/a' }

# --- Nom PC ---
$compName = $env:COMPUTERNAME

# --- Matériel ---
$cs   = Get-CimInstance -ClassName Win32_ComputerSystem
$bios = Get-CimInstance -ClassName Win32_BIOS
$cpu  = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
$memGB = [int]($cs.TotalPhysicalMemory / 1GB)

$enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue
$assetTag  = ($enclosure.SMBIOSAssetTag | Select-Object -First 1)
if ([string]::IsNullOrWhiteSpace($assetTag)) { $assetTag = $bios.SerialNumber }

$pd = $null
try { $pd = Get-PhysicalDisk | Sort-Object FriendlyName | Select-Object -First 1 } catch {}
if ($pd) {
    $diskModel = $pd.FriendlyName
    $diskSN    = if ($pd.SerialNumber) { $pd.SerialNumber } else { "n/a" }
    $diskSize  = if ($pd.Size) { "{0:N0}GB" -f ($pd.Size / 1GB) } else { "n/a" }
} else {
    $dd = Get-CimInstance -ClassName Win32_DiskDrive | Sort-Object Index | Select-Object -First 1
    $diskModel = $dd.Model
    $diskSN    = if ($dd.SerialNumber) { $dd.SerialNumber } else { "n/a" }
    $diskSize  = Format-GB $dd.Size
}

# --- Logiciel / OS ---
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$caption = $os.Caption
$ver     = $os.Version
$arch    = $os.OSArchitecture
$install = $os.InstallDate
$boot    = $os.LastBootUpTime

try {
    $prod = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion
} catch { $prod = $null }

if ([string]::IsNullOrWhiteSpace($prod)) { $displayVersion = $ver } else { $displayVersion = $prod }

$uptime = (Get-Date) - $boot
$uptimeStr = ("{0} day(s), {1} hour(s), {2} minute(s), {3} second(s)." -f `
              [int]$uptime.Days, [int]$uptime.Hours, [int]$uptime.Minutes, [int]$uptime.Seconds)

# --- Affichage ---
Write-Separator
Write-Host ("  My IP Address1 is : ") -NoNewline -ForegroundColor Gray; Write-Host ($ip1) -ForegroundColor Yellow
Write-Host ("  My IP Address2 is : ") -NoNewline -ForegroundColor Gray; Write-Host ($ip2) -ForegroundColor Yellow
Write-Host ""
Write-Host ("  My ComputerName is : ") -NoNewline -ForegroundColor Gray; Write-Host ($compName) -ForegroundColor Cyan
Write-Separator

Write-Host ("Hardware    make|model : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} | {1}" -f $cs.Manufacturer, $cs.Model)
Write-Host ("           tag|bios : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} | {1}" -f $assetTag, $bios.SMBIOSBIOSVersion) -ForegroundColor Red
Write-Host ("           memory : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0}gb" -f $memGB)
Write-Host ("           cpu : ") -NoNewline -ForegroundColor Gray; Write-Host ($cpu.Name -replace '\s+', ' ')
Write-Host ("           disk1 : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} - {1} - s/n: {2}" -f $diskSize, $diskModel, $diskSN)

Write-Host ("-" * 62) -ForegroundColor Gray
Write-Section "Software"; Write-Host ""
Write-Host ("           OS : ") -NoNewline -ForegroundColor Gray
$osLine = "{0}, Version {1}, {2}" -f $caption, $displayVersion, $arch
Write-Host $osLine -ForegroundColor Green

Write-Host ("   install date : ") -NoNewline -ForegroundColor Gray; Write-Host ($install.ToString("yyyy/MM/dd HH:mm"))
Write-Host ("        uptime : ") -NoNewline -ForegroundColor Gray; Write-Host ($uptimeStr)

Write-Separator
Write-Host "Press Enter to close" -ForegroundColor Gray
[Console]::ReadLine() | Out-Null j'aimerais rajouter sa dans mon post 
Aurel-AD/Scripts-Administration-PowerShell
<#  WhatIsMyIP_Address.ps1 (fix compatibilité PS 5.1)
#>

function Format-GB($bytes) {
    if (-not $bytes) { return "n/a" }
    "{0:N0}GB" -f ($bytes / 1GB)
}
function Write-Separator { Write-Host ("=" * 62) -ForegroundColor Gray }
function Write-Section($title) { Write-Host ("{0,-13}: " -f $title) -NoNewline -ForegroundColor Gray }

# --- IP ---
$ipv4 = Get-NetIPAddress -AddressFamily IPv4 -PrefixOrigin Dhcp,Manual -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notmatch '^169\.254\.' -and $_.IPAddress -ne '127.0.0.1' } |
        Sort-Object InterfaceMetric, SkipAsSource |
        Select-Object -ExpandProperty IPAddress
$ip1 = $ipv4[0]; if (-not $ip1) { $ip1 = 'n/a' }
$ip2 = $ipv4[1]; if (-not $ip2) { $ip2 = 'n/a' }

# --- Nom PC ---
$compName = $env:COMPUTERNAME

# --- Matériel ---
$cs   = Get-CimInstance -ClassName Win32_ComputerSystem
$bios = Get-CimInstance -ClassName Win32_BIOS
$cpu  = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
$memGB = [int]($cs.TotalPhysicalMemory / 1GB)

$enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue
$assetTag  = ($enclosure.SMBIOSAssetTag | Select-Object -First 1)
if ([string]::IsNullOrWhiteSpace($assetTag)) { $assetTag = $bios.SerialNumber }

$pd = $null
try { $pd = Get-PhysicalDisk | Sort-Object FriendlyName | Select-Object -First 1 } catch {}
if ($pd) {
    $diskModel = $pd.FriendlyName
    $diskSN    = if ($pd.SerialNumber) { $pd.SerialNumber } else { "n/a" }
    $diskSize  = if ($pd.Size) { "{0:N0}GB" -f ($pd.Size / 1GB) } else { "n/a" }
} else {
    $dd = Get-CimInstance -ClassName Win32_DiskDrive | Sort-Object Index | Select-Object -First 1
    $diskModel = $dd.Model
    $diskSN    = if ($dd.SerialNumber) { $dd.SerialNumber } else { "n/a" }
    $diskSize  = Format-GB $dd.Size
}

# --- Logiciel / OS ---
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$caption = $os.Caption
$ver     = $os.Version
$arch    = $os.OSArchitecture
$install = $os.InstallDate
$boot    = $os.LastBootUpTime

try {
    $prod = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion
} catch { $prod = $null }

if ([string]::IsNullOrWhiteSpace($prod)) { $displayVersion = $ver } else { $displayVersion = $prod }

$uptime = (Get-Date) - $boot
$uptimeStr = ("{0} day(s), {1} hour(s), {2} minute(s), {3} second(s)." -f `
              [int]$uptime.Days, [int]$uptime.Hours, [int]$uptime.Minutes, [int]$uptime.Seconds)

# --- Affichage ---
Write-Separator
Write-Host ("  My IP Address1 is : ") -NoNewline -ForegroundColor Gray; Write-Host ($ip1) -ForegroundColor Yellow
Write-Host ("  My IP Address2 is : ") -NoNewline -ForegroundColor Gray; Write-Host ($ip2) -ForegroundColor Yellow
Write-Host ""
Write-Host ("  My ComputerName is : ") -NoNewline -ForegroundColor Gray; Write-Host ($compName) -ForegroundColor Cyan
Write-Separator

Write-Host ("Hardware    make|model : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} | {1}" -f $cs.Manufacturer, $cs.Model)
Write-Host ("           tag|bios : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} | {1}" -f $assetTag, $bios.SMBIOSBIOSVersion) -ForegroundColor Red
Write-Host ("           memory : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0}gb" -f $memGB)
Write-Host ("           cpu : ") -NoNewline -ForegroundColor Gray; Write-Host ($cpu.Name -replace '\s+', ' ')
Write-Host ("           disk1 : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} - {1} - s/n: {2}" -f $diskSize, $diskModel, $diskSN)

Write-Host ("-" * 62) -ForegroundColor Gray
Write-Section "Software"; Write-Host ""
Write-Host ("           OS : ") -NoNewline -ForegroundColor Gray
$osLine = "{0}, Version {1}, {2}" -f $caption, $displayVersion, $arch
Write-Host $osLine -ForegroundColor Green

Write-Host ("   install date : ") -NoNewline -ForegroundColor Gray; Write-Host ($install.ToString("yyyy/MM/dd HH:mm"))
Write-Host ("        uptime : ") -NoNewline -ForegroundColor Gray; Write-Host ($uptimeStr)

Write-Separator
Write-Host "Press Enter to close" -ForegroundColor Gray
[Console]::ReadLine() | Out-Null j'aimerais rajouter sa dans mon post 
Aurel-AD/Scripts-Administration-PowerShell
<#  WhatIsMyIP_Address.ps1 (fix compatibilité PS 5.1)
#>

function Format-GB($bytes) {
    if (-not $bytes) { return "n/a" }
    "{0:N0}GB" -f ($bytes / 1GB)
}
function Write-Separator { Write-Host ("=" * 62) -ForegroundColor Gray }
function Write-Section($title) { Write-Host ("{0,-13}: " -f $title) -NoNewline -ForegroundColor Gray }

# --- IP ---
$ipv4 = Get-NetIPAddress -AddressFamily IPv4 -PrefixOrigin Dhcp,Manual -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notmatch '^169\.254\.' -and $_.IPAddress -ne '127.0.0.1' } |
        Sort-Object InterfaceMetric, SkipAsSource |
        Select-Object -ExpandProperty IPAddress
$ip1 = $ipv4[0]; if (-not $ip1) { $ip1 = 'n/a' }
$ip2 = $ipv4[1]; if (-not $ip2) { $ip2 = 'n/a' }

# --- Nom PC ---
$compName = $env:COMPUTERNAME

# --- Matériel ---
$cs   = Get-CimInstance -ClassName Win32_ComputerSystem
$bios = Get-CimInstance -ClassName Win32_BIOS
$cpu  = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
$memGB = [int]($cs.TotalPhysicalMemory / 1GB)

$enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue
$assetTag  = ($enclosure.SMBIOSAssetTag | Select-Object -First 1)
if ([string]::IsNullOrWhiteSpace($assetTag)) { $assetTag = $bios.SerialNumber }

$pd = $null
try { $pd = Get-PhysicalDisk | Sort-Object FriendlyName | Select-Object -First 1 } catch {}
if ($pd) {
    $diskModel = $pd.FriendlyName
    $diskSN    = if ($pd.SerialNumber) { $pd.SerialNumber } else { "n/a" }
    $diskSize  = if ($pd.Size) { "{0:N0}GB" -f ($pd.Size / 1GB) } else { "n/a" }
} else {
    $dd = Get-CimInstance -ClassName Win32_DiskDrive | Sort-Object Index | Select-Object -First 1
    $diskModel = $dd.Model
    $diskSN    = if ($dd.SerialNumber) { $dd.SerialNumber } else { "n/a" }
    $diskSize  = Format-GB $dd.Size
}

# --- Logiciel / OS ---
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$caption = $os.Caption
$ver     = $os.Version
$arch    = $os.OSArchitecture
$install = $os.InstallDate
$boot    = $os.LastBootUpTime

try {
    $prod = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion
} catch { $prod = $null }

if ([string]::IsNullOrWhiteSpace($prod)) { $displayVersion = $ver } else { $displayVersion = $prod }

$uptime = (Get-Date) - $boot
$uptimeStr = ("{0} day(s), {1} hour(s), {2} minute(s), {3} second(s)." -f `
              [int]$uptime.Days, [int]$uptime.Hours, [int]$uptime.Minutes, [int]$uptime.Seconds)

# --- Affichage ---
Write-Separator
Write-Host ("  My IP Address1 is : ") -NoNewline -ForegroundColor Gray; Write-Host ($ip1) -ForegroundColor Yellow
Write-Host ("  My IP Address2 is : ") -NoNewline -ForegroundColor Gray; Write-Host ($ip2) -ForegroundColor Yellow
Write-Host ""
Write-Host ("  My ComputerName is : ") -NoNewline -ForegroundColor Gray; Write-Host ($compName) -ForegroundColor Cyan
Write-Separator

Write-Host ("Hardware    make|model : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} | {1}" -f $cs.Manufacturer, $cs.Model)
Write-Host ("           tag|bios : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} | {1}" -f $assetTag, $bios.SMBIOSBIOSVersion) -ForegroundColor Red
Write-Host ("           memory : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0}gb" -f $memGB)
Write-Host ("           cpu : ") -NoNewline -ForegroundColor Gray; Write-Host ($cpu.Name -replace '\s+', ' ')
Write-Host ("           disk1 : ") -NoNewline -ForegroundColor Gray; Write-Host ("{0} - {1} - s/n: {2}" -f $diskSize, $diskModel, $diskSN)

Write-Host ("-" * 62) -ForegroundColor Gray
Write-Section "Software"; Write-Host ""
Write-Host ("           OS : ") -NoNewline -ForegroundColor Gray
$osLine = "{0}, Version {1}, {2}" -f $caption, $displayVersion, $arch
Write-Host $osLine -ForegroundColor Green

Write-Host ("   install date : ") -NoNewline -ForegroundColor Gray; Write-Host ($install.ToString("yyyy/MM/dd HH:mm"))
Write-Host ("        uptime : ") -NoNewline -ForegroundColor Gray; Write-Host ($uptimeStr)

Write-Separator
Write-Host "Press Enter to close" -ForegroundColor Gray
[Console]::ReadLine() | Out-Null
