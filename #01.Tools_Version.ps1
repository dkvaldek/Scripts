Server 2008
_______________________________________________
# ============================
# Versões RSAT (GPMC, ADUC, DNS)
#   - Somente saída na tela
# ============================

$sys32 = Join-Path $env:SystemRoot "System32"

# Arquivos-alvo
$gpmcFiles = @(
    (Join-Path $sys32 "gpmc.msc"),
    (Join-Path $sys32 "gpmgmt.dll")
)

$aducFiles = @(
    (Join-Path $sys32 "dsa.msc"),
    (Join-Path $sys32 "dsadmin.dll"),
    (Join-Path $sys32 "dsquery.dll")
)

$dnsFiles = @(
    (Join-Path $sys32 "dnsmgmt.msc"),
    (Join-Path $sys32 "dnsmgmt.dll")
)

function Get-ToolVersions {
    param(
        [string]$ToolName,
        [string[]]$Files
    )

    $results = @()

    foreach ($file in $Files) {
        if (Test-Path $file) {
            $fi = Get-Item $file
            $results += [PSCustomObject]@{
                Tool          = $ToolName
                File          = [System.IO.Path]::GetFileName($file)
                FullPath      = $file
                FileVersion   = $fi.VersionInfo.FileVersion
                ProductVersion= $fi.VersionInfo.ProductVersion
            }
        }
    }

    return $results
}

Write-Host "Verificando versões dos componentes RSAT (GPMC, ADUC, DNS)..." -ForegroundColor Cyan

$gpmcVers = Get-ToolVersions -ToolName "GPMC" -Files $gpmcFiles
$aducVers = Get-ToolVersions -ToolName "ADUC" -Files $aducFiles
$dnsVers  = Get-ToolVersions -ToolName "DNS Manager" -Files $dnsFiles

Write-Host ""

# GPMC
Write-Host "=== GPMC (Group Policy Management Console) ===" -ForegroundColor Yellow
if ($gpmcVers.Count -gt 0) {
    $gpmcVers |
        Sort-Object File |
        Format-Table File,FileVersion,ProductVersion,FullPath -AutoSize
} else {
    Write-Host "GPMC não encontrada (gpmc.msc/gpmgmt.dll não localizados em $sys32)." -ForegroundColor DarkYellow
}

Write-Host ""

# ADUC
Write-Host "=== ADUC (Active Directory Users and Computers) ===" -ForegroundColor Yellow
if ($aducVers.Count -gt 0) {
    $aducVers |
        Sort-Object File |
        Format-Table File,FileVersion,ProductVersion,FullPath -AutoSize
} else {
    Write-Host "ADUC não encontrado (dsa.msc/dsadmin.dll/dsquery.dll não localizados em $sys32)." -ForegroundColor DarkYellow
}

Write-Host ""

# DNS
Write-Host "=== DNS Manager ===" -ForegroundColor Yellow
if ($dnsVers.Count -gt 0) {
    $dnsVers |
        Sort-Object File |
        Format-Table File,FileVersion,ProductVersion,FullPath -AutoSize
} else {
    Write-Host "DNS Manager não encontrado (dnsmgmt.msc/dnsmgmt.dll não localizados em $sys32)." -ForegroundColor DarkYellow

    #2016
    #____________
    # ============================
# Versões RSAT (GPMC, ADUC, DNS) + DNS Server
#   - Somente saída na tela
# ============================

$sys32 = [System.Environment]::SystemDirectory   # pega o System32 real (64 bits)

Write-Host "Usando System32 em: $sys32`n" -ForegroundColor Cyan

# Arquivos-alvo
$gpmcFiles = @(
    (Join-Path $sys32 "gpmc.msc"),
    (Join-Path $sys32 "gpmgmt.dll")
)

$aducFiles = @(
    (Join-Path $sys32 "dsa.msc"),
    (Join-Path $sys32 "dsadmin.dll"),
    (Join-Path $sys32 "dsquery.dll")
)

$dnsFiles = @(
    (Join-Path $sys32 "dnsmgmt.msc"),
    (Join-Path $sys32 "dnsmgmt.dll"),
    (Join-Path $sys32 "dns.exe")        # binário do DNS Server
)

function Get-ToolVersions {
    param(
        [string]$ToolName,
        [string[]]$Files
    )

    $results = @()

    foreach ($file in $Files) {
        if (Test-Path $file) {
            $fi = Get-Item $file
            $results += [PSCustomObject]@{
                Tool           = $ToolName
                File           = [System.IO.Path]::GetFileName($file)
                FullPath       = $file
                FileVersion    = $fi.VersionInfo.FileVersion
                ProductVersion = $fi.VersionInfo.ProductVersion
            }
        }
    }

    return $results
}

Write-Host "Verificando versões dos componentes RSAT (GPMC, ADUC, DNS)..." -ForegroundColor Cyan
Write-Host ""

# GPMC
$gpmcVers = Get-ToolVersions -ToolName "GPMC" -Files $gpmcFiles
Write-Host "=== GPMC (Group Policy Management Console) ===" -ForegroundColor Yellow
if ($gpmcVers.Count -gt 0) {
    $gpmcVers |
        Sort-Object File |
        Format-Table File,FileVersion,ProductVersion,FullPath -AutoSize
} else {
    Write-Host "GPMC não encontrada (gpmc.msc/gpmgmt.dll não localizados em $sys32)." -ForegroundColor DarkYellow
}
Write-Host ""

# ADUC
$aducVers = Get-ToolVersions -ToolName "ADUC" -Files $aducFiles
Write-Host "=== ADUC (Active Directory Users and Computers) ===" -ForegroundColor Yellow
if ($aducVers.Count -gt 0) {
    $aducVers |
        Sort-Object File |
        Format-Table File,FileVersion,ProductVersion,FullPath -AutoSize
} else {
    Write-Host "ADUC não encontrado (dsa.msc/dsadmin.dll/dsquery.dll não localizados em $sys32)." -ForegroundColor DarkYellow
}
Write-Host ""

# DNS
$dnsVers  = Get-ToolVersions -ToolName "DNS" -Files $dnsFiles
Write-Host "=== DNS (Manager + DNS Server) ===" -ForegroundColor Yellow
if ($dnsVers.Count -gt 0) {
    $dnsVers |
        Sort-Object File |
        Format-Table File,FileVersion,ProductVersion,FullPath -AutoSize
} else {
    Write-Host "DNS Manager / DNS Server não encontrados (dnsmgmt.msc/dnsmgmt.dll/dns.exe não localizados em $sys32)." -ForegroundColor DarkYellow
}
a