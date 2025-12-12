Import-Module ActiveDirectory

# Pasta de export (a mesma que você já está usando)
$exportPath = "C:\AD-Export"
New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

Write-Host "Coletando Sites e Subnets do AD..." -ForegroundColor Cyan

# Pega o Configuration NC da floresta
try {
    $rootDse  = Get-ADRootDSE -ErrorAction Stop
    $configNC = $rootDse.ConfigurationNamingContext
}
catch {
    Write-Host "ERRO: Não foi possível obter o RootDSE. Verifique conectividade com o DC / serviços ADWS." -ForegroundColor Red
    Write-Host $_
    return
}

########## SITES ##########

# força array com @(...)
$sites = @(
    Get-ADObject `
        -SearchBase ("CN=Sites," + $configNC) `
        -LDAPFilter "(objectClass=site)" `
        -Properties *
)

# Export detalhado
$sites |
    Select-Object `
        Name,
        DistinguishedName,
        @{Name = "Location"; Expression = { $_.location }},
        Description |
    Export-Csv "$exportPath\sites_all_raw.csv" -NoTypeInformation -Encoding UTF8

# Resumo na tela
Write-Host ""
Write-Host "=== SITES (RESUMO NA TELA) ===" -ForegroundColor Yellow

if ($sites -and $sites.Count -gt 0) {
    $sites |
        Select-Object `
            Name,
            @{Name = "Location"; Expression = { $_.location }},
            Description |
        Sort-Object Name |
        Format-Table -AutoSize
} else {
    Write-Host "Nenhum site encontrado." -ForegroundColor DarkYellow
}

########## SUBNETS ##########

$subnets = @(
    Get-ADObject `
        -SearchBase ("CN=Subnets,CN=Sites," + $configNC) `
        -LDAPFilter "(objectClass=subnet)" `
        -Properties siteObject,location,description
)

# Monta objeto com SiteName (também forçando array)
$subnetsExpanded = @(
    $subnets | ForEach-Object {
        $siteName = $null
        if ($_.siteObject) {
            $siteObj = Get-ADObject -Identity $_.siteObject -ErrorAction SilentlyContinue
            if ($siteObj) { $siteName = $siteObj.Name }
        }

        [PSCustomObject]@{
            Name              = $_.Name
            DistinguishedName = $_.DistinguishedName
            SiteDN            = $_.siteObject
            SiteName          = $siteName
            Location          = $_.location
            Description       = $_.Description
        }
    }
)

# Export detalhado
$subnetsExpanded |
    Export-Csv "$exportPath\subnets_all_raw.csv" -NoTypeInformation -Encoding UTF8

# Resumo na tela
Write-Host ""
Write-Host "=== SUBNETS (RESUMO NA TELA) ===" -ForegroundColor Yellow

if ($subnetsExpanded -and $subnetsExpanded.Count -gt 0) {
    $subnetsExpanded |
        Select-Object `
            Name,
            SiteName,
            Location,
            Description |
        Sort-Object Name |
        Format-Table -AutoSize
} else {
    Write-Host "Nenhuma subnet encontrado." -ForegroundColor DarkYellow
}

Write-Host ""
Write-Host "Arquivos gerados:" -ForegroundColor Cyan
Write-Host "  - $exportPath\sites_all_raw.csv"
Write-Host "  - $exportPath\subnets_all_raw.csv"
