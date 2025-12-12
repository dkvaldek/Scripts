# ============================
# Coleta de SITES e SUBNETS (sem Get-AD*)
# - Resumo na tela
# - Detalhes em:
#    C:\AD-Export\sites_all_raw.csv
#    C:\AD-Export\subnets_all_raw.csv
# ============================

$exportPath = "C:\AD-Export"
if (-not (Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
}

Write-Host "Coletando Sites e Subnets via LDAP/.NET (sem AD Web Services)..." -ForegroundColor Cyan

# RootDSE via LDAP (não usa Get-ADRootDSE)
try {
    $rootDse  = [ADSI]"LDAP://RootDSE"
    $configNC = $rootDse.configurationNamingContext
}
catch {
    Write-Host "ERRO: não foi possível obter o RootDSE. Esta máquina está no domínio e alcança um DC por LDAP?" -ForegroundColor Red
    Write-Host $_
    return
}

function Get-LdapObjects {
    param(
        [string]$SearchBase,
        [string]$Filter,
        [string[]]$Properties
    )

    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.SearchRoot  = [ADSI]("LDAP://$SearchBase")
    $searcher.Filter      = $Filter
    $searcher.SearchScope = "Subtree"
    $searcher.PageSize    = 1000
    $searcher.PropertiesToLoad.Clear()

    foreach ($p in $Properties) {
        [void]$searcher.PropertiesToLoad.Add($p)
    }

    $results = $searcher.FindAll()
    $list = @()

    foreach ($r in $results) {
        $objProps = @{}
        foreach ($p in $Properties) {
            if ($r.Properties[$p].Count -gt 0) {
                $objProps[$p] = $r.Properties[$p][0]
            } else {
                $objProps[$p] = $null
            }
        }
        $list += New-Object psobject -Property $objProps
    }

    $results.Dispose()
    return $list
}

########## SITES ##########

Write-Host ""
Write-Host "=== SITES ===" -ForegroundColor Yellow

$sitesBase = "CN=Sites,$configNC"

$sites = Get-LdapObjects `
    -SearchBase $sitesBase `
    -Filter "(objectClass=site)" `
    -Properties @("name","distinguishedName","location","description")

$sitesOut = $sites | Select-Object `
    @{Name="Name";               Expression = { $_.name }},
    @{Name="DistinguishedName";  Expression = { $_.distinguishedName }},
    @{Name="Location";           Expression = { $_.location }},
    @{Name="Description";        Expression = { $_.description }}

# Resumo na tela
Write-Host "Total de sites encontrados: $($sitesOut.Count)"
if ($sitesOut.Count -gt 0) {
    $sitesOut |
        Select-Object Name,Location |
        Sort-Object Name |
        Format-Table -AutoSize
} else {
    Write-Host "Nenhum site encontrado em CN=Sites,$configNC" -ForegroundColor DarkYellow
}

# CSV detalhado
$sitesCsv = Join-Path $exportPath "sites_all_raw.csv"
$sitesOut | Export-Csv $sitesCsv -NoTypeInformation -Encoding UTF8
Write-Host "Detalhes de sites exportados para: $sitesCsv"


########## SUBNETS ##########

Write-Host ""
Write-Host "=== SUBNETS ===" -ForegroundColor Yellow

$subnetsBase = "CN=Subnets,CN=Sites,$configNC"

$subnets = Get-LdapObjects `
    -SearchBase $subnetsBase `
    -Filter "(objectClass=subnet)" `
    -Properties @("name","distinguishedName","siteObject","location","description")

$subnetsOut = $subnets | Select-Object `
    @{Name="Name";              Expression = { $_.name }},                 # Ex: 10.0.0.0/24
    @{Name="DistinguishedName"; Expression = { $_.distinguishedName }},
    @{Name="SiteDN";            Expression = { $_.siteObject }},
    @{Name="SiteName";          Expression = {
            if ($_.siteObject) {
                ($_.siteObject -split ",")[0] -replace '^CN='
            } else {
                $null
            }
        }},
    @{Name="Location";          Expression = { $_.location }},
    @{Name="Description";       Expression = { $_.description }}

# Resumo na tela
Write-Host "Total de subnets encontradas: $($subnetsOut.Count)"
if ($subnetsOut.Count -gt 0) {
    $subnetsOut |
        Select-Object Name,SiteName |
        Sort-Object Name |
        Format-Table -AutoSize
} else {
    Write-Host "Nenhuma subnet encontrada em CN=Subnets,CN=Sites,$configNC" -ForegroundColor DarkYellow
}

# CSV detalhado
$subnetsCsv = Join-Path $exportPath "subnets_all_raw.csv"
$subnetsOut | Export-Csv $subnetsCsv -NoTypeInformation -Encoding UTF8
Write-Host "Detalhes de subnets exportados para: $subnetsCsv"

Write-Host "`nColeta de Sites e Subnets concluída." -ForegroundColor Green
    