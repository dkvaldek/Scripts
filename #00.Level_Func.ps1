# ============================
# Nível funcional de Floresta e Domínio(s)
#   - Resumo na tela
#   - Detalhes em CSV na pasta 01_Inventario_DomainControllers
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "01_Inventario_DomainControllers"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Import-Module ActiveDirectory

Write-Host "Coletando nível funcional de Floresta e Domínios..." -ForegroundColor Cyan

# ---------- FLORESTA ----------
$forest = Get-ADForest

$forestInfo = [PSCustomObject]@{
    ForestName   = $forest.Name
    RootDomain   = $forest.RootDomain
    ForestMode   = $forest.ForestMode.ToString()
    DomainsCount = $forest.Domains.Count
    Domains      = ($forest.Domains -join ';')
}

Write-Host ""
Write-Host "=== NÍVEL FUNCIONAL DA FLORESTA ===" -ForegroundColor Yellow
$forestInfo |
    Select-Object ForestName,RootDomain,ForestMode,DomainsCount |
    Format-Table -AutoSize

$forestFile = Join-Path $OutFolder "forest_functional_level.csv"
$forestInfo | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $forestFile

# ---------- DOMÍNIOS ----------
$domains = $forest.Domains
$domainInfos = @()

foreach ($d in $domains) {
    $dom = Get-ADDomain -Identity $d

    $domainInfos += [PSCustomObject]@{
        DomainName = $dom.DNSRoot
        NetBIOS    = $dom.NetBIOSName
        DomainMode = $dom.DomainMode.ToString()
        DomainDN   = $dom.DistinguishedName
    }
}

Write-Host ""
Write-Host "=== NÍVEL FUNCIONAL DOS DOMÍNIOS ===" -ForegroundColor Yellow
$domainInfos |
    Select-Object DomainName,NetBIOS,DomainMode |
    Sort-Object DomainName |
    Format-Table -AutoSize

$domainFile = Join-Path $OutFolder "domain_functional_levels.csv"
$domainInfos | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $domainFile

Write-Host ""
Write-Host "Arquivos gerados:" -ForegroundColor Cyan
Write-Host "  - $forestFile"
Write-Host "  - $domainFile"
