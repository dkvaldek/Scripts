# ============================
# Nível funcional de Floresta e Domínio(s) - via .NET (sem ADWS)
#   - Resumo na tela
#   - Detalhes em CSV na pasta 01_Inventario_DomainControllers
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "01_Inventario_DomainControllers"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Write-Host "Coletando nível funcional de Floresta e Domínios (usando .NET, sem AD Web Services)..." -ForegroundColor Cyan

# ---------- FLORESTA ----------
try {
    $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
}
catch {
    Write-Host "ERRO: não foi possível obter informações da floresta via .NET. Verifique conectividade com o AD." -ForegroundColor Red
    Write-Host $_
    return
}

# Pega lista de domínios como strings
$domainNames = $forest.Domains | ForEach-Object { $_.Name }

$forestInfo = [PSCustomObject]@{
    ForestName   = $forest.Name
    RootDomain   = $forest.RootDomain.Name
    ForestMode   = $forest.ForestMode.ToString()
    DomainsCount = $domainNames.Count
    Domains      = ($domainNames -join ';')
}

Write-Host ""
Write-Host "=== NÍVEL FUNCIONAL DA FLORESTA ===" -ForegroundColor Yellow
$forestInfo |
    Select-Object ForestName,RootDomain,ForestMode,DomainsCount |
    Format-Table -AutoSize

$forestFile = Join-Path $OutFolder "forest_functional_level.csv"
$forestInfo | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $forestFile

# ---------- DOMÍNIOS ----------
Write-Host ""
Write-Host "Coletando nível funcional dos domínios..." -ForegroundColor Cyan

$domainInfos = @()

# RootDSE para achar a partição de Partitions (pra pegar NetBIOS)
$rootDse = [ADSI]"LDAP://RootDSE"
$confNC  = $rootDse.configurationNamingContext
$partitionsPath = "LDAP://CN=Partitions,$confNC"

foreach ($domObj in $forest.Domains) {
    try {
        # Objeto Domain (System.DirectoryServices.ActiveDirectory.Domain)
        $domName = $domObj.Name
        $domMode = $domObj.DomainMode.ToString()

        # DN do domínio pela entry LDAP
        $de = $domObj.GetDirectoryEntry()
        $dn = $de.distinguishedName

        # Descobrir NetBIOS via crossRef em CN=Partitions
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry($partitionsPath)
        $searcher.Filter = "(&(objectClass=crossRef)(nCName=$dn))"
        $searcher.PropertiesToLoad.Add("nETBIOSName") | Out-Null

        $result  = $searcher.FindOne()
        $netbios = $null
        if ($result -and $result.Properties["nETBIOSName"].Count -gt 0) {
            $netbios = $result.Properties["nETBIOSName"][0]
        }

        $domainInfos += [PSCustomObject]@{
            DomainName = $domName
            NetBIOS    = $netbios
            DomainMode = $domMode
            DomainDN   = $dn
        }
    }
    catch {
        Write-Warning "Falha ao obter informações do domínio $($domObj.Name) : $_"
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
