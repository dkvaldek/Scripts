Import-Module ActiveDirectory

$basePath  = "C:\AD-Export"
$outFolder = Join-Path $basePath "PowerBI"

if (-not (Test-Path $outFolder)) {
    New-Item -Path $outFolder -ItemType Directory -Force | Out-Null
}

$forest = Get-ADForest

Write-Host "Coletando nível funcional de floresta e domínios: $($forest.Name)..." -ForegroundColor Cyan

$rows = @()

# Linha da FLORESTA
$rows += [PSCustomObject]@{
    LevelType      = 'Forest'
    ForestName     = $forest.Name
    RootDomain     = $forest.RootDomain
    DomainFqdn     = $null
    DomainNetBIOS  = $null
    FunctionalMode = $forest.ForestMode.ToString()
}

# Linhas dos DOMÍNIOS dessa floresta
foreach ($domName in $forest.Domains) {
    $dom = Get-ADDomain -Identity $domName

    $rows += [PSCustomObject]@{
        LevelType      = 'Domain'
        ForestName     = $forest.Name
        RootDomain     = $forest.RootDomain
        DomainFqdn     = $dom.DNSRoot
        DomainNetBIOS  = $dom.NetBIOSName
        FunctionalMode = $dom.DomainMode.ToString()
    }
}

$funcFile = Join-Path $outFolder ("FunctionalLevels_{0}.csv" -f $forest.Name.Replace('.','_'))
$rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $funcFile

Write-Host "Arquivo gerado: $funcFile" -ForegroundColor Green
