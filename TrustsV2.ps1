# ============================
# Coleta de Relações de Confiança (Trusts)
# - Domínios e florestas relacionados
# - Resumo na tela
# - Detalhado em C:\AD-Export\trusts_all_raw.csv
# - Versão SEM Get-AD*
# ============================

# Pasta de export
$exportPath = "C:\AD-Export"
New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

Write-Host "Coletando informações de confiança (trusts) da floresta via .NET..." -ForegroundColor Cyan

$allTrusts = @()

# -----------------------------------------
# 1) Obter floresta e domínios via .NET
# -----------------------------------------
try {
    $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
}
catch {
    Write-Host "ERRO: não foi possível obter a floresta via .NET. Esta máquina está no domínio e alcança um DC?" -ForegroundColor Red
    Write-Host $_
    return
}

$domains = $forest.Domains

# -----------------------------------------
# 2) Trusts por DOMÍNIO
# -----------------------------------------
foreach ($dom in $domains) {
    $domainName = $dom.Name
    Write-Host "`nConsultando trusts no domínio (via .NET): $domainName" -ForegroundColor Yellow

    try {
        $domTrusts = $dom.GetAllTrustRelationships()

        foreach ($t in $domTrusts) {
            # $t é TrustRelationshipInformation
            $allTrusts += [PSCustomObject]@{
                SourceDomain        = $domainName
                Name                = $t.TargetName
                Target              = $t.TargetName
                Direction           = $t.TrustDirection.ToString()  # Inbound/Outbound/Bidirectional
                TrustType           = $t.TrustType.ToString()       # External, Forest, Kerberos, etc.
                TrustScope          = "DomainLevel"
                ForestName          = $forest.Name
                # Campos que precisaríamos do Get-ADTrust -> deixamos nulos/N/A
                TrustAttributes     = $null
                ForestTransitive    = $null
                IntraForest         = $null
                SIDFilteringEnabled = $null
                Created             = $null
                Modified            = $null
                DistinguishedName   = "(Domain-level trust via .NET: $domainName -> $($t.TargetName))"
            }
        }
    }
    catch {
        Write-Warning "Falha ao consultar trusts no domínio $domainName via .NET: $_"
    }
}

# -----------------------------------------
# 3) Trusts de FLORESTA
# -----------------------------------------
try {
    Write-Host "`nConsultando trusts de FLORESTA (API .NET)..." -ForegroundColor Yellow
    $forestTrusts = $forest.GetAllTrustRelationships()

    foreach ($ft in $forestTrusts) {
        $allTrusts += [PSCustomObject]@{
            SourceDomain        = $forest.Name + " (ForestRoot)"
            Name                = $ft.TargetName
            Target              = $ft.TargetName
            Direction           = $ft.TrustDirection.ToString()
            TrustType           = $ft.TrustType.ToString()
            TrustScope          = "ForestLevel"
            ForestName          = $forest.Name
            TrustAttributes     = $null
            ForestTransitive    = $null
            IntraForest         = $null
            SIDFilteringEnabled = $null
            Created             = $null
            Modified            = $null
            DistinguishedName   = "(Forest-level trust via .NET: $($forest.Name) -> $($ft.TargetName))"
        }
    }
}
catch {
    Write-Warning "Falha ao consultar trusts de floresta via .NET: $_"
}

Write-Host "`n=== RELAÇÕES DE CONFIANÇA (RESUMO) ===`n" -ForegroundColor Cyan

if ($allTrusts.Count -gt 0) {
    $allTrusts |
        Select-Object SourceDomain, Target, Direction, TrustType, TrustScope, ForestName |
        Sort-Object SourceDomain, Target |
        Format-Table -AutoSize

    # Export detalhado
    $csvFile = Join-Path $exportPath "trusts_all_raw.csv"
    $allTrusts |
        Export-Csv $csvFile -NoTypeInformation -Encoding UTF8

    Write-Host "`nExport concluído: $csvFile" -ForegroundColor Green
} else {
    Write-Warning "Nenhum trust encontrado (ou erro ao consultar)."
}
