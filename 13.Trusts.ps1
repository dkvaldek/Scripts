# ============================
# Coleta de Relações de Confiança (Trusts)
# - Domínios e florestas relacionados
# - Resumo na tela
# - Detalhado em C:\AD-Export\trusts_all_raw.csv
# ============================

Import-Module ActiveDirectory

# Pasta de export
$exportPath = "C:\AD-Export"
New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

Write-Host "Coletando informações de confiança (trusts) da floresta..." -ForegroundColor Cyan

# Pega floresta e domínios
$forest  = Get-ADForest
$domains = $forest.Domains

$allTrusts = @()

foreach ($domain in $domains) {
    Write-Host "`nConsultando trusts no domínio: $domain" -ForegroundColor Yellow

    try {
        # Trusts vistos a partir desse domínio
        $trusts = Get-ADTrust -Server $domain -Filter * -ErrorAction Stop

        foreach ($t in $trusts) {
            $allTrusts += [PSCustomObject]@{
                SourceDomain        = $domain
                Name                = $t.Name
                Target              = $t.Target
                Direction           = $t.Direction          # Inbound / Outbound / Bidirectional
                TrustType           = $t.TrustType          # External / Forest / Shortcut / Realm
                TrustAttributes     = $t.TrustAttributes
                ForestTransitive    = $t.ForestTransitive
                IntraForest         = $t.IntraForest
                SIDFilteringEnabled = $t.SIDFilteringEnabled
                Created             = $t.Created
                Modified            = $t.Modified
                DistinguishedName   = $t.DistinguishedName
            }
        }
    }
    catch {
        Write-Warning "Falha ao consultar trusts no domínio $domain : $_"
    }
}

# Também coleta trusts em nível de floresta usando .NET (complementar)
try {
    Write-Host "`nConsultando trusts de FLORESTA (API .NET)..." -ForegroundColor Yellow
    $forestObj     = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $forestTrusts  = $forestObj.GetAllTrustRelationships()

    foreach ($ft in $forestTrusts) {
        $allTrusts += [PSCustomObject]@{
            SourceDomain        = $forestObj.Name + " (ForestRoot)"
            Name                = $ft.TargetName
            Target              = $ft.TargetName
            Direction           = $ft.TrustDirection
            TrustType           = $ft.TrustType
            TrustAttributes     = $null
            ForestTransitive    = $null
            IntraForest         = $null
            SIDFilteringEnabled = $null
            Created             = $null
            Modified            = $null
            DistinguishedName   = "(Forest-level, via GetAllTrustRelationships())"
        }
    }
}
catch {
    Write-Warning "Falha ao consultar trusts de floresta via .NET: $_"
}

Write-Host "`n=== RELAÇÕES DE CONFIANÇA (RESUMO) ===`n" -ForegroundColor Cyan

if ($allTrusts.Count -gt 0) {
    $allTrusts |
        Select-Object SourceDomain, Target, Direction, TrustType, IntraForest, ForestTransitive |
        Sort-Object SourceDomain, Target |
        Format-Table -AutoSize

    # Export detalhado
    $allTrusts |
        Export-Csv "$exportPath\trusts_all_raw.csv" -NoTypeInformation -Encoding UTF8

    Write-Host "`nExport concluído: $exportPath\trusts_all_raw.csv" -ForegroundColor Green
} else {
    Write-Warning "Nenhum trust encontrado (ou erro ao consultar)."
}
