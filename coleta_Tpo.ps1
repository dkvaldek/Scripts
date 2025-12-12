Import-Module ActiveDirectory

# Pasta base
$basePath  = "C:\AD-Export"
$outFolder = Join-Path $basePath "AD_Topologia_Excel"

if (-not (Test-Path $outFolder)) {
    New-Item -Path $outFolder -ItemType Directory -Force | Out-Null
}

$domain   = Get-ADDomain
$domName  = $domain.DNSRoot
$domDN    = $domain.DistinguishedName

Write-Host "Coletando topologia AD para o domínio: $domName" -ForegroundColor Cyan

# ----------------------------------------------------
# 1) OUs
# ----------------------------------------------------
Write-Host "  - Coletando OUs..." -ForegroundColor Yellow

$ousRaw = Get-ADOrganizationalUnit -Filter * `
    -SearchBase $domDN `
    -SearchScope Subtree `
    -Properties Name,DistinguishedName,Description,ManagedBy

function Get-ParentDn {
    param([string]$dn)
    $parts = $dn -split ",", 2
    if ($parts.Count -eq 2) { return $parts[1] }
    return $null
}

$ouTable = $ousRaw | ForEach-Object {
    $parentDn = Get-ParentDn -dn $_.DistinguishedName

    # “nível” da OU na hierarquia (conta quantos “OU=” existem no DN)
    $level = ( $_.DistinguishedName -split ',' | Where-Object { $_ -like 'OU=*' } ).Count

    [PSCustomObject]@{
        Domain        = $domName
        OUName        = $_.Name
        OUDN          = $_.DistinguishedName
        ParentDN      = $parentDn
        Level         = $level
        Description   = $_.Description
        ManagedBy     = $_.ManagedBy
    }
}

$ouCsv = Join-Path $outFolder ("01_OUs_{0}.csv" -f $domName.Replace(".","_"))
$ouTable | Sort-Object Level, OUName | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ouCsv
Write-Host "    -> OUs exportadas em: $ouCsv" -ForegroundColor Green

# ----------------------------------------------------
# 2) Sites
# ----------------------------------------------------
Write-Host "  - Coletando Sites e DCs..." -ForegroundColor Yellow

$sitesRaw = Get-ADReplicationSite -Filter * -Properties *
$dcsRaw   = Get-ADDomainController -Filter * | 
            Select-Object Name,HostName,Site,IPv4Address,OperatingSystem,OperatingSystemVersion,IsGlobalCatalog,OperationMasterRoles

# Tabela de Sites
$siteTable = $sitesRaw | ForEach-Object {
    $siteName = $_.Name
    $siteDCs  = $dcsRaw | Where-Object { $_.Site -eq $siteName -or $_.Site -eq $_.DistinguishedName }

    [PSCustomObject]@{
        Domain       = $domName
        SiteName     = $siteName
        SiteDN       = $_.DistinguishedName
        Location     = $_.Location
        Description  = $_.Description
        DCCount      = $siteDCs.Count
    }
}

$sitesCsv = Join-Path $outFolder ("02_Sites_{0}.csv" -f $domName.Replace(".","_"))
$siteTable | Sort-Object SiteName | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $sitesCsv
Write-Host "    -> Sites exportados em: $sitesCsv" -ForegroundColor Green

# Tabela de DCs
$dcTable = $dcsRaw | ForEach-Object {
    [PSCustomObject]@{
        Domain       = $domName
        DCName       = $_.Name
        DCFQDN       = $_.HostName
        SiteName     = $_.Site
        IPv4Address  = $_.IPv4Address
        OperatingSystem        = $_.OperatingSystem
        OperatingSystemVersion = $_.OperatingSystemVersion
        IsGlobalCatalog        = $_.IsGlobalCatalog
        OperationMasterRoles   = ($_.OperationMasterRoles -join ';')
    }
}

$dcsCsv = Join-Path $outFolder ("03_DCs_{0}.csv" -f $domName.Replace(".","_"))
$dcTable | Sort-Object SiteName, DCName | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $dcsCsv
Write-Host "    -> DCs exportados em: $dcsCsv" -ForegroundColor Green

# ----------------------------------------------------
# 3) Site Links
# ----------------------------------------------------
Write-Host "  - Coletando Site Links..." -ForegroundColor Yellow

$siteLinksRaw = Get-ADReplicationSiteLink -Filter * -Properties SitesIncluded,Cost,ReplicationFrequencyInMinutes,Description,Name

$linkTable = @()

foreach ($sl in $siteLinksRaw) {
    $sitesIncluded = $sl.SitesIncluded
    if (-not $sitesIncluded -or $sitesIncluded.Count -lt 2) { continue }

    # converte DNs em nomes
    $siteNames = @()
    foreach ($sdn in $sitesIncluded) {
        $s = $sitesRaw | Where-Object { $_.DistinguishedName -eq $sdn }
        if ($s) { $siteNames += $s.Name }
    }

    for ($i = 0; $i -lt $siteNames.Count; $i++) {
        for ($j = $i + 1; $j -lt $siteNames.Count; $j++) {
            $siteA = $siteNames[$i]
            $siteB = $siteNames[$j]

            $linkTable += [PSCustomObject]@{
                Domain        = $domName
                LinkName      = $sl.Name
                SiteA         = $siteA
                SiteB         = $siteB
                Cost          = $sl.Cost
                FrequencyMin  = $sl.ReplicationFrequencyInMinutes
                Description   = $sl.Description
            }
        }
    }
}

$linksCsv = Join-Path $outFolder ("04_SiteLinks_{0}.csv" -f $domName.Replace(".","_"))
$linkTable | Sort-Object LinkName, SiteA, SiteB | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $linksCsv
Write-Host "    -> Site Links exportados em: $linksCsv" -ForegroundColor Green

Write-Host ""
Write-Host "Coleta concluída. Arquivos gerados em: $outFolder" -ForegroundColor Cyan

# ----------------------------------------------------
# 4) OPCIONAL: gerar workbook .xlsx se houver módulo ImportExcel
# ----------------------------------------------------
if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host "Módulo ImportExcel encontrado. Gerando workbook .XLSX..." -ForegroundColor Cyan

    $xlsxPath = Join-Path $outFolder ("AD_Topologia_{0}.xlsx" -f $domName.Replace(".","_"))

    # Remove se já existir
    if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }

    $ouTable    | Export-Excel -Path $xlsxPath -WorksheetName 'OUs'       -AutoSize -TableName 'OUs'       -BoldTopRow
    $siteTable  | Export-Excel -Path $xlsxPath -WorksheetName 'Sites'     -AutoSize -TableName 'Sites'     -BoldTopRow
    $dcTable    | Export-Excel -Path $xlsxPath -WorksheetName 'DCs'       -AutoSize -TableName 'DCs'       -BoldTopRow
    $linkTable  | Export-Excel -Path $xlsxPath -WorksheetName 'SiteLinks' -AutoSize -TableName 'SiteLinks' -BoldTopRow

    Write-Host "Workbook Excel gerado em: $xlsxPath" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "OBS: módulo ImportExcel não encontrado." -ForegroundColor Yellow
    Write-Host "     Você já tem CSVs prontos; se quiser um .xlsx automático," 
    Write-Host "     instale o módulo com:" 
    Write-Host "         Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Cyan
}
