Import-Module ActiveDirectory

# Pasta de export
$basePath  = "C:\AD-Export"
$outFolder = Join-Path $basePath "Visio"

if (-not (Test-Path $outFolder)) {
    New-Item -Path $outFolder -ItemType Directory -Force | Out-Null
}

$domain  = Get-ADDomain
$domName = $domain.DNSRoot

Write-Host "Gerando CSVs de Sites / DCs / SiteLinks para Visio (domínio: $domName)..." -ForegroundColor Cyan

# Sites de replicação
$sites = Get-ADReplicationSite -Filter * -Properties * |
    Sort-Object Name

# Site Links (precisamos de SitesIncluded, Cost, ReplicationFrequencyInMinutes)
$siteLinks = Get-ADReplicationSiteLink -Filter * -Properties SitesIncluded,Cost,ReplicationFrequencyInMinutes,Description,Name

# DCs (propriedades padrão já trazem HostName, Site, IPv4, OS...)
$dcs = Get-ADDomainController -Filter * |
    Select-Object Name,HostName,Site,IPv4Address,OperatingSystem,OperatingSystemVersion

# --- 1) Sites + DCs (para shapes e contagem) ---

$siteRows = @()
$dcRows   = @()

foreach ($s in $sites) {
    $siteName = $s.Name
    $siteDN   = $s.DistinguishedName

    $siteDCs = $dcs | Where-Object { $_.Site -eq $siteName -or $_.Site -eq $siteDN }

    $siteRows += [PSCustomObject]@{
        Domain       = $domName
        SiteName     = $siteName
        SiteDN       = $siteDN
        Location     = $s.Location
        Description  = $s.Description
        DCCount      = $siteDCs.Count
    }

    foreach ($dc in $siteDCs) {
        $dcRows += [PSCustomObject]@{
            Domain      = $domName
            SiteName    = $siteName
            DCName      = $dc.Name
            DCFQDN      = $dc.HostName
            IPv4        = $dc.IPv4Address
            OS          = $dc.OperatingSystem
            OSVersion   = $dc.OperatingSystemVersion
        }
    }
}

$sitesFile = Join-Path $outFolder ("sites_dc_visio_{0}.csv" -f $domName.Replace(".","_"))

# Junta info de site + DCs (opcionalmente você pode usar só $siteRows se quiser algo mais macro)
$combined = @()

foreach ($sr in $siteRows) {
    $combined += $sr
}

$dcRows | ForEach-Object {
    $combined += $_
}

$combined | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $sitesFile

Write-Host "CSV de Sites/DCs gerado em: $sitesFile" -ForegroundColor Green

# --- 2) Site Links (pra você desenhar as conexões) ---

$linkRows = @()

foreach ($sl in $siteLinks) {
    $sitesIncluded = $sl.SitesIncluded
    if (-not $sitesIncluded -or $sitesIncluded.Count -lt 2) { continue }

    # Transforma DNs em nomes amigáveis
    $siteNames = @()
    foreach ($sdn in $sitesIncluded) {
        $siteObj = $sites | Where-Object { $_.DistinguishedName -eq $sdn }
        if ($siteObj) { $siteNames += $siteObj.Name }
    }

    # Pra Visio, é mais prático ter linha por "par"
    for ($i = 0; $i -lt $siteNames.Count; $i++) {
        for ($j = $i + 1; $j -lt $siteNames.Count; $j++) {
            $siteA = $siteNames[$i]
            $siteB = $siteNames[$j]

            $linkRows += [PSCustomObject]@{
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

$linksFile = Join-Path $outFolder ("site_links_visio_{0}.csv" -f $domName.Replace(".","_"))
$linkRows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $linksFile

Write-Host "CSV de SiteLinks gerado em: $linksFile" -ForegroundColor Green
Write-Host "Concluído." -ForegroundColor Cyan
