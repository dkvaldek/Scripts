    Import-Module ActiveDirectory

    $basePath  = "C:\AD-Export"
    $outFolder = Join-Path $basePath "Visio"

    if (-not (Test-Path $outFolder)) {
        New-Item -Path $outFolder -ItemType Directory -Force | Out-Null
    }

    $domain  = Get-ADDomain
    $domName = $domain.DNSRoot

    Write-Host "Gerando TXT de Sites + DCs para Visio (Organization Chart) - domínio: $domName..." -ForegroundColor Cyan

    # Sites
    $sites = Get-ADReplicationSite -Filter * -Properties * | Sort-Object Name

    # DCs (compatível com versões antigas: sem -Properties)
    $dcs = Get-ADDomainController -Filter *  # HostName, Site e Name já vêm por padrão

    # Monta as linhas tab-delimited
    $lines = @()

    # Cabeçalho no formato que o Visio entende
    $lines += "Name`tReportsTo`tType`tExtra"

    # Nó raiz = domínio
    $lines += ("{0}`t`tDomain`t{1}" -f $domName, $domName)

    # Sites: pendurados direto no domínio
    foreach ($s in $sites) {
        $siteName = $s.Name
        $extra    = $s.Description
        if (-not $extra) { $extra = $s.Location }

        $lines += ("{0}`t{1}`tSite`t{2}" -f $siteName, $domName, $extra)
    }

    # DCs: pendurados em seus respectivos sites
    foreach ($dc in $dcs) {
        # nome pra exibir
        $dcName = if ($dc.HostName) { $dc.HostName } else { $dc.Name }

        # em algumas versões, Site é string com nome do site
        $siteName = $dc.Site

        if (-not $siteName) {
            # se não tiver site, pendura direto no domínio
            $lines += ("{0}`t{1}`tDC`t(no site)" -f $dcName, $domName)
        } else {
            $lines += ("{0}`t{1}`tDC`t{2}" -f $dcName, $siteName, $dc.IPv4Address)
        }
    }

    $outFile = Join-Path $outFolder ("sites_dcs_orgchart_{0}.txt" -f $domName.Replace(".","_"))

    # Encoding "Default" (ANSI) pra não ter problema de BOM
    $lines | Set-Content -Path $outFile -Encoding Default

    Write-Host "Arquivo TXT de Sites/DCs para Visio gerado em: $outFile" -ForegroundColor Green
