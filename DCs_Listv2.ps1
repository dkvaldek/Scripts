# ============================
# Listar TODOS os DCs da FLORESTA (sem ADWS)
# - Saída simples na tela
# - Detalhado em C:\AD-Export\dcs_all_raw.csv
# ============================

$exportPath = "C:\AD-Export"
if (-not (Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
}

Write-Host "Coletando informações de Domain Controllers via .NET..." -ForegroundColor Cyan

try {
    # Pega floresta atual via .NET (não usa ADWS)
    $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
}
catch {
    Write-Host "ERRO: não foi possível obter a floresta via .NET. Verifique se o servidor está no domínio." -ForegroundColor Red
    Write-Host $_
    return
}

$allDCs = @()

foreach ($dom in $forest.Domains) {
    Write-Host "Consultando DCs do domínio: $($dom.Name)" -ForegroundColor Yellow

    foreach ($dc in $dom.DomainControllers) {
        # $dc é um objeto System.DirectoryServices.ActiveDirectory.DomainController
        $allDCs += [PSCustomObject]@{
            HostName      = $dc.Name
            Domain        = $dom.Name
            IPAddress     = $dc.IPAddress
            Site          = $dc.SiteName
            OSVersion     = $dc.OSVersion
            Forest        = $forest.Name
        }
    }
}

Write-Host "`n=== DOMAIN CONTROLLERS DA FLORESTA (RESUMO) ===`n"

$allDCs |
    Select-Object HostName,Domain,IPAddress,Site,OSVersion |
    Sort-Object Domain,HostName |
    Format-Table -AutoSize

# Export detalhado
$dcsFile = Join-Path $exportPath "dcs_all_raw.csv"

$allDCs |
    Export-Csv $dcsFile -NoTypeInformation -Encoding UTF8

Write-Host "`nExport concluído: $dcsFile" -ForegroundColor Green
