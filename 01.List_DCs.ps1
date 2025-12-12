# ============================
# Listar TODOS os DCs da FLORESTA
# - Saída simples na tela
# - Detalhado em C:\AD-Export\dcs_all_raw.csv
# ============================

# Carrega módulo AD
Import-Module ActiveDirectory

# Pasta de export
$exportPath = "C:\AD-Export"
New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

Write-Host "Coletando domínios da floresta..." -ForegroundColor Cyan

# Pega floresta e domínios
$forest  = Get-ADForest
$domains = $forest.Domains

$dcs = @()

foreach ($domain in $domains) {
    Write-Host "Consultando DCs do domínio: $domain"

    $dcs += Get-ADDomainController -Server $domain -Filter *
}

Write-Host "`n=== DOMAIN CONTROLLERS DA FLORESTA (RESUMO) ===`n"

$dcs |
    Select-Object HostName,
                  Domain,
                  IPv4Address,
                  Site,
                  OperatingSystem,
                  OperatingSystemVersion |
    Sort-Object Domain, HostName |
    Format-Table -AutoSize

# Export detalhado
$dcs |
    Select-Object * |
    Export-Csv "$exportPath\dcs_all_raw.csv" -NoTypeInformation -Encoding UTF8

Write-Host "`nExport concluído: $exportPath\dcs_all_raw.csv" -ForegroundColor Green
