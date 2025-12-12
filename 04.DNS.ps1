# ============================
# Coleta de DNS - Zonas e SRV _msdcs
# ============================

$exportPath = "C:\AD-Export"
New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

# Servidor DNS alvo (local)
$dnsServer = $env:COMPUTERNAME

Write-Host "Coletando informações de DNS no servidor $dnsServer..." -ForegroundColor Cyan

##############################
# 1) ZONAS DNS -> dns_zones_raw.csv
##############################

$zones = Get-WmiObject -Namespace "root\MicrosoftDNS" -Class MicrosoftDNS_Zone -ComputerName $dnsServer

Write-Host ""
Write-Host "=== ZONAS DNS NO SERVIDOR $dnsServer (RESUMO) ===" -ForegroundColor Yellow

$zones |
    Select-Object `
        Name, `
        ZoneType, `
        @{Name = "DsIntegrated"; Expression = { $_.DsIntegrated }}, `
        @{Name = "ReverseLookup"; Expression = { $_.ReverseLookupZone }} |
    Format-Table -AutoSize

# Exporta todas as propriedades (raw)
$zones |
    Select-Object * |
    Export-Csv "$exportPath\dns_zones_raw.csv" -NoTypeInformation -Encoding UTF8


##############################
# 2) SRV relacionados a _msdcs -> dns_msdcs_srv_all_raw.csv
##############################

$msdcsSrv = Get-WmiObject -Namespace "root\MicrosoftDNS" -Class MicrosoftDNS_SRVType -ComputerName $dnsServer |
    Where-Object {
        $_.OwnerName   -like "*_msdcs.*" -or
        $_.OwnerName   -like "*._msdcs"  -or
        $_.DomainName  -like "*_msdcs.*"
    }

Write-Host ""
Write-Host "=== REGISTROS SRV RELACIONADOS A _msdcs EM $dnsServer (RESUMO) ===" -ForegroundColor Yellow

$msdcsSrv |
    Select-Object `
        OwnerName, `
        Priority, `
        Weight, `
        Port, `
        SRVDomainName |
    Sort-Object OwnerName, Priority |
    Format-Table -AutoSize

# Exporta todas as propriedades (raw)
$msdcsSrv |
    Select-Object * |
    Export-Csv "$exportPath\dns_msdcs_srv_all_raw.csv" -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Arquivos gerados em $exportPath :" -ForegroundColor Cyan
Write-Host " - dns_zones_raw.csv"
Write-Host " - dns_msdcs_srv_all_raw.csv"
