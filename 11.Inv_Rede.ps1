# ============================
# 11 - Rede / Serviços dependentes
#   - Inventário de rede (servers)
#   - Inventário de aplicativos (roles/features)
#   - Resumo na tela + CSV detalhado
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "11_Rede_Servicos_dependentes"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Import-Module ActiveDirectory

Write-Host "Coletando inventário de rede e aplicativos (roles/features) dos servidores..." -ForegroundColor Cyan

# Lista de servers no AD
$servers = Get-ADComputer -Filter "OperatingSystem -like '*Server*'" `
    -Properties OperatingSystem,IPv4Address,LastLogonDate,DNSHostName

if (-not $servers -or $servers.Count -eq 0) {
    Write-Host "Nenhum computador com OperatingSystem contendo 'Server' encontrado no AD." -ForegroundColor Yellow
    return
}

$networkResults  = @()
$featureResults  = @()
$offlineServers  = @()

$idx    = 0
$total  = $servers.Count

foreach ($s in $servers) {
    $idx++

    $dnsName = if ($s.DNSHostName) { $s.DNSHostName } else { $s.Name }
    Write-Host ("[{0}/{1}] Coletando de {2} ..." -f $idx, $total, $dnsName) -ForegroundColor Yellow

    try {
        # ---- WMI: OS ----
        $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $dnsName -ErrorAction Stop

        # ---- WMI: Network ----
        $nets = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $dnsName `
                -Filter "IPEnabled=TRUE" -ErrorAction SilentlyContinue

        $ipList = @()
        $gwList = @()

        if ($nets) {
            foreach ($n in $nets) {
                if ($n.IPAddress) {
                    $ipList += $n.IPAddress | Where-Object { $_ -match '\.' }  # só IPv4
                }
                if ($n.DefaultIPGateway) {
                    $gwList += $n.DefaultIPGateway | Where-Object { $_ -match '\.' }
                }
            }
        }

        # Converte LastBootUpTime
        $bootTime = $null
        if ($os.LastBootUpTime) {
            $bootTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
        }

        $networkResults += [PSCustomObject]@{
            ComputerName      = $s.Name
            DNSHostName       = $dnsName
            OperatingSystem   = $s.OperatingSystem
            OSCaption         = $os.Caption
            OSVersion         = $os.Version
            LastLogonDate     = $s.LastLogonDate
            LastBootUpTime    = $bootTime
            IPAddresses       = ($ipList -join ';')
            DefaultGateways   = ($gwList -join ';')
        }

        # ---- WMI: Roles / Features ----
        $features = Get-WmiObject -Class Win32_ServerFeature -ComputerName $dnsName -ErrorAction SilentlyContinue
        if ($features) {
            foreach ($f in $features) {
                $featureResults += [PSCustomObject]@{
                    ComputerName = $s.Name
                    DNSHostName  = $dnsName
                    FeatureID    = $f.ID
                    FeatureName  = $f.Name
                    ParentID     = $f.ParentID
                }
            }
        }
    }
    catch {
        Write-Host ("  ! Servidor OFFLINE ou inacessível: {0} ({1})" -f $dnsName, $_.Exception.Message) -ForegroundColor Red
        $offlineServers += $dnsName
        continue
    }
}

# ----- Exporta CSVs -----

$netFile  = Join-Path $OutFolder "network_inventory_servers.csv"
$appFile  = Join-Path $OutFolder "apps_server_features_inventory.csv"
$offFile  = Join-Path $OutFolder "network_inventory_offline_servers.txt"

if ($networkResults.Count -gt 0) {
    $networkResults |
        Sort-Object DNSHostName |
        Export-Csv -NoTypeInformation -Encoding UTF8 -Path $netFile
}

if ($featureResults.Count -gt 0) {
    $featureResults |
        Sort-Object DNSHostName,FeatureID |
        Export-Csv -NoTypeInformation -Encoding UTF8 -Path $appFile
}

if ($offlineServers.Count -gt 0) {
    $offlineServers |
        Sort-Object |
        Out-File -FilePath $offFile -Encoding UTF8
}

# ----- Resumo na tela -----

Write-Host ""
Write-Host "=== RESUMO DE INVENTÁRIO DE REDE / APLICATIVOS ===" -ForegroundColor Cyan

Write-Host ("Total de servers encontrados no AD: {0}" -f $servers.Count)
Write-Host ("Servers com inventário coletado:    {0}" -f $networkResults.Count)
Write-Host ("Servers offline/inacessíveis:       {0}" -f $offlineServers.Count)

Write-Host ""
Write-Host "Top sistemas operacionais (por contagem):" -ForegroundColor Yellow
$networkResults |
    Group-Object OperatingSystem |
    Sort-Object Count -Descending |
    Select-Object -First 10 |
    Format-Table Name,Count -AutoSize

if ($featureResults.Count -gt 0) {
    Write-Host ""
    Write-Host "Top roles/features instalados (por número de servidores):" -ForegroundColor Yellow
    $featureResults |
        Group-Object FeatureName |
        Sort-Object Count -Descending |
        Select-Object -First 15 |
        Format-Table Name,Count -AutoSize
}

Write-Host ""
Write-Host "Arquivos detalhados gerados em: $OutFolder" -ForegroundColor Cyan
Write-Host "  - network_inventory_servers.csv" -ForegroundColor DarkCyan
Write-Host "  - apps_server_features_inventory.csv" -ForegroundColor DarkCyan
if ($offlineServers.Count -gt 0) {
    Write-Host "  - network_inventory_offline_servers.txt" -ForegroundColor DarkCyan
}
