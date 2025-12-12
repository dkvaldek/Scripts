param(
    # Quantos dias considerar como "ativo" pelo último logon
    [int]$DiasLogon = 90,

    # Pasta onde será gerado o CSV
    [string]$ExportFolder = "C:\AD-Export\Legacy_Clients",

    # Se quiser pular o teste de ping, use -SemPing
    [switch]$SemPing
)

# Garante pasta de export
if (-not (Test-Path $ExportFolder)) {
    New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
}

$csvPath = Join-Path $ExportFolder "legacy_clients_report.csv"

Write-Host "Importando módulo ActiveDirectory..." -ForegroundColor Yellow
Import-Module ActiveDirectory -ErrorAction Stop

Write-Host "Procurando computadores com sistemas operacionais legados (até Windows 8.1)..." -ForegroundColor Yellow

# Filtro de OS legados
$filter = @"
OperatingSystem -like '*Windows XP*' -or
OperatingSystem -like '*Windows 7*'  -or
OperatingSystem -like '*Windows 8*'  -or
OperatingSystem -like '*Windows 8.1*'
"@

# Busca no AD
$computers = Get-ADComputer -Filter $filter `
    -Properties OperatingSystem,OperatingSystemVersion,LastLogonDate,LastLogonTimestamp,IPv4Address,Enabled,DNSHostName

if (-not $computers -or $computers.Count -eq 0) {
    Write-Host "Nenhum computador legado (XP/7/8/8.1) encontrado no AD com esse filtro." -ForegroundColor Cyan
    return
}

Write-Host ("Total de computadores legados encontrados no AD: {0}" -f $computers.Count) -ForegroundColor Cyan

# Data de corte para considerar "ativo"
$cutoff = (Get-Date).AddDays(-$DiasLogon)

Write-Host ("Considerando como 'logon recente' máquinas com LastLogonDate >= {0}" -f $cutoff.ToString("dd/MM/yyyy")) -ForegroundColor Yellow

$resultados = @()
$idx = 0

foreach ($comp in $computers) {
    $idx++

    $name     = $comp.Name
    $dns      = $comp.DNSHostName
    $os       = $comp.OperatingSystem
    $osVer    = $comp.OperatingSystemVersion
    $enabled  = $comp.Enabled
    $ip       = $comp.IPv4Address

    # LastLogonDate já é calculado pelo módulo AD com base no lastLogonTimestamp
    $lastLogon = $comp.LastLogonDate

    $logonRecente = $false
    if ($lastLogon) {
        if ($lastLogon -ge $cutoff) {
            $logonRecente = $true
        }
    }

    # Teste de ping (opcional)
    $online = $null
    if ($SemPing) {
        $online = $null
    } else {
        $alvoPing = if ($dns) { $dns } else { $name }

        try {
            $online = Test-Connection -ComputerName $alvoPing -Count 1 -Quiet -ErrorAction SilentlyContinue
        } catch {
            $online = $false
        }
    }

    # Classificação simples
    $status = ""

    if (-not $enabled) {
        $status = "Conta de computador desabilitada"
    }
    elseif ($logonRecente -and $online) {
        $status = "Provavelmente ATIVA (logon recente + responde ping)"
    }
    elseif ($logonRecente -and (-not $online)) {
        $status = "Possivelmente ativa (logon recente, não responde ping)"
    }
    elseif ((-not $logonRecente) -and $online) {
        $status = "Possivelmente obsoleta (sem logon recente, mas online)"
    }
    else {
        $status = "Provavelmente INATIVA (sem logon recente e não responde ping)"
    }

    $result = [PSCustomObject]@{
        Name               = $name
        DNSHostName        = $dns
        IPv4Address        = $ip
        OperatingSystem    = $os
        OperatingSystemVer = $osVer
        Enabled            = $enabled
        LastLogonDate      = $lastLogon
        LogonRecente       = $logonRecente
        OnlinePing         = $online
        Status             = $status
    }

    $resultados += $result
}

# Exibe um resumo na tela (top 50)
Write-Host ""
Write-Host "=== RESUMO (primeiros 50 resultados) ===" -ForegroundColor Cyan
$resultados |
    Sort-Object Status, OperatingSystem, Name |
    Select-Object -First 50 Name,OperatingSystem,IPv4Address,LastLogonDate,LogonRecente,OnlinePing,Status |
    Format-Table -AutoSize

# Exporta tudo para CSV
$resultados |
    Sort-Object Status, OperatingSystem, Name |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

Write-Host ""
Write-Host "Relatório completo exportado para:" -ForegroundColor Green
Write-Host "  $csvPath"
Write-Host ""
Write-Host "Dica: abra o CSV no Excel e filtre por coluna 'Status' / 'OperatingSystem' para evidenciar máquinas ativas/inativas." -ForegroundColor Yellow
