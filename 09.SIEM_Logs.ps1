# ============================
# 09 - Logs SIEM / Security
#   - Security logs 7 e 30 dias (filtrados por EventID)
#   - Resumo na tela + CSV detalhado
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "09_Logs_SIEM_Eventos"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Import-Module ActiveDirectory

Write-Host "Coletando Security Logs (7 e 30 dias, filtrados por EventID) de todos os Domain Controllers..." -ForegroundColor Cyan

# Lista de DCs (hostname/FQDN)
$dcList = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName

# EventIDs relevantes (ajuste se quiser)
$eventIds = @(
    4624, # logon bem-sucedido
    4625, # logon falho
    4634, # logoff
    4647, # logoff iniciado
    4672, # logon com privilégios especiais
    4768, # TGT emitido
    4769, # TGS emitido
    4771, # falha Kerberos
    4776, # NTLM logon
    4740, # conta bloqueada
    4720, # conta criada
    4722, # conta habilitada
    4723, # tentativa mudança de senha
    4724, # reset de senha
    4725, # conta desabilitada
    4726, # conta deletada
    4728, # adição a grupo global
    4732, # adição a grupo local
    4756  # adição a grupo universal
)

$summary = @()
$start30 = (Get-Date).AddDays(-30)
$start7  = (Get-Date).AddDays(-7)

foreach ($dc in $dcList) {
    $safeName = $dc -replace '[\\\/:\*\?\"<>\|]', '_'

    $out7  = Join-Path $OutFolder "security_events_7d_$safeName.csv"
    $out30 = Join-Path $OutFolder "security_events_30d_$safeName.csv"

    Write-Host ""
    Write-Host "DC: $dc" -ForegroundColor Yellow

    try {
        Write-Host "  → Coletando últimos 30 dias (filtrado por EventID)..." -ForegroundColor Green

        # UMA única leitura por DC (30 dias + EventIDs filtrados)
        $events = Get-WinEvent -ComputerName $dc -FilterHashtable @{
            LogName   = 'Security'
            StartTime = $start30
            Id        = $eventIds
        } -MaxEvents 50000 -ErrorAction Stop |
        Select-Object TimeCreated, Id, ProviderName, LevelDisplayName, MachineName, Message

        $count30 = $events.Count

        # Separa os últimos 7 dias a partir do mesmo conjunto
        $events7 = $events | Where-Object { $_.TimeCreated -ge $start7 }
        $count7  = $events7.Count

        # Exporta
        $events7  | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $out7
        $events   | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $out30

        $summary += [PSCustomObject]@{
            DomainController = $dc
            Events7Days      = $count7
            Events30Days     = $count30
            File7Days        = $out7
            File30Days       = $out30
        }

        Write-Host ("  ✓ 7 dias: {0} eventos | 30 dias: {1} eventos" -f $count7, $count30) -ForegroundColor Cyan
        Write-Host "    Arquivos: " -NoNewline
        Write-Host "$out7"  -ForegroundColor DarkCyan
        Write-Host "              $out30" -ForegroundColor DarkCyan
    }
    catch {
        Write-Host ("  ! Falha ao coletar eventos do DC {0}: {1}" -f $dc, $_) -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=== RESUMO DE EVENTOS DE SEGURANÇA (7 / 30 DIAS, FILTRADOS) ===" -ForegroundColor Cyan

if ($summary.Count -gt 0) {
    $summary |
        Select-Object DomainController,Events7Days,Events30Days |
        Sort-Object DomainController |
        Format-Table -AutoSize
} else {
    Write-Host "Nenhum dado coletado. Verifique conectividade/permissões nos DCs." -ForegroundColor Red
}

Write-Host ""
Write-Host "Arquivos detalhados gerados em: $OutFolder" -ForegroundColor Cyan
Write-Host "  - security_events_7d_<DC>.csv" -ForegroundColor DarkCyan
Write-Host "  - security_events_30d_<DC>.csv" -ForegroundColor DarkCyan
Write-Host ""
