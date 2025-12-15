param(
    # Quantos dias considerar como "ativo" pelo último logon
    [int]$DiasLogon = 90,

    # Pasta onde será gerado o CSV
    [string]$ExportFolder = "C:\AD-Export\Legacy_Clients",

    # Se quiser pular o teste de ping, use -SemPing
    [switch]$SemPing
)

# ---------------------------
# Funções de apoio
# ---------------------------
function Convert-ADFileTimeToDateTime {
    param([object]$value)

    if (-not $value) { return $null }

    try {
        $fileTime = [int64]$value
        if ($fileTime -le 0) { return $null }
        return [DateTime]::FromFileTime($fileTime)
    }
    catch {
        return $null
    }
}

# Tenta resolver IPv4 a partir do DNS
function Get-IPv4FromDns {
    param([string]$DnsName)

    if (-not $DnsName) { return $null }

    try {
        $ips = [System.Net.Dns]::GetHostAddresses($DnsName)
        if ($ips -and $ips.Length -gt 0) {
            foreach ($ip in $ips) {
                if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
                    return $ip.IPAddressToString
                }
            }
        }
    } catch {
        return $null
    }
    return $null
}

# ---------------------------
# Garante pasta de export
# ---------------------------
if (-not (Test-Path $ExportFolder)) {
    New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
}

$csvPath = Join-Path $ExportFolder "legacy_clients_report.csv"

Write-Host "Inventário de clientes legados (XP/7/8/8.1) via LDAP/ADSI - compatível com PS 2.0 / 2008" -ForegroundColor Yellow

# ---------------------------
# Descobre o contexto de domínio (SearchBase)
# ---------------------------
try {
    $rootDse = [ADSI]"LDAP://RootDSE"
} catch {
    Write-Host "ERRO: não foi possível ler LDAP://RootDSE. Esta máquina está no domínio?" -ForegroundColor Red
    Write-Host $_
    return
}

$defaultNC = $rootDse.defaultNamingContext
Write-Host ("SearchBase LDAP: {0}" -f $defaultNC) -ForegroundColor Cyan

# ---------------------------
# Monta pesquisa LDAP de computadores legados
# ---------------------------
Write-Host "Procurando computadores com sistemas operacionais legados (até Windows 8.1)..." -ForegroundColor Yellow

# Filtro LDAP (computador + OS legado)
$ldapFilter = "(&(objectClass=computer)(|(operatingSystem=*Windows XP*)(operatingSystem=*Windows 7*)(operatingSystem=*Windows 8*)(operatingSystem=*Windows 8.1*)))"

$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.SearchRoot  = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$defaultNC")
$searcher.SearchScope = "Subtree"
$searcher.PageSize    = 1000
$searcher.Filter      = $ldapFilter

$searcher.PropertiesToLoad.Clear()
@("name","dNSHostName","operatingSystem","operatingSystemVersion","lastLogonTimestamp","userAccountControl") | ForEach-Object {
    [void]$searcher.PropertiesToLoad.Add($_)
}

$results = $searcher.FindAll()

if (-not $results -or $results.Count -eq 0) {
    Write-Host "Nenhum computador legado (XP/7/8/8.1) encontrado no AD com esse filtro." -ForegroundColor Cyan
    return
}

Write-Host ("Total de computadores legados encontrados no AD: {0}" -f $results.Count) -ForegroundColor Cyan

# ---------------------------
# Lógica de classificação
# ---------------------------
$cutoff = (Get-Date).AddDays(-$DiasLogon)
Write-Host ("Considerando como 'logon recente' máquinas com LastLogonTimestamp >= {0}" -f $cutoff.ToString("dd/MM/yyyy")) -ForegroundColor Yellow

$resultados = @()

foreach ($r in $results) {

    $nameProp = $r.Properties["name"]
    if (-not $nameProp -or $nameProp.Count -eq 0) { continue }
    $name = [string]$nameProp[0]

    $dnsProp = $r.Properties["dnshostname"]
    $dns     = $null
    if ($dnsProp -and $dnsProp.Count -gt 0) { $dns = [string]$dnsProp[0] }

    $osProp  = $r.Properties["operatingsystem"]
    $os      = $null
    if ($osProp -and $osProp.Count -gt 0) { $os = [string]$osProp[0] }

    $osVerProp = $r.Properties["operatingsystemversion"]
    $osVer     = $null
    if ($osVerProp -and $osVerProp.Count -gt 0) { $osVer = [string]$osVerProp[0] }

    $uacProp = $r.Properties["useraccountcontrol"]
    $uac     = 0
    if ($uacProp -and $uacProp.Count -gt 0) { $uac = [int]$uacProp[0] }

    # Disabled = flag ACCOUNTDISABLE (0x0002)
    $enabled = (($uac -band 0x2) -eq 0)

    # LastLogonTimestamp
    $lltProp = $r.Properties["lastlogontimestamp"]
    $lastLogon = $null
    if ($lltProp -and $lltProp.Count -gt 0) {
        $lastLogon = Convert-ADFileTimeToDateTime $lltProp[0]
    }

    $logonRecente = $false
    if ($lastLogon) {
        if ($lastLogon -ge $cutoff) {
            $logonRecente = $true
        }
    }

    # IP (opcional, tenta resolver)
    $ip = Get-IPv4FromDns $dns

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

    $result = New-Object PSObject
    $result | Add-Member NoteProperty Name               $name
    $result | Add-Member NoteProperty DNSHostName        $dns
    $result | Add-Member NoteProperty IPv4Address        $ip
    $result | Add-Member NoteProperty OperatingSystem    $os
    $result | Add-Member NoteProperty OperatingSystemVer $osVer
    $result | Add-Member NoteProperty Enabled            $enabled
    $result | Add-Member NoteProperty LastLogonDate      $lastLogon
    $result | Add-Member NoteProperty LogonRecente       $logonRecente
    $result | Add-Member NoteProperty OnlinePing         $online
    $result | Add-Member NoteProperty Status             $status

    $resultados += $result
}

# ---------------------------
# Resumo na tela
# ---------------------------
Write-Host ""
Write-Host "=== RESUMO (primeiros 50 resultados) ===" -ForegroundColor Cyan
$resultados |
    Sort-Object Status, OperatingSystem, Name |
    Select-Object -First 50 Name,OperatingSystem,IPv4Address,LastLogonDate,LogonRecente,OnlinePing,Status |
    Format-Table -AutoSize

# ---------------------------
# Exporta tudo para CSV
# ---------------------------
$resultados |
    Sort-Object Status, OperatingSystem, Name |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

Write-Host ""
Write-Host "Relatório completo exportado para:" -ForegroundColor Green
Write-Host "  $csvPath"
Write-Host ""
Write-Host "Dica: abra o CSV no Excel e filtre por coluna 'Status' / 'OperatingSystem' para evidenciar máquinas ativas/inativas." -ForegroundColor Yellow
