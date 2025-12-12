<#
.SYNOPSIS
  Coleta automatizada (somente leitura) de itens básicos do Active Directory para o Assessment.
.DESCRIPTION
  Script para execução a partir de uma estação administrativa com módulos do RSAT e permissão de leitura no AD.
  NÃO ALTERAR Configurações de produção. O script exporta CSV/TXT para a estrutura de pastas especificada.
.PARAMETER BasePath
  Caminho base onde os arquivos serão salvos. Padrão: C:\Evidencias_Assessment_HospitalCare
.EXAMPLE
  .\main_collect.ps1 -BasePath "C:\Evidencias_Assessment_HospitalCare"
.NOTES
  - Revisar permissões antes de executar.
  - Recomendado executar sob conta de auditoria com privilégios de leitura.
#>
param(
    [string]$BasePath = "C:\Evidencias_Assessment_HospitalCare",
    [switch]$SkipGPOBackup
)

function Ensure-Path {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

# Cria estrutura básica
$folders = @(
    "00_Meta","01_Inventario_DomainControllers","02_AD_DS_Objects","03_FSMO_Replicacao",
    "04_DNS","05_DHCP","06_GPOs\backup_gpos","06_GPOs\GPO_Reports","07_Autenticacao_Metodos",
    "08_EntraSync_Azure","09_Logs_SIEM_Eventos","10_Seguranca_Permissoes_RBAC","11_Rede_Servicos_dependentes","12_Anexos_Extras"
)
foreach ($f in $folders) {
    Ensure-Path (Join-Path $BasePath $f)
}

# Registro inicial
$logFile = Join-Path $BasePath "00_Meta\00_log_execucao.txt"
"StartTime: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" | Out-File -FilePath $logFile -Append
"Host: $env:COMPUTERNAME" | Out-File -FilePath $logFile -Append
"User: $env:USERNAME" | Out-File -FilePath $logFile -Append

# Import modules
Try {
    Import-Module ActiveDirectory -ErrorAction Stop
    "Module ActiveDirectory loaded" | Out-File -FilePath $logFile -Append
} Catch {
    "ActiveDirectory module not available: $_" | Out-File -FilePath $logFile -Append
    Write-Warning "ActiveDirectory module is required. Stop."
    return
}

# 1) DC Inventory
$dcOut = Join-Path $BasePath "01_Inventario_DomainControllers\DCs_list_raw.csv"
Get-ADDomainController -Filter * |
    Select Name,HostName,IPv4Address,Site,OperatingSystem,IsGlobalCatalog,OSVersion |
    Sort-Object Site,Name |
    Export-Csv -NoTypeInformation -Path $dcOut
"Exported DC list to $dcOut" | Out-File -FilePath $logFile -Append

# 2) FSMO roles
$fsmoOut = Join-Path $BasePath "01_Inventario_DomainControllers\fsmo_raw.txt"
netdom query fsmo | Out-File -FilePath $fsmoOut
"Exported FSMO roles to $fsmoOut" | Out-File -FilePath $logFile -Append

# 3) dcdiag and repadmin (may require being on a DC or tools installed)
$dcdiagOut   = Join-Path $BasePath "01_Inventario_DomainControllers\dcdiag_all_raw.txt"
$repadminOut = Join-Path $BasePath "01_Inventario_DomainControllers\repadmin_replsummary.txt"
try {
    dcdiag /v /c /d /e > $dcdiagOut
    "dcdiag exported to $dcdiagOut" | Out-File -FilePath $logFile -Append
} catch {
    "dcdiag failed: $_" | Out-File -FilePath $logFile -Append
}
try {
    repadmin /replsummary > $repadminOut
    "repadmin replsummary exported to $repadminOut" | Out-File -FilePath $logFile -Append
} catch {
    "repadmin failed: $_" | Out-File -FilePath $logFile -Append
}

# 4) AD Objects
$usersOut = Join-Path $BasePath "02_AD_DS_Objects\users_all_raw.csv"
Get-ADUser -Filter * -Properties Enabled,LastLogonDate,PasswordLastSet,whenCreated,proxyAddresses |
    Select SamAccountName,Name,Enabled,LastLogonDate,PasswordLastSet,whenCreated,DistinguishedName,proxyAddresses |
    Export-Csv -NoTypeInformation -Path $usersOut
"Exported users to $usersOut" | Out-File -FilePath $logFile -Append

$computersOut = Join-Path $BasePath "02_AD_DS_Objects\computers_all_raw.csv"
Get-ADComputer -Filter * -Properties OperatingSystem,LastLogonDate,whenCreated |
    Select Name,OperatingSystem,LastLogonDate,whenCreated,DistinguishedName |
    Export-Csv -NoTypeInformation -Path $computersOut
"Exported computers to $computersOut" | Out-File -FilePath $logFile -Append

$groupsOut = Join-Path $BasePath "02_AD_DS_Objects\groups_all_raw.csv"
Get-ADGroup -Filter * -Properties GroupCategory,GroupScope,whenCreated |
    Select Name,GroupScope,GroupCategory,whenCreated,DistinguishedName |
    Export-Csv -NoTypeInformation -Path $groupsOut
"Exported groups to $groupsOut" | Out-File -FilePath $logFile -Append

# 5) Privileged groups members
$privOut = Join-Path $BasePath "02_AD_DS_Objects\privileged_group_members.csv"
$admins  = @('Domain Admins','Enterprise Admins','Schema Admins','Administrators')
if (Test-Path $privOut) { Remove-Item $privOut -Force }
foreach ($g in $admins) {
    Try {
        Get-ADGroupMember -Identity $g -Recursive |
            Select @{n='Group';e={$g}},Name,SamAccountName,DistinguishedName |
            Export-Csv -NoTypeInformation -Path $privOut -Append
    } Catch {
        # <-- LINHA CORRIGIDA AQUI
        ("Failed to enumerate group {0}: {1}" -f $g, $_) | Out-File -FilePath $logFile -Append
    }
}
"Exported privileged groups to $privOut" | Out-File -FilePath $logFile -Append

# 6) GPOs (list + reports). Skip backup if SkipGPOBackup switch used.
Try {
    Import-Module GroupPolicy -ErrorAction Stop
    $gpolist = Join-Path $BasePath "06_GPOs\gpo_list.csv"
    if (Test-Path $gpolist) { Remove-Item $gpolist -Force }
    $gpos = Get-GPO -All
    foreach ($g in $gpos) {
        $safeName = ($g.DisplayName -replace '[\\/:*?"<>|]','_')
        if (-not $SkipGPOBackup) {
            Try {
                Backup-GPO -Guid $g.Id -Path (Join-Path $BasePath "06_GPOs\backup_gpos")
            } Catch {
                "Failed to backup GPO $($g.DisplayName): $_" | Out-File -FilePath $logFile -Append
            }
        }
        $reportHtml = Join-Path $BasePath "06_GPOs\GPO_Reports\$safeName.html"
        Try {
            Get-GPOReport -Guid $g.Id -ReportType Html -Path $reportHtml
        } Catch {
            "Failed to generate report for $($g.DisplayName): $_" | Out-File -FilePath $logFile -Append
        }
        $g | Select DisplayName,Id,CreatedTime,ModifiedTime |
            Export-Csv -NoTypeInformation -Path $gpolist -Append
    }
    "GPO export completed" | Out-File -FilePath $logFile -Append
} Catch {
    "GroupPolicy module not available: $_" | Out-File -FilePath $logFile -Append
}

# 7) DNS zones (attempt using first DC)
Try {
    $firstDC = (Get-ADDomainController -Filter * | Select -First 1 -ExpandProperty HostName)
    $dnsOut  = Join-Path $BasePath "04_DNS\dns_zones_raw.csv"
    Get-DnsServerZone -ComputerName $firstDC |
        Select ZoneName,ZoneType,IsDsIntegrated,ReplicationScope,IsPaused |
        Export-Csv -NoTypeInformation -Path $dnsOut
    "Exported DNS zones to $dnsOut (using $firstDC)" | Out-File -FilePath $logFile -Append
} Catch {
    "DNS export failed: $_" | Out-File -FilePath $logFile -Append
}

# 8) AD Connect (if module present)
Try {
    Import-Module ADSync -ErrorAction Stop
    $connectorsOut = Join-Path $BasePath "08_EntraSync_Azure\connectors.csv"
    Get-ADSyncConnector | Export-Csv -NoTypeInformation -Path $connectorsOut
    "Exported ADSync connectors to $connectorsOut" | Out-File -FilePath $logFile -Append
} Catch {
    "ADSync not available or module not present" | Out-File -FilePath $logFile -Append
}

# 9) Sample event export (Security events last 7 days) - careful with volume
try {
    $dcList = (Get-ADDomainController -Filter * | Select -ExpandProperty HostName)
    foreach ($dc in $dcList) {
        $evOut = Join-Path $BasePath "09_Logs_SIEM_Eventos\security_events_$dc.csv"
        try {
            Get-WinEvent -ComputerName $dc -FilterHashtable @{LogName='Security';StartTime=(Get-Date).AddDays(-7)} -ErrorAction Stop |
                Export-Csv -NoTypeInformation -Path $evOut
            "Exported security events for $dc to $evOut" | Out-File -FilePath $logFile -Append
        } catch {
            # <-- LINHA CORRIGIDA AQUI
            ("Failed to export events from {0}: {1}" -f $dc, $_) | Out-File -FilePath $logFile -Append
        }
    }
} catch {
    "Event export section failed: $_" | Out-File -FilePath $logFile -Append
}

# Final log entry
"EndTime: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" | Out-File -FilePath $logFile -Append
"Script finished." | Out-File -FilePath $logFile -Append

Write-Output "Collection complete. Check $BasePath for outputs."
