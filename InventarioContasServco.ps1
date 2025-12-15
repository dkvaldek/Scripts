# ==========================================================
# Inventário de contas de serviço / contas usadas em serviços
# Compatível com Windows Server 2008+ / PowerShell 2.0+
# - NÃO usa módulo ActiveDirectory
# - Usa LDAP (.NET) para consultar usuários
# - Usa WMI (Win32_Service) no servidor local para achar contas
#   usadas em serviços (StartName)
# ==========================================================

param(
    # Pasta para exportar o CSV
    [string]$ExportFolder = "C:\AD-Export\ServiceAccounts",

    # Opcional: DN base para busca (se vazio, usa o domínio padrão)
    [string]$SearchBase = $null
)

function Write-Yellow {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Yellow
}

function Write-Cyan {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Cyan
}

function Convert-ADFileTimeToDateTime {
    param([object]$value)

    if (-not $value) { return $null }

    try {
        # lastLogonTimestamp vem como inteiro de FileTime
        $fileTime = [int64]$value
        if ($fileTime -le 0) { return $null }
        return [DateTime]::FromFileTime($fileTime)
    }
    catch {
        return $null
    }
}

# -----------------------------------------
# 1) Garantir pasta de export
# -----------------------------------------
if (-not (Test-Path $ExportFolder)) {
    New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
}
$csvPath = Join-Path $ExportFolder "service_accounts_report.csv"

Write-Yellow "Inventário de contas de serviço / contas usadas em serviços (AD + serviços locais)"

# -----------------------------------------
# 2) Descobrir SearchBase (domínio)
# -----------------------------------------
try {
    $rootDse = [ADSI]"LDAP://RootDSE"
} catch {
    Write-Host "ERRO: não foi possível ler LDAP://RootDSE. Esta máquina está no domínio?" -ForegroundColor Red
    Write-Host $_
    return
}

if ([string]::IsNullOrEmpty($SearchBase)) {
    $defaultNC = $rootDse.defaultNamingContext
    $SearchBase = $defaultNC
}

Write-Cyan ("SearchBase LDAP: " + $SearchBase)

# -----------------------------------------
# 3) Mapear contas usadas em serviços nesse servidor
# -----------------------------------------
Write-Yellow "Coletando serviços locais e contas associadas (Win32_Service)..."

$serviceAccounts = @{}   # chave: samAccountName (lower), valor: lista de serviços

try {
    $services = Get-WmiObject -Class Win32_Service -ErrorAction SilentlyContinue
} catch {
    Write-Host "ERRO ao consultar Win32_Service via WMI: $_" -ForegroundColor Red
    $services = @()
}

if ($services -and $services.Count -gt 0) {

    foreach ($svc in $services) {
        $startName = $svc.StartName

        if (-not $startName) { continue }

        # Ignora contas padrão do sistema
        $snLower = $startName.ToLower()
        if ($snLower -like "localsystem" -or
            $snLower -like "nt authority\*" -or
            $snLower -like "localservice" -or
            $snLower -like "networkservice") {
            continue
        }

        $samFromService = $null

        # Formato DOMAIN\user
        if ($startName -match '^(?<dom>[^\\]+)\\(?<user>.+)$') {
            $samFromService = $matches['user']
        }
        # Formato user@dominio
        elseif ($startName -match '^(?<user>[^@]+)@(?<dom>.+)$') {
            $samFromService = $matches['user']
        }

        if ($samFromService) {
            $key = $samFromService.ToLower()
            if (-not $serviceAccounts.ContainsKey($key)) {
                $serviceAccounts[$key] = New-Object System.Collections.ArrayList
            }
            [void]$serviceAccounts[$key].Add($svc.Name + " (" + $svc.DisplayName + ")")
        }
    }
}

Write-Cyan ("Contas diferentes encontradas em serviços locais: " + $serviceAccounts.Keys.Count)

# -----------------------------------------
# 4) Buscar usuários no AD (via LDAP)
# -----------------------------------------
Write-Yellow "Consultando usuários no AD (LDAP) e identificando possíveis contas de serviço..."

$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searchRootPath = "LDAP://$SearchBase"
$searcher.SearchRoot  = New-Object System.DirectoryServices.DirectoryEntry($searchRootPath)
$searcher.SearchScope = "Subtree"
$searcher.PageSize    = 1000

# Somente usuários (sem computadores)
$searcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(objectClass=computer)))"

$searcher.PropertiesToLoad.Clear()
# atributos que vamos usar
$props = @(
    "samAccountName",
    "userPrincipalName",
    "displayName",
    "distinguishedName",
    "userAccountControl",
    "servicePrincipalName",
    "lastLogonTimestamp",
    "description",
    "whenCreated"
)
foreach ($p in $props) {
    [void]$searcher.PropertiesToLoad.Add($p)
}

$results = $searcher.FindAll()

if (-not $results -or $results.Count -eq 0) {
    Write-Host "Nenhum usuário retornado da consulta LDAP (verifique permissões / filtro)." -ForegroundColor Red
    return
}

Write-Cyan ("Total de usuários retornados: " + $results.Count)

$serviceLike = @()
$idx = 0

foreach ($r in $results) {
    $idx++

    $samProp   = $r.Properties["samaccountname"]
    if (-not $samProp -or $samProp.Count -eq 0) { continue }
    $sam       = [string]$samProp[0]

    $upnProp   = $r.Properties["userprincipalname"]
    $upn       = $null
    if ($upnProp -and $upnProp.Count -gt 0) { $upn = [string]$upnProp[0] }

    $dnProp    = $r.Properties["distinguishedname"]
    $dn        = $null
    if ($dnProp -and $dnProp.Count -gt 0) { $dn = [string]$dnProp[0] }

    $dispProp  = $r.Properties["displayname"]
    $display   = $null
    if ($dispProp -and $dispProp.Count -gt 0) { $display = [string]$dispProp[0] }

    $descProp  = $r.Properties["description"]
    $desc      = $null
    if ($descProp -and $descProp.Count -gt 0) { $desc = [string]$descProp[0] }

    $wcProp    = $r.Properties["whencreated"]
    $whenCreated = $null
    if ($wcProp -and $wcProp.Count -gt 0) { $whenCreated = [datetime]$wcProp[0] }

    $uacProp   = $r.Properties["useraccountcontrol"]
    $uac       = 0
    if ($uacProp -and $uacProp.Count -gt 0) { $uac = [int]$uacProp[0] }

    $spnProp   = $r.Properties["serviceprincipalname"]
    $spnCount  = 0
    if ($spnProp) { $spnCount = $spnProp.Count }

    $lltProp   = $r.Properties["lastlogontimestamp"]
    $lastLogonDate = $null
    if ($lltProp -and $lltProp.Count -gt 0) {
        $lastLogonDate = Convert-ADFileTimeToDateTime $lltProp[0]
    }

    # Flags de userAccountControl
    # 0x0002  = ACCOUNTDISABLE
    # 0x10000 = DONT_EXPIRE_PASSWORD
    $isDisabled      = (($uac -band 0x2) -ne 0)
    $pwdNoExpire     = (($uac -band 0x10000) -ne 0)

    $samLower = $sam.ToLower()
    $usedOnThisServer = $false
    $svcListText      = $null

    if ($serviceAccounts.ContainsKey($samLower)) {
        $usedOnThisServer = $true
        $svcListText = ($serviceAccounts[$samLower] -join "; ")
    }

    # Heurística: conta de serviço "provável" se:
    # - Tem SPN, OU
    # - Senha não expira, OU
    # - Está sendo usada em serviços nesse servidor
    $isServiceLike = $false
    if ($spnCount -gt 0 -or $pwdNoExpire -or $usedOnThisServer) {
        $isServiceLike = $true
    }

    if (-not $isServiceLike) {
        continue
    }

    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "SamAccountName"           -Value $sam
    $obj | Add-Member -MemberType NoteProperty -Name "DisplayName"             -Value $display
    $obj | Add-Member -MemberType NoteProperty -Name "UserPrincipalName"       -Value $upn
    $obj | Add-Member -MemberType NoteProperty -Name "DistinguishedName"       -Value $dn
    $obj | Add-Member -MemberType NoteProperty -Name "Enabled"                 -Value (-not $isDisabled)
    $obj | Add-Member -MemberType NoteProperty -Name "LastLogonDate"           -Value $lastLogonDate
    $obj | Add-Member -MemberType NoteProperty -Name "HasSPN"                  -Value ($spnCount -gt 0)
    $obj | Add-Member -MemberType NoteProperty -Name "SPNCount"                -Value $spnCount
    $obj | Add-Member -MemberType NoteProperty -Name "PasswordNoExpire"        -Value $pwdNoExpire
    $obj | Add-Member -MemberType NoteProperty -Name "UsedAsServiceOnThisHost" -Value $usedOnThisServer
    $obj | Add-Member -MemberType NoteProperty -Name "ServicesOnThisHost"      -Value $svcListText
    $obj | Add-Member -MemberType NoteProperty -Name "WhenCreated"             -Value $whenCreated
    $obj | Add-Member -MemberType NoteProperty -Name "Description"             -Value $desc

    $serviceLike += $obj
}

Write-Host ""
Write-Cyan ("Total de contas de serviço 'prováveis' encontradas: " + $serviceLike.Count)

if ($serviceLike.Count -eq 0) {
    Write-Host "Nenhuma conta de serviço candidata encontrada com os critérios atuais (SPN / senha não expira / usada em serviços locais)." -ForegroundColor Yellow
    return
}

# -----------------------------------------
# 5) Resumo na tela
# -----------------------------------------
Write-Host ""
Write-Host "=== RESUMO (primeiras 50 contas) ===" -ForegroundColor Cyan

$serviceLike |
    Sort-Object -Property UsedAsServiceOnThisHost, HasSPN, SamAccountName -Descending |
    Select-Object -First 50 SamAccountName,Enabled,LastLogonDate,HasSPN,PasswordNoExpire,UsedAsServiceOnThisHost |
    Format-Table -AutoSize

# -----------------------------------------
# 6) Exportar CSV
# -----------------------------------------
$serviceLike |
    Sort-Object -Property SamAccountName |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

Write-Host ""
Write-Host "Relatório completo exportado para:" -ForegroundColor Green
Write-Host "  $csvPath"
Write-Host ""
Write-Host "Dica: abra no Excel e filtre por Enabled, LastLogonDate, HasSPN, PasswordNoExpire e UsedAsServiceOnThisHost." -ForegroundColor Yellow
