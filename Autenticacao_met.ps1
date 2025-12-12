# ============================
# 07 - Métodos de Autenticação (sem Get-AD*)
#   - SPN list
#   - Default domain password policy
#   - Fine-grained password policies
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "07_Autenticacao_Metodos"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Write-Host "Coletando informações de SPN e políticas de autenticação via LDAP/.NET..." -ForegroundColor Cyan

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

# RootDSE via LDAP (sem Get-ADRootDSE)
try {
    $rootDse   = [ADSI]"LDAP://RootDSE"
    $defaultNC = $rootDse.defaultNamingContext
    $configNC  = $rootDse.configurationNamingContext
}
catch {
    Write-Host "ERRO: não foi possível obter RootDSE. Esta máquina está no domínio e alcança um DC por LDAP?" -ForegroundColor Red
    Write-Host $_
    return
}

function Convert-LargeIntegerToTimeSpan {
    param([object]$large)

    if (-not $large) { return $null }

    try {
        $high  = $large.HighPart
        $low   = $large.LowPart
        $ticks = [int64]$high -shl 32 -bor ([uint32]$low)

        # AD armazena valores negativos -> usamos o módulo
        $ticks = [math]::Abs($ticks)

        return [TimeSpan]::FromTicks($ticks)
    }
    catch {
        return $null
    }
}

# ------------------------------------------------------------------
# 1) SPN LIST
# ------------------------------------------------------------------

Write-Host ""
Write-Host "=== SPN LIST (RESUMO) ===" -ForegroundColor Yellow

$spnFile = Join-Path $OutFolder "spn_all_raw.csv"

# Busca todos objetos com servicePrincipalName
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.SearchRoot  = [ADSI]("LDAP://$defaultNC")
$searcher.Filter      = "(servicePrincipalName=*)"
$searcher.SearchScope = "Subtree"
$searcher.PageSize    = 1000
$searcher.PropertiesToLoad.Clear()
$searcher.PropertiesToLoad.Add("servicePrincipalName") | Out-Null
$searcher.PropertiesToLoad.Add("sAMAccountName")       | Out-Null
$searcher.PropertiesToLoad.Add("objectClass")          | Out-Null
$searcher.PropertiesToLoad.Add("distinguishedName")    | Out-Null

$results = $searcher.FindAll()
$spnObjects = @()

foreach ($r in $results) {
    $spns = @()
    if ($r.Properties["serviceprincipalname"].Count -gt 0) {
        $spns = @($r.Properties["serviceprincipalname"])
    }

    $sam = $null
    if ($r.Properties["samaccountname"].Count -gt 0) {
        $sam = $r.Properties["samaccountname"][0]
    }

    $cls = $null
    if ($r.Properties["objectclass"].Count -gt 0) {
        # último objectClass é o mais específico
        $cls = $r.Properties["objectclass"][-1]
    }

    $dn = $null
    if ($r.Properties["distinguishedname"].Count -gt 0) {
        $dn = $r.Properties["distinguishedname"][0]
    }

    $spnObjects += [PSCustomObject]@{
        sAMAccountName       = $sam
        ObjectClass          = $cls
        DistinguishedName    = $dn
        ServicePrincipalName = $spns
    }
}
$results.Dispose()

# Resumo na tela: top 20 objetos com mais SPNs
$spnObjects |
    Select-Object `
        sAMAccountName,
        ObjectClass,
        @{Name="SPNCount"; Expression={ ($_.ServicePrincipalName).Count }} |
    Sort-Object SPNCount -Descending |
    Select-Object -First 20 |
    Format-Table -AutoSize

# CSV detalhado
$spnObjects |
    Select-Object `
        sAMAccountName,
        ObjectClass,
        DistinguishedName,
        @{Name="ServicePrincipalNames"; Expression={ ($_.ServicePrincipalName -join ";") }} |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $spnFile

Write-Host "Arquivo SPN detalhado: $spnFile" -ForegroundColor Green


# ------------------------------------------------------------------
# 2) DEFAULT DOMAIN PASSWORD / KERBEROS POLICY
# ------------------------------------------------------------------

Write-Host ""
Write-Host "=== DEFAULT DOMAIN PASSWORD / KERBEROS POLICY (RESUMO) ===" -ForegroundColor Yellow

$domainEntry = [ADSI]("LDAP://$defaultNC")

# Atributos de tempo (LargeInteger) -> usar .Value
$maxPwdAgeTS  = if ($domainEntry.maxPwdAge -ne $null)  { Convert-LargeIntegerToTimeSpan $domainEntry.maxPwdAge.Value }  else { $null }
$minPwdAgeTS  = if ($domainEntry.minPwdAge -ne $null)  { Convert-LargeIntegerToTimeSpan $domainEntry.minPwdAge.Value }  else { $null }
$lockDurTS    = if ($domainEntry.lockoutDuration -ne $null) { Convert-LargeIntegerToTimeSpan $domainEntry.lockoutDuration.Value } else { $null }
$lockObsTS    = if ($domainEntry.lockoutObservationWindow -ne $null) { Convert-LargeIntegerToTimeSpan $domainEntry.lockoutObservationWindow.Value } else { $null }

# Atributos inteiros -> usar .Value
$minPwdLength      = if ($domainEntry.minPwdLength      -ne $null) { [int]$domainEntry.minPwdLength.Value }      else { $null }
$pwdHistoryLength  = if ($domainEntry.pwdHistoryLength  -ne $null) { [int]$domainEntry.pwdHistoryLength.Value }  else { $null }
$lockoutThreshold  = if ($domainEntry.lockoutThreshold  -ne $null) { [int]$domainEntry.lockoutThreshold.Value }  else { $null }

# pwdProperties também é coleção -> usar .Value
$pwdProps = 0
if ($domainEntry.pwdProperties -ne $null) {
    $pwdProps = [int]$domainEntry.pwdProperties.Value
}

# bits de pwdProperties:
# 1  -> complexidade
# 16 -> armazenar senha reversível
$complexEnabled = (($pwdProps -band 1)  -ne 0)
$reversible     = (($pwdProps -band 16) -ne 0)

$defaultPolicyObj = [PSCustomObject]@{
    DomainDN                    = $defaultNC
    MinPasswordLength           = $minPwdLength
    PasswordHistoryCount        = $pwdHistoryLength
    MaxPasswordAge_Days         = if ($maxPwdAgeTS) { [int]$maxPwdAgeTS.TotalDays } else { $null }
    MinPasswordAge_Days         = if ($minPwdAgeTS) { [int]$minPwdAgeTS.TotalDays } else { $null }
    ComplexityEnabled           = $complexEnabled
    ReversibleEncryptionEnabled = $reversible
    LockoutThreshold            = $lockoutThreshold
    LockoutDuration_Minutes     = if ($lockDurTS) { [int]$lockDurTS.TotalMinutes } else { $null }
    LockoutObservation_Minutes  = if ($lockObsTS) { [int]$lockObsTS.TotalMinutes } else { $null }
    Raw_pwdProperties           = $pwdProps
}

# Resumo na tela
$defaultPolicyObj |
    Select-Object `
        MinPasswordLength,
        PasswordHistoryCount,
        MaxPasswordAge_Days,
        MinPasswordAge_Days,
        ComplexityEnabled,
        ReversibleEncryptionEnabled,
        LockoutThreshold,
        LockoutDuration_Minutes,
        LockoutObservation_Minutes |
    Format-List

# Detalhado em TXT
$defaultPolicyFile = Join-Path $OutFolder "auth_default_domain_password_policy.txt"
$defaultPolicyObj |
    Format-List * |
    Out-File -FilePath $defaultPolicyFile -Encoding UTF8

Write-Host "Arquivo de policy padrão: $defaultPolicyFile" -ForegroundColor Green


# ------------------------------------------------------------------
# 3) FINE-GRAINED PASSWORD POLICIES (FGPP)
# ------------------------------------------------------------------

Write-Host ""
Write-Host "=== FINE-GRAINED PASSWORD POLICIES (RESUMO) ===" -ForegroundColor Yellow

$fineFile = Join-Path $OutFolder "auth_fine_grained_policies.csv"

# Container padrão de FGPP
$fgppBase = "CN=Password Settings Container,CN=System,$defaultNC"

$finePolicies = @()

try {
    $fgEntry = [ADSI]("LDAP://$fgppBase")  # se não existir, vai dar erro
}
catch {
    $fgEntry = $null
}

if (-not $fgEntry) {
    Write-Host "Nenhuma fine-grained password policy encontrada (container não existe)." -ForegroundColor DarkYellow
}
else {
    $searcherFG = New-Object System.DirectoryServices.DirectorySearcher
    $searcherFG.SearchRoot  = $fgEntry
    $searcherFG.Filter      = "(objectClass=msDS-PasswordSettings)"
    $searcherFG.SearchScope = "Subtree"
    $searcherFG.PageSize    = 1000
    $searcherFG.PropertiesToLoad.Clear()
    $propsFG = @(
        "name",
        "msDS-PasswordSettingsPrecedence",
        "minPwdLength",
        "pwdHistoryLength",
        "maxPwdAge",
        "minPwdAge",
        "lockoutThreshold",
        "lockoutDuration",
        "lockoutObservationWindow",
        "msDS-PSOAppliesTo"
    )
    foreach ($p in $propsFG) { [void]$searcherFG.PropertiesToLoad.Add($p) }

    $resultsFG = $searcherFG.FindAll()

    foreach ($r in $resultsFG) {
        $name   = $r.Properties["name"][0]
        $prec   = if ($r.Properties["msdspasswordsettingsprecedence"].Count -gt 0) { [int]$r.Properties["msdspasswordsettingsprecedence"][0] } else { $null }
        $minLen = if ($r.Properties["minpwdlength"].Count -gt 0) { [int]$r.Properties["minpwdlength"][0] } else { $null }
        $hist   = if ($r.Properties["pwdhistorylength"].Count -gt 0) { [int]$r.Properties["pwdhistorylength"][0] } else { $null }

        $maxTS  = $null
        if ($r.Properties["maxpwdage"].Count -gt 0) {
            $maxTS = Convert-LargeIntegerToTimeSpan $r.Properties["maxpwdage"][0]
        }

        $lockTh = if ($r.Properties["lockoutthreshold"].Count -gt 0) { [int]$r.Properties["lockoutthreshold"][0] } else { $null }

        $lockDurFG = $null
        if ($r.Properties["lockoutduration"].Count -gt 0) {
            $lockDurFG = Convert-LargeIntegerToTimeSpan $r.Properties["lockoutduration"][0]
        }

        $lockObsFG = $null
        if ($r.Properties["lockoutobservationwindow"].Count -gt 0) {
            $lockObsFG = Convert-LargeIntegerToTimeSpan $r.Properties["lockoutobservationwindow"][0]
        }

        $appliesTo = @()
        if ($r.Properties["msds-psoappliesto"].Count -gt 0) {
            $appliesTo = @($r.Properties["msds-psoappliesto"])
        }

        $finePolicies += [PSCustomObject]@{
            Name                        = $name
            Precedence                  = $prec
            MinPasswordLength           = $minLen
            PasswordHistoryCount        = $hist
            MaxPasswordAge_Days         = if ($maxTS) { [int]$maxTS.TotalDays } else { $null }
            LockoutThreshold            = $lockTh
            LockoutDuration_Minutes     = if ($lockDurFG) { [int]$lockDurFG.TotalMinutes } else { $null }
            LockoutObservation_Minutes  = if ($lockObsFG) { [int]$lockObsFG.TotalMinutes } else { $null }
            AppliesTo_DN                = $appliesTo -join ";"
        }
    }

    $resultsFG.Dispose()

    if (-not $finePolicies -or $finePolicies.Count -eq 0) {
        Write-Host "Nenhuma fine-grained password policy encontrada." -ForegroundColor DarkYellow
    }
    else {
        # Resumo na tela
        $finePolicies |
            Select-Object `
                Name,
                Precedence,
                MinPasswordLength,
                PasswordHistoryCount,
                MaxPasswordAge_Days,
                LockoutThreshold |
            Sort-Object Precedence |
            Format-Table -AutoSize

        # CSV detalhado
        $finePolicies |
            Export-Csv -NoTypeInformation -Encoding UTF8 -Path $fineFile

        Write-Host "Arquivo de fine-grained policies: $fineFile" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "Coleta de Métodos de Autenticação concluída. Verifique a pasta: $OutFolder" -ForegroundColor Cyan
