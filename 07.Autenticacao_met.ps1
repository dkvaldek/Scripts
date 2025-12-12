# ============================
# 07 - Métodos de Autenticação
#   - SPN list
#   - Auth policy evidence
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "07_Autenticacao_Metodos"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Import-Module ActiveDirectory

Write-Host "Coletando informações de SPN e políticas de autenticação..." -ForegroundColor Cyan

# -----------------------------------------------------------
# 1) SPN LIST
# -----------------------------------------------------------
$spnFile = Join-Path $OutFolder "spn_all_raw.csv"

Write-Host ""
Write-Host "=== SPN LIST (RESUMO) ===" -ForegroundColor Yellow

$spnObjects = Get-ADObject -LDAPFilter "(servicePrincipalName=*)" -Properties servicePrincipalName,sAMAccountName,objectClass,distinguishedName

# Resumo na tela: quem tem mais SPNs
$spnObjects |
    Select-Object `
        @{Name="SamAccountName"; Expression={$_.sAMAccountName}},
        ObjectClass,
        @{Name="SPNCount"; Expression={($_.servicePrincipalName).Count}} |
    Sort-Object SPNCount -Descending |
    Select-Object -First 20 |
    Format-Table -AutoSize

# Detalhado em CSV
$spnObjects |
    Select-Object `
        sAMAccountName,
        ObjectClass,
        DistinguishedName,
        @{Name="ServicePrincipalNames"; Expression={ $_.servicePrincipalName -join ";" }} |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $spnFile

Write-Host "Arquivo SPN detalhado: $spnFile" -ForegroundColor Green


# -----------------------------------------------------------
# 2) AUTH POLICY EVIDENCE
#    - Default domain password policy
#    - Fine-grained password policies (se houver)
# -----------------------------------------------------------

Write-Host ""
Write-Host "=== DEFAULT DOMAIN PASSWORD / KERBEROS POLICY (RESUMO) ===" -ForegroundColor Yellow

# Default domain password policy
$defaultPolicy = Get-ADDefaultDomainPasswordPolicy

# Resumo na tela
$defaultPolicy |
    Select-Object `
        MinPasswordLength,
        PasswordHistoryCount,
        MaxPasswordAge,
        MinPasswordAge,
        ComplexityEnabled,
        ReversibleEncryptionEnabled,
        LockoutThreshold,
        LockoutObservationWindow,
        LockoutDuration |
    Format-List

# Detalhado em TXT
$defaultPolicy |
    Format-List * |
    Out-File -FilePath (Join-Path $OutFolder "auth_default_domain_password_policy.txt") -Encoding UTF8

Write-Host "Arquivo de policy padrão: $(Join-Path $OutFolder 'auth_default_domain_password_policy.txt')" -ForegroundColor Green


Write-Host ""
Write-Host "=== FINE-GRAINED PASSWORD POLICIES (RESUMO) ===" -ForegroundColor Yellow

$fineFile = Join-Path $OutFolder "auth_fine_grained_policies.csv"

$finePolicies = Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction SilentlyContinue

if (-not $finePolicies) {
    Write-Host "Nenhuma fine-grained password policy encontrada." -ForegroundColor DarkYellow
} else {
    # Resumo na tela
    $finePolicies |
        Select-Object `
            Name,
            Precedence,
            MinPasswordLength,
            PasswordHistoryCount,
            MaxPasswordAge,
            LockoutThreshold |
        Format-Table -AutoSize

    # Detalhado em CSV
    $finePolicies |
        Select-Object * |
        Export-Csv -NoTypeInformation -Encoding UTF8 -Path $fineFile

    Write-Host "Arquivo de fine-grained policies: $fineFile" -ForegroundColor Green
}

Write-Host ""
Write-Host "Coleta de Métodos de Autenticação concluída. Verifique a pasta: $OutFolder" -ForegroundColor Cyan
