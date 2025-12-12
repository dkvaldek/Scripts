# ============================
# 10 - Segurança / Permissões RBAC
#   - Grupos privilegiados (RBAC) + membros
#   - ACLs de Domain Root, AdminSDHolder e OUs
#   - Resumo na tela + CSV detalhado
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "10_Seguranca_Permissoes_RBAC"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Import-Module ActiveDirectory

Write-Host "Coletando Permissões RBAC (grupos privilegiados + ACLs AD)..." -ForegroundColor Cyan

# -----------------------------------------------------------
# 1) GRUPOS PRIVILEGIADOS / RBAC
# -----------------------------------------------------------

$rbacGroups = @(
    'Domain Admins',
    'Enterprise Admins',
    'Schema Admins',
    'Administrators',
    'Account Operators',
    'Backup Operators',
    'Server Operators',
    'Print Operators',
    'DnsAdmins',
    'DHCP Administrators',
    'Group Policy Creator Owners'
)

$rbacFile    = Join-Path $OutFolder "rbac_privileged_groups_members.csv"
$rbacSummary = @()

if (Test-Path $rbacFile) { Remove-Item $rbacFile -Force }

Write-Host ""
Write-Host "=== GRUPOS RBAC / PRIVILEGIADOS (RESUMO) ===" -ForegroundColor Yellow

foreach ($g in $rbacGroups) {
    try {
        $group = Get-ADGroup -Identity $g -ErrorAction Stop
    } catch {
        Write-Host "Grupo não encontrado: $g (ignorando)" -ForegroundColor DarkYellow
        continue
    }

    try {
        $members = Get-ADGroupMember -Identity $group -Recursive
        $count   = $members.Count

        foreach ($m in $members) {
            # Tenta enriquecer com SamAccountName e ObjectClass
            $obj = $null
            try {
                $obj = Get-ADObject -Identity $m.DistinguishedName -Properties sAMAccountName,objectClass -ErrorAction SilentlyContinue
            } catch { }

            $sam  = $null
            $type = $null
            if ($obj) {
                $sam  = $obj.sAMAccountName
                $type = $obj.objectClass -join ','
            }

            [PSCustomObject]@{
                GroupName          = $group.Name
                GroupDistinguishedName = $group.DistinguishedName
                MemberName         = $m.Name
                MemberSamAccount   = $sam
                MemberObjectClass  = $type
                MemberDistinguishedName = $m.DistinguishedName
            } | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $rbacFile -Append
        }

        $rbacSummary += [PSCustomObject]@{
            GroupName = $group.Name
            Members   = $count
        }
    } catch {
        Write-Host ("  ! Falha ao enumerar membros do grupo {0}: {1}" -f $g, $_) -ForegroundColor Red
    }
}

if ($rbacSummary.Count -gt 0) {
    $rbacSummary |
        Sort-Object Members -Descending |
        Format-Table GroupName,Members -AutoSize
    Write-Host ""
    Write-Host "Detalhe de membros salvo em: $rbacFile" -ForegroundColor Green
} else {
    Write-Host "Nenhum grupo privilegiado encontrado / sem membros." -ForegroundColor DarkYellow
}

# -----------------------------------------------------------
# 2) ACLs AD - Domain Root, AdminSDHolder e OUs
# -----------------------------------------------------------

function Get-AdAclEntries {
    param(
        [string]$DistinguishedName,
        [string]$TargetLabel
    )

    $path = "AD:$DistinguishedName"
    try {
        $acl = Get-Acl $path
    } catch {
        Write-Host ("  ! Falha ao obter ACL de {0}: {1}" -f $TargetLabel, $_) -ForegroundColor Red
        return @()
    }

    $entries = @()

    foreach ($ace in $acl.Access) {
        $entries += [PSCustomObject]@{
            TargetLabel         = $TargetLabel
            TargetDistinguishedName = $DistinguishedName
            IdentityReference   = $ace.IdentityReference.ToString()
            ActiveDirectoryRights = $ace.ActiveDirectoryRights.ToString()
            AccessControlType   = $ace.AccessControlType.ToString()
            IsInherited         = $ace.IsInherited
            InheritanceType     = $ace.InheritanceType.ToString()
            ObjectType          = $ace.ObjectType
            InheritedObjectType = $ace.InheritedObjectType
            PropagationFlags    = $ace.PropagationFlags.ToString()
        }
    }

    return $entries
}

$domain = Get-ADDomain
$domainDN = $domain.DistinguishedName

Write-Host ""
Write-Host "=== ACLs do Domain Root e AdminSDHolder (RESUMO) ===" -ForegroundColor Yellow

# Domain Root
$domainAclEntries = Get-AdAclEntries -DistinguishedName $domainDN -TargetLabel ("Domain Root: " + $domain.DNSRoot)

# AdminSDHolder
$adminSDHolderDN  = "CN=AdminSDHolder,CN=System,$domainDN"
$adminAclEntries  = Get-AdAclEntries -DistinguishedName $adminSDHolderDN -TargetLabel "AdminSDHolder"

$aclDomainAdminFile = Join-Path $OutFolder "rbac_acl_domain_adminsdholder.csv"

$allCriticalAcl = $domainAclEntries + $adminAclEntries
if ($allCriticalAcl.Count -gt 0) {
    $allCriticalAcl | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $aclDomainAdminFile

    Write-Host "Top identidades (Domain Root):" -ForegroundColor Cyan
    $domainAclEntries |
        Group-Object IdentityReference |
        Sort-Object Count -Descending |
        Select-Object -First 10 |
        Format-Table Name,Count -AutoSize

    Write-Host ""
    Write-Host "Top identidades (AdminSDHolder):" -ForegroundColor Cyan
    $adminAclEntries |
        Group-Object IdentityReference |
        Sort-Object Count -Descending |
        Select-Object -First 10 |
        Format-Table Name,Count -AutoSize

    Write-Host ""
    Write-Host "ACLs detalhadas de Domain Root + AdminSDHolder salvas em: $aclDomainAdminFile" -ForegroundColor Green
} else {
    Write-Host "Nenhum ACE encontrado em Domain Root/AdminSDHolder (algo estranho aí...)" -ForegroundColor DarkYellow
}

# OUs
Write-Host ""
Write-Host "Coletando ACLs de todas as OUs..." -ForegroundColor Cyan

$ouAclFile = Join-Path $OutFolder "rbac_acl_ous.csv"
if (Test-Path $ouAclFile) { Remove-Item $ouAclFile -Force }

$ouList = Get-ADOrganizationalUnit -Filter * -Properties Name,DistinguishedName
$ouCount = $ouList.Count
$idx = 0

$ouAclAll = @()

foreach ($ou in $ouList) {
    $idx++
    $label = "OU: " + $ou.Name
    Write-Host ("  [{0}/{1}] {2}" -f $idx, $ouCount, $label)

    $entries = Get-AdAclEntries -DistinguishedName $ou.DistinguishedName -TargetLabel $label
    $ouAclAll += $entries
}

if ($ouAclAll.Count -gt 0) {
    $ouAclAll | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ouAclFile
    Write-Host ""
    Write-Host ("Total de ACEs em OUs: {0}" -f $ouAclAll.Count) -ForegroundColor Cyan

    Write-Host "Top identidades com permissões em OUs (top 15):" -ForegroundColor Cyan
    $ouAclAll |
        Group-Object IdentityReference |
        Sort-Object Count -Descending |
        Select-Object -First 15 |
        Format-Table Name,Count -AutoSize

    Write-Host ""
    Write-Host "ACLs de OUs salvas em: $ouAclFile" -ForegroundColor Green
} else {
    Write-Host "Nenhuma ACL de OU coletada." -ForegroundColor DarkYellow
}

Write-Host ""
Write-Host "Coleta de Permissões RBAC concluída. Verifique a pasta: $OutFolder" -ForegroundColor Cyan
