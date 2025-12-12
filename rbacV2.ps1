# ============================
# 10 - Segurança / Permissões RBAC
#   - Grupos privilegiados (RBAC) + membros
#   - ACLs de Domain Root, AdminSDHolder e OUs
#   - Resumo na tela + CSV detalhado
#   - Versão SEM Get-AD* e SEM drive AD:
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "10_Seguranca_Permissoes_RBAC"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Write-Host "Coletando Permissões RBAC (grupos privilegiados + ACLs AD) via LDAP/.NET..." -ForegroundColor Cyan

# --------------------------------------------------------------------
# RootDSE / domínio
# --------------------------------------------------------------------
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

# Função pra converter DN -> nome DNS (ex: DC=hospitalveracruz,DC=com,DC=br -> hospitalveracruz.com.br)
function Convert-DNToDnsName {
    param([string]$DN)

    if (-not $DN) { return $null }

    return (
        $DN -split "," |
        Where-Object { $_ -like "DC=*" } |
        ForEach-Object { $_.Substring(3) }
    ) -join "."
}

$domainDN  = $defaultNC
$domainDNS = Convert-DNToDnsName $domainDN

# --------------------------------------------------------------------
# 1) GRUPOS PRIVILEGIADOS / RBAC (via LDAP)
# --------------------------------------------------------------------

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
$rbacMembers = @()

if (Test-Path $rbacFile) { Remove-Item $rbacFile -Force }

Write-Host ""
Write-Host "=== GRUPOS RBAC / PRIVILEGIADOS (RESUMO) ===" -ForegroundColor Yellow

# Função helper para localizar um grupo por nome (sAMAccountName ou CN)
function Find-LdapGroupByName {
    param(
        [string]$GroupName,
        [string]$SearchBase
    )

    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.SearchRoot  = [ADSI]("LDAP://$SearchBase")
    $searcher.SearchScope = "Subtree"
    # tenta por sAMAccountName OU cn
    $escaped = $GroupName.Replace("(", "\28").Replace(")", "\29") # escape simples
    $searcher.Filter = "(&(objectClass=group)(|(sAMAccountName=$escaped)(cn=$escaped)))"
    $searcher.PageSize = 1000
    $searcher.PropertiesToLoad.Clear()
    "distinguishedName","name","sAMAccountName","objectClass" | ForEach-Object {
        [void]$searcher.PropertiesToLoad.Add($_)
    }

    $result = $searcher.FindOne()
    if (-not $result) { return $null }

    $dn   = $result.Properties["distinguishedname"][0]
    $name = $result.Properties["name"][0]
    $sam  = $null
    if ($result.Properties["samaccountname"].Count -gt 0) {
        $sam = $result.Properties["samaccountname"][0]
    }

    return [PSCustomObject]@{
        DistinguishedName = $dn
        Name              = $name
        SamAccountName    = $sam
    }
}

# Recursão de membros (equivalente ao -Recursive do Get-ADGroupMember)
function Get-LdapGroupMembersRecursive {
    param(
        [string]$GroupDN,
        [System.Collections.Generic.HashSet[string]]$Visited
    )

    $members = @()

    if ($Visited.Contains($GroupDN)) {
        return $members
    }

    $Visited.Add($GroupDN) | Out-Null

    try {
        $groupDE = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$GroupDN")
    } catch {
        return $members
    }

    $memberProp = $groupDE.Properties["member"]
    if (-not $memberProp -or $memberProp.Count -eq 0) {
        return $members
    }

    foreach ($mDN in $memberProp) {
        try {
            $mDE = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$mDN")
        }
        catch {
            continue
        }

        $nameProp = $mDE.Properties["name"]
        $samProp  = $mDE.Properties["sAMAccountName"]
        $ocProp   = $mDE.Properties["objectClass"]

        $mName = if ($nameProp -and $nameProp.Count -gt 0) { $nameProp.Value } else { $null }
        $mSam  = if ($samProp  -and $samProp.Count  -gt 0) { $samProp.Value }  else { $null }
        $mType = $null
        if ($ocProp -and $ocProp.Count -gt 0) {
            $mType = $ocProp[$ocProp.Count - 1]  # último objectClass (mais específico)
        }

        $members += [PSCustomObject]@{
            MemberName              = $mName
            MemberSamAccount        = $mSam
            MemberObjectClass       = $mType
            MemberDistinguishedName = $mDN
        }

        # Se o membro também é grupo, recursão (como -Recursive)
        if ($mType -eq "group") {
            $members += Get-LdapGroupMembersRecursive -GroupDN $mDN -Visited $Visited
        }
    }

    return $members
}

foreach ($g in $rbacGroups) {
    $group = Find-LdapGroupByName -GroupName $g -SearchBase $defaultNC

    if (-not $group) {
        Write-Host "Grupo não encontrado (no domínio): $g (ignorando)" -ForegroundColor DarkYellow
        continue
    }

    Write-Host "Grupo encontrado: $($group.Name) [$($group.DistinguishedName)]"

    # HashSet de DNs já visitados (para evitar loop)
    $visited = New-Object 'System.Collections.Generic.HashSet[string]'

    $members = Get-LdapGroupMembersRecursive -GroupDN $group.DistinguishedName -Visited $visited
    $count   = $members.Count

    foreach ($m in $members) {
        $rbacMembers += [PSCustomObject]@{
            GroupName               = $group.Name
            GroupDistinguishedName  = $group.DistinguishedName
            MemberName              = $m.MemberName
            MemberSamAccount        = $m.MemberSamAccount
            MemberObjectClass       = $m.MemberObjectClass
            MemberDistinguishedName = $m.MemberDistinguishedName
        }
    }

    $rbacSummary += [PSCustomObject]@{
        GroupName = $group.Name
        Members   = $count
    }
}

if ($rbacSummary.Count -gt 0) {
    $rbacMembers |
        Export-Csv -NoTypeInformation -Encoding UTF8 -Path $rbacFile

    $rbacSummary |
        Sort-Object Members -Descending |
        Format-Table GroupName,Members -AutoSize

    Write-Host ""
    Write-Host "Detalhe de membros salvo em: $rbacFile" -ForegroundColor Green
} else {
    Write-Host "Nenhum grupo privilegiado encontrado / sem membros." -ForegroundColor DarkYellow
}

# --------------------------------------------------------------------
# 2) ACLs AD - Domain Root, AdminSDHolder e OUs
# --------------------------------------------------------------------

function Get-AdAclEntriesLdap {
    param(
        [string]$DistinguishedName,
        [string]$TargetLabel
    )

    $entries = @()

    try {
        $de  = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DistinguishedName")
        $acl = $de.ObjectSecurity
    }
    catch {
        Write-Host ("  ! Falha ao obter ACL de {0}: {1}" -f $TargetLabel, $_.Exception.Message) -ForegroundColor Red
        return $entries
    }

    foreach ($ace in $acl.Access) {
        $entries += [PSCustomObject]@{
            TargetLabel             = $TargetLabel
            TargetDistinguishedName = $DistinguishedName
            IdentityReference       = $ace.IdentityReference.ToString()
            ActiveDirectoryRights   = $ace.ActiveDirectoryRights.ToString()
            AccessControlType       = $ace.AccessControlType.ToString()
            IsInherited             = $ace.IsInherited
            InheritanceType         = $ace.InheritanceType.ToString()
            ObjectType              = $ace.ObjectType
            InheritedObjectType     = $ace.InheritedObjectType
            PropagationFlags        = $ace.PropagationFlags.ToString()
        }
    }

    return $entries
}

Write-Host ""
Write-Host "=== ACLs do Domain Root e AdminSDHolder (RESUMO) ===" -ForegroundColor Yellow

# Domain Root
$domainLabel       = "Domain Root: $domainDNS"
$domainAclEntries  = Get-AdAclEntriesLdap -DistinguishedName $domainDN -TargetLabel $domainLabel

# AdminSDHolder
$adminSDHolderDN   = "CN=AdminSDHolder,CN=System,$domainDN"
$adminAclEntries   = Get-AdAclEntriesLdap -DistinguishedName $adminSDHolderDN -TargetLabel "AdminSDHolder"

$aclDomainAdminFile = Join-Path $OutFolder "rbac_acl_domain_adminsdholder.csv"

$allCriticalAcl = $domainAclEntries + $adminAclEntries
if ($allCriticalAcl.Count -gt 0) {
    $allCriticalAcl |
        Export-Csv -NoTypeInformation -Encoding UTF8 -Path $aclDomainAdminFile

    if ($domainAclEntries.Count -gt 0) {
        Write-Host "Top identidades (Domain Root):" -ForegroundColor Cyan
        $domainAclEntries |
            Group-Object IdentityReference |
            Sort-Object Count -Descending |
            Select-Object -First 10 |
            Format-Table Name,Count -AutoSize
    }

    if ($adminAclEntries.Count -gt 0) {
        Write-Host ""
        Write-Host "Top identidades (AdminSDHolder):" -ForegroundColor Cyan
        $adminAclEntries |
            Group-Object IdentityReference |
            Sort-Object Count -Descending |
            Select-Object -First 10 |
            Format-Table Name,Count -AutoSize
    }

    Write-Host ""
    Write-Host "ACLs detalhadas de Domain Root + AdminSDHolder salvas em: $aclDomainAdminFile" -ForegroundColor Green
} else {
    Write-Host "Nenhum ACE encontrado em Domain Root/AdminSDHolder (ou falha ao ler ACL)." -ForegroundColor DarkYellow
}

# --------------------------------------------------------------------
# OUs
# --------------------------------------------------------------------

Write-Host ""
Write-Host "Coletando ACLs de todas as OUs..." -ForegroundColor Cyan

$ouAclFile = Join-Path $OutFolder "rbac_acl_ous.csv"
if (Test-Path $ouAclFile) { Remove-Item $ouAclFile -Force }

# Enumerar OUs via LDAP
$ouSearcher = New-Object System.DirectoryServices.DirectorySearcher
$ouSearcher.SearchRoot  = [ADSI]("LDAP://$domainDN")
$ouSearcher.SearchScope = "Subtree"
$ouSearcher.Filter      = "(objectClass=organizationalUnit)"
$ouSearcher.PageSize    = 1000
$ouSearcher.PropertiesToLoad.Clear()
"distinguishedName","name" | ForEach-Object {
    [void]$ouSearcher.PropertiesToLoad.Add($_)
}

$ouResults = $ouSearcher.FindAll()
$ouList = @()

foreach ($r in $ouResults) {
    $ouDN   = $r.Properties["distinguishedname"][0]
    $ouName = $r.Properties["name"][0]

    $ouList += [PSCustomObject]@{
        Name              = $ouName
        DistinguishedName = $ouDN
    }
}
$ouResults.Dispose()

$ouCount  = $ouList.Count
$idx      = 0
$ouAclAll = @()

foreach ($ou in $ouList) {
    $idx++
    $label = "OU: " + $ou.Name
    Write-Host ("  [{0}/{1}] {2}" -f $idx, $ouCount, $label)

    $entries = Get-AdAclEntriesLdap -DistinguishedName $ou.DistinguishedName -TargetLabel $label
    $ouAclAll += $entries
}

if ($ouAclAll.Count -gt 0) {
    $ouAclAll |
        Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ouAclFile

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
