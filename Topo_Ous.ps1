Import-Module ActiveDirectory

$basePath  = "C:\AD-Export"
$outFolder = Join-Path $basePath "Visio"

if (-not (Test-Path $outFolder)) {
    New-Item -Path $outFolder -ItemType Directory -Force | Out-Null
}

$domain   = Get-ADDomain
$domainDN = $domain.DistinguishedName
$domName  = $domain.DNSRoot

Write-Host "Gerando TXT de OUs para Visio (Organization Chart) - domínio: $domName..." -ForegroundColor Cyan

# Todas as OUs
$ous = Get-ADOrganizationalUnit -Filter * `
    -SearchBase $domainDN `
    -SearchScope Subtree `
    -Properties Name,DistinguishedName

function Get-ParentDn {
    param([string]$dn)
    $parts = $dn -split ",", 2
    if ($parts.Count -eq 2) { return $parts[1] }
    return $null
}

# Vamos guardar um map DN -> Name pra achar o nome do pai rápido
$ouByDn = @{}
foreach ($ou in $ous) {
    $ouByDn[$ou.DistinguishedName] = $ou.Name
}

# Monta as linhas do arquivo de texto (tab-delimited)
$lines = @()

# Cabeçalho no formato que o Visio entende
# Name = “nome da caixinha”
# ReportsTo = “quem é o chefe / pai”
# DN = informação extra (só pra referência)
$lines += "Name`tReportsTo`tDN"

# Nó raiz = domínio (sem ReportsTo)
$lines += ("{0}`t`t{1}" -f $domName, $domainDN)

foreach ($ou in $ous) {
    $name = $ou.Name
    $dn   = $ou.DistinguishedName
    $parentDn = Get-ParentDn -dn $dn

    # Se o pai é o próprio domínio, ReportsTo = nome do domínio
    if ($parentDn -eq $domainDN) {
        $reportsTo = $domName
    } else {
        # Se o pai é outra OU, pega o nome dessa OU
        if ($ouByDn.ContainsKey($parentDn)) {
            $reportsTo = $ouByDn[$parentDn]
        } else {
            # fallback - se não achar, deixa vazio
            $reportsTo = ""
        }
    }

    $line = ("{0}`t{1}`t{2}" -f $name, $reportsTo, $dn)
    $lines += $line
}

$outFile = Join-Path $outFolder ("ou_visio_orgchart_{0}.txt" -f $domName.Replace(".","_"))

# Encoding "Default" (ANSI) evita frescura de BOM e é bem aceito pelo Visio
$lines | Set-Content -Path $outFile -Encoding Default

Write-Host "Arquivo TXT para Visio gerado em: $outFile" -ForegroundColor Green
