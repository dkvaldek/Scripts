Import-Module ActiveDirectory

# Pasta de export (a mesma que você já está usando)
$exportPath = "C:\AD-Export"
New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

# Pega o Configuration NC da floresta
$rootDse  = Get-ADRootDSE
$configNC = $rootDse.ConfigurationNamingContext

########## SITES ##########

$sites = Get-ADObject `
    -SearchBase ("CN=Sites," + $configNC) `
    -LDAPFilter "(objectClass=site)" `
    -Properties *

$sites |
    Select-Object `
        Name,
        DistinguishedName,
        @{Name = "Location";    Expression = { $_.location }},
        Description |
    Export-Csv "$exportPath\sites_all_raw.csv" -NoTypeInformation -Encoding UTF8

########## SUBNETS ##########

$subnets = Get-ADObject `
    -SearchBase ("CN=Subnets,CN=Sites," + $configNC) `
    -LDAPFilter "(objectClass=subnet)" `
    -Properties siteObject,location,description

$subnets |
    Select-Object `
        Name,                   # Ex: 10.0.0.0/24
        DistinguishedName,
        @{Name = "SiteDN";   Expression = { $_.siteObject }},
        @{Name = "SiteName"; Expression = {
            (Get-ADObject -Identity $_.siteObject -ErrorAction SilentlyContinue).Name
        }},
        @{Name = "Location"; Expression = { $_.location }},
        Description |
    Export-Csv "$exportPath\subnets_all_raw.csv" -NoTypeInformation -Encoding UTF8
