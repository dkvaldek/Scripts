########## FSMO DE FLORESTA ##########

$forest = Get-ADForest

$fsmoForest = @(
    [PSCustomObject]@{
        Scope      = 'Forest'
        Role       = 'SchemaMaster'
        FSMOHolder = $forest.SchemaMaster
        ServerName = ($forest.SchemaMaster -split '\.')[0]
    }
    [PSCustomObject]@{
        Scope      = 'Forest'
        Role       = 'DomainNamingMaster'
        FSMOHolder = $forest.DomainNamingMaster
        ServerName = ($forest.DomainNamingMaster -split '\.')[0]
    }
)

# Mostra na tela (pra print)
"=== FSMO de FLORESTA ==="
$fsmoForest | Format-Table Scope,Role,ServerName,FSMOHolder -AutoSize

# Salva em CSV
$fsmoForest | Export-Csv "$exportPath\fsmo_forest_roles.csv" -NoTypeInformation -Encoding UTF8


########## FSMO DE DOMÍNIO(S) ##########

$fsmoDomain = @()

foreach ($domainName in $forest.Domains) {
    $d = Get-ADDomain $domainName

    $fsmoDomain += [PSCustomObject]@{
        Domain     = $domainName
        Role       = 'PDCEmulator'
        FSMOHolder = $d.PDCEmulator
        ServerName = ($d.PDCEmulator -split '\.')[0]
    }
    $fsmoDomain += [PSCustomObject]@{
        Domain     = $domainName
        Role       = 'RIDMaster'
        FSMOHolder = $d.RIDMaster
        ServerName = ($d.RIDMaster -split '\.')[0]
    }
    $fsmoDomain += [PSCustomObject]@{
        Domain     = $domainName
        Role       = 'InfrastructureMaster'
        FSMOHolder = $d.InfrastructureMaster
        ServerName = ($d.InfrastructureMaster -split '\.')[0]
    }
}

""
"=== FSMO de DOMÍNIO ==="
$fsmoDomain | Format-Table Domain,Role,ServerName,FSMOHolder -AutoSize

$fsmoDomain | Export-Csv "$exportPath\fsmo_domain_roles.csv" -NoTypeInformation -Encoding UTF8
