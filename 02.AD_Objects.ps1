Import-Module ActiveDirectory

Write-Host "Coletando totais de objetos no AD..." -ForegroundColor Cyan

# Usuários
$usersTotal     = (Get-ADUser -Filter * -SearchBase (Get-ADDomain).DistinguishedName | Measure-Object).Count
$usersEnabled   = (Get-ADUser -Filter "enabled -eq 'true'"  -SearchBase (Get-ADDomain).DistinguishedName | Measure-Object).Count
$usersDisabled  = (Get-ADUser -Filter "enabled -eq 'false'" -SearchBase (Get-ADDomain).DistinguishedName | Measure-Object).Count

# Grupos
$groupsTotal    = (Get-ADGroup -Filter * -SearchBase (Get-ADDomain).DistinguishedName | Measure-Object).Count

# Computadores
$computersTotal    = (Get-ADComputer -Filter * -SearchBase (Get-ADDomain).DistinguishedName | Measure-Object).Count
$computersEnabled  = (Get-ADComputer -Filter "enabled -eq 'true'"  -SearchBase (Get-ADDomain).DistinguishedName | Measure-Object).Count
$computersDisabled = (Get-ADComputer -Filter "enabled -eq 'false'" -SearchBase (Get-ADDomain).DistinguishedName | Measure-Object).Count

Write-Host ""
Write-Host "=== RESUMO DE OBJETOS AD ===" -ForegroundColor Yellow

Write-Host ""
Write-Host "USUÁRIOS:"
Write-Host "  Total      : $usersTotal"
Write-Host "  Ativos     : $usersEnabled"
Write-Host "  Inativos   : $usersDisabled"

Write-Host ""
Write-Host "GRUPOS:"
Write-Host "  Total      : $groupsTotal"

Write-Host ""
Write-Host "COMPUTADORES:"
Write-Host "  Total      : $computersTotal"
Write-Host "  Ativos     : $computersEnabled"
Write-Host "  Inativos   : $computersDisabled"

Write-Host ""
Write-Host "Contagem concluída." -ForegroundColor Cyan
