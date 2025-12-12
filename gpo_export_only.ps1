<#
.SYNOPSIS
  Script menor para backup/export apenas de GPOs e geração de reports HTML.
.DESCRIPTION
  Use quando estiver com tempo limitado para rodar o conjunto completo.
.PARAMETER BasePath
  Caminho base para salvar os backups e reports.
#>
param(
    [string]$BasePath = "C:\Evidencias_Assessment_HospitalCare"
)

function Ensure-Path {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

Ensure-Path (Join-Path $BasePath "06_GPOs")
Ensure-Path (Join-Path $BasePath "06_GPOs\backup_gpos")
Ensure-Path (Join-Path $BasePath "06_GPOs\GPO_Reports")

Import-Module GroupPolicy -ErrorAction Stop

$gpos = Get-GPO -All
foreach ($g in $gpos) {
    $safeName = ($g.DisplayName -replace '[\\/:*?"<>|]','_')
    Try {
        Backup-GPO -Guid $g.Id -Path (Join-Path $BasePath "06_GPOs\backup_gpos")
    } Catch {
        Write-Warning "Backup failed for $($g.DisplayName): $_"
    }
    Try {
        Get-GPOReport -Guid $g.Id -ReportType Html -Path (Join-Path $BasePath "06_GPOs\GPO_Reports\$safeName.html")
    } Catch {
        Write-Warning "Report failed for $($g.DisplayName): $_"
    }
}
Write-Output "GPO export complete."
