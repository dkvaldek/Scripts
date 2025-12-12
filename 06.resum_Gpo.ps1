        Import-Module GroupPolicy

        Write-Host "Coletando resumo de GPOs..." -ForegroundColor Cyan

        # Pega todas as GPOs
        $gpos = Get-GPO -All

        if (-not $gpos -or $gpos.Count -eq 0) {
            Write-Host "Nenhuma GPO encontrada." -ForegroundColor Yellow
            return
        }

        $totalGpos = $gpos.Count

        # GpoStatus pode ser:
        # AllSettingsEnabled, UserSettingsDisabled, ComputerSettingsDisabled, AllSettingsDisabled

        # Usuário habilitado: AllSettingsEnabled OU ComputerSettingsDisabled
        $userEnabled  = $gpos | Where-Object {
            $_.GpoStatus -eq 'AllSettingsEnabled' -or
            $_.GpoStatus -eq 'ComputerSettingsDisabled'
        } | Measure-Object | Select-Object -ExpandProperty Count

        # Usuário desabilitado: UserSettingsDisabled OU AllSettingsDisabled
        $userDisabled = $gpos | Where-Object {
            $_.GpoStatus -eq 'UserSettingsDisabled' -or
            $_.GpoStatus -eq 'AllSettingsDisabled'
        } | Measure-Object | Select-Object -ExpandProperty Count

        # Computador habilitado: AllSettingsEnabled OU UserSettingsDisabled
        $compEnabled  = $gpos | Where-Object {
            $_.GpoStatus -eq 'AllSettingsEnabled' -or
            $_.GpoStatus -eq 'UserSettingsDisabled'
        } | Measure-Object | Select-Object -ExpandProperty Count

        # Computador desabilitado: ComputerSettingsDisabled OU AllSettingsDisabled
        $compDisabled = $gpos | Where-Object {
            $_.GpoStatus -eq 'ComputerSettingsDisabled' -or
            $_.GpoStatus -eq 'AllSettingsDisabled'
        } | Measure-Object | Select-Object -ExpandProperty Count

        Write-Host ""
        Write-Host "=== RESUMO DE GPOs ===" -ForegroundColor Yellow

        Write-Host ""
        Write-Host "GPOs:"
        Write-Host "  Total                     : $totalGpos"

        Write-Host ""
        Write-Host "Configurações de USUÁRIO:"
        Write-Host "  GPOs com User config ATIVA   : $userEnabled"
        Write-Host "  GPOs com User config INATIVA : $userDisabled"

        Write-Host ""
        Write-Host "Configurações de COMPUTADOR:"
        Write-Host "  GPOs com Computer config ATIVA   : $compEnabled"
        Write-Host "  GPOs com Computer config INATIVA : $compDisabled"

        Write-Host ""
        Write-Host "Resumo de GPOs concluído." -ForegroundColor Cyan
