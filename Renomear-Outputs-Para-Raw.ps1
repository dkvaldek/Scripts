param(
    # Pasta raiz onde a busca começa (mude se quiser)
    [string]$RootPath = "C:\",

    # Use -WhatIf para só simular sem alterar nada
    [switch]$WhatIf
)

if (-not (Test-Path $RootPath)) {
    Write-Host "Caminho não existe: $RootPath" -ForegroundColor Red
    return
}

Write-Host "Procurando pastas chamadas 'outputs' em: $RootPath" -ForegroundColor Yellow

# Encontra apenas diretórios com o nome 'outputs'
$dirs = Get-ChildItem -Path $RootPath -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ieq "outputs" }

if (-not $dirs) {
    Write-Host "Nenhuma pasta 'outputs' encontrada." -ForegroundColor Yellow
    return
}

foreach ($dir in $dirs) {
    $oldPath = $dir.FullName
    $parent  = $dir.Parent.FullName
    $newPath = Join-Path $parent "raw"

    if (Test-Path $newPath) {
        Write-Warning "Já existe uma pasta 'raw' em: $parent  (pulando $oldPath)"
        continue
    }

    if ($WhatIf) {
        Write-Host "SIMULAÇÃO: renomearia '$oldPath' para '$newPath'" -ForegroundColor Cyan
    } else {
        Write-Host "Renomeando '$oldPath' para '$newPath'..." -ForegroundColor Cyan
        Rename-Item -LiteralPath $oldPath -NewName "raw" -ErrorAction Continue
    }
}

Write-Host "Concluído." -ForegroundColor Green
