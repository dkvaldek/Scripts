param(
    [string]$RootPath
)

if (-not $RootPath) {
    Write-Host "Uso:" -ForegroundColor Yellow
    Write-Host "  .\Gerar-Readme-Dominios.ps1 -RootPath ""CAMINHO_DOS_DOMINIOS""" -ForegroundColor Yellow
    return
}

if (-not (Test-Path $RootPath)) {
    Write-Host "Caminho raiz NÃO existe:" -ForegroundColor Red
    Write-Host "  $RootPath"
    return
}

Write-Host "Gerando Readme.txt em cada pasta 01_..12_ de cada domínio em: $RootPath" -ForegroundColor Yellow

# Pastas de domínio diretamente dentro de $RootPath (ex: csh.corp, hes.local)
$domainFolders = Get-ChildItem -Path $RootPath -Directory -ErrorAction SilentlyContinue

if (-not $domainFolders -or $domainFolders.Count -eq 0) {
    Write-Host "Nenhuma pasta de domínio encontrada em: $RootPath" -ForegroundColor Red
    return
}

foreach ($domainFolder in $domainFolders) {

    $domainName = $domainFolder.Name
    $domainPath = $domainFolder.FullName

    # Data de criação da pasta do domínio
    $creationDate = $domainFolder.CreationTime.ToString("dd/MM/yyyy")

    Write-Host ""
    Write-Host "Domínio: $domainName (Criado em: $creationDate)" -ForegroundColor Cyan

    # Subpastas numeradas dentro do domínio (01_xxx, 02_xxx, ..., 12_xxx)
    $subfolders = Get-ChildItem -Path $domainPath -Directory -ErrorAction SilentlyContinue

    foreach ($sub in $subfolders) {
        $name = $sub.Name

        # Formato NN_algo (ex: 08_EntraSync_Azure)
        if ($name -match '^(\d{2})_(.+)$') {
            $num   = $matches[1]
            $resto = $matches[2]

            # Tira underscores para virar texto legível
            $descricao = $resto -replace '_',' '

            # Pequenas correções de português/formatação
            $descricao = $descricao -replace 'Inventario','Inventário'
            $descricao = $descricao -replace 'DomainControllers','Domain Controllers'

            # Texto final do campo Conteúdo (sem número e sem underline)
            $conteudo = "informações sobre $descricao"

            $readmePath = Join-Path $sub.FullName "Readme.txt"

            $linhasArquivo = @()
            $linhasArquivo += "Nome do consultor : Klaus"
            $linhasArquivo += "Data: $creationDate"
            $linhasArquivo += "Conteúdo: $conteudo"
            $linhasArquivo += "Observações:"

            # Usa Encoding Default para evitar problemas de acentuação na visualização
            $linhasArquivo | Set-Content -Path $readmePath -Encoding Default

            Write-Host "  Readme.txt criado em: $readmePath" -ForegroundColor Green
        }
    }
}

Write-Host ""
Write-Host "Processo concluído." -ForegroundColor Green
