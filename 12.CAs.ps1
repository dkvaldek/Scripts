# ============================
# 12 - Anexos Extras
#   - Autoridades Certificadoras Enterprise
#   - Info, templates e (se possível) certificados emitidos
#   - Resumo na tela + artefatos em arquivos
# ============================

$BasePath  = "C:\AD-Export"
$OutFolder = Join-Path $BasePath "12_Anexos_Extras"

if (-not (Test-Path $OutFolder)) {
    New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null
}

Import-Module ActiveDirectory

Write-Host "Descobrindo CAs Enterprise no AD (pKIEnrollmentService)..." -ForegroundColor Cyan

# ----- 1) Descobrir CAs Enterprise no AD -----

$rootDse   = Get-ADRootDSE
$configNC  = $rootDse.ConfigurationNamingContext
$enrollBase = "CN=Enrollment Services,CN=Public Key Services,CN=Services,$configNC"

$caObjects = Get-ADObject -LDAPFilter "(objectClass=pKIEnrollmentService)" `
               -SearchBase $enrollBase `
               -Properties dNSHostName,cn,certificateTemplates,distinguishedName

if (-not $caObjects -or $caObjects.Count -eq 0) {
    Write-Host "Nenhuma Enterprise CA encontrada no AD (pKIEnrollmentService)." -ForegroundColor Red
    return
}

$caList = @()
$idx    = 0

foreach ($ca in $caObjects) {
    $idx++
    $configString = "{0}\{1}" -f $ca.dNSHostName, $ca.cn
    $safeName     = $configString -replace '[\\\/:\*\?\"<>\|]','_'

    $templateCount = 0
    if ($ca.certificateTemplates) {
        if ($ca.certificateTemplates -is [array]) {
            $templateCount = $ca.certificateTemplates.Count
        } else {
            $templateCount = 1
        }
    }

    $caList += [PSCustomObject]@{
        Index         = $idx
        CAName        = $ca.cn
        DNSHostName   = $ca.dNSHostName
        ConfigString  = $configString
        SafeName      = $safeName
        TemplatesCount = $templateCount
        DistinguishedName = $ca.DistinguishedName
    }
}

# Exporta lista básica
$caCsv = Join-Path $OutFolder "ca_enterprise_list.csv"
$caList | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $caCsv

Write-Host ""
Write-Host "=== ENTERPRISE CAs ENCONTRADAS NO AD (RESUMO) ===" -ForegroundColor Yellow
$caList |
    Select-Object Index,CAName,DNSHostName,TemplatesCount |
    Sort-Object Index |
    Format-Table -AutoSize

Write-Host ""
Write-Host "Lista detalhada salva em: $caCsv" -ForegroundColor Green

# ----- 2) Para cada CA, tentar ping + coletar artefatos se online -----

$summary = @()

foreach ($ca in $caList) {
    Write-Host ""
    Write-Host ("[{0}] CA: {1} ({2})" -f $ca.Index, $ca.CAName, $ca.ConfigString) -ForegroundColor Cyan

    $config   = $ca.ConfigString
    $safeName = $ca.SafeName

    # Teste de ping
    $pingOk = $false
    $pingMsg = $null
    $pingFile = Join-Path $OutFolder ("ca_{0}_ping_raw.txt" -f $safeName)

    try {
        $pingOutput = certutil -config "$config" -ping 2>&1
        $pingOutput | Out-File -FilePath $pingFile -Encoding UTF8

        if ($pingOutput -join "`n" -match "Successfully") {
            $pingOk = $true
            $pingMsg = "Ping OK"
            Write-Host "  - Ping da CA OK." -ForegroundColor Green
        } else {
            $pingMsg = "Ping retornou sem 'Successfully' (ver TXT)."
            Write-Host "  ! Ping da CA não confirmou sucesso. Veja: $pingFile" -ForegroundColor DarkYellow
        }
    } catch {
        $pingMsg = "Ping falhou: $($_.Exception.Message)"
        Write-Host ("  ! Ping da CA falhou: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }

    # Se ping não OK, não tenta ca.info / DB para não travar
    if (-not $pingOk) {
        $summary += [PSCustomObject]@{
            CAName       = $ca.CAName
            DNSHostName  = $ca.DNSHostName
            ConfigString = $ca.ConfigString
            Online       = $false
            PingMessage  = $pingMsg
        }
        continue
    }

    # 2.1 ca.info
    $infoFile = Join-Path $OutFolder ("ca_{0}_info.txt" -f $safeName)
    try {
        & certutil -config "$config" -ca.info 2>&1 |
            Out-File -FilePath $infoFile -Encoding UTF8
        Write-Host ("  - ca.info salvo em: {0}" -f $infoFile) -ForegroundColor Green
    } catch {
        Write-Host ("  ! Falha ao coletar ca.info da CA {0}: {1}" -f $config, $_) -ForegroundColor Red
    }

    # 2.2 Templates publicados (via certutil)
    $tplFile = Join-Path $OutFolder ("ca_{0}_templates.txt" -f $safeName)
    try {
        & certutil -config "$config" -catemplates 2>&1 |
            Out-File -FilePath $tplFile -Encoding UTF8
        Write-Host ("  - Templates publicados salvos em: {0}" -f $tplFile) -ForegroundColor Green
    } catch {
        Write-Host ("  ! Falha ao coletar templates via certutil da CA {0}: {1}" -f $config, $_) -ForegroundColor Red
    }

    # 2.3 Certificados emitidos (pode ser pesado)
    $issuedFile = Join-Path $OutFolder ("ca_{0}_issued_certs.txt" -f $safeName)

    Write-Host "  - Coletando lista de certificados emitidos (Disposition=20)..." -ForegroundColor Yellow
    Write-Host "    (isso pode demorar em CAs com base grande)" -ForegroundColor DarkYellow

    try {
        & certutil -config "$config" `
            -view `
            -restrict "Disposition=20" `
            -out "RequestID,SerialNumber,NotBefore,NotAfter,Request.CommonName,CertificateTemplate" 2>&1 |
            Out-File -FilePath $issuedFile -Encoding UTF8

        Write-Host ("    Certificados emitidos salvos em: {0}" -f $issuedFile) -ForegroundColor Green
    } catch {
        Write-Host ("  ! Falha ao consultar DB de certificados da CA {0}: {1}" -f $config, $_) -ForegroundColor Red
    }

    $summary += [PSCustomObject]@{
        CAName       = $ca.CAName
        DNSHostName  = $ca.DNSHostName
        ConfigString = $ca.ConfigString
        Online       = $true
        PingMessage  = $pingMsg
    }
}

# ----- 3) Resumo simples na tela -----

Write-Host ""
Write-Host "=== RESUMO DE CAs (ONLINE / OFFLINE) ===" -ForegroundColor Cyan

$summary |
    Sort-Object CAName |
    Select-Object CAName,DNSHostName,Online,PingMessage |
    Format-Table -AutoSize

$summaryFile = Join-Path $OutFolder "ca_enterprise_status.csv"
$summary | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $summaryFile

Write-Host ""
Write-Host "Artefatos detalhados gerados em: $OutFolder" -ForegroundColor Cyan
Write-Host "  - ca_enterprise_list.csv" -ForegroundColor DarkCyan
Write-Host "  - ca_enterprise_status.csv" -ForegroundColor DarkCyan
Write-Host "  - ca_<CA>_ping_raw.txt" -ForegroundColor DarkCyan
Write-Host "  - ca_<CA>_info.txt (quando ping OK)" -ForegroundColor DarkCyan
Write-Host "  - ca_<CA>_templates.txt (quando ping OK)" -ForegroundColor DarkCyan
Write-Host "  - ca_<CA>_issued_certs.txt (quando ping OK)" -ForegroundColor DarkCyan
