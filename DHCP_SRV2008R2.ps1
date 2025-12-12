# NÃO usa módulo DhcpServer (não existe no 2008 R2)
# Precisa rodar o PowerShell como Administrador
# E o servidor precisa ter o papel DHCP instalado

# Pasta de export
$exportPath = "C:\AD-Export"
New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

Write-Host "=== DHCP SERVERS AUTORIZADOS NO AD (via netsh dhcp show server) ===`n"

# Saída bruta do netsh
$rawServers = netsh dhcp show server 2>&1

$dhcpServers = @()

# Tenta descobrir o domínio atual só pra preencher a coluna Domain
$domain = $null
try {
    $domain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
} catch {
    $domain = $null
}

foreach ($line in $rawServers) {
    # Exemplo de linha:
    # Server [savdaldc02.savilltech.net] Address [192.168.1.11] Ds location: ...
    if ($line -match 'Server\s+\[(?<Dns>[^\]]+)\]\s+Address\s+\[(?<IP>[0-9\.]+)\]') {
        $dhcpServers += [PSCustomObject]@{
            DnsName   = $matches['Dns']
            IPAddress = $matches['IP']
            Domain    = $domain
        }
    }
}

if (-not $dhcpServers -or $dhcpServers.Count -eq 0) {
    Write-Warning "Nenhum servidor DHCP encontrado em 'netsh dhcp show server'. Verifique se o papel DHCP está instalado."
    return
}

# RESUMO NA TELA - SERVIDORES
$dhcpServers |
    Select-Object DnsName, IPAddress, Domain |
    Format-Table -AutoSize

# DETALHADO EM CSV - SERVIDORES
$dhcpServers |
    Export-Csv "$exportPath\dhcp_servers_all_raw.csv" -NoTypeInformation -Encoding UTF8


Write-Host "`n=== ESCOPOS IPv4 POR SERVIDOR (RESUMO NA TELA) ===`n"

$allScopes = @()

foreach ($srv in $dhcpServers) {

    Write-Host "`nServidor: $($srv.DnsName) [$($srv.IPAddress)]"

    try {
        # Saída bruta dos escopos IPv4 desse servidor
        # Exemplo de linhas úteis (StackOverflow):
        #  Scope Address  - Subnet Mask    - State    - Scope Name    - Comment
        #  10.5.116.0     - 255.255.255.0  -Active   -LAN 1          -VLAN 1
        $rawScopes = netsh dhcp server \\$($srv.DnsName) show scope 2>&1

        $scopesThisServer = @()

        foreach ($line in $rawScopes) {
            if ($line -match '^\s*(?<Scope>[0-9\.]+)\s+-\s+(?<Mask>[0-9\.]+)\s+-\s*(?<State>\S+)\s+-\s*(?<Name>.+?)(\s+-\s*(?<Comment>.*))?$') {
                $scopeId    = $matches['Scope']
                $subnetMask = $matches['Mask']
                $state      = $matches['State']
                $name       = $matches['Name'].Trim()
                $comment    = $matches['Comment']

                $obj = [PSCustomObject]@{
                    DhcpServer = $srv.DnsName
                    ScopeId    = $scopeId
                    SubnetMask = $subnetMask
                    State      = $state
                    Name       = $name
                    Comment    = $comment
                }

                $scopesThisServer += $obj
                $allScopes        += $obj
            }
        }

        if ($scopesThisServer.Count -gt 0) {
            # RESUMO NA TELA - ESCOPOS
            $scopesThisServer |
                Select-Object ScopeId, SubnetMask, State, Name, Comment |
                Format-Table -AutoSize
        } else {
            Write-Warning "Nenhum escopo retornado para $($srv.DnsName)."
        }

    } catch {
        Write-Warning "Nao foi possivel consultar escopos no servidor $($srv.DnsName): $_"
    }
}

# DETALHADO EM CSV - ESCOPOS
if ($allScopes.Count -gt 0) {
    $allScopes |
        Export-Csv "$exportPath\dhcp_scopes_v4_all_raw.csv" -NoTypeInformation -Encoding UTF8
    Write-Host "`nEscopos IPv4 exportados para $exportPath\dhcp_scopes_v4_all_raw.csv"
} else {
    Write-Warning "`nNenhum escopo IPv4 encontrado ou consulta falhou."
}

Write-Host "`nServidores DHCP exportados para $exportPath\dhcp_servers_all_raw.csv"
