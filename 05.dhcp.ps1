        Import-Module DhcpServer

        # Pasta de export
        $exportPath = "C:\AD-Export"
        New-Item -ItemType Directory -Path $exportPath -ErrorAction SilentlyContinue | Out-Null

        Write-Host "=== DHCP SERVERS AUTORIZADOS NO AD ===`n"

        # Servidores DHCP autorizados na floresta
        $dhcpServers = Get-DhcpServerInDC

        # RESUMO NA TELA - SERVIDORES
        $dhcpServers |
            Select-Object DnsName, IPAddress, Domain |
            Format-Table -AutoSize

        # DETALHADO EM CSV - SERVIDORES
        $dhcpServers |
            Select-Object * |
            Export-Csv "$exportPath\dhcp_servers_all_raw.csv" -NoTypeInformation -Encoding UTF8


        Write-Host "`n=== ESCOPOS IPv4 POR SERVIDOR (RESUMO NA TELA) ===`n"

        $allScopes = @()

        foreach ($srv in $dhcpServers) {

            Write-Host "`nServidor: $($srv.DnsName) [$($srv.IPAddress)]"

            try {
                # Escopos IPv4 do servidor
                $scopes = Get-DhcpServerv4Scope -ComputerName $srv.DnsName -ErrorAction Stop

                # RESUMO NA TELA - ESCOPOS
                $scopes |
                    Select-Object ScopeId, Name, State, SubnetMask, StartRange, EndRange, LeaseDuration |
                    Format-Table -AutoSize

                # Acrescenta info do servidor e guarda para CSV
                $scopes | ForEach-Object {
                    $obj = $_ | Select-Object *
                    Add-Member -InputObject $obj -NotePropertyName DhcpServer -NotePropertyValue $srv.DnsName
                    $allScopes += $obj
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
