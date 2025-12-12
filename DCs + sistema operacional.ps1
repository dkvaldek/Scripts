                    Import-Module ActiveDirectory

                    $basePath  = "C:\AD-Export"
                    $outFolder = Join-Path $basePath "PowerBI"

                    if (-not (Test-Path $outFolder)) {
                        New-Item -Path $outFolder -ItemType Directory -Force | Out-Null
                    }

                    $forest = Get-ADForest
                    $domain = Get-ADDomain

                    Write-Host "Coletando DCs da floresta: $($forest.Name) / domínio: $($domain.DNSRoot)..." -ForegroundColor Cyan

                    $dcs = Get-ADDomainController -Filter *

                    $rows = foreach ($dc in $dcs) {
                        [PSCustomObject]@{
                            ForestName   = $forest.Name
                            RootDomain   = $forest.RootDomain
                            DomainFqdn   = $dc.Domain
                            DomainNetBIOS = $domain.NetBIOSName
                            DCName       = $dc.Name
                            DCFQDN       = $dc.HostName
                            Site         = $dc.Site
                            IPv4         = $dc.IPv4Address
                            OS           = $dc.OperatingSystem
                            OSVersion    = $dc.OperatingSystemVersion
                            IsGlobalCatalog = $dc.IsGlobalCatalog
                            FSMORoles    = ($dc.OperationMasterRoles -join ';')
                        }
                    }

                    $dcFile = Join-Path $outFolder ("DCs_{0}.csv" -f $domain.DNSRoot.Replace('.','_'))
                    $rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $dcFile

                    Write-Host "Arquivo gerado: $dcFile" -ForegroundColor Green
                        