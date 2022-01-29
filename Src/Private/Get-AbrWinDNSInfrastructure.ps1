function Get-AbrWinDNSInfrastructure {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Windows Domain Name System Infrastructure information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.3.0
        Author:         Jonathan Colon
        Twitter:        @jcolonfzenpr
        Github:         rebelinux
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows
    #>
    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "DHCP InfoLevel set at $($InfoLevel.DNS)."
        Write-PscriboMessage "Collecting Host DNS Server information."    }

    process {
        try {
            $DNSSetting = Get-DnsServerSetting -CimSession $TempCIMSession
            if ($DNSSetting) {
                $OutObj = @()
                try {
                    $inObj = [ordered] @{
                        'Build Number' = ConvertTo-EmptyToFiller $DNSSetting.BuildNumber
                        'IPv6' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $DNSSetting.EnableIPv6)
                        'DnsSec' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $DNSSetting.EnableDnsSec)
                        'ReadOnly DC' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $DNSSetting.IsReadOnlyDC)
                        'Listening IP' = $DNSSetting.ListeningIPAddress
                        'All IPs' = $DNSSetting.AllIPAddress
                    }
                    $OutObj += [pscustomobject]$inobj
                }
                catch {
                    Write-PscriboMessage -IsWarning " $($_.Exception.Message) (Infrastructure Summary)"
                }
            }

            $TableParams = @{
                Name = "DNS Servers Settings - $($System.toUpper().split(".")[0])"
                List = $true
                ColumnWidths = 40, 60
            }
            if ($Report.ShowTableCaptions) {
                $TableParams['Caption'] = "- $($TableParams.Name)"
            }
            $OutObj | Sort-Object -Property 'DC Name' | Table @TableParams
            #---------------------------------------------------------------------------------------------#
            #                                 DNS IP Section                                              #
            #---------------------------------------------------------------------------------------------#
            if ($InfoLevel.DNS -ge 2) {
                try {
                    $DNSIPSetting = Get-NetAdapter -CimSession $TempCIMSession | Get-DnsClientServerAddress -CimSession $TempCIMSession -AddressFamily IPv4
                    if ($DNSIPSetting) {
                        Section -Style Heading3 "Domain Controller DNS IP Configuration" {
                            $OutObj = @()
                            try {
                                $inObj = [ordered] @{
                                    'Interface' = $DNSIPSetting.InterfaceAlias
                                    'DNS IP 1' = ConvertTo-EmptyToFiller $DNSIPSetting.ServerAddresses[0]
                                    'DNS IP 2' = ConvertTo-EmptyToFiller $DNSIPSetting.ServerAddresses[1]
                                    'DNS IP 3' = ConvertTo-EmptyToFiller $DNSIPSetting.ServerAddresses[2]
                                    'DNS IP 4' = ConvertTo-EmptyToFiller $DNSIPSetting.ServerAddresses[3]
                                }
                                $OutObj = [pscustomobject]$inobj

                                if ($HealthCheck.DNS.DP) {
                                    $OutObj | Where-Object { $_.'DNS IP 1' -eq "127.0.0.1"} | Set-Style -Style Warning -Property 'DNS IP 1'
                                }

                                $TableParams = @{
                                    Name = "IP Configuration - $($System.toUpper().split(".")[0])"
                                    List = $false
                                    ColumnWidths = 20, 20, 20, 20, 20
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $OutObj | Table @TableParams
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning "$($_.Exception.Message) (DNS IP Configuration Table)"
                }
            }
            <#
            #---------------------------------------------------------------------------------------------#
            #                                 DNS Scanvenging Section                                     #
            #---------------------------------------------------------------------------------------------#
            if ($InfoLevel.DNS -ge 2) {
                try {
                    Section -Style Heading6 "Scavenging Options" {
                        $OutObj = @()
                        foreach ($DC in $DCs) {
                            Write-PscriboMessage "Collecting Scavenging Options information from $($DC)."
                            try {
                                $DNSSetting = Invoke-Command -Session $Session {Get-DnsServerScavenging -ComputerName $using:DC}
                                $inObj = [ordered] @{
                                    'DC Name' = $($DC.ToString().ToUpper().Split(".")[0])
                                    'NoRefresh Interval' = ConvertTo-EmptyToFiller $DNSSetting.NoRefreshInterval
                                    'Refresh Interval' = ConvertTo-EmptyToFiller $DNSSetting.RefreshInterval
                                    'Scavenging Interval' = ConvertTo-EmptyToFiller $DNSSetting.ScavengingInterval
                                    'Last Scavenge Time' = Switch ($DNSSetting.LastScavengeTime) {
                                        "" {"-"; break}
                                        $Null {"-"; break}
                                        default {ConvertTo-EmptyToFiller ($DNSSetting.LastScavengeTime.ToString("MM/dd/yyyy"))}
                                    }
                                    'Scavenging State' = Switch ($DNSSetting.ScavengingState) {
                                        "True" {"Enabled"}
                                        "False" {"Disabled"}
                                        default {ConvertTo-EmptyToFiller $DNSSetting.ScavengingState}
                                    }
                                }
                                $OutObj += [pscustomobject]$inobj
                            }
                            catch {
                                Write-PscriboMessage -IsWarning "$($_.Exception.Message) (Scavenging Item)"
                            }
                        }

                        $TableParams = @{
                            Name = "Scavenging - $($Domain.ToString().ToUpper())"
                            List = $false
                            ColumnWidths = 25, 15, 15, 15, 15, 15
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Sort-Object -Property 'DC Name' | Table @TableParams
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning "$($_.Exception.Message) (Scavenging Table)"
                }
            }
            #---------------------------------------------------------------------------------------------#
            #                                 DNS Forwarder Section                                       #
            #---------------------------------------------------------------------------------------------#
            try {
                Section -Style Heading6 "Forwarder Options" {
                    $OutObj = @()
                    foreach ($DC in $DCs) {
                        Write-PscriboMessage "Collecting Forwarder Options information from $($DC)."
                        try {
                            $DNSSetting = Invoke-Command -Session $Session {Get-DnsServerForwarder -ComputerName $using:DC}
                            $Recursion = Invoke-Command -Session $Session {Get-DnsServerRecursion -ComputerName $using:DC | Select-Object -ExpandProperty Enable}
                            $inObj = [ordered] @{
                                'DC Name' = $($DC.ToString().ToUpper().Split(".")[0])
                                'IP Address' = $DNSSetting.IPAddress
                                'Timeout' = ("$($DNSSetting.Timeout)/s")
                                'Use Root Hint' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $DNSSetting.UseRootHint)
                                'Use Recursion' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Recursion)
                            }
                            $OutObj += [pscustomobject]$inobj
                        }
                        catch {
                            Write-PscriboMessage -IsWarning "$($_.Exception.Message) (Forwarder Item)"
                        }
                    }
                    $TableParams = @{
                        Name = "Forwarders - $($Domain.ToString().ToUpper())"
                        List = $false
                        ColumnWidths = 35, 15, 15, 15, 20
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Sort-Object -Property 'DC Name' | Table @TableParams
                }
            }
            catch {
                Write-PscriboMessage -IsWarning "$($_.Exception.Message) (Forwarder Table)"
            }#>
        }
        catch {
            Write-PscriboMessage -IsWarning "$($_.Exception.Message) (DNS Infrastructure Section)"
        }
    }

    end {}

}