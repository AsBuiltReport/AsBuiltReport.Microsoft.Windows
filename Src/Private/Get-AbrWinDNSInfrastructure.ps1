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
        Write-PScriboMessage "DNS InfoLevel set at $($InfoLevel.DNS)."
        Write-PscriboMessage "Collecting Host DNS Server information."
    }

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
                        Section -Style Heading3 "DNS IP Configuration" {
                            Paragraph "The following table details DNS Server IP Configuration Settings"
                            Blankline
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

            #---------------------------------------------------------------------------------------------#
            #                                 DNS Scanvenging Section                                     #
            #---------------------------------------------------------------------------------------------#
            if ($InfoLevel.DNS -ge 2) {
                try {
                    $DNSSetting = Get-DnsServerScavenging -CimSession $TempCIMSession
                    if ($DNSSetting) {
                        Section -Style Heading3 "Scavenging Options" {
                            Paragraph "The following table details scavenging configuration settings"
                            Blankline
                            $OutObj = @()
                            try {
                                $inObj = [ordered] @{
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
                                $TableParams = @{
                                    Name = "Scavenging - $($System.toUpper().split(".")[0])"
                                    List = $false
                                    ColumnWidths = 20, 20, 20, 20, 20
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $OutObj | Table @TableParams
                            }
                            catch {
                                Write-PscriboMessage -IsWarning "$($_.Exception.Message) (Scavenging Item)"
                            }
                        }
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
                Section -Style Heading3 "Forwarder Options" {
                    Paragraph "The following table details forwarder configuration settings"
                    Blankline
                    $OutObj = @()
                    try {
                        $DNSSetting = Get-DnsServerForwarder -CimSession $TempCIMSession
                        $Recursion = Get-DnsServerRecursion -CimSession $TempCIMSession
                        $inObj = [ordered] @{
                            'IP Address' = $DNSSetting.IPAddress -join ","
                            'Timeout' = ("$($DNSSetting.Timeout)/s")
                            'Use Root Hint' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $DNSSetting.UseRootHint)
                            'Use Recursion' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Recursion.Enable)
                        }
                        $OutObj += [pscustomobject]$inobj
                        $TableParams = @{
                            Name = "Forwarders - $($System.toUpper().split(".")[0])"
                            List = $false
                            ColumnWidths = 25, 25, 25, 25
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Table @TableParams
                    }
                    catch {
                        Write-PscriboMessage -IsWarning "$($_.Exception.Message) (Forwarder Item)"
                    }
                }
            }
            catch {
                Write-PscriboMessage -IsWarning "$($_.Exception.Message) (Forwarder Table)"
            }
        }
        catch {
            Write-PscriboMessage -IsWarning "$($_.Exception.Message) (DNS Infrastructure Section)"
        }
    }

    end {}

}