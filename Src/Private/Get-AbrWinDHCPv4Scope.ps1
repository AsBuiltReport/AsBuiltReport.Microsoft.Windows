function Get-AbrWinDHCPv4Scope {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Windows DHCP Servers Scopes.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.2
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
        Write-PScriboMessage "DHCP InfoLevel set at $($InfoLevel.DHCP)."
        Write-PScriboMessage "Collecting Host DHCP Server information."
    }

    process {
        try {
            $DHCPScopes = Get-DhcpServerv4Scope -CimSession $TempCIMSession
            Write-PScriboMessage "Discovered '$(($DHCPScopes | Measure-Object).Count)' DHCP SCopes in $($System.split(".")[0])."
            if ($DHCPScopes) {
                Section -Style Heading3 "Scopes" {
                    Paragraph "The following section provides detailed information of the Scope configuration."
                    BlankLine
                    $OutObj = @()
                    foreach ($Scope in $DHCPScopes) {
                        try {
                            Write-PScriboMessage "Collecting DHCP Server $($Scope.ScopeId) Scope"
                            $SubnetMask = Convert-IpAddressToMaskLength $Scope.SubnetMask
                            $inObj = [ordered] @{
                                'Scope Id' = "$($Scope.ScopeId)/$($SubnetMask)"
                                'Scope Name' = $Scope.Name
                                'Scope Range' = "$($Scope.StartRange) - $($Scope.EndRange)"
                                'Lease Duration' = Switch ($Scope.LeaseDuration) {
                                    "10675199.02:48:05.4775807" { "Unlimited" }
                                    default { $Scope.LeaseDuration }
                                }
                                'State' = $Scope.State
                            }
                            $OutObj += [pscustomobject]$inobj
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    $TableParams = @{
                        Name = "Scopes - $($System.toUpper().split(".")[0])"
                        List = $false
                        ColumnWidths = 20, 20, 35, 15, 10
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Sort-Object -Property 'Scope Id' | Table @TableParams
                    try {
                        $DHCPStatistics = Get-DhcpServerv4ScopeStatistics -CimSession $TempCIMSession
                        if ($DHCPStatistics) {
                            Section -Style Heading4 "Scope Statistics" {
                                $OutObj = @()
                                foreach ($DHCPStatistic in $DHCPStatistics) {
                                    try {
                                        Write-PScriboMessage "Collecting DHCP Server $($DHCPStatistic.ScopeId) scope statistics"
                                        $inObj = [ordered] @{
                                            'Scope Id' = $DHCPStatistic.ScopeId
                                            'Free IP' = $DHCPStatistic.Free
                                            'In Use IP' = $DHCPStatistic.InUse
                                            'Percentage In Use' = [math]::Round($DHCPStatistic.PercentageInUse, 0)
                                            'Reserved IP' = $DHCPStatistic.Reserved
                                        }
                                        $OutObj += [pscustomobject]$inobj
                                    } catch {
                                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (IPv4 Scope Statistics Item)"
                                    }
                                }

                                if ($HealthCheck.DHCP.Statistics) {
                                    $OutObj | Where-Object { $_.'Percentage In Use' -gt '95' } | Set-Style -Style Warning -Property 'Percentage In Use'
                                }

                                $TableParams = @{
                                    Name = "Scope Statistics - $($System.toUpper().split(".")[0])"
                                    List = $false
                                    ColumnWidths = 20, 20, 20, 20, 20
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $OutObj | Sort-Object -Property 'Scope Id' | Table @TableParams
                            }
                        }
                    } catch {
                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (IPv4 Scope Statistics Table)"
                    }
                    try {
                        $DHCPv4Failovers = Get-DhcpServerv4Failover -CimSession $TempCIMSession
                        if ($DHCPv4Failovers) {
                            Section -Style Heading4 "Scope Failover" {
                                $OutObj = @()
                                foreach ($DHCPv4Failover in $DHCPv4Failovers) {
                                    try {
                                        Write-PScriboMessage "Collecting DHCP Server $($DHCPv4Failover.ScopeId) scope failover setting"
                                        $inObj = [ordered] @{
                                            'Partner DHCP Server' = $DHCPv4Failover.PartnerServer
                                            'Mode' = $DHCPv4Failover.Mode
                                            'LoadBalance Percent' = ConvertTo-EmptyToFiller ([math]::Round($DHCPv4Failover.LoadBalancePercent, 0))
                                            'Server Role' = ConvertTo-EmptyToFiller $DHCPv4Failover.ServerRole
                                            'Reserve Percent' = ConvertTo-EmptyToFiller ([math]::Round($DHCPv4Failover.ReservePercent, 0))
                                            'Max Client Lead Time' = ConvertTo-EmptyToFiller $DHCPv4Failover.MaxClientLeadTime
                                            'State Switch Interval' = ConvertTo-EmptyToFiller $DHCPv4Failover.StateSwitchInterval
                                            'Scope Ids' = $DHCPv4Failover.ScopeId
                                            'State' = $DHCPv4Failover.State
                                            'Auto State Transition' = ConvertTo-TextYN $DHCPv4Failover.AutoStateTransition
                                            'Authetication Enable' = ConvertTo-TextYN $DHCPv4Failover.EnableAuth
                                        }
                                        $OutObj = [pscustomobject]$inobj
                                    } catch {
                                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (IPv4 Scope Failover Item)"
                                    }
                                    if ($HealthCheck.DHCP.BP) {
                                        $OutObj | Where-Object { $_.'Authetication Enable' -eq 'No' } | Set-Style -Style Warning -Property 'Authetication Enable'
                                        $OutObj | Where-Object { $_.'State' -ne 'Normal' } | Set-Style -Style Warning -Property 'State'
                                    }

                                    $TableParams = @{
                                        Name = "Scope Failover Cofiguration - $($System.split(".", 2).ToUpper()[0])"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $OutObj | Table @TableParams
                                }
                            }
                        }
                    } catch {
                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (IPv4 Scope Failover Table)"
                    }
                    try {
                        $DHCPv4Bindings = Get-DhcpServerv4Binding -CimSession $TempCIMSession
                        if ($DHCPv4Bindings) {
                            Section -Style Heading4 "Network Interface Binding" {
                                $OutObj = @()
                                foreach ($DHCPv4Binding in $DHCPv4Bindings) {
                                    try {
                                        Write-PScriboMessage "Collecting DHCP Server $($DHCPv4Binding.InterfaceAlias) binding."
                                        $SubnetMask = Convert-IpAddressToMaskLength $DHCPv4Binding.SubnetMask
                                        $inObj = [ordered] @{
                                            'Interface Alias' = $DHCPv4Binding.InterfaceAlias
                                            'IP Address' = $DHCPv4Binding.IPAddress
                                            'Subnet Mask' = $DHCPv4Binding.SubnetMask
                                            'State' = Switch ($DHCPv4Binding.BindingState) {
                                                "" { "-"; break }
                                                $Null { "-"; break }
                                                "True" { "Enabled" }
                                                "False" { "Disabled" }
                                                default { $DHCPv4Binding.BindingState }
                                            }
                                        }
                                        $OutObj += [pscustomobject]$inobj
                                    } catch {
                                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (IPv4 Network Interface binding Item)"
                                    }
                                }
                                $TableParams = @{
                                    Name = "Network Interface binding - $($System.split(".", 2).ToUpper()[0])"
                                    List = $false
                                    ColumnWidths = 25, 25, 25, 25
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $OutObj | Table @TableParams
                            }
                        }
                    } catch {
                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (IPv4 Network Interface binding Table)"
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (IPv4 Scope Summary)"
        }
    }
    end {}
}