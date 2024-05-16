function Get-AbrWinDNSZone {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Windows Domain Name System Zone information.
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
        Write-PScriboMessage "DNS InfoLevel set at $($InfoLevel.DNS)."
        Write-PScriboMessage "Collecting Host DNS Server information."
    }

    process {
        try {
            $DNSSetting = Get-DnsServerZone -CimSession $TempCIMSession | Where-Object { $_.IsReverseLookupZone -like "False" -and $_.ZoneType -notlike "Forwarder" }
            if ($DNSSetting) {
                Section -Style Heading3 "DNS Zone Configuration" {
                    Paragraph "The following table details zones configuration settings"
                    BlankLine
                    $OutObj = @()
                    foreach ($Zones in $DNSSetting) {
                        try {
                            Write-PScriboMessage "Collecting Actve Directory DNS Zone: '$($Zones.ZoneName)' on $DC"
                            $inObj = [ordered] @{
                                'Zone Name' = ConvertTo-EmptyToFiller $Zones.ZoneName
                                'Zone Type' = ConvertTo-EmptyToFiller $Zones.ZoneType
                                'Replication Scope' = ConvertTo-EmptyToFiller $Zones.ReplicationScope
                                'Dynamic Update' = ConvertTo-EmptyToFiller $Zones.DynamicUpdate
                                'DS Integrated' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Zones.IsDsIntegrated)
                                'Read Only' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Zones.IsReadOnly)
                                'Signed' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Zones.IsSigned)
                            }
                            $OutObj += [pscustomobject]$inobj
                        } catch {
                            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Domain Name System Zone Item)"
                        }
                    }

                    $TableParams = @{
                        Name = "Zones - $($System.toUpper().split(".")[0])"
                        List = $false
                        ColumnWidths = 25, 15, 12, 12, 12, 12, 12
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Sort-Object -Property 'Zone Name' | Table @TableParams

                    if ($InfoLevel.DNS -ge 2) {
                        try {
                            $DNSSetting = Invoke-Command -Session $TempPssSession { Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DNS Server\Zones\*" | Get-ItemProperty | Where-Object { $_ -match 'SecondaryServers' } }
                            if ($DNSSetting) {
                                Section -Style Heading4 "Zone Transfers" {
                                    Paragraph "The following table details zone transfer configuration settings"
                                    BlankLine
                                    $OutObj = @()
                                    foreach ($Zone in $DNSSetting) {
                                        try {
                                            $inObj = [ordered] @{
                                                'Zone Name' = $Zone.PSChildName
                                                'Secondary Servers' = ConvertTo-EmptyToFiller ($Zone.SecondaryServers -join ", ")
                                                'Notify Servers' = ConvertTo-EmptyToFiller $Zone.NotifyServers
                                                'Secure Secondaries' = Switch ($Zone.SecureSecondaries) {
                                                    "0" { "Send zone transfers to all secondary servers that request them." }
                                                    "1" { "Send zone transfers only to name servers that are authoritative for the zone." }
                                                    "2" { "Send zone transfers only to servers you specify in Secondary Servers." }
                                                    "3" { "Do not send zone transfers." }
                                                    default { $Zone.SecureSecondaries }
                                                }
                                            }
                                            $OutObj = [pscustomobject]$inobj

                                            $TableParams = @{
                                                Name = "Zone Transfers - $($Zone.PSChildName.toUpper())"
                                                List = $true
                                                ColumnWidths = 40, 60
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $OutObj | Table @TableParams
                                        } catch {
                                            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Zone Transfers Item)"
                                        }
                                    }
                                }
                            }
                        } catch {
                            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Zone Transfers Table)"
                        }
                    }
                    try {
                        $DNSSetting = Get-DnsServerZone -CimSession $TempCIMSession | Where-Object { $_.IsReverseLookupZone -like "True" }
                        if ($DNSSetting) {
                            Section -Style Heading4 "Reverse Lookup Zone Configuration" {
                                Paragraph "The following table details reverse looup zone configuration settings"
                                BlankLine
                                $OutObj = @()
                                foreach ($Zones in $DNSSetting) {
                                    try {
                                        Write-PScriboMessage "Collecting Actve Directory DNS Zone: '$($Zones.ZoneName)'"
                                        $inObj = [ordered] @{
                                            'Zone Name' = ConvertTo-EmptyToFiller $Zones.ZoneName
                                            'Zone Type' = ConvertTo-EmptyToFiller $Zones.ZoneType
                                            'Replication Scope' = ConvertTo-EmptyToFiller $Zones.ReplicationScope
                                            'Dynamic Update' = ConvertTo-EmptyToFiller $Zones.DynamicUpdate
                                            'DS Integrated' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Zones.IsDsIntegrated)
                                            'Read Only' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Zones.IsReadOnly)
                                            'Signed' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Zones.IsSigned)
                                        }
                                        $OutObj += [pscustomobject]$inobj
                                    } catch {
                                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Reverse Lookup Zone Configuration Item)"
                                    }
                                }

                                $TableParams = @{
                                    Name = "Zones - $($System.toUpper().split(".")[0])"
                                    List = $false
                                    ColumnWidths = 25, 15, 12, 12, 12, 12, 12
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $OutObj | Sort-Object -Property 'Zone Name' | Table @TableParams
                            }
                        }
                    } catch {
                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Reverse Lookup Zone Configuration Table)"
                    }
                    try {
                        $DNSSetting = Get-DnsServerZone -CimSession $TempCIMSession | Where-Object { $_.IsReverseLookupZone -like "False" -and $_.ZoneType -like "Forwarder" }
                        if ($DNSSetting) {
                            Section -Style Heading4 "Conditional Forwarder" {
                                Paragraph "The following table details conditional forwarder configuration settings"
                                BlankLine
                                $OutObj = @()
                                foreach ($Zones in $DNSSetting) {
                                    try {
                                        Write-PScriboMessage "Collecting Actve Directory DNS Zone: '$($Zones.ZoneName)'"
                                        $inObj = [ordered] @{
                                            'Zone Name' = $Zones.ZoneName
                                            'Zone Type' = $Zones.ZoneType
                                            'Replication Scope' = $Zones.ReplicationScope
                                            'Master Servers' = $Zones.MasterServers
                                            'DS Integrated' = ConvertTo-TextYN $Zones.IsDsIntegrated
                                        }
                                        $OutObj += [pscustomobject]$inobj
                                    } catch {
                                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Conditional Forwarder Item)"
                                    }
                                }

                                $TableParams = @{
                                    Name = "Conditional Forwarders - $($System.toUpper().split(".")[0])"
                                    List = $false
                                    ColumnWidths = 25, 20, 20, 20, 15
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $OutObj | Sort-Object -Property 'Zone Name' | Table @TableParams
                            }
                        }
                    } catch {
                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Conditional Forwarder Table)"
                    }
                    if ($InfoLevel.DNS -ge 2) {
                        try {
                            $DNSSetting = Get-DnsServerZone -CimSession $TempCIMSession | Where-Object { $_.IsReverseLookupZone -like "False" -and $_.ZoneType -eq "Primary" } | Select-Object -ExpandProperty ZoneName
                            $Zones = Get-DnsServerZoneAging -CimSession $TempCIMSession -Name $DNSSetting
                            if ($Zones) {
                                Section -Style Heading4 "Zone Scope Aging Properties" {
                                    Paragraph "The following table details zone configuration aging settings"
                                    BlankLine
                                    $OutObj = @()
                                    foreach ($Settings in $Zones) {
                                        try {
                                            Write-PScriboMessage "Collecting Actve Directory DNS Zone: '$($Settings.ZoneName)'"
                                            $inObj = [ordered] @{
                                                'Zone Name' = ConvertTo-EmptyToFiller $Settings.ZoneName
                                                'Aging Enabled' = ConvertTo-EmptyToFiller (ConvertTo-TextYN $Settings.AgingEnabled)
                                                'Refresh Interval' = ConvertTo-EmptyToFiller $Settings.RefreshInterval
                                                'NoRefresh Interval' = ConvertTo-EmptyToFiller $Settings.NoRefreshInterval
                                                'Available For Scavenge' = Switch ($Settings.AvailForScavengeTime) {
                                                    "" { "--"; break }
                                                    $Null { "--"; break }
                                                    default { (ConvertTo-EmptyToFiller ($Settings.AvailForScavengeTime).ToUniversalTime().toString("r")); break }
                                                }
                                            }
                                            $OutObj += [pscustomobject]$inobj
                                        } catch {
                                            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Zone Scope Aging Item)"
                                        }
                                    }

                                    if ($HealthCheck.DNS.Aging) {
                                        $OutObj | Where-Object { $_.'Aging Enabled' -ne 'Yes' } | Set-Style -Style Warning -Property 'Aging Enabled'
                                    }

                                    $TableParams = @{
                                        Name = "Zone Aging Properties - $($System.toUpper().split(".")[0])"
                                        List = $false
                                        ColumnWidths = 25, 10, 15, 15, 35
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $OutObj | Sort-Object -Property 'Zone Name' | Table @TableParams
                                }
                            }
                        } catch {
                            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Zone Scope Aging Table)"
                        }
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Global DNS Zone Information)"
        }
    }

    end {}

}