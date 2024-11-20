function Get-AbrWinHyperVNetworking {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Hyper-V Networking information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.6
        Author:         Andrew Ramsay
        Editor:         Jonathan Colon
        Twitter:        @asbuiltreport
        Github:         AsBuiltReport
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows
    #>
    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "Hyper-V InfoLevel set at $($InfoLevel.HyperV)."
        Write-PScriboMessage "Collecting Hyper-V Networking information."
    }

    process {
        if ($InfoLevel.HyperV -ge 1) {
            try {
                try {
                    Section -Style Heading3 "Hyper-V MAC Pool settings" {
                        $OutObj = @()
                        $inObj = [ordered] @{
                            'Mac Address Minimum' = Switch (($VmHost.MacAddressMinimum).Length) {
                                0 { "--" }
                                12 { $VmHost.MacAddressMinimum -replace '..(?!$)', '$&:' }
                                default { $VmHost.MacAddressMinimum }
                            }
                            'Mac Address Maximum' = Switch (($VmHost.MacAddressMaximum).Length) {
                                0 { "--" }
                                12 { $VmHost.MacAddressMaximum -replace '..(?!$)', '$&:' }
                                default { $VmHost.MacAddressMinimum }
                            }
                        }
                        $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)

                        $TableParams = @{
                            Name = "Host MAC Pool"
                            List = $false
                            ColumnWidths = 50, 50
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj |  Table @TableParams
                    }
                } catch {
                    Write-PScriboMessage -IsWarning $_.Exception.Message
                }
                try {
                    $VmOsAdapters = Invoke-Command -Session $TempPssSession { Get-VMNetworkAdapter -ManagementOS | Select-Object -Property * }
                    if ($VmOsAdapters) {
                        Section -Style Heading3 "Hyper-V Management OS Adapters" {
                            Paragraph 'The following table details the Management OS Virtual Adapters created on Virtual Switches'
                            BlankLine
                            $OutObj = @()
                            Foreach ($VmOsAdapter in $VmOsAdapters) {
                                try {
                                    $AdapterVlan = Invoke-Command -Session $TempPssSession { Get-VMNetworkAdapterVlan -ManagementOS -VMNetworkAdapterName ($using:VmOsAdapter).Name | Select-Object -Property * }
                                    $inObj = [ordered] @{
                                        'Name' = $VmOsAdapter.Name
                                        'Switch Name' = $VmOsAdapter.SwitchName
                                        'Mac Address' = $VmOsAdapter.MacAddress
                                        'IPv4 Address' = $VmOsAdapter.IPAddresses
                                        'Adapter Mode' = $AdapterVlan.OperationMode
                                        'Vlan ID' = $AdapterVlan.AccessVlanId
                                    }
                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                } catch {
                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                }
                            }
                            $TableParams = @{
                                Name = "VM Management OS Adapters"
                                List = $true
                                ColumnWidths = 50, 50
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $OutObj | Sort-Object -Property 'Name' | Table @TableParams
                        }
                    }
                } catch {
                    Write-PScriboMessage -IsWarning $_.Exception.Message
                }
                try {
                    $VmSwitches = Invoke-Command -Session $TempPssSession { Get-VMSwitch }
                    if ($VmSwitches) {
                        Section -Style Heading3 "Hyper-V vSwitch Settings" {
                            Paragraph 'The following table provide a summary of Hyper-V configured vSwitches'
                            BlankLine
                            $OutObj = @()
                            ForEach ($VmSwitch in $VmSwitches) {
                                try {
                                    $inObj = [ordered] @{
                                        'Switch Name' = $VmSwitch.Name
                                        'Switch Type' = $VmSwitch.SwitchType
                                        'Embedded Team' = $VmSwitch.EmbeddedTeamingEnabled
                                        'Interface Description' = $VmSwitch.NetAdapterInterfaceDescription
                                    }
                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                } catch {
                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                }
                            }

                            $TableParams = @{
                                Name = "Virtual Switch Summary"
                                List = $false
                                ColumnWidths = 30, 20, 20, 30
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $OutObj | Sort-Object -Property 'Switch Name' | Table @TableParams

                            Foreach ($VmSwitch in $VmSwitches) {
                                try {
                                    Section -ExcludeFromTOC -Style NOTOCHeading4 ($VmSwitch.Name) {
                                        $OutObj = @()
                                        $inObj = [ordered] @{
                                            'Switch Name' = $VmSwitch.Name
                                            'Switch Type' = $VmSwitch.SwitchType
                                            'Switch Embedded Teaming Status' = $VmSwitch.EmbeddedTeamingEnabled
                                            'Bandwidth Reservation Mode' = $VmSwitch.BandwidthReservationMode
                                            'Bandwidth Reservation Percentage' = $VmSwitch.Percentage
                                            'Management OS Allowed' = $VmSwitch.AllowManagementOS
                                            'Physical Adapters' = $VmSwitch.NetAdapterInterfaceDescriptions -Join ","
                                            'IOV Support' = $VmSwitch.IovSupport
                                            'IOV Support Reasons' = $VmSwitch.IovSupportReasons
                                            'Available VM Queues' = $VmSwitch.AvailableVMQueues
                                            'Packet Direct Enabled' = $VmSwitch.PacketDirectinUse
                                        }
                                        $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)


                                        $TableParams = @{
                                            Name = "VM Switch Details"
                                            List = $true
                                            ColumnWidths = 50, 50
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $OutObj | Table @TableParams
                                    }
                                } catch {
                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                }
                            }
                        }
                    }
                } catch {
                    Write-PScriboMessage -IsWarning $_.Exception.Message
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}
