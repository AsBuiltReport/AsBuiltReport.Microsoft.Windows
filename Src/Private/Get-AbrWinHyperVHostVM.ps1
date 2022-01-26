function Get-AbrWinHyperVHostVM {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Hyper-V VM information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
        Author:         Andrew Ramsay
        Editor:         Jonathan Colon
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
        Write-PScriboMessage "Hyper-V InfoLevel set at $($InfoLevel.HyperV)."
        Write-PscriboMessage "Collecting Hyper-V VM information."
    }

    process {
        if ($InfoLevel.HyperV -ge 1) {
            $Vms = Get-VM -CimSession $TempCimSession
            if ($Vms) {
                try {
                    Section -Style Heading2 'Hyper-V VMs' {
                        Paragraph 'The following section details the Hyper-V VMs running on this host'
                        Blankline
                        $VmSummary = @()
                        foreach ($Vm in $Vms) {
                            try {
                                $TempVmSummary = [PSCustomObject]@{
                                    'VM Name' = $Vm.Name
                                    'vCPU Count' = $Vm.ProcessorCount
                                    'Memory (GB)' = [Math]::Round($Vm.MemoryAssigned / 1gb)
                                    'Memory Type' = $Vm.DynamicMemoryEnabled
                                    'Generation' = $Vm.Generation
                                    'Version' = $Vm.Version
                                    'Numa Aligned' = $Vm.NumaAligned
                                }
                                $VmSummary += $TempVmSummary
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Virtual Machines"
                            List = $false
                            ColumnWidths = 50, 50
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $VmSummary | Sort-Object 'VM Name' | Table @TableParams
                        foreach ($Vm in $Vms) {
                            try {
                                Section -Style Heading3 ($Vm.Name) {
                                    Paragraph 'The following sections detail the VM configuration settings'
                                    Blankline
                                    try {
                                        Section -Style Heading4 'Virtual Machine Configuration' {
                                            Blankline
                                            $VmConfiguration = [PSCustomObject]@{
                                                'VM id' = $Vm.VMid
                                                'VM Path' = $Vm.Path
                                                'Uptime' = $Vm.Uptime
                                                'vCPU Count' = $Vm.ProcessorCount
                                                'Memory Assigned (GB)' = [Math]::Round($Vm.MemoryAssigned / 1gb)
                                                'Dynamic Memory Enabled' = $Vm.DynamicMemoryEnabled
                                                'Memory Startup (GB)' = [Math]::Round($Vm.MemoryStartup / 1gb)
                                                'Memory Minimum (GB)' = [Math]::Round($Vm.MemoryMinimum / 1gb)
                                                'Memory Maximum (GB)' = [Math]::Round($Vm.MemoryMaximum / 1gb)
                                                'Numa Aligned' = $Vm.NumaAligned
                                                'Nuber of Numa Nodes' = $Vm.NumaNodesCount
                                                'Number of Numa Sockets' = $Vm.NumaSocketCount
                                                'Check Point Type' = $Vm.CheckpointType
                                                'Parent Snapshot Id' = $Vm.ParentSnapshotId
                                                'Parent Snapshot Name' = $Vm.ParentSnapshotName
                                                'Generation' = $Vm.Generation
                                                'DVD Drives' = $Vm.DVDDrives -Join ","
                                            }
                                            $TableParams = @{
                                                Name = "Virtual Machines"
                                                List = $true
                                                ColumnWidths = 40, 60
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VmConfiguration | Table @TableParams
                                        }
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                    try {
                                        Section -Style Heading4 'Virtual Machine Guest Integration Service' {
                                            Paragraph 'The following section details the status of Integration Services'
                                            Blankline
                                            $VmIntegrationServiceSummary = @()
                                            Foreach ($Service in ($Vm.VMIntegrationService)) {
                                                try {
                                                    $TempVmIntegrationServiceSummary = [PSCustomObject]@{
                                                        'Service Name' = $Service.Name
                                                        'Service State' = $Service.Enabled
                                                        'Primary Status' = $Service.PrimaryStatusDescription
                                                    }
                                                    $VmIntegrationServiceSummary += $TempVmIntegrationServiceSummary
                                                }
                                                catch {
                                                    Write-PscriboMessage -IsWarning $_.Exception.Message
                                                }
                                            }
                                            $TableParams = @{
                                                Name = "Integration Service"
                                                List = $false
                                                ColumnWidths = 40, 30, 30
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VmIntegrationServiceSummary | Table @TableParams
                                        }
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                    try {
                                        $VmNetworkAdapters = Get-VMNetworkAdapter -CimSession $TempCimSession -VMName $VM.Name
                                        if ($VmNetworkAdapters) {
                                            Section -Style Heading4 'VM Network Adapters' {
                                                Paragraph 'The following table details the network adapter details'
                                                BlankLine
                                                $VmNetworkAdapterReport = @()
                                                ForEach ($Adapter in $VmNetworkAdapters) {
                                                    try {
                                                        $TempVmNetworkAdapter = [PSCustomObject]@{
                                                            'Name' = $Adapter.Name
                                                            'Mac Address' = $Adapter.MacAddress
                                                            'IP Address' = $Adapter.IPAddresses[0]
                                                            'Switch Name' = $Adapter.SwitchName
                                                        }
                                                        $VmNetworkAdapterReport += $TempVmNetworkAdapter
                                                    }
                                                    catch {
                                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                                    }
                                                }

                                                $TableParams = @{
                                                    Name = "VM Network Adapters"
                                                    List = $false
                                                    ColumnWidths = 25, 25, 25, 25
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VmNetworkAdapterReport | Sort-Object -Property 'Name' | Table @TableParams
                                            }
                                        }
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                    try {
                                        $VmAdapterVlan = Get-VMNetworkAdapterVlan -CimSession $TempCimSession -VMName $VM.Name
                                        if ($VmAdapterVlan) {
                                            Section -Style Heading4 'VM Network Adpater VLANs' {
                                                Paragraph 'The following section details the VLAN configuration of VM Network Adapters'
                                                BlankLine
                                                $VmAdapterVlanReport = @()
                                                ForEach ($Adapter in $VmAdapterVlan) {
                                                    try {
                                                        $TempVmAdapterVlanReport = [PSCustomObject]@{
                                                            'Adapter Name' = $Adapter.ParentAdapter.Name
                                                            'Operation Mode' = $Adapter.OperationMode
                                                            'Vlan ID' = $Adapter.AccessVlanId
                                                            'Trunk Vlans' = $Adapter.AllowedVlanIdList -Join ","
                                                        }
                                                        $VmAdapterVlanReport += $TempVmAdapterVlanReport
                                                    }
                                                    catch {
                                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                                    }
                                                }

                                                $TableParams = @{
                                                    Name = "VM Network Adapter Vlans"
                                                    List = $false
                                                    ColumnWidths = 25, 25, 25, 25
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VmAdapterVlanReport | Sort-Object -Property 'Adapter Name' | Table @TableParams
                                            }
                                        }
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                    try {
                                        $VmHardDisks = Get-VMHardDiskDrive -CimSession $TempCimSession -VMName $VM.Name
                                        if ($VmHardDisks) {
                                            Section -Style Heading4 'VM Hard Disks' {
                                                Paragraph 'The following table details the VM hard disks'
                                                BlankLine
                                                $VmDiskReport = @()
                                                foreach ($VmHardDisk in $VMHardDisks) {
                                                    try {
                                                        $VmVhd = Get-VHD -CimSession $TempCimSession -Path $VmHardDisk.Path
                                                        $TempVmDiskReport = [PSCustomObject]@{
                                                            'Disk Path' = $VmVhd.Path
                                                            'Disk Format' = $VmVhd.VhdFormat
                                                            'Disk Type' = $VmVhd.VhdType
                                                            'Disk Used(GB)' = [Math]::Round($VmVhd.FileSize / 1gb)
                                                            'Disk Max(GB)' = [Math]::Round($VmVhd.Size / 1gb)
                                                            'Bus Type' = $VmHardDisk.ControllerType
                                                            'Bus No' = $VmHardDisk.ControllerNumber
                                                            'Bus Location' = $VmHardDisk.ControllerLocation
                                                        }
                                                        $VmDiskReport += $TempVmDiskReport
                                                    }
                                                    catch {
                                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                                    }
                                                }

                                                $TableParams = @{
                                                    Name = "VM Hard disks"
                                                    List = $false
                                                    ColumnWidths = 30, 10, 10, 10, 10, 10, 10, 10
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VmDiskReport | Sort-Object -Property 'Disk Path' | Table @TableParams
                                            }
                                        }
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
        }
    }

    end {}
}