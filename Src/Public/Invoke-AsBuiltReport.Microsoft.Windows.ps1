function Invoke-AsBuiltReport.Microsoft.Windows {
    <#
    .SYNOPSIS
        PowerShell script to document the configuration of Microsoft Windows Server in Word/HTML/Text formats
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

    # Do not remove or add to these parameters
    param (
        [String[]] $Target,
        [PSCredential] $Credential
    )

    # Import Report Configuration
    $Report = $ReportConfig.Report
    $InfoLevel = $ReportConfig.InfoLevel
    $Options = $ReportConfig.Options

    # Used to set values to TitleCase where required
    $TextInfo = (Get-Culture).TextInfo

    # Update/rename the $System variable and build out your code within the ForEach loop. The ForEach loop enables AsBuiltReport to generate an as built configuration against multiple defined targets.

    #region foreach loop
    foreach ($System in $Target) {
        Section -Style Heading1 $System {
            Paragraph "The following table details the Windows Host $System"
            BlankLine
            try {
                $script:TempPssSession = New-PSSession $System -Credential $Credential -Authentication Default -ErrorAction stop
                $script:TempCimSession = New-CimSession $System -Credential $Credential -Authentication Default -ErrorAction stop
            }
            catch {
                Write-PScriboMessage -IsWarning  "Unable to connect to $($System)"
                throw
            }
            $script:HostInfo = Invoke-Command -Session $TempPssSession { Get-ComputerInfo }
            $script:HostCPU = Get-CimInstance -Class Win32_Processor -CimSession $TempCimSession
            $script:HostComputer = Get-CimInstance -Class Win32_ComputerSystem -CimSession $TempCimSession
            $script:HostBIOS = Get-CimInstance -Class Win32_Bios -CimSession $TempCimSession
            $script:HostLicense =  Get-CimInstance -Query 'Select * from SoftwareLicensingProduct' -CimSession $TempCimSession | Where-Object { $_.LicenseStatus -eq 1 }
            #Host Hardware
            Get-AbrWinHostHWSummary
            #Host OS
            try {
                Section -Style Heading2 'Host Operating System' {
                    Paragraph 'The following settings details host OS Settings'
                    Blankline
                    #Host OS Configuration
                    Get-AbrWinOSConfig
                    #Host Hotfixes
                    Get-AbrWinOSHotfix
                    #Host Drivers
                    Get-AbrWinOSDriver
                    #Host Roles and Features
                    Get-AbrWinOSRoleFeature
                    #Host 3rd Party Applications
                    Get-AbrWinApplication
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
            #Local Users and Groups
            try {
                Section -Style Heading2 'Local Users and Groups' {
                    Paragraph 'The following section details local users and groups configured'
                    Blankline
                    #Local Users
                    Get-AbrWinLocalUser
                    #Local Groups
                    Get-AbrWinLocalGroup
                    #Local Administrators
                    Get-AbrWinLocalAdmin
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
            #Host Firewall
            Get-AbrWinNetFirewall
            #Host Networking
            try {
                Section -Style Heading2 'Host Networking' {
                    Paragraph 'The following section details Host Network Configuration'
                    Blankline
                    Get-AbrWinNetAdapter
                    Get-AbrWinNetIPAddress
                    Get-AbrWinNetDNSClient
                    Get-AbrWinNetDNSServer
                    Get-AbrWinNetTeamInterface
                    Get-AbrWinNetAdapterMTU
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
            #Host Storage
            try {
                Section -Style Heading2 'Host Storage' {
                    Paragraph 'The following section details the storage configuration of the host'
                    #Local Disks
                    Get-AbrWinHostStorage
                    #Local Volumes
                    Get-AbrWinHostStorageVolume
                    #iSCSI
                    Get-AbrWinHostStorageISCSI
                    #MPIO
                    Get-AbrWinHostStorageMPIO
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
            #HyperV Configuration
            try {
                $HyperVInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*Hyper-V*" } }
                if ($HyperVInstalledCheck.InstallState -eq "Installed") {
                    Section -Style Heading2 "Hyper-V Configuration Settings" {
                        Paragraph 'The following table details the Hyper-V Server Settings'
                        Blankline
                        $VmHost = Invoke-Command -Session $TempPssSession { Get-VMHost }
                        if ($VmHost) {
                            $VmHostReport = [PSCustomObject]@{
                                'Logical Processor Count' = $VmHost.LogicalProcessorCount
                                'Memory Capacity (GB)' = [Math]::Round($VmHost.MemoryCapacity / 1gb)
                                'VM Default Path' = $VmHost.VirtualMachinePath
                                'VM Disk Default Path' = $VmHost.VirtualHardDiskPath
                                'Supported VM Versions' = $VmHost.SupportedVmVersions -Join ","
                                'Numa Spannning Enabled' = ConvertTo-TextYN $VmHost.NumaSpanningEnabled
                                'Iov Support' = ConvertTo-TextYN $VmHost.IovSupport
                                'VM Migrations Enabled' = ConvertTo-TextYN $VmHost.VirtualMachineMigrationEnabled
                                'Allow any network for Migrations' = ConvertTo-TextYN $VmHost.UseAnyNetworkForMigration
                                'VM Migration Authentication Type' = $VmHost.VirtualMachineMigrationAuthenticationType
                                'Max Concurrent Storage Migrations' = $VmHost.MaximumStorageMigrations
                                'Max Concurrent VM Migrations' = $VmHost.MaximumStorageMigrations
                            }
                            $TableParams = @{
                                Name = "Hyper-V Host Settings"
                                List = $true
                                ColumnWidths = 50, 50
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmHostReport | Table @TableParams
                        }
                        try {
                            $VmHostNumaNodes = Invoke-Command -Session $TempPssSession { Get-VMHostNumaNode }
                            if ($VmHostNumaNodes) {
                                Section -Style Heading3 "Hyper-V NUMA Boundaries" {
                                    Paragraph 'The following table details the NUMA nodes on the host'
                                    Blankline
                                    [array]$VmHostNumaReport = @()
                                    foreach ($Node in $VmHostNumaNodes) {
                                        try {
                                            $TempVmHostNumaReport = [PSCustomObject]@{
                                                'Numa Node Id' = $Node.NodeId
                                                'Memory Available(GB)' = ($Node.MemoryAvailable)/1024
                                                'Memory Total(GB)' = ($Node.MemoryTotal)/1024
                                            }
                                            $VmHostNumaReport += $TempVmHostNumaReport
                                        }
                                        catch {
                                            Write-PscriboMessage -IsWarning $_.Exception.Message
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "Host NUMA Nodes"
                                        List = $false
                                        ColumnWidths = 34, 33, 33
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VmHostNumaReport | Sort-Object -Property 'Numa Node Id' | Table @TableParams
                                }
                            }
                        }
                        catch {
                            Write-PscriboMessage -IsWarning $_.Exception.Message
                        }
                        try {
                            Section -Style Heading3 "Hyper-V MAC Pool settings" {
                                Paragraph 'The following table details the Hyper-V MAC Pool'
                                Blankline
                                $VmHostMacPool = [PSCustomObject]@{
                                    'Mac Address Minimum' = Switch (($VmHost.MacAddressMinimum).Length) {
                                        0 {"-"}
                                        12 {$VmHost.MacAddressMinimum -replace '..(?!$)', '$&:'}
                                        default {$VmHost.MacAddressMinimum}
                                    }
                                    'Mac Address Maximum' = Switch (($VmHost.MacAddressMaximum).Length) {
                                        0 {"-"}
                                        12 {$VmHost.MacAddressMaximum -replace '..(?!$)', '$&:'}
                                        default {$VmHost.MacAddressMinimum}
                                    }
                                }
                                $TableParams = @{
                                    Name = "Host MAC Pool"
                                    List = $false
                                    ColumnWidths = 50, 50
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $VmHostMacPool |  Table @TableParams
                            }
                        }
                        catch {
                            Write-PscriboMessage -IsWarning $_.Exception.Message
                        }
                        try {
                            $VmOsAdapters = Get-VMNetworkAdapter -CimSession $TempCimSession -ManagementOS
                            if ($VmOsAdapters) {
                                Section -Style Heading3 "Hyper-V Management OS Adapters" {
                                    Paragraph 'The following table details the Management OS Virtual Adapters created on Virtual Switches'
                                    Blankline
                                    $VmOsAdapterReport = @()
                                    Foreach ($VmOsAdapter in $VmOsAdapters) {
                                        try {
                                            $AdapterVlan = Get-VMNetworkAdapterVlan -CimSession $TempCimSession -ManagementOS -VMNetworkAdapterName $VmOsAdapter.Name
                                            $TempVmOsAdapterReport = [PSCustomObject]@{
                                                'Name' = $VmOsAdapter.Name
                                                'Switch Name' = $VmOsAdapter.SwitchName
                                                'Mac Address' = $VmOsAdapter.MacAddress
                                                'IPv4 Address' = $VmOsAdapter.IPAddresses -Join ","
                                                'Adapter Mode' = $AdapterVlan.OperationMode
                                                'Vlan ID' = $AdapterVlan.AccessVlanId
                                            }
                                            $VmOsAdapterReport += $TempVmOsAdapterReport
                                        }
                                        catch {
                                            Write-PscriboMessage -IsWarning $_.Exception.Message
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "VM Management OS Adapters"
                                        List = $false
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VmOsAdapterReport | Sort-Object -Property 'Name' | Table @TableParams
                                }
                            }
                        }
                        catch {
                            Write-PscriboMessage -IsWarning $_.Exception.Message
                        }
                        $VmSwitches = Invoke-Command -Session $TempPssSession { Get-VMSwitch }
                        if ($VmSwitches) {
                            Section -Style Heading3 "Hyper-V vSwitch Settings" {
                                Paragraph 'The following table details the Hyper-V vSwitches configured'
                                Blankline
                                $VmSwitchesReport = @()
                                ForEach ($VmSwitch in $VmSwitches) {
                                    try {
                                        $TempVmSwitchesReport = [PSCustomObject]@{
                                            'Switch Name' = $VmSwitch.Name
                                            'Switch Type' = $VmSwitch.SwitchType
                                            'Embedded Team' = $VmSwitch.EmbeddedTeamingEnabled
                                            'Interface Description' = $VmSwitch.NetAdapterInterfaceDescription
                                        }
                                        $VmSwitchesReport += $TempVmSwitchesReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
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
                                $VmSwitchesReport | Sort-Object -Property 'Switch Name' | Table @TableParams

                                Foreach ($VmSwitch in $VmSwitches) {
                                    try {
                                        Section -Style Heading4 ($VmSwitch.Name) {
                                            Paragraph 'The following table details the Hyper-V vSwitch'
                                            Blankline
                                            $VmSwitchReport = [PSCustomObject]@{
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

                                            $TableParams = @{
                                                Name = "VM Switch Details"
                                                List = $true
                                                ColumnWidths = 50, 50
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VmSwitchReport | Table @TableParams
                                        }
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                            }
                        }
                    }
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
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
        }
        Remove-PSSession $TempPssSession
        Remove-CimSession $TempCimSession
    }
    #endregion foreach loop
}
