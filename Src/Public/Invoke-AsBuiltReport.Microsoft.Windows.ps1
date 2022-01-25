function Invoke-AsBuiltReport.Microsoft.Windows {
    <#
    .SYNOPSIS
        PowerShell script to document the configuration of Microsoft Windows Server in Word/HTML/Text formats
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.1.0
        Author:         Andrew Ramsay
        Twitter:
        Github:
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
            $script:TempPssSession = New-PSSession $System -Credential $Credential -Authentication Default
            $script:TempCimSession = New-CimSession $System -Credential $Credential -Authentication Default
            $script:HostInfo = Invoke-Command -Session $TempPssSession { Get-ComputerInfo }
            $script:HostCPU = Get-CimInstance -Class Win32_Processor -CimSession $TempCimSession
            $script:HostComputer = Get-CimInstance -Class Win32_ComputerSystem -CimSession $TempCimSession
            $script:HostBIOS = Get-CimInstance -Class Win32_Bios -CimSession $TempCimSession
            $script:HostLicense =  Get-CimInstance -Query 'Select * from SoftwareLicensingProduct' -CimSession $TempCimSession | Where-Object { $_.LicenseStatus -eq 1 }
            #Host Hardware
            try {
                Section -Style Heading2 'Host Hardware Settings' {
                    Paragraph 'The following section details hardware settings for the host'
                    $HostHardware = [PSCustomObject] @{
                        'Manufacturer' = $HostComputer.Manufacturer
                        'Model' = $HostComputer.Model
                        'Product ID' = $HostComputer.SystemSKUNumbe
                        'Serial Number' = $HostBIOS.SerialNumber
                        'BIOS Version' = $HostBIOS.Version
                        'Processor Manufacturer' = $HostCPU[0].Manufacturer
                        'Processor Model' = $HostCPU[0].Name
                        'Number of Processors' = $HostCPU.Length
                        'Number of CPU Cores' = $HostCPU[0].NumberOfCores
                        'Number of Logical Cores' = $HostCPU[0].NumberOfLogicalProcessors
                        'Physical Memory (GB)' = [Math]::Round($HostComputer.TotalPhysicalMemory / 1Gb)
                    }
                    $HostHardware | Table -Name 'Host Hardware Specifications' -List -ColumnWidths 50, 50
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
            #HP iLO Configuration
            if ($HostComputer.Manufacturer -eq "HPE") {
                Section -Style Heading2 'Host iLO Settings' {
                    Paragraph 'The following section details HPE iLO settings for the host'
                    }
            }
            #Dell iDRAC Configuration
            if ($HostComputer.Manufacturer -eq "Dell"){
                Section -Style Heading2 'Host iLO Settings' {
                    Paragraph 'The following section details HPE iLO settings for the host'
                    }
            }
            #Host OS
            try {
                Section -Style Heading2 'Host OS' {
                    Paragraph 'The following settings details host OS Settings'
                    Section -Style Heading3 'OS Configuration' {
                        Paragraph 'The following section details hos OS configuration'
                        $HostOSReport = [PSCustomObject] @{
                        'Windows Product Name' = $HostInfo.WindowsProductName
                        'Windows Version' = $HostInfo.WindowsCurrentVersion
                        'Windows Build Number' = $HostInfo.OsVersion
                        'Windows Install Type' = $HostInfo.WindowsInstallationType
                        'AD Domain' = $HostInfo.CsDomain
                        'Windows Installation Date' = switch (($HostInfo.OsInstallDate).count) {
                            0 {'-'}
                            default {$HostInfo.OsInstallDate.ToShortDateString()}
                        }
                        'Time Zone' = $HostInfo.TimeZone
                        'License Type' = $HostLicense.ProductKeyChannel
                        'Partial Product Key' = $HostLicense.PartialProductKey
                        }
                        $HostOSReport | Table -Name 'Host OS Settings' -List -ColumnWidths 50, 50
                    }
                    #Host Hotfixes
                    try {
                        $HotFixes = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-HotFix }
                        if ($HotFixes) {
                            Section -Style Heading3 'Host Hotfixes' {
                                Paragraph 'The following table details the OS Hotfixes installed'
                                $HotFixReport = @()
                                Foreach ($HotFix in $HotFixes) {
                                    try {
                                        $TempHotFix = [PSCustomObject] @{
                                            'Hotfix ID' = $HotFix.HotFixID
                                            'Description' = $HotFix.Description
                                            'Installation Date' = $HotFix.InstalledOn.ToShortDateString()
                                        }
                                        $HotfixReport += $TempHotFix
                                    } catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $HotFixReport | Table -Name 'Host Hotfixes' -ColumnWidths 34, 33, 33
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    #Host Drivers
                    try {
                        $HostDriversList = Invoke-Command -Session $TempPssSession { Get-WindowsDriver -Online }
                        if ($HostDriversList) {
                            Section -Style Heading3 'Host Drivers' {
                                Paragraph 'The following section details host drivers'
                                Invoke-Command -Session $TempPssSession { Import-Module DISM }
                                $HostDriverReport = @()
                                ForEach ($HostDriver in $HostDriversList) {
                                    try {
                                        $TempDriver = [PSCustomObject] @{
                                            'Class Description' = $HostDriver.ClassDescription
                                            'Provider Name' = $HostDriver.ProviderName
                                            'Driver Version' = $HostDriver.Version
                                            'Version Date' = $HostDriver.Date
                                        }
                                        $HostDriverReport += $TempDriver
                                    } catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $HostDriverReport | Table -Name 'Host Drivers' -ColumnWidths 30, 30, 20, 20
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    #Host Roles and Features
                    try {
                        $HostRolesAndFeatures = Invoke-Command -Session $TempPssSession -ScriptBlock{Get-WindowsFeature | Where-Object { $_.Installed -eq $True }}
                        if ($HostRolesAndFeatures) {
                            Section -Style Heading3 'Roles and Features' {
                                Paragraph 'The following settings details host roles and features installed'
                                [array]$HostRolesAndFeaturesReport = @()
                                ForEach ($HostRoleAndFeature in $HostRolesAndFeatures) {
                                    try {
                                        $TempHostRolesAndFeaturesReport = [PSCustomObject] @{
                                            'Feature Name' = $HostRoleAndFeature.DisplayName
                                            'Feature Type' = $HostRoleAndFeature.FeatureType
                                            'Description' = $HostRoleAndFeature.Description
                                        }
                                        $HostRolesAndFeaturesReport += $TempHostRolesAndFeaturesReport
                                    } catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $HostRolesAndFeaturesReport | Table -Name 'Roles and Features' -ColumnWidths 20, 10, 70
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    #Host 3rd Party Applications
                    try {
                        [array]$AddRemove = @()
                        $AddRemove += Invoke-Command -Session $TempPssSession { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* }
                        $AddRemove += Invoke-Command -Session $TempPssSession { Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* }
                        if ($AddRemove) {
                            Section -Style Heading3 'Installed Applications' {
                                Paragraph 'The following settings details applications listed in Add/Remove Programs'
                                #$AddRemove = Get-WmiObject -Class Win32_Product -ComputerName $System
                                [array]$AddRemoveReport = @()
                                ForEach ($App in $AddRemove) {
                                        try {
                                        $TempAddRemoveReport = [PSCustomObject]@{
                                            'Application Name' = $App.DisplayName
                                            'Publisher' = $App.Publisher
                                            'Version' = $App.Version
                                            'Install Date' = $App.InstallDate
                                        }
                                        $AddRemoveReport += $TempAddRemoveReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $AddRemoveReport | Where-Object { $_.'Application Name' -notlike $null } | Sort-Object  'Application Name' | Table -Name 'Installed Applications'
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
            #Local Users and Groups
            try {
                Section -Style Heading2 'Local Users and Groups' {
                    Paragraph 'The following section details local users and groups configured'
                    try {
                        $LocalUsers = Invoke-Command -Session $TempPssSession { Get-LocalUser }
                        if ($LocalUsers) {
                            Section -Style Heading3 'Local Users' {
                                Paragraph 'The following table details local users'
                                $LocalUsersReport = @()
                                ForEach ($LocalUser in $LocalUsers) {
                                    try {
                                        $TempLocalUsersReport = [PSCustomObject]@{
                                            'User Name' = $LocalUser.Name
                                            'Description' = $LocalUser.Description
                                            'Account Enabled' = $LocalUser.Enabled
                                            'Last Logon Date' = $LocalUser.LastLogon
                                        }
                                        $LocalUsersReport += $TempLocalUsersReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $LocalUsersReport | Table -Name 'Local Users' -ColumnWidths 20, 40, 10, 30
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    try {
                        $LocalGroups = Invoke-Command -Session $TempPssSession { Get-LocalGroup }
                        if ($LocalGroups) {
                            Section -Style Heading3 'Local Groups' {
                                Paragraph 'The following table details local groups configured'
                                $LocalGroupsReport = @()
                                ForEach ($LocalGroup in $LocalGroups) {
                                    try {
                                        $TempLocalGroupsReport = [PSCustomObject]@{
                                            'Group Name' = $LocalGroup.Name
                                            'Description' = $LocalGroup.Description
                                        }
                                        $LocalGroupsReport += $TempLocalGroupsReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $LocalGroupsReport | Table -Name 'Local Group Summary' -ColumnWidths 40, 60
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    try {
                        $LocalAdmins = Invoke-Command -Session $TempPssSession { Get-LocalGroupMember -Name 'Administrators' }
                        if ($LocalAdmins) {
                            Section -Style Heading3 'Local Administrators' {
                                Paragraph 'The following table lists Local Administrators'
                                $LocalAdminsReport = @()
                                ForEach ($LocalAdmin in $LocalAdmins) {
                                    try {
                                        $TempLocalAdminsReport = [PSCustomObject]@{
                                            'Account Name' = $LocalAdmin.Name
                                            'Account Type' = $LocalAdmin.ObjectClass
                                            'Account Source' = $LocalAdmin.PrincipalSource
                                        }
                                        $LocalAdminsReport += $TempLocalAdminsReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $LocalAdminsReport | Table -Name 'Local Administrators'
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
            #Host Firewall
            try {
                $NetFirewallProfile = Get-NetFirewallProfile -CimSession $TempCimSession
                if ($NetFirewallProfile) {
                    Section -Style Heading2 'Windows Firewall' {
                        Paragraph 'The Following table is a the Windowss Firewall Summary'
                        $NetFirewallProfileReport = @()
                        Foreach ($FirewallProfile in $NetFireWallProfile) {
                            try {
                                $TempNetFirewallProfileReport = [PSCustomObject]@{
                                    'Profile' = $FirewallProfile.Name
                                    'Profile Enabled' = $FirewallProfile.Enabled
                                    'Inbound Action' = $FirewallProfile.DefaultInboundAction
                                    'Outbound Action' = $FirewallProfile.DefaultOutboundAction
                                }
                                $NetFirewallProfileReport += $TempNetFirewallProfileReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $NetFirewallProfileReport | Table -Name 'Windows Firewall Profiles'
                    }
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
            #Host Networking
            try {
                Section -Style Heading2 'Host Networking' {
                    Paragraph 'The following section details Host Network Configuration'
                    try {
                        $HostAdapters = Invoke-Command -Session $TempPssSession { Get-NetAdapter }
                        if ($HostAdapters) {
                            Section -Style Heading3 'Network Adapters' {
                                Paragraph 'The Following table details host network adapters'
                                $HostAdaptersReport = @()
                                ForEach ($HostAdapter in $HostAdapters) {
                                    try {
                                        $TempHostAdaptersReport = [PSCustomObject]@{
                                            'Adapter Name' = $HostAdapter.Name
                                            'Adapter Description' = $HostAdapter.InterfaceDescription
                                            'Mac Address' = $HostAdapter.MacAddress
                                            'Link Speed' = $HostAdapter.LinkSpeed
                                        }
                                        $HostAdaptersReport += $TempHostAdaptersReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $HostAdaptersReport | Table -Name 'Network Adapters' -ColumnWidths 20, 40, 20, 20
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    try {
                        $NetIPs = Invoke-Command -Session $TempPssSession { Get-NetIPConfiguration | Where-Object -FilterScript { ($_.NetAdapter.Status -Eq "Up") } }
                        if ($NetIPs) {
                            Section -Style Heading3 'IP Addresses' {
                                Paragraph 'The following table details IP Addresses assigned to hosts'
                                $NetIpsReport = @()
                                ForEach ($NetIp in $NetIps) {
                                    try {
                                        $TempNetIpsReport = [PSCustomObject]@{
                                            'Interface Name' = $NetIp.InterfaceAlias
                                            'Interface Description' = $NetIp.InterfaceDescription
                                            'IPv4 Addresses' = $NetIp.IPv4Address -Join ","
                                            'Subnet Mask' = $NetIp.IPv4Address[0].PrefixLength
                                            'IPv4 Gateway' = $NetIp.IPv4DefaultGateway.NextHop
                                        }
                                        $NetIpsReport += $TempNetIpsReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $NetIpsReport | Table -Name 'Net IP Addresses'
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    try {
                        $DnsClient = Invoke-Command -Session $TempPssSession { Get-DnsClientGlobalSetting }
                        if ($DnsClient) {
                            Section -Style Heading3 'DNS Client' {
                                Paragraph 'The following table details the DNS Seach Domains'
                                $DnsClientReport = [PSCustomObject]@{
                                    'DNS Suffix' = $DnsClient.SuffixSearchList -Join ","
                                    'Use Suffix Search List' = ConvertTo-TextYN $DnsClient.UseSuffixSearchList
                                    'Use Devolution' = ConvertTo-TextYN $DnsClient.UseDevolution
                                    'Devolution Level' = $DnsClient.DevolutionLevel
                                }
                                $DnsClientReport | Table -Name "DNS Seach Domain" -ColumnWidths 40, 20, 20, 20
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    try {
                        $DnsServers = Invoke-Command -Session $TempPssSession { Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.ServerAddresses -notlike $null -and $_.InterfaceAlias -notlike "*isatap*" } }
                        if ($DnsServers) {
                            Section -Style Heading3 'DNS Servers' {
                                Paragraph 'The following table details the DNS Server Addresses Configured'
                                $DnsServerReport = @()
                                ForEach ($DnsServer in $DnsServers) {
                                    try {
                                        $TempDnsServerReport = [PSCustomObject]@{
                                            'Interface' = $DnsServer.InterfaceAlias
                                            'Server Address' = $DnsServer.ServerAddresses -Join ","
                                        }
                                        $DnsServerReport += $TempDnsServerReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $DnsServerReport | Table -Name 'DNS Server Addresses' -ColumnWidths 40, 60
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    try {
                        $NetworkTeamCheck = Invoke-Command -Session $TempPssSession { Get-NetLbfoTeam }
                        if ($NetworkTeamCheck) {
                            Section -Style Heading3 'Network Team Interfaces' {
                                Paragraph 'The following table details Network Team Interfaces'
                                $NetTeams = Invoke-Command -Session $TempPssSession { Get-NetLbfoTeam }
                                $NetTeamReport = @()
                                ForEach ($NetTeam in $NetTeams) {
                                    try {
                                        $TempNetTeamReport = [PSCustomObject]@{
                                            'Team Name' = $NetTeam.Name
                                            'Team Mode' = $NetTeam.tm
                                            'Load Balancing' = $NetTeam.lba
                                            'Network Adapters' = $NetTeam.Members -Join ","
                                        }
                                        $NetTeamReport += $TempNetTeamReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $NetTeamReport | Table -Name 'Network Team Interfaces'
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    try {
                        $NetMtus = Invoke-Command -Session $TempPssSession { Get-NetAdapterAdvancedProperty | Where-Object { $_.DisplayName -eq 'Jumbo Packet' } }
                        if ($NetMtus) {
                            Section -Style Heading3 'Network Adapter MTU' {
                                Paragraph 'The following table lists Network Adapter MTU settings'
                                $NetMtuReport = @()
                                ForEach ($NetMtu in $NetMtus) {
                                    try {
                                        $TempNetMtuReport = [PSCustomObject]@{
                                            'Adapter Name' = $NetMtu.Name
                                            'MTU Size' = $NetMtu.DisplayValue
                                        }
                                        $NetMtuReport += $TempNetMtuReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $NetMtuReport | Table -Name 'Network Adapter MTU' -ColumnWidths 50, 50
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
            #Host Storage
            try {
                Section -Style Heading2 'Host Storage' {
                    Paragraph 'The following section details the storage configuration of the host'
                    #Local Disks
                    try {
                        $HostDisks = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Disk }
                        if ($HostDisks) {
                            Section -Style Heading3 'Local Disks' {
                                Paragraph 'The following table details physical disks installed in the host'
                                $LocalDiskReport = @()
                                ForEach ($Disk in $HostDisks) {
                                    try {
                                        $TempLocalDiskReport = [PSCustomObject]@{
                                            'Disk Number' = $Disk.Number
                                            'Model' = $Disk.Model
                                            'Serial Number' = $Disk.SerialNumber
                                            'Partition Style' = $Disk.PartitionStyle
                                            'Disk Size(GB)' = [Math]::Round($Disk.Size / 1Gb)
                                        }
                                        $LocalDiskReport += $TempLocalDiskReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $LocalDiskReport | Sort-Object -Property 'Disk Number' | Table -Name 'Local Disks'
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    #Report any SAN Disks if they exist
                    try {
                        $SanDisks = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Disk | Where-Object { $_.BusType -Eq "iSCSI" } }
                        if ($SanDisks) {
                            Section -Style Heading3 'SAN Disks' {
                                Paragraph 'The following section details SAN disks connected to the host'
                                $SanDiskReport = @()
                                ForEach ($Disk in $SanDisks) {
                                    try {
                                        $TempSanDiskReport = [PSCustomObject]@{
                                            'Disk Number' = $Disk.Number
                                            'Model' = $Disk.Model
                                            'Serial Number' = $Disk.SerialNumber
                                            'Partition Style' = $Disk.PartitionStyle
                                            'Disk Size(GB)' = [Math]::Round($Disk.Size / 1Gb)
                                        }
                                        $SanDiskReport += $TempSanDiskReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $SanDiskReport | Sort-Object -Property 'Disk Number' | Table -Name 'Local Disks'
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    #Local Volumes
                    try {
                        $HostVolumes = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Volume }
                        if ($HostVolumes) {
                            Section -Style Heading3 'Host Volumes' {
                                Paragraph 'The following section details local volumes on the host'
                                $HostVolumeReport = @()
                                ForEach ($HostVolume in $HostVolumes) {
                                    try {
                                        $TempHostVolumeReport = [PSCustomObject]@{
                                            'Drive Letter' = $HostVolume.DriveLetter
                                            'File System Label' = $HostVolume.FileSystemLabel
                                            'File System' = $HostVolume.FileSystem
                                            'Size (GB)' = [Math]::Round($HostVolume.Size / 1gb)
                                            'Free Space(GB)' = [Math]::Round($HostVolume.SizeRemaining / 1gb)
                                        }
                                        $HostVolumeReport += $TempHostVolumeReport
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                }
                                $HostVolumeReport | Sort-Object 'Drive Letter' | Table -Name 'Host Volumes'
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                    #iSCSI
                    $iSCSICheck = Invoke-Command -Session $TempPssSession { Get-Service -Name 'MSiSCSI' }
                    try {
                        if ($iSCSICheck.Status -eq 'Running') {
                            Section -Style Heading3 'Host iSCSI Settings' {
                                Paragraph 'The following section details the iSCSI configuration for the host'
                                try {
                                    $HostInitiator = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-InitiatorPort }
                                    Paragraph 'The following table details the hosts iSCI IQN'
                                    $HostInitiator | Select-Object NodeAddress | Table -Name 'Host IQN'
                                    $HostIscsiTargetServer = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiTargetPortal }
                                    if($HostIscsiTargetServer){
                                        Section -Style Heading4 'iSCSI Target Server' {
                                            Paragraph 'The following table details iSCSI Target Server details'
                                            $HostIscsiTargetServer | Select-Object TargetPortalAddress, TargetPortalPortNumber | Table -Name 'iSCSI Target Servers' -ColumnWidths 50, 50
                                        }
                                    }
                                    $HostIscsiTargetVolumes = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiTarget }
                                }
                                catch {
                                    Write-PscriboMessage -IsWarning $_.Exception.Message
                                }
                                try {
                                    if($HostIscsiTargetVolumes){
                                        Section -Style Heading4 'iSCIS Target Volumes' {
                                            Paragraph 'The following table details iSCSI target volumes'
                                            $HostIscsiTargetVolumeReport = @()
                                            ForEach ($HostIscsiTargetVolume in $HostIscsiTargetVolumes) {
                                                try {
                                                    $TempHostIscsiTargetVolumeReport = [PSCustomObject]@{
                                                        'Node Address' = $HostIscsiTargetVolume.NodeAddress
                                                        'Node Connected' = $HostIscsiTargetVolume.IsConnected
                                                    }
                                                    $HostIscsiTargetVolumeReport += $TempHostIscsiTargetVolumeReport
                                                }
                                                catch {
                                                    Write-PscriboMessage -IsWarning $_.Exception.Message
                                                }
                                            }
                                            $HostIscsiTargetVolumeReport | Table -Name 'iSCSI Target Volumes' -ColumnWidths 80, 20
                                        }
                                    }
                                }
                                catch {
                                    Write-PscriboMessage -IsWarning $_.Exception.Message
                                }
                                try {
                                    $HostIscsiConnections = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiConnection }
                                    if($HostIscsiConnections){
                                        Section -Style Heading4 'iSCSI Connections' {
                                            Paragraph 'The following table details iSCSI Connections'
                                            $HostIscsiConnections | Select-Object ConnectionIdentifier, InitiatorAddress, TargetAddress | Table -Name 'iSCSI Connections'
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
                    #MPIO
                    try {
                        $MPIOInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "Multipath*" } }
                        if ($MPIOInstalledCheck.InstallState -eq "Installed") {
                            try {
                                Section -Style Heading3 'Host MPIO Settings' {
                                    Paragraph 'The following section details host MPIO Settings'
                                    [string]$MpioLoadBalance = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-MSDSMGlobalDefaultLoadBalancePolicy }
                                    Paragraph "The default load balancing policy is: $MpioLoadBalance"
                                    $MpioAutoClaim = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-MSDSMAutomaticClaimSettings | Select-Object -ExpandProperty Keys }
                                    if ($MpioAutoClaim) {
                                        Section -Style Heading4 'Multipath  I/O AutoClaim' {
                                            Paragraph 'The Following table details the BUS types MPIO will automatically claim for'
                                            $MpioAutoClaimReport = @()
                                            foreach ($key in $MpioAutoClaim) {
                                                try {
                                                    $Temp = "" | Select-Object BusType, State
                                                    $Temp.BusType = $key
                                                    $Temp.State = 'Enabled'
                                                    $MpioAutoClaimReport += $Temp
                                                }
                                                catch {
                                                    Write-PscriboMessage -IsWarning $_.Exception.Message
                                                }
                                            }
                                            $MpioAutoClaimReport | Table -Name 'Multipath I/O Auto Claim Settings'
                                        }
                                    }
                                    try {
                                        $MpioAvailableHw = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-MPIOAvailableHw }
                                        if ($MpioAvailableHw) {
                                            Section -Style Heading4 'MPIO Detected Hardware' {
                                                Paragraph 'The following table details the hardware detected and claimed by MPIO'
                                                $MpioAvailableHw | Select-Object VendorId, ProductId, BusType, IsMultipathed | Table -Name 'MPIO Available Hardware'
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
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
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
                        $VmHost = Invoke-Command -Session $TempPssSession { Get-VMHost }
                        if ($VmHost) {
                            $VmHostReport = [PSCustomObject]@{
                                'Logical Processor Count' = $VmHost.LogicalProcessorCount
                                'Memory Capacity (GB)' = [Math]::Round($VmHost.MemoryCapacity / 1gb)
                                'VM Default Path' = $VmHost.VirtualMachinePath
                                'VM Disk Default Path' = $VmHost.VirtualHardDiskPath
                                'Supported VM Versions' = $VmHost.SupportedVmVersions -Join ","
                                'Numa Spannning Enabled' = $VmHost.NumaSpanningEnabled
                                'Iov Support' = $VmHost.IovSupport
                                'VM Migrations Enabled' = $VmHost.VirtualMachineMigrationEnabled
                                'Allow any network for Migrations' = $VmHost.UseAnyNetworkForMigrations
                                'VM Migration Authentication Type' = $VmHost.VirtualMachineMigrationAuthenticationType
                                'Max Concurrent Storage Migrations' = $VmHost.MaximumStorageMigrations
                                'Max Concurrent VM Migrations' = $VmHost.MaximumStorageMigrations
                            }
                            $VmHostReport | Table -Name 'Hyper-V Host Settings' -List -ColumnWidths 50, 50
                        }
                        try {
                            $VmHostNumaNodes = Get-VMHostNumaNode -CimSession $TempCimSession
                            if ($VmHostNumaNodes) {
                                Section -Style Heading3 "Hyper-V NUMA Boundaries" {
                                    Paragraph 'The following table details the NUMA nodes on the host'
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
                                    $VmHostNumaReport | Table -Name 'Host NUMA Nodes'
                                }
                            }
                        }
                        catch {
                            Write-PscriboMessage -IsWarning $_.Exception.Message
                        }
                        try {
                            Section -Style Heading3 "Hyper-V MAC Pool settings" {
                                'The following table details the Hyper-V MAC Pool'
                                $VmHostMacPool = [PSCustomObject]@{
                                    'Mac Address Minimum' = $VmHost.MacAddressMinimum
                                    'Mac Address Maximum' = $VmHost.MacAddressMaximum
                                }
                                $VmHostMacPool | Table -Name 'MAC Address Pool' -ColumnWidths 50, 50
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
                                    $VmOsAdapterReport = @()
                                    Foreach ($VmOsAdapter in $VmOsAdapters) {
                                        try {
                                            $AdapterVlan = Get-VMNetworkAdapterVlan -CimSession $TempCimSession -ManagementOS -VMNetworkAdapterName $VmOsAdapter.Name
                                            $VmOsAdapterReport += $TempVmOsAdapaterReport
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
                                    if ($NULL -ne $VmOsAdapterReport) {
                                        $VmOsAdapterReport | Table -Name 'VM Management OS Adapters' -List -ColumnWidths 50, 50
                                    }
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
                                $VmSwitchesReport | Table -Name 'Virtual Switch Summary' -ColumnWidths 40, 10, 10, 40
                                Foreach ($VmSwitch in $VmSwitches) {
                                    try {
                                        Section -Style Heading4 ($VmSwitch.Name) {
                                            Paragraph 'The following table details the Hyper-V vSwitch'
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
                                            $VmSwitchReport | Table -Name 'VM Switch Details' -List -ColumnWidths 50, 50
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
                        Section -Style Heading2 'Hyper-V VMs' {
                            Paragraph 'The following section details the Hyper-V VMs running on this host'
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
                            $VmSummary | Sort-Object 'VM Name' | Table -Name 'Virtual Machines'
                            foreach ($Vm in $Vms) {
                                try {
                                    Section -Style Heading3 ($Vm.Name) {
                                        Paragraph 'The following sections detail the VM configuration settings'
                                        try {
                                            Section -Style Heading4 'Virtual Machine Configuration' {
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
                                                $VmConfiguration | Table -List -ColumnWidths 40, 60
                                            }
                                        }
                                        catch {
                                            Write-PscriboMessage -IsWarning $_.Exception.Message
                                        }
                                        try {
                                            Section -Style Heading4 'Virtual Machine Guest Integration Service' {
                                                Paragraph 'The following section details the status of Integration Services'
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
                                                $VmIntegrationServiceSummary | Table -Name 'Integration Service' -ColumnWidths 40, 30, 30
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
                                                    $VmNetworkAdapterReport | Table -Name 'VM Network Adapters'
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
                                                    $VmAdapterVlanReport | Table -Name 'VM Network Adapter Vlans'
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
                                                    $VmDiskReport | Table 'VM Hard disks' -ColumnWidths 30, 10, 10, 10, 10, 10, 10, 10
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
