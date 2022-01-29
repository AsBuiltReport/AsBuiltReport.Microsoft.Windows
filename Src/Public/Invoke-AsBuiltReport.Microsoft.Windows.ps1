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
        Twitter:        @asbuiltreport
        Github:         AsBuiltReport
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

    #region foreach loop
    foreach ($System in $Target) {
        Section -Style Heading1 $System {
            Paragraph "The following table details the Windows Host $System"
            BlankLine
            try {
                $script:TempPssSession = New-PSSession $System -Credential $Credential -Authentication Negotiate -ErrorAction stop
                $script:TempCimSession = New-CimSession $System -Credential $Credential -Authentication Negotiate -ErrorAction stop
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
            if ($InfoLevel.OperatingSystem -ge 1) {
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
                        # Host Service Status
                        Get-AbrWinOSService
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            #Local Users and Groups
            if ($InfoLevel.Account -ge 1) {
                try {
                    $LocalUsers = Invoke-Command -Session $TempPssSession { Get-LocalUser | Where-Object {$_.PrincipalSource -ne "ActiveDirectory"} }
                    $LocalGroups = Invoke-Command -Session $TempPssSession { Get-LocalGroup | Where-Object {$_.PrincipalSource -ne "ActiveDirectory" }}
                    $LocalAdmins = Invoke-Command -Session $TempPssSession { Get-LocalGroupMember -Name 'Administrators' -ErrorAction SilentlyContinue }
                    if ($LocalUsers -or $LocalGroups -or $LocalAdmins) {
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
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            #Host Firewall
            Get-AbrWinNetFirewall
            #Host Networking
            if ($InfoLevel.Networking -ge 1) {
                try {
                    Section -Style Heading2 'Host Networking' {
                        Paragraph 'The following section details Host Network Configuration'
                        Blankline
                        #Host Network Adapter
                        Get-AbrWinNetAdapter
                        #Host Network IP Address
                        Get-AbrWinNetIPAddress
                        #Host DNS Client Setting
                        Get-AbrWinNetDNSClient
                        #Host DNS Server Setting
                        Get-AbrWinNetDNSServer
                        #Host Network Teaming
                        Get-AbrWinNetTeamInterface
                        #Host Network Adapter MTU
                        Get-AbrWinNetAdapterMTU
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            #Host Storage
            if ($InfoLevel.Storage -ge 1) {
                try {
                    Section -Style Heading2 'Host Storage' {
                        Paragraph 'The following section details the storage configuration of the host'
                        #Local Disks
                        Get-AbrWinHostStorage
                        #Local Volumes
                        Get-AbrWinHostStorageVolume
                        #iSCSI Settings
                        Get-AbrWinHostStorageISCSI
                        #MPIO Setting
                        Get-AbrWinHostStorageMPIO
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            #HyperV Configuration
            if ($InfoLevel.HyperV -ge 1) {
                try {
                    $HyperVInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*Hyper-V*" } }
                    if ($HyperVInstalledCheck.InstallState -eq "Installed") {
                        $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'vmms' -ErrorAction SilentlyContinue }
                        if ($Status.Status -eq "Running") {
                            Section -Style Heading2 "Hyper-V Configuration Settings" {
                                Paragraph 'The following table details the Hyper-V Server Settings'
                                Blankline
                                # Hyper-V Configuration
                                Get-AbrWinHyperVSummary
                                # Hyper-V Numa Information
                                Get-AbrWinHyperVNuma
                                # Hyper-V Networking
                                Get-AbrWinHyperVNetworking
                                # Hyper-V VM Information (Buggy as hell)
                                #Get-AbrWinHyperVHostVM
                            }
                        }
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            if ($InfoLevel.IIS -ge 1) {
                try {
                    $IISInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*Web-Server*" } }
                    if ($IISInstalledCheck.InstallState -eq "Installed") {
                        $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'W3SVC' -ErrorAction SilentlyContinue }
                        if ($Status.Status -eq "Running") {
                            Section -Style Heading2 "IIS Configuration Settings" {
                                Paragraph 'The following table details the IIS Server Settings'
                                Blankline
                                # IIS Configuration
                                Get-AbrWinIISSummary
                                # IIS Web Application Pools
                                Get-AbrWinIISWebAppPool
                                # IIS Web Site
                                Get-AbrWinIISWebSite

                            }
                        }
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            if ($InfoLevel.SMB -ge 1) {
                try {
                    $SMBInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*FileAndStorage-Services*" } }
                    if ($SMBInstalledCheck.InstallState -eq "Installed") {
                        $Global:SMBShares = Invoke-Command -Session $TempPssSession { Get-SmbShare | Where-Object {$_.Special -like 'False'} }
                        if ($SMBShares) {
                            Section -Style Heading2 "File Server Configuration Settings" {
                                Paragraph 'The following table details the File Server Settings'
                                Blankline
                                # SMB Server Configuration
                                Get-AbrWinSMBSummary
                                # SMB Server Network Interface
                                Get-AbrWinSMBNetworkInterface
                                # SMB Shares
                                Get-AbrWinSMBShare


                            }
                        }
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            if ($InfoLevel.DHCP -ge 1) {
                try {
                    $DHCPInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*DHCP*" } }
                    if ($DHCPInstalledCheck.InstallState -eq "Installed") {
                        $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'DHCPServer' -ErrorAction SilentlyContinue }
                        if ($Status.Status -eq "Running") {
                            Section -Style Heading2 "DHCP Server Configuration Settings" {
                                Paragraph 'The following table details the DHCP Server Settings'
                                Blankline
                                # DHCP Server Configuration
                                Get-AbrWinDHCPInfrastructure
                                # DHCP Server Stats
                                Get-AbrWinDHCPv4Statistic
                                # DHCP Server Scope Info
                                Get-AbrWinDHCPv4Scope
                                # DHCP Server Scope Settings
                                Get-AbrWinDHCPv4ScopeServerSetting
                                # DHCP Server Per Scope Info
                                Get-AbrWinDHCPv4PerScopeSetting
                            }
                        }
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            if ($InfoLevel.DNS -ge 1) {
                try {
                    $DHCPInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*DNS*" } }
                    if ($DHCPInstalledCheck.InstallState -eq "Installed") {
                        $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'DNS' -ErrorAction SilentlyContinue }
                        if ($Status.Status -eq "Running") {
                            Section -Style Heading2 "DNS Server Configuration Settings" {
                                Paragraph 'The following table details the DNS Server Settings'
                                Blankline
                                # DNS Server Configuration
                                Get-AbrWinDNSInfrastructure
                                # DNS Zones Configuration
                                Get-AbrWinDNSZone
                            }
                        }
                    }
                }
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
        }
        Remove-PSSession $TempPssSession
        Remove-CimSession $TempCimSession
    }
    #endregion foreach loop
}
