function Invoke-AsBuiltReport.Microsoft.Windows {
    <#
    .SYNOPSIS
        PowerShell script to document the configuration of Microsoft Windows Server in Word/HTML/Text formats
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.4.0
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

    Write-PScriboMessage -IsWarning "Please refer to the AsBuiltReport.Microsoft.Windows github website for more detailed information about this project."
    Write-PScriboMessage -IsWarning "Do not forget to update your report configuration file after each new release."
    Write-PScriboMessage -IsWarning "Documentation: https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows"
    Write-PScriboMessage -IsWarning "Issues or bug reporting: https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows/issues"

    Try {
        $InstalledVersion = Get-Module -ListAvailable -Name AsBuiltReport.Microsoft.Windows -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending | Select-Object -First 1 -ExpandProperty Version

        if ($InstalledVersion) {
            Write-PScriboMessage -IsWarning "AsBuiltReport.Microsoft.Windows $($InstalledVersion.ToString()) is currently installed."
            $LatestVersion = Find-Module -Name AsBuiltReport.Microsoft.Windows -Repository PSGallery -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version
            if ($LatestVersion -gt $InstalledVersion) {
                Write-PScriboMessage -IsWarning "AsBuiltReport.Microsoft.Windows $($LatestVersion.ToString()) is available."
                Write-PScriboMessage -IsWarning "Run 'Update-Module -Name AsBuiltReport.Microsoft.Windows -Force' to install the latest version."
            }
        }
    } Catch {
            Write-PscriboMessage -IsWarning $_.Exception.Message
        }

        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())


    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

        throw "The requested operation requires elevation: Run PowerShell console as administrator"
    }

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

            #Validate Required Modules and Features
            $script:OSType = Invoke-Command -Session $TempPssSession { (Get-ComputerInfo).OsProductType }
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
                $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'vmms' -ErrorAction SilentlyContinue }
                if ($Status) {
                    try {
                        if (Get-RequiredFeature -Name Hyper-V-PowerShell -OSType $OSType.Value -Status) {
                            Section -Style Heading2 "Hyper-V Configuration" {
                                Paragraph 'The following table details the Hyper-V server settings'
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
                        } else {
                            Get-RequiredFeature -Name Hyper-V-PowerShell -OSType $OSType.Value -Service "Hyper-V"
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                } else {
                    Write-PScriboMessage "No HyperV service detected. Disabling HyperV server section"
                }
            }
            if ($InfoLevel.IIS -ge 1) {
                $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'W3SVC' -ErrorAction SilentlyContinue }
                if ($Status) {
                    try {
                        if (((Get-RequiredFeature -Name web-mgmt-console -OSType $OSType.Value -Status) -and (Get-RequiredFeature -Name Web-Scripting-Tools -OSType $OSType.Value -Status)) -or ((Get-RequiredFeature -Name IIS-WebServerRole -OSType $OSType.Value -Status) -and (Get-RequiredFeature -Name WebServerManagementTools -OSType $OSType.Value -Status) -and (Get-RequiredFeature -Name IIS-ManagementScriptingTools -OSType $OSType.Value -Status))) {
                            Section -Style Heading2 "IIS Configuration" {
                                Paragraph 'The following table details the IIS server settings'
                                Blankline
                                # IIS Configuration
                                Get-AbrWinIISSummary
                                # IIS Web Application Pools
                                Get-AbrWinIISWebAppPool
                                # IIS Web Site
                                Get-AbrWinIISWebSite
                            }
                        } else  {
                            If ($OSType -eq 'Server' -or $OSType -eq 'DomainController') {
                                Get-RequiredFeature -Name web-mgmt-console -OSType $OSType.Value -Service "IIS"
                                Get-RequiredFeature -Name Web-Scripting-Tools -OSType $OSType.Value -Service "IIS"
                            } else {
                                Get-RequiredFeature -Name IIS-WebServerRole -OSType $OSType.Value -Service "IIS"
                                Get-RequiredFeature -Name WebServerManagementTools -OSType $OSType.Value -Service "IIS"
                                Get-RequiredFeature -Name IIS-ManagementScriptingTools -OSType $OSType.Value -Service "IIS"
                            }
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                } else {
                    Write-PScriboMessage "No W3SVC service detected. Disabling IIS server section"
                }
            }
            if ($InfoLevel.SMB -ge 1) {
                try {
                    $Global:SMBShares = Invoke-Command -Session $TempPssSession { Get-SmbShare | Where-Object {$_.Special -like 'False'} }
                    if ($SMBShares) {
                        Section -Style Heading2 "File Server Configuration" {
                            Paragraph 'The following table details the File Server settings'
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
                catch {
                    Write-PscriboMessage -IsWarning $_.Exception.Message
                }
            }
            if ($InfoLevel.DHCP -ge 1 -and $OSType.Value -ne 'WorkStation') {
                $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'DHCPServer' -ErrorAction SilentlyContinue }
                if ($Status) {
                    try {
                        if (Get-RequiredFeature -Name RSAT-DHCP -OSType $OSType.Value -Status) {
                            Section -Style Heading2 "DHCP Server Configuration" {
                                Paragraph 'The following table details the DHCP server configurations'
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
                        } else {
                            Get-RequiredFeature -Name RSAT-DHCP -OSType $OSType.Value -Service "DHCP Server"
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                } else {
                    Write-PScriboMessage "No DHCPServer service detected. Disabling Dhcp server section"
                }
            }
            if ($InfoLevel.DNS -ge 1 -and $OSType.Value -ne 'WorkStation') {
                $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'DNS' -ErrorAction SilentlyContinue }
                if ($Status) {
                    try {
                        if (Get-RequiredFeature -Name RSAT-DNS-Server -OSType $OSType.Value -Status) {
                            Section -Style Heading2 "DNS Server Configuration" {
                                Paragraph 'The following table details the DNS server settings'
                                Blankline
                                # DNS Server Configuration
                                Get-AbrWinDNSInfrastructure
                                # DNS Zones Configuration
                                Get-AbrWinDNSZone
                            }
                        } else {
                            Get-RequiredFeature -Name RSAT-DNS-Server -OSType $OSType.Value -Service "DNS Server"
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                } else {
                    Write-PScriboMessage "No DNS Server service detected. Disabling DNS server section"
                }
            }

            if ($InfoLevel.FailOverCluster -ge 1 -and $OSType.Value -ne 'WorkStation') {
                $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service 'ClusSvc' -ErrorAction SilentlyContinue }
                if ($Status.Status -eq "Running") {
                    try {
                        if (Get-RequiredFeature -Name RSAT-Clustering-PowerShell -OSType $OSType.Value -Status) {
                            Section -Style Heading2 "Failover Cluster Configuration" {
                                Paragraph 'The following table details the Failover Cluster Settings'
                                Blankline
                                # # DNS Server Configuration
                                # Get-AbrWinDNSInfrastructure
                                # # DNS Zones Configuration
                                # Get-AbrWinDNSZone
                            }
                        }
                        else {
                            Get-RequiredFeature -Name RSAT-Clustering-PowerShell -OSType $OSType.Value -Service "FailOver Cluster"
                        }
                    }
                    catch {
                        Write-PscriboMessage -IsWarning $_.Exception.Message
                    }
                } else {
                    Write-PScriboMessage "No FailOver Cluster service detected. Disabling FailOver Cluster section"
                }
            }

        }
        Remove-PSSession $TempPssSession
        Remove-CimSession $TempCimSession
    }
    #endregion foreach loop
}
