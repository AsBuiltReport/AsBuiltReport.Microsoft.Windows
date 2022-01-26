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
            #Host Storage
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
            #HyperV Configuration
            try {
                $HyperVInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*Hyper-V*" } }
                if ($HyperVInstalledCheck.InstallState -eq "Installed") {
                    Section -Style Heading2 "Hyper-V Configuration Settings" {
                        Paragraph 'The following table details the Hyper-V Server Settings'
                        Blankline
                        # Hyper-V Configuration
                        Get-AbrWinHyperVSummary
                        # Hyper-V Numa Information
                        Get-AbrWinHyperVNuma
                        # Hyper-V Networking
                        Get-AbrWinHyperVNetworking
                        # Hyper-V VM Information
                        Get-AbrWinHyperVHostVM
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
