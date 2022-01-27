function Get-AbrWinHostStorageISCSI {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Host Storage ISCSI information.
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
    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "Storage InfoLevel set at $($InfoLevel.Storage)."
        Write-PscriboMessage "Collecting Host Storage ISCSI information."
    }

    process {
        if ($InfoLevel.Storage -ge 1) {
            $iSCSICheck = Invoke-Command -Session $TempPssSession { Get-Service -Name 'MSiSCSI' }
            try {
                if ($iSCSICheck.Status -eq 'Running') {
                    Section -Style Heading3 'Host iSCSI Settings' {
                        Paragraph 'The following section details the iSCSI configuration for the host'
                        Blankline
                        try {
                            $HostInitiator = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-InitiatorPort }
                            if ($HostInitiator) {
                                Section -Style Heading4 'iSCSI Target Server' {
                                    Paragraph 'The following table details the hosts iSCI IQN'
                                    Blankline
                                    $HostInitiatorReport = @()
                                    try {
                                        $TempHostInitiator = [PSCustomObject]@{
                                            'Node Address' = $HostInitiator.NodeAddress
                                            'Operational Status' = Switch ($HostInitiator.OperationalStatus) {
                                                1 {'Unknown'}
                                                2 {'Operational'}
                                                3 {'User Offline'}
                                                4 {'Bypassed'}
                                                5 {'In diagnostics mode'}
                                                6 {'Link Down'}
                                                7 {'Port Error'}
                                                8 {'Loopback'}
                                                default {$HostInitiator.OperationalStatus}
                                            }
                                        }
                                    }
                                    catch {
                                        Write-PscriboMessage -IsWarning $_.Exception.Message
                                    }
                                    $HostInitiatorReport += $TempHostInitiator

                                    $TableParams = @{
                                        Name = "Host IQN"
                                        List = $false
                                        ColumnWidths = 60, 40
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $HostInitiatorReport | Table @TableParams
                                }
                            }

                            $HostIscsiTargetServers = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiTargetPortal }
                            if($HostIscsiTargetServers){
                                Section -Style Heading4 'iSCSI Target Server' {
                                    Paragraph 'The following table details iSCSI Target Server details'
                                    Blankline
                                    $HostIscsiTargetServerReport = @()
                                    ForEach ($HostIscsiTargetServer in $HostIscsiTargetServers) {
                                        try {
                                            $TempHostIscsiTargetServerReport = [PSCustomObject]@{
                                                'Target Portal Address' = $HostIscsiTargetServer.TargetPortalAddress
                                                'Target Portal Port Number' = $HostIscsiTargetServer.TargetPortalPortNumber
                                            }
                                            $HostIscsiTargetServerReport += $TempHostIscsiTargetServerReport
                                        }
                                        catch {
                                            Write-PscriboMessage -IsWarning $_.Exception.Message
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "iSCSI Target Servers"
                                        List = $false
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $HostIscsiTargetServerReport | Sort-Object -Property 'Target Portal Address' | Table @TableParams
                                }
                            }
                        }
                        catch {
                            Write-PscriboMessage -IsWarning $_.Exception.Message
                        }
                        try {
                            $HostIscsiTargetVolumes = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiTarget }
                            if($HostIscsiTargetVolumes){
                                Section -Style Heading4 'iSCIS Target Volumes' {
                                    Paragraph 'The following table details iSCSI target volumes'
                                    Blankline
                                    $HostIscsiTargetVolumeReport = @()
                                    ForEach ($HostIscsiTargetVolume in $HostIscsiTargetVolumes) {
                                        try {
                                            $TempHostIscsiTargetVolumeReport = [PSCustomObject]@{
                                                'Node Address' = $HostIscsiTargetVolume.NodeAddress
                                                'Node Connected' = ConvertTo-TextYN $HostIscsiTargetVolume.IsConnected
                                            }
                                            $HostIscsiTargetVolumeReport += $TempHostIscsiTargetVolumeReport
                                        }
                                        catch {
                                            Write-PscriboMessage -IsWarning $_.Exception.Message
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "iSCIS Target Volumes"
                                        List = $false
                                        ColumnWidths = 80, 20
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $HostIscsiTargetVolumeReport | Sort-Object -Property 'Node Address' | Table @TableParams
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
                                    Blankline
                                    $HostIscsiConnectionsReport = @()
                                    ForEach ($HostIscsiConnection in $HostIscsiConnections) {
                                        try {
                                            $TempHostIscsiConnectionsReport = [PSCustomObject]@{
                                                'Connection Identifier' = $HostIscsiConnection.ConnectionIdentifier
                                                'Initiator Address' = $HostIscsiConnection.InitiatorAddress
                                                'Target Address' = $HostIscsiConnection.TargetAddress
                                            }
                                            $HostIscsiConnectionsReport += $TempHostIscsiConnectionsReport
                                        }
                                        catch {
                                            Write-PscriboMessage -IsWarning $_.Exception.Message
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "iSCSI Connections"
                                        List = $false
                                        ColumnWidths = 34, 33, 33
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $HostIscsiConnectionsReport | Sort-Object -Property 'Connection Identifier' | Table @TableParams
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
    end {}
}