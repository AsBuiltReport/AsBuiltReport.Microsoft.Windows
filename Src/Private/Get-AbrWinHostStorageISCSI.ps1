function Get-AbrWinHostStorageISCSI {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Host Storage ISCSI information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.2
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
        Write-PScriboMessage "Collecting Host Storage ISCSI information."
    }

    process {
        if ($InfoLevel.Storage -ge 1) {
            $iSCSICheck = Invoke-Command -Session $TempPssSession { Get-Service -Name 'MSiSCSI' }
            try {
                if ($iSCSICheck.Status -eq 'Running') {
                    Section -Style Heading3 'Host iSCSI Settings' {
                        Paragraph 'The following section details the iSCSI configuration for the host'
                        BlankLine
                        try {
                            $HostInitiator = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-InitiatorPort }
                            if ($HostInitiator) {
                                Section -Style Heading4 'iSCSI Host Initiator' {
                                    Paragraph 'The following table details the hosts iSCSI IQN'
                                    BlankLine
                                    $OutObj = @()
                                    try {
                                        $inObj = [ordered] @{
                                            'Node Address' = $HostInitiator.NodeAddress
                                            'Operational Status' = Switch ($HostInitiator.OperationalStatus) {
                                                1 { 'Unknown' }
                                                2 { 'Operational' }
                                                3 { 'User Offline' }
                                                4 { 'Bypassed' }
                                                5 { 'In diagnostics mode' }
                                                6 { 'Link Down' }
                                                7 { 'Port Error' }
                                                8 { 'Loopback' }
                                                default { $HostInitiator.OperationalStatus }
                                            }
                                        }
                                    } catch {
                                        Write-PScriboMessage -IsWarning $_.Exception.Message
                                    }
                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)

                                    if ($HealthCheck.Storage.BP) {
                                        $OutObj | Where-Object { $_.'Operational Status' -ne 'Operational' } | Set-Style -Style Warning -Property 'Operational Status'
                                    }

                                    $TableParams = @{
                                        Name = "iSCSI Host Initiator"
                                        List = $false
                                        ColumnWidths = 60, 40
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $OutObj | Table @TableParams
                                }
                            }

                            $HostIscsiTargetServers = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiTargetPortal }
                            if ($HostIscsiTargetServers) {
                                Section -Style Heading4 'iSCSI Target Server' {
                                    Paragraph 'The following table details iSCSI Target Server details'
                                    BlankLine
                                    $OutObj = @()
                                    ForEach ($HostIscsiTargetServer in $HostIscsiTargetServers) {
                                        try {
                                            $inObj = [ordered] @{
                                                'Target Portal Address' = $HostIscsiTargetServer.TargetPortalAddress
                                                'Target Portal Port Number' = $HostIscsiTargetServer.TargetPortalPortNumber
                                            }
                                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                        } catch {
                                            Write-PScriboMessage -IsWarning $_.Exception.Message
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
                                    $OutObj | Sort-Object -Property 'Target Portal Address' | Table @TableParams
                                }
                            }
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                        try {
                            $HostIscsiTargetVolumes = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiTarget }
                            if ($HostIscsiTargetVolumes) {
                                Section -Style Heading4 'iSCIS Target Volumes' {
                                    Paragraph 'The following table details iSCSI target volumes'
                                    BlankLine
                                    $OutObj = @()
                                    ForEach ($HostIscsiTargetVolume in $HostIscsiTargetVolumes) {
                                        try {
                                            $inObj = [ordered] @{
                                                'Node Address' = $HostIscsiTargetVolume.NodeAddress
                                                'Node Connected' = $HostIscsiTargetVolume.IsConnected
                                            }
                                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                        } catch {
                                            Write-PScriboMessage -IsWarning $_.Exception.Message
                                        }
                                    }

                                    if ($HealthCheck.Storage.BP) {
                                        $OutObj | Where-Object { $_.'Node Connected' -ne 'Yes' } | Set-Style -Style Warning -Property 'Node Connected'
                                    }

                                    $TableParams = @{
                                        Name = "iSCIS Target Volumes"
                                        List = $false
                                        ColumnWidths = 80, 20
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $OutObj | Sort-Object -Property 'Node Address' | Table @TableParams
                                }
                            }
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                        try {
                            $HostIscsiConnections = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-IscsiConnection }
                            if ($HostIscsiConnections) {
                                Section -Style Heading4 'iSCSI Connections' {
                                    Paragraph 'The following table details iSCSI Connections'
                                    BlankLine
                                    $OutObj = @()
                                    ForEach ($HostIscsiConnection in $HostIscsiConnections) {
                                        try {
                                            $inObj = [ordered] @{
                                                'Connection Identifier' = $HostIscsiConnection.ConnectionIdentifier
                                                'Initiator Address' = $HostIscsiConnection.InitiatorAddress
                                                'Target Address' = $HostIscsiConnection.TargetAddress
                                            }
                                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                        } catch {
                                            Write-PScriboMessage -IsWarning $_.Exception.Message
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
                                    $OutObj | Sort-Object -Property 'Connection Identifier' | Table @TableParams
                                }
                            }
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}