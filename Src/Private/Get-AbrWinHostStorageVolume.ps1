function Get-AbrWinHostStorageVolume {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Host Storage Volume information.
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
        Write-PScriboMessage "Storage InfoLevel set at $($InfoLevel.Storage)."
        Write-PscriboMessage "Collecting Host Storage Volume information."
    }

    process {
        if ($InfoLevel.Storage -ge 1) {
            try {
                $HostVolumes = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Volume | Where-Object {$_.DriveType -ne "CD-ROM"}}
                if ($HostVolumes) {
                    Section -Style Heading3 'Host Volumes' {
                        Paragraph 'The following section details local volumes on the host'
                        Blankline
                        $HostVolumeReport = @()
                        ForEach ($HostVolume in $HostVolumes) {
                            try {
                                $TempHostVolumeReport = [PSCustomObject]@{
                                    'Drive Letter' = $HostVolume.DriveLetter
                                    'File System Label' = $HostVolume.FileSystemLabel
                                    'File System' = $HostVolume.FileSystem
                                    'Size (GB)' = [Math]::Round($HostVolume.Size / 1gb)
                                    'Free Space(GB)' = [Math]::Round($HostVolume.SizeRemaining / 1gb)
                                    'Health Status' = $HostVolume.HealthStatus
                                }
                                $HostVolumeReport += $TempHostVolumeReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Host Volumes"
                            List = $false
                            ColumnWidths = 15, 15, 15, 20, 20, 15
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $HostVolumeReport | Sort-Object -Property 'Drive Letter' | Table @TableParams
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