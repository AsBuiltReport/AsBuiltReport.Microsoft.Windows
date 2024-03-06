function Get-AbrWinSQLBackupDevice {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows SQL Server backup device information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.3
        Author:         Jonathan Colon
        Twitter:        @rebelinux
        Github:         AsBuiltReport
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows
    #>
    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "SQL Server Backup Device InfoLevel set at $($InfoLevel.SQLServer)."
    }

    process {
        try {
            Write-PScriboMessage "Collecting SQL Server Backup Device information."
            $SQLBackUpDevices = Get-DbaBackupDevice -SqlInstance $SQLServer | Sort-Object -Property Name
            if ($SQLBackUpDevices) {
                Write-PScriboMessage "Collecting SQL Server Backup Device information."
                Section -Style Heading4 'Backup Device' {
                    $ItemInfo = @()
                    foreach ($Item in $SQLBackUpDevices) {
                        try {
                            $InObj = [Ordered]@{
                                'Name' = $Item.Name
                                'Backup Device Type' = $Item.BackupDeviceType
                                'Physical Location' = $Item.PhysicalLocation
                                'Skip Tape Label' = ConvertTo-TextYN $Item.SkipTapeLabel
                            }
                            $ItemInfo += [PSCustomObject]$InObj
                        } catch {
                            Write-PScriboMessage -IsWarning "SQL Server System Backup Device Section: $($_.Exception.Message)"
                        }
                    }

                    if ($InfoLevel.SQLServer -ge 2) {
                        Paragraph "The following sections detail the configuration of the backup device."
                        foreach ($Item in $ItemInfo) {
                            Section -Style NOTOCHeading5 -ExcludeFromTOC "$($Item.Name)" {
                                $TableParams = @{
                                    Name = "Backup Device - $($Item.Name)"
                                    List = $true
                                    ColumnWidths = 50, 50
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $Item | Table @TableParams
                            }
                        }
                    } else {
                        Paragraph "The following table summarises the configuration of the backup device."
                        BlankLine
                        $TableParams = @{
                            Name = "Backup Devices"
                            List = $false
                            Columns = 'Name', 'Backup Device Type', 'Physical Location'
                            ColumnWidths = 25, 25, 50
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $ItemInfo | Table @TableParams
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning "SQL Server Backup Device Section: $($_.Exception.Message)"
        }
    }
    end {}
}