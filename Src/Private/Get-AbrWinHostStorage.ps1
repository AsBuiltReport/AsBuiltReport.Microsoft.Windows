function Get-AbrWinHostStorage {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Host Storage information.
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
        Write-PscriboMessage "Collecting Host Storage information."
    }

    process {
        if ($InfoLevel.Storage -ge 1) {
            try {
                $HostDisks = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Disk }
                if ($HostDisks) {
                    Section -Style Heading3 'Local Disks' {
                        Paragraph 'The following table details physical disks installed in the host'
                        Blankline
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
                        $TableParams = @{
                            Name = "Local Disks"
                            List = $false
                            ColumnWidths = 20, 20, 20, 20, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $LocalDiskReport | Sort-Object -Property 'Disk Number' | Table @TableParams
                    }
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
            #Report any SAN Disks if they exist
            try {
                $SanDisks = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Disk | Where-Object { $_.BusType -Eq "iSCSI" -or $_.BusType -Eq "FCP" } }
                if ($SanDisks) {
                    Section -Style Heading3 'SAN Disks' {
                        Paragraph 'The following section details SAN disks connected to the host'
                        Blankline
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
                        $TableParams = @{
                            Name = "SAN Disks"
                            List = $false
                            ColumnWidths = 20, 20, 20, 20, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $SanDiskReport | Sort-Object -Property 'Disk Number' | Table @TableParams
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