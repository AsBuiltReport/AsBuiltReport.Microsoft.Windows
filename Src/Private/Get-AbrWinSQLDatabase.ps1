function Get-AbrWinSQLDatabase {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows SQL Server database information.
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
        Write-PScriboMessage "SQL Server Dstabases InfoLevel set at $($InfoLevel.SQLServer)."
    }

    process {
        Section -Style Heading3 'Databases' {
            $CompatibilityHash = @{
                'Version60' = 'SQL Server 6.0'
                'Version65' = 'SQL Server 6.5'
                'Version70' = 'SQL Server 7.0'
                'Version80' = 'SQL Server 2000'
                'Version90' = 'SQL Server 2005'
                'Version100' = 'SQL Server 2008'
                'Version110' = 'SQL Server 2012'
                'Version120' = 'SQL Server 2014'
                'Version130' = 'SQL Server 2016'
                'Version140' = 'SQL Server 2017'
                'Version150' = 'SQL Server 2019'
                'Version160' = 'SQL Server 2022'

            }
            try {
                Write-PScriboMessage "Collecting SQL Server databases information."
                $SQLDBs = Get-DbaDatabase -SqlInstance $SQLServer -ExcludeUser | Sort-Object -Property Name
                if ($SQLDBs) {
                    Write-PScriboMessage "Collecting SQL Server system databases information."
                    Section -Style Heading4 'System Databases' {
                        $SQLDBInfo = @()
                        foreach ($SQLDB in $SQLDBs) {
                            try {
                                $InObj = [Ordered]@{
                                    'Name' = $SQLDB.Name
                                    'Status' = $SQLDB.Status
                                    'Is Accessible?' = ConvertTo-TextYN $SQLDB.IsAccessible
                                    'Recovery Model' = $SQLDB.RecoveryModel
                                    'Size' = Switch ([string]::IsNullOrEmpty($SQLDB.SizeMB)) {
                                        $true { '--' }
                                        $false { "$($SQLDB.SizeMB) MB" }
                                        default { 'Unknown' }
                                    }
                                    'Compatibility' = $CompatibilityHash[[string]$SQLDB.Compatibility]
                                    'Collation' = $SQLDB.Collation
                                    'Encrypted' = ConvertTo-TextYN $SQLDB.Encrypted
                                    'Last Full Backup' = Switch ($SQLDB.LastFullBackup) {
                                        '01/01/0001 00:00:00' { "Never" }
                                        $null { '--' }
                                        default { $SQLDB.LastFullBackup }
                                    }
                                    'Last Log Backup' = Switch ($SQLDB.LastLogBackup) {
                                        '01/01/0001 00:00:00' { "Never" }
                                        $null { '--' }
                                        default { $SQLDB.LastLogBackup }
                                    }
                                    'Owner' = $SQLDB.Owner
                                }
                                $SQLDBInfo += [PSCustomObject]$InObj
                            } catch {
                                Write-PScriboMessage -IsWarning "SQL Server System Database table: $($_.Exception.Message)"
                            }
                        }

                        if ($InfoLevel.SQLServer -ge 2) {
                            Paragraph "The following sections detail the configuration of the system databases."
                            foreach ($SQLDB in $SQLDBInfo) {
                                Section -Style NOTOCHeading5 -ExcludeFromTOC "$($SQLDB.Name)" {
                                    $TableParams = @{
                                        Name = "System Database - $($SQLDB.Name)"
                                        List = $true
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $SQLDB | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the system databases."
                            BlankLine
                            $TableParams = @{
                                Name = "System Databases"
                                List = $false
                                Columns = 'Name', 'Owner', 'Status', 'Recovery Model', 'Size'
                                ColumnWidths = 32, 32, 12, 12, 12
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $SQLDBInfo | Table @TableParams
                        }
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning "SQL Server System Database Section: $($_.Exception.Message)"
            }
            try {
                $SQLDBs = Get-DbaDatabase -SqlInstance $SQLServer -ExcludeSystem | Sort-Object -Property Name
                if ($SQLDBs) {
                    Write-PScriboMessage "Collecting SQL Server user databases information."
                    Section -Style Heading4 'User Databases' {
                        $SQLDBInfo = @()
                        foreach ($SQLDB in $SQLDBs) {
                            try {
                                $InObj = [Ordered]@{
                                    'Name' = $SQLDB.Name
                                    'Status' = $SQLDB.Status
                                    'Is Accessible?' = ConvertTo-TextYN $SQLDB.IsAccessible
                                    'Recovery Model' = $SQLDB.RecoveryModel
                                    'Size' = Switch ([string]::IsNullOrEmpty($SQLDB.SizeMB)) {
                                        $true { '--' }
                                        $false { "$($SQLDB.SizeMB) MB" }
                                        default { 'Unknown' }
                                    }
                                    'Compatibility' = $CompatibilityHash[[string]$SQLDB.Compatibility]
                                    'Collation' = $SQLDB.Collation
                                    'Encrypted' = ConvertTo-TextYN $SQLDB.Encrypted
                                    'Last Full Backup' = Switch ($SQLDB.LastFullBackup) {
                                        '01/01/0001 00:00:00' { "Never" }
                                        $null { '--' }
                                        default { $SQLDB.LastFullBackup }
                                    }
                                    'Last Log Backup' = Switch ($SQLDB.LastLogBackup) {
                                        '01/01/0001 00:00:00' { "Never" }
                                        $null { '--' }
                                        default { $SQLDB.LastLogBackup }
                                    }
                                    'Owner' = $SQLDB.Owner
                                }
                                $SQLDBInfo += [PSCustomObject]$InObj
                            } catch {
                                Write-PScriboMessage -IsWarning "SQL Server User Database table: $($_.Exception.Message)"
                            }
                        }

                        if ($InfoLevel.SQLServer -ge 2) {
                            Paragraph "The following sections detail the configuration of the user databases within $($SQLServer.Name)."
                            foreach ($SQLDB in $SQLDBInfo) {
                                Section -Style NOTOCHeading5 -ExcludeFromTOC "$($SQLDB.Name)" {
                                    $TableParams = @{
                                        Name = "User Database - $($SQLDB.Name)"
                                        List = $true
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $SQLDB | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the databases within $($SQLServer.Name)."
                            BlankLine
                            $TableParams = @{
                                Name = "User Databases"
                                List = $false
                                Columns = 'Name', 'Owner', 'Status', 'Recovery Model', 'Size'
                                ColumnWidths = 32, 32, 12, 12, 12
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $SQLDBInfo | Table @TableParams
                        }
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning "SQL Server User Database Section: $($_.Exception.Message)"
            }
        }
    }
    end {}
}