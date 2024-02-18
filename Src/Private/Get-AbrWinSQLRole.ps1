function Get-AbrWinSQLRole {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows SQL Server roles information.
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
        Write-PScriboMessage "SQL Server Roles InfoLevel set at $($InfoLevel.SQLServer)."
    }

    process {
        try {
            Write-PScriboMessage "Collecting SQL Server roles information."
            $SQLRoles = Get-DbaServerRole -SqlInstance $SQLServer | Sort-Object -Property Role
            if ($SQLRoles) {
                Write-PScriboMessage "Collecting SQL Server roles information."
                Section -Style Heading4 'Roles' {
                    $ItemInfo = @()
                    foreach ($Item in $SQLRoles) {
                        try {
                            $InObj = [Ordered]@{
                                'Name' = $Item.Role
                                'Owner' = $Item.Owner
                                'Login' = Switch ([string]::IsNullOrEmpty($Item.Login)) {
                                    $true { '--' }
                                    $false { $Item.Login }
                                    default { 'Unknown' }
                                }
                                'Fixed Role' = ConvertTo-TextYN $Item.IsFixedRole
                                'Create Date' = $Item.DateCreated
                            }
                            $ItemInfo += [PSCustomObject]$InObj
                        } catch {
                            Write-PScriboMessage -IsWarning "SQL Server System Roles Section: $($_.Exception.Message)"
                        }
                    }

                    if ($InfoLevel.SQLServer -ge 2) {
                        Paragraph "The following sections detail the configuration of the security roles $($SQLServer.Name)."
                        foreach ($Item in $ItemInfo) {
                            Section -Style NOTOCHeading5 -ExcludeFromTOC "$($Item.Name)" {
                                $TableParams = @{
                                    Name = "Role - $($Item.Name)"
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
                        Paragraph "The following table summarises the configuration of the security role within $($SQLServer.Name)."
                        BlankLine
                        $TableParams = @{
                            Name = "Roles"
                            List = $false
                            Columns = 'Name', 'Owner', 'Login'
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
            Write-PScriboMessage -IsWarning "SQL Server Role Section: $($_.Exception.Message)"
        }
    }
    end {}
}