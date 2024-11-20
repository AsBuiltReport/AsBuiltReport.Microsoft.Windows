function Get-AbrWinSQLLogin {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows SQL Server login information.
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
        Write-PScriboMessage "SQL Server Logins InfoLevel set at $($InfoLevel.SQLServer)."
    }

    process {
        try {
            Write-PScriboMessage "Collecting SQL Server logins information."
            $SQLLogins = Get-DbaLogin -SqlInstance $SQLServer | Sort-Object -Property Name
            if ($SQLLogins) {
                Write-PScriboMessage "Collecting SQL Server logins information."
                Section -Style Heading4 'Logins' {
                    $ItemInfo = @()
                    foreach ($Item in $SQLLogins) {
                        $ServerRoles = try { Get-DbaServerRoleMember -SqlInstance $SQLServer -Login $Item.Name } catch { Out-Null }
                        try {
                            $InObj = [Ordered]@{
                                'Name' = $Item.Name
                                'Login Type' = $Item.LoginType -creplace '(?<=\w)([A-Z])', ' $1'
                                'Server Roles' = Switch ([string]::IsNullOrEmpty($ServerRoles.Role)) {
                                    $true { '--' }
                                    $false { $ServerRoles.Role }
                                    default { 'Unknown' }
                                }
                                'Create Date' = $Item.CreateDate
                                'Last Login' = Switch ([string]::IsNullOrEmpty($Item.LastLogin)) {
                                    $true { 'Never' }
                                    $false { $Item.LastLogin }
                                    default { 'Unknown' }
                                }
                                'Has Access?' = $Item.HasAccess
                                'Is Locked?' = Switch ([string]::IsNullOrEmpty($Item.IsLocked)) {
                                    $true { 'No' }
                                    $false { $Item.IsLocked }
                                    default { 'Unknown' }
                                }
                                'Is Disabled?' = $Item.IsDisabled
                                'Must Change Password' = Switch ([string]::IsNullOrEmpty($Item.MustChangePassword)) {
                                    $true { 'No' }
                                    $false { $Item.MustChangePassword }
                                    default { 'Unknown' }
                                }


                            }
                            $ItemInfo += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning "SQL Server System Login Section: $($_.Exception.Message)"
                        }
                    }

                    if ($InfoLevel.SQLServer -ge 2) {
                        Paragraph "The following sections detail the configuration of the security login."
                        foreach ($Item in $ItemInfo) {
                            Section -Style NOTOCHeading5 -ExcludeFromTOC "$($Item.Name)" {
                                $TableParams = @{
                                    Name = "Login - $($Item.Name)"
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
                        Paragraph "The following table summarises the configuration of the security login."
                        BlankLine
                        $TableParams = @{
                            Name = "Logins"
                            List = $false
                            Columns = 'Name', 'Login Type', 'Server Roles', 'Has Access?', 'Is Locked?', 'Is Disabled?'
                            ColumnWidths = 30, 17, 17, 12, 12, 12
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $ItemInfo | Table @TableParams
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning "SQL Server System Security Login Section: $($_.Exception.Message)"
        }
    }
    end {}
}