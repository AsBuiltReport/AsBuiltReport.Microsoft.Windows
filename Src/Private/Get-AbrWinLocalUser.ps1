function Get-AbrWinLocalUser {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Local Users information.
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
        Write-PScriboMessage "Account InfoLevel set at $($InfoLevel.Account)."
        Write-PscriboMessage "Collecting Local Users information."
    }

    process {
        if ($InfoLevel.Account -ge 1) {
            try {
                $LocalUsers = Invoke-Command -Session $TempPssSession { Get-LocalUser }
                if ($LocalUsers) {
                    Section -Style Heading3 'Local Users' {
                        Paragraph 'The following table details local users'
                        Blankline
                        $LocalUsersReport = @()
                        ForEach ($LocalUser in $LocalUsers) {
                            try {
                                $TempLocalUsersReport = [PSCustomObject]@{
                                    'User Name' = $LocalUser.Name
                                    'Description' = $LocalUser.Description
                                    'Account Enabled' = ConvertTo-TextYN $LocalUser.Enabled
                                    'Last Logon Date' = Switch (($LocalUser.LastLogon).count) {
                                        0 {"-"}
                                        default {$LocalUser.LastLogon.ToShortDateString()}
                                    }
                                }
                                $LocalUsersReport += $TempLocalUsersReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Local Users"
                            List = $false
                            ColumnWidths = 20, 40, 10, 30
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $LocalUsersReport | Sort-Object -Property 'User Name' | Table @TableParams
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