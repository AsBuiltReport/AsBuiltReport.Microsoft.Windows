function Get-AbrWinLocalUser {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Local Users information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.6
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
        Write-PScriboMessage "Account InfoLevel set at $($InfoLevel.Account)."
        Write-PScriboMessage "Collecting Local Users information."
    }

    process {
        if ($InfoLevel.Account -ge 1) {
            try {
                if ($LocalUsers) {
                    Section -Style Heading3 'Local Users' {
                        $OutObj = @()
                        ForEach ($LocalUser in $LocalUsers) {
                            try {
                                $inObj = [ordered] @{
                                    'User Name' = $LocalUser.Name
                                    'Description' = $LocalUser.Description
                                    'Account Enabled' = $LocalUser.Enabled
                                    'Last Logon Date' = Switch (($LocalUser.LastLogon).count) {
                                        0 { "--" }
                                        default { $LocalUser.LastLogon.ToShortDateString() }
                                    }
                                }
                                $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
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
                        $OutObj | Sort-Object -Property 'User Name' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}