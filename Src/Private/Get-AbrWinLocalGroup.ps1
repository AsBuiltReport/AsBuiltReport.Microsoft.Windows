function Get-AbrWinLocalGroup {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Local Groups information.
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
        Write-PscriboMessage "Collecting Local Groups information."
    }

    process {
        if ($InfoLevel.Account -ge 1) {
            try {
                $LocalGroups = Invoke-Command -Session $TempPssSession { Get-LocalGroup }
                if ($LocalGroups) {
                    Section -Style Heading3 'Local Groups' {
                        Paragraph 'The following table details local groups configured'
                        Blankline
                        $LocalGroupsReport = @()
                        ForEach ($LocalGroup in $LocalGroups) {
                            try {
                                $TempLocalGroupsReport = [PSCustomObject]@{
                                    'Group Name' = $LocalGroup.Name
                                    'Description' = $LocalGroup.Description
                                }
                                $LocalGroupsReport += $TempLocalGroupsReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Local Group Summary"
                            List = $false
                            ColumnWidths = 40, 60
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $LocalGroupsReport | Sort-Object -Property 'Group Name' | Table @TableParams
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