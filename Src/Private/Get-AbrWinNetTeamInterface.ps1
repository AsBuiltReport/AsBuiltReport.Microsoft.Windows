function Get-AbrWinNetTeamInterface {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network Team Interfaces information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
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
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
        Write-PscriboMessage "Collecting Network Team Interfaces information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $NetworkTeamCheck = Invoke-Command -Session $TempPssSession { Get-NetLbfoTeam }
                if ($NetworkTeamCheck) {
                    Section -Style Heading3 'Network Team Interfaces' {
                        Paragraph 'The following table details Network Team Interfaces'
                        Blankline
                        $NetTeams = Invoke-Command -Session $TempPssSession { Get-NetLbfoTeam }
                        $NetTeamReport = @()
                        ForEach ($NetTeam in $NetTeams) {
                            try {
                                $TempNetTeamReport = [PSCustomObject]@{
                                    'Team Name' = $NetTeam.Name
                                    'Team Mode' = $NetTeam.tm
                                    'Load Balancing' = $NetTeam.lba
                                    'Network Adapters' = $NetTeam.Members -Join ","
                                }
                                $NetTeamReport += $TempNetTeamReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Network Team Interfaces"
                            List = $false
                            ColumnWidths = 20, 20, 20, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $NetTeamReport | Sort-Object -Property 'Team Name' | Table @TableParams
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