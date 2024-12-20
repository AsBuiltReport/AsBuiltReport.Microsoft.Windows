function Get-AbrWinLocalGroup {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Local Groups information.
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
        Write-PScriboMessage "Collecting Local Groups information."
    }

    process {
        if ($InfoLevel.Account -ge 1) {
            try {
                if ($LocalGroups) {
                    Section -Style Heading3 'Local Groups' {
                        $OutObj = @()
                        ForEach ($LocalGroup in $LocalGroups) {
                            try {
                                $inObj = [ordered] @{
                                    'Group Name' = $LocalGroup.GroupName
                                    'Description' = Switch ([string]::IsNullOrEmpty($LocalGroup.Description)) {
                                        $true { "--" }
                                        $false { $LocalGroup.Description }
                                        default { "Unknown" }
                                    }
                                    'Members' = Switch ([string]::IsNullOrEmpty($LocalGroup.Members)) {
                                        $true { "--" }
                                        $false { $LocalGroup.Members }
                                        default { "Unknown" }
                                    }
                                }
                                $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Local Group Summary"
                            List = $false
                            ColumnWidths = 30, 40, 30
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Sort-Object -Property 'Group Name' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}
