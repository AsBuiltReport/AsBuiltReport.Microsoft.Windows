function Get-AbrWinNetAdapterMTU {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network Adapter Interface MTU information.
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
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
        Write-PscriboMessage "Collecting Network Adapter Interfaces MTU information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $NetMtus = Invoke-Command -Session $TempPssSession { Get-NetAdapterAdvancedProperty | Where-Object { $_.DisplayName -eq 'Jumbo Packet' } }
                if ($NetMtus) {
                    Section -Style Heading3 'Network Adapter MTU' {
                        Paragraph 'The following table lists Network Adapter MTU settings'
                        Blankline
                        $NetMtuReport = @()
                        ForEach ($NetMtu in $NetMtus) {
                            try {
                                $TempNetMtuReport = [PSCustomObject]@{
                                    'Adapter Name' = $NetMtu.Name
                                    'MTU Size' = $NetMtu.DisplayValue
                                }
                                $NetMtuReport += $TempNetMtuReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Network Adapter MTU"
                            List = $false
                            ColumnWidths = 50, 50
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $NetMtuReport | Sort-Object -Property 'Adapter Name' | Table @TableParams
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