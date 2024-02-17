function Get-AbrWinNetAdapterMTU {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network Adapter Interface MTU information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.2
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
        Write-PScriboMessage "Collecting Network Adapter Interfaces MTU information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $NetMtus = Invoke-Command -Session $TempPssSession { Get-NetAdapterAdvancedProperty | Where-Object { $_.DisplayName -eq 'Jumbo Packet' } }
                if ($NetMtus) {
                    Section -Style Heading3 'Network Adapter MTU' {
                        $NetMtuReport = @()
                        ForEach ($NetMtu in $NetMtus) {
                            try {
                                $TempNetMtuReport = [PSCustomObject]@{
                                    'Adapter Name' = $NetMtu.Name
                                    'MTU Size' = $NetMtu.DisplayValue
                                }
                                $NetMtuReport += $TempNetMtuReport
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
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
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}