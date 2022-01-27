function Get-AbrWinNetAdapter {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network Adapter information.
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
        Write-PscriboMessage "Collecting Network Adapter information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $HostAdapters = Invoke-Command -Session $TempPssSession { Get-NetAdapter }
                if ($HostAdapters) {
                    Section -Style Heading3 'Network Adapters' {
                        Paragraph 'The Following table details host network adapters'
                        Blankline
                        $HostAdaptersReport = @()
                        ForEach ($HostAdapter in $HostAdapters) {
                            try {
                                $TempHostAdaptersReport = [PSCustomObject]@{
                                    'Adapter Name' = $HostAdapter.Name
                                    'Adapter Description' = $HostAdapter.InterfaceDescription
                                    'Mac Address' = $HostAdapter.MacAddress
                                    'Link Speed' = $HostAdapter.LinkSpeed
                                }
                                $HostAdaptersReport += $TempHostAdaptersReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Network Adapters"
                            List = $false
                            ColumnWidths = 30, 35, 20, 15
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $HostAdaptersReport | Sort-Object -Property 'Adapter Name' | Table @TableParams
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