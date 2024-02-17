function Get-AbrWinSMBSummary {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server File Server Summary information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.2
        Author:         Jonathan Colon
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
        Write-PScriboMessage "SMB InfoLevel set at $($InfoLevel.SMB)."
        Write-PScriboMessage "Collecting File Server Summary information."
    }

    process {
        if ($InfoLevel.SMB -ge 1) {
            try {
                $SMBSummary = Get-SmbServerConfiguration -CimSession $TempCimSession | Select-Object AutoShareServer, EnableLeasing, EnableMultiChannel, EnableOplocks, KeepAliveTime, EnableSMB1Protocol, EnableSMB2Protocol
                if ($SMBSummary) {
                    $SMBSummaryReport = [PSCustomObject]@{
                        'Auto Share Server' = ConvertTo-TextYN $SMBSummary.AutoShareServer
                        'Enable Leasing' = ConvertTo-TextYN $SMBSummary.EnableLeasing
                        'Enable MultiChannel' = ConvertTo-TextYN $SMBSummary.EnableMultiChannel
                        'Enable Oplocks' = ConvertTo-TextYN $SMBSummary.EnableOplocks
                        'Keep Alive Time' = $SMBSummary.KeepAliveTime
                        'SMB1 Protocol' = ConvertTo-TextYN $SMBSummary.EnableSMB1Protocol
                        'SMB2 Protocol' = ConvertTo-TextYN $SMBSummary.EnableSMB2Protocol
                    }

                    if ($HealthCheck.SMB.BP) {
                        $SMBSummaryReport | Where-Object { $_.'SMB1 Protocol' -eq 'Yes' } | Set-Style -Style Warning -Property 'SMB1 Protocol'
                    }

                    $TableParams = @{
                        Name = "SMB Server Settings"
                        List = $true
                        ColumnWidths = 40, 60
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $SMBSummaryReport | Table @TableParams
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}