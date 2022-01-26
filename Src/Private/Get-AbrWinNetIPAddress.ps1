function Get-AbrWinNetIPAddress {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network IP Address information.
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
        Write-PscriboMessage "Collecting Network IP Address information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $NetIPs = Invoke-Command -Session $TempPssSession { Get-NetIPConfiguration | Where-Object -FilterScript { ($_.NetAdapter.Status -Eq "Up") } }
                if ($NetIPs) {
                    Section -Style Heading3 'IP Addresses' {
                        Paragraph 'The following table details IP Addresses assigned to hosts'
                        Blankline
                        $NetIpsReport = @()
                        ForEach ($NetIp in $NetIps) {
                            try {
                                $TempNetIpsReport = [PSCustomObject]@{
                                    'Interface Name' = $NetIp.InterfaceAlias
                                    'Interface Description' = $NetIp.InterfaceDescription
                                    'IPv4 Addresses' = $NetIp.IPv4Address.IPAddress -Join ","
                                    'Subnet Mask' = $NetIp.IPv4Address[0].PrefixLength
                                    'IPv4 Gateway' = $NetIp.IPv4DefaultGateway.NextHop
                                }
                                $NetIpsReport += $TempNetIpsReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Net IP Addresse"
                            List = $false
                            ColumnWidths = 25, 25, 20, 10, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $NetIpsReport | Sort-Object -Property 'Interface Name' | Table @TableParams
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