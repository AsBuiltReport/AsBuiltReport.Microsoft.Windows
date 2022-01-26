function Get-AbrWinNetDNSServer {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network DNS Server information.
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
        Write-PscriboMessage "Collecting Network DNS Server information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $DnsServers = Invoke-Command -Session $TempPssSession { Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.ServerAddresses -notlike $null -and $_.InterfaceAlias -notlike "*isatap*" } }
                if ($DnsServers) {
                    Section -Style Heading3 'DNS Servers' {
                        Paragraph 'The following table details the DNS Server Addresses Configured'
                        Blankline
                        $DnsServerReport = @()
                        ForEach ($DnsServer in $DnsServers) {
                            try {
                                $TempDnsServerReport = [PSCustomObject]@{
                                    'Interface' = $DnsServer.InterfaceAlias
                                    'Server Address' = $DnsServer.ServerAddresses -Join ","
                                }
                                $DnsServerReport += $TempDnsServerReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "DNS Servers"
                            List = $false
                            ColumnWidths = 40, 60
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $DnsServerReport | Sort-Object -Property 'Interface' | Table @TableParams
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