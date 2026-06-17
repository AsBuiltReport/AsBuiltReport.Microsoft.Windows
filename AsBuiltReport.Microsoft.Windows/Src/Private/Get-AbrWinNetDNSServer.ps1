function Get-AbrWinNetDNSServer {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network DNS Server information.
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
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
        Write-PScriboMessage "Collecting Network DNS Server information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $DnsServers = Invoke-Command -Session $TempPssSession { Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.ServerAddresses -notlike $null -and $_.InterfaceAlias -notlike "*isatap*" } }
                if ($DnsServers) {
                    Section -Style Heading3 'DNS Servers' {
                        Paragraph 'The following table details the DNS Server Addresses Configured'
                        BlankLine
                        $OutObj = @()
                        ForEach ($DnsServer in $DnsServers) {
                            try {
                                $inObj = [ordered] @{
                                    'Interface' = $DnsServer.InterfaceAlias
                                    'Server Address' = $DnsServer.ServerAddresses -Join ","
                                }
                                $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
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
                        $OutObj | Sort-Object -Property 'Interface' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}