function Get-AbrWinNetDNSClient {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Network DNS Client information.
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
        Write-PScriboMessage "Collecting Network DNS Client information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $DnsClient = Invoke-Command -Session $TempPssSession { Get-DnsClientGlobalSetting }
                if ($DnsClient) {
                    Section -Style Heading3 'DNS Client' {
                        Paragraph 'The following table details the DNS Search Domains'
                        BlankLine
                        $OutObj = @()
                        $inObj = [ordered] @{
                            'DNS Suffix' = $DnsClient.SuffixSearchList -Join ","
                            'Use Suffix Search List' = $DnsClient.UseSuffixSearchList
                            'Use Devolution' = $DnsClient.UseDevolution
                            'Devolution Level' = $DnsClient.DevolutionLevel
                        }

                        $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)

                        $TableParams = @{
                            Name = "DNS Search Domain"
                            List = $false
                            ColumnWidths = 40, 20, 20, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Sort-Object -Property 'DNS Suffix' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}