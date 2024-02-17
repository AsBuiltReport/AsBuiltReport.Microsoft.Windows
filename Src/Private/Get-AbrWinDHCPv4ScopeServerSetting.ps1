function Get-AbrWinDHCPv4ScopeServerSetting {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Windows DHCP Servers Scopes Server Options from DHCP Servers
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
        Write-PScriboMessage "DHCP InfoLevel set at $($InfoLevel.DHCP)."
        Write-PScriboMessage "Collecting Host DHCP Server information."
    }

    process {
        $DHCPScopeOptions = Get-DhcpServerv4OptionValue -CimSession $TempCIMSession
        if ($DHCPScopeOptions) {
            Section -Style Heading3 "Scope Server Options" {
                Paragraph "The following section provides a summary of the DHCP servers Scope Server Options information."
                BlankLine
                $OutObj = @()
                Write-PScriboMessage "Discovered '$(($DHCPScopeOptions | Measure-Object).Count)' DHCP scopes server opions."
                foreach ($Option in $DHCPScopeOptions) {
                    try {
                        Write-PScriboMessage "Collecting DHCP Server Scope Server Option value $($Option.OptionId)"
                        $inObj = [ordered] @{
                            'Name' = $Option.Name
                            'Option Id' = $Option.OptionId
                            'Value' = $Option.Value
                            'Policy Name' = ConvertTo-EmptyToFiller $Option.PolicyName
                        }
                        $OutObj += [pscustomobject]$inobj
                    } catch {
                        Write-PScriboMessage -IsWarning "$($_.Exception.Message) (DHCP scopes server opions item)"
                    }
                }
                $TableParams = @{
                    Name = "Scopes Server Options - $($System.split(".", 2).ToUpper()[0])"
                    List = $false
                    ColumnWidths = 40, 15, 20, 25
                }
                if ($Report.ShowTableCaptions) {
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                }
                $OutObj | Sort-Object -Property 'Option Id' | Table @TableParams
                try {
                    $DHCPScopeOptions = Get-DhcpServerv4DnsSetting -CimSession $TempCIMSession
                    if ($DHCPScopeOptions) {
                        Section -Style Heading4 "Scope DNS Setting" {
                            Paragraph "The following section provides a summary of the DHCP servers Scope DNS Setting information."
                            BlankLine
                            $OutObj = @()
                            foreach ($Option in $DHCPScopeOptions) {
                                try {
                                    Write-PScriboMessage "Collecting DHCP Server Scope DNS Setting."
                                    $inObj = [ordered] @{
                                        'Dynamic Updates' = $Option.DynamicUpdates
                                        'Dns Suffix' = ConvertTo-EmptyToFiller $Option.DnsSuffix
                                        'Name Protection' = ConvertTo-EmptyToFiller $Option.NameProtection
                                        'Update Dns RR For Older Clients' = ConvertTo-EmptyToFiller $Option.UpdateDnsRRForOlderClients
                                        'Disable Dns Ptr RR Update' = ConvertTo-EmptyToFiller $Option.DisableDnsPtrRRUpdate
                                        'Delete Dns RR On Lease Expiry' = ConvertTo-EmptyToFiller $Option.DeleteDnsRROnLeaseExpiry
                                    }
                                    $OutObj += [pscustomobject]$inobj
                                } catch {
                                    Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Scope DNS Setting Item)"
                                }
                            }

                            $TableParams = @{
                                Name = "Scopes DNS Setting - $($System.toUpper().split(".", 2)[0])"
                                List = $true
                                ColumnWidths = 40, 60
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $OutObj | Table @TableParams
                        }
                    }
                } catch {
                    Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Scope DNS Setting Table)"
                }
            }
        }
    }

    end {}

}