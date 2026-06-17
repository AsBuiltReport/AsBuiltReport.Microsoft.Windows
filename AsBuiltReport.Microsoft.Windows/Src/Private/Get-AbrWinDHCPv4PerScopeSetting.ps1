function Get-AbrWinDHCPv4PerScopeSetting {
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
        Write-PScriboMessage "Collecting Host DHCP Server Scope information."
    }

    process {
        try {
            $DHCPScopes = Get-DhcpServerv4Scope -CimSession $TempCimSession | Select-Object -ExpandProperty ScopeId
            if ($DHCPScopes) {
                Section -Style Heading3 "Per Scope Options" {
                    Paragraph "The following section provides a summary of the DHCP servers Scope Server Options information."
                    BlankLine
                    foreach ($Scope in $DHCPScopes) {
                        try {
                            $DHCPScopeOptions = Get-DhcpServerv4OptionValue -CimSession $TempCIMSession -ScopeId $Scope
                            if ($DHCPScopeOptions) {
                                Section -Style Heading4 "$Scope" {
                                    Paragraph "The following table details Scope Server Options Settings."
                                    BlankLine
                                    $OutObj = @()
                                    foreach ($Option in $DHCPScopeOptions) {
                                        try {
                                            Write-PScriboMessage "Collecting DHCP Server Scope Server Option value $($Option.OptionId)"
                                            $inObj = [ordered] @{
                                                'Name' = $Option.Name
                                                'Option Id' = $Option.OptionId
                                                'Value' = $Option.Value
                                                'Policy Name' = $Option.PolicyName
                                            }
                                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                        } catch {
                                            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Scope Options Item)"
                                        }
                                    }

                                    $TableParams = @{
                                        Name = "Scopes Options - $Scope"
                                        List = $false
                                        ColumnWidths = 40, 15, 20, 25
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $OutObj | Sort-Object -Property 'Option Id' | Table @TableParams
                                }
                            }
                        } catch {
                            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Scope Options Section)"
                        }
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning "$($_.Exception.Message) (Scope Options Section)"
        }
    }

    end {}

}