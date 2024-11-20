function Get-AbrWinNetFirewall {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Host Firewall information.
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
        Write-PscriboMessage "Collecting Host Firewall information."
    }

    process {
        if ($InfoLevel.Networking -ge 1) {
            try {
                $NetFirewallProfile = Get-NetFirewallProfile -CimSession $TempCimSession
                if ($NetFirewallProfile) {
                    Section -Style Heading2 'Windows Firewall' {
                        $OutObj = @()
                        Foreach ($FirewallProfile in $NetFireWallProfile) {
                            try {
                                $inObj = [ordered] @{
                                    'Profile' = $FirewallProfile.Name
                                    'Profile Enabled' = $FirewallProfile.Enabled
                                    'Inbound Action' = $FirewallProfile.DefaultInboundAction
                                    'Outbound Action' = $FirewallProfile.DefaultOutboundAction
                                }
                                $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }

                        if ($HealthCheck.Networking.Firewall) {
                            $OutObj | Where-Object { $_.'Profile Enabled' -notlike 'Yes'} | Set-Style -Style Warning -Property 'Profile Enabled'
                        }

                        $TableParams = @{
                            Name = "Windows Firewall Profiles"
                            List = $false
                            ColumnWidths = 25, 25, 25, 25
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Sort-Object -Property 'Profile' | Table @TableParams
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