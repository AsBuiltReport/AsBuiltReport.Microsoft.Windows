function Get-AbrWinFOClusterNetwork {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft FailOver Cluster Networks
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.0
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
        Write-PScriboMessage "FailOverCluster InfoLevel set at $($InfoLevel.FailOverCluster)."
        Write-PscriboMessage "Collecting Host FailOver Cluster Networks information."
    }

    process {
        try {
            $Settings = Invoke-Command -Session $TempPssSession { Get-ClusterNetwork -Cluster $using:Cluster} | Sort-Object -Property Name
            if ($Settings) {
                Section -Style Heading3 "Networks" {
                    $OutObj = @()
                    foreach  ($Setting in $Settings) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $Setting.Name
                                'State' = $Setting.State
                                'Role' = $Setting.Role
                                'Address' = "$($Setting.Address)/$($Setting.AddressMask)"
                            }
                            $OutObj += [pscustomobject]$inobj
                        }
                        catch {
                            Write-PscriboMessage -IsWarning $_.Exception.Message
                        }
                    }


                    if ($HealthCheck.FailOverCluster.Network) {
                        $OutObj | Where-Object { $_.'State' -ne 'UP'} | Set-Style -Style Warning -Property 'State'
                    }

                    $TableParams = @{
                        Name = "Networks - $($Cluster.toUpper().split(".")[0])"
                        List = $false
                        ColumnWidths = 30, 15, 20, 35
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Table @TableParams
                    # Cluster Network Interfaces
                    Get-AbrWinFOClusterNetworkInterface
                }
            }
        }
        catch {
            Write-PscriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}