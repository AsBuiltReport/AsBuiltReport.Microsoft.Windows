function Get-AbrWinFOClusterNetworkInterface {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft FailOver Cluster Network Interfaces
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
        Write-PScriboMessage "FailOverCluster InfoLevel set at $($InfoLevel.FailOverCluster)."
        Write-PScriboMessage "Collecting Host FailOver Cluster Network Interface information."
    }

    process {
        try {
            $Settings = Invoke-Command -Session $TempPssSession { Get-ClusterNetworkInterface -Cluster $using:Cluster } | Sort-Object -Property Name
            if ($Settings) {
                Section -Style Heading3 "Interfaces" {
                    $OutObj = @()
                    foreach ($Setting in $Settings) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $Setting.Name
                                'Node' = $Setting.Node
                                'Network' = $Setting.Network
                                'State' = $Setting.State
                            }
                            $OutObj += [pscustomobject]$inobj
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }


                    if ($HealthCheck.FailOverCluster.Network) {
                        $OutObj | Where-Object { $_.'State' -ne 'UP' } | Set-Style -Style Warning -Property 'State'
                    }

                    $TableParams = @{
                        Name = "Interfaces - $($Cluster.toUpper().split(".")[0])"
                        List = $false
                        ColumnWidths = 30, 25, 30, 15
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Table @TableParams
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}