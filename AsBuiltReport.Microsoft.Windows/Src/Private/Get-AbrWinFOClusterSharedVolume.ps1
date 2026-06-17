function Get-AbrWinFOClusterSharedVolume {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft FailOver Cluster Shared Volume
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.5
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
        Write-PScriboMessage "Collecting Host FailOver Cluster Shared Volume information."
    }

    process {
        try {
            $Settings = Invoke-Command -Session $TempPssSession { Get-ClusterSharedVolume | Select-Object -Property * } | Sort-Object -Property Name
            if ($Settings) {
                Section -Style Heading3 "Cluster Shared Volume" {
                    $OutObj = @()
                    foreach ($Setting in $Settings) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $Setting.Name
                                'Owner Node' = $Setting.OwnerNode
                                'Shared Volume' = Switch ([string]::IsNullOrEmpty($Setting.SharedVolumeInfo.FriendlyVolumeName)) {
                                    $true { "Unknown" }
                                    $false { $Setting.SharedVolumeInfo.FriendlyVolumeName }
                                    default { "--" }
                                }
                                'State' = $Setting.State
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }


                    if ($HealthCheck.FailOverCluster.ClusterSharedVolume) {
                        $OutObj | Where-Object { $_.'State' -notlike 'Online' } | Set-Style -Style Warning -Property 'State'
                    }

                    $TableParams = @{
                        Name = "Cluster Shared Volume - $($Cluster)"
                        List = $false
                        ColumnWidths = 25, 25, 35, 15
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Table @TableParams
                    #Cluster Shared Volume State
                    Get-AbrWinFOClusterSharedVolumeState
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}