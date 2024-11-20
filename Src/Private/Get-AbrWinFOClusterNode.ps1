function Get-AbrWinFOClusterNode {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft FailOver Cluster Nodes
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
        Write-PScriboMessage "FailOverCluster InfoLevel set at $($InfoLevel.Infrastructure.FailOverCluster)."
    }

    process {
        try {
            $Settings = Invoke-Command -Session $TempPssSession { Get-ClusterNode } | Sort-Object -Property Identity
            if ($Settings) {
                Write-PScriboMessage "Collecting Host FailOver Cluster Permissions Settings information."
                Section -Style Heading3 'Nodes' {
                    $OutObj = @()
                    foreach ($Setting in $Settings) {
                        $inObj = [ordered] @{
                            'Name' = $Setting.Name
                            'State' = $Setting.State
                            'Type' = $Setting.Type
                            'Cluster' = $Setting.Cluster
                            'Fault Domain' = $Setting.FaultDomain
                            'Model' = $Setting.Model
                            'Manufacturer' = $Setting.Manufacturer
                            'Description' = $Setting.Description

                        }
                        $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                    }

                    if ($HealthCheck.FailOverCluster.Nodes) {
                        $OutObj | Where-Object { $_.'State' -ne 'UP' } | Set-Style -Style Warning -Property 'State'
                    }

                    if ($InfoLevel.FailOverCluster -ge 2) {
                        Paragraph "The following sections detail the configuration of the Failover Cluster Nodes."
                        foreach ($Setting in $OutObj) {
                            Section -ExcludeFromTOC -Style NOTOCHeading4 "$($Setting.Name)" {
                                $TableParams = @{
                                    Name = "Node - $($Setting.Name)"
                                    List = $true
                                    ColumnWidths = 50, 50
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $Setting | Table @TableParams
                            }
                        }
                    } else {
                        Paragraph "The following table summarizes the configuration of the Failover Cluster Nodes."
                        BlankLine
                        $TableParams = @{
                            Name = "Nodes - $($Cluster)"
                            List = $false
                            Columns = 'Name', 'State', 'Type'
                            ColumnWidths = 40, 30, 30
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Table @TableParams
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning "FailOver Cluster Nodes Section: $($_.Exception.Message)"
        }
    }

    end {}
}