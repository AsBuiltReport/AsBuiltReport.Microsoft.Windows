function Get-AbrWinFOClusterFaultDomain {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft FailOver Cluster Fault Domain
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
        Write-PscriboMessage "Collecting Host FailOver Cluster Fault Domain information."
    }

    process {
        try {
            $Settings = Get-ClusterFaultDomain -CimSession $TempCimSession | Sort-Object -Property Name
            if ($Settings) {
                Section -Style Heading3 "Fault Domain" {
                    $OutObj = @()
                    foreach  ($Setting in $Settings) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $Setting.Name
                                'Type' = $Setting.Type
                                'Parent Name' = $Setting.ParentName
                                'Children Names' = $Setting.ChildrenNames
                                'Location' = ConvertTo-EmptyToFiller $Setting.Location
                            }
                            $OutObj += [pscustomobject]$inobj
                        }
                        catch {
                            Write-PscriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    $TableParams = @{
                        Name = "Fault Domain - $($Cluster.toUpper().split(".")[0])"
                        List = $false
                        ColumnWidths = 20, 20, 20, 20, 20
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Table @TableParams
                }
            }
        }
        catch {
            Write-PscriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}