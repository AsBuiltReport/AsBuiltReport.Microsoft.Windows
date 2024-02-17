function Get-AbrWinFOClusterPermission {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft FailOver Cluster Permissions
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
        Write-PScriboMessage "Collecting Host FailOver Cluster Permissions Settings information."
    }

    process {
        try {
            $Settings = Invoke-Command -Session $TempPssSession { Get-ClusterAccess } | Sort-Object -Property Identity
            if ($Settings) {
                Section -Style Heading3 "Access Permissions" {
                    $OutObj = @()
                    foreach ($Setting in $Settings) {
                        try {
                            $inObj = [ordered] @{
                                'Identity' = $Setting.IdentityReference
                                'Access Control Type' = $Setting.AccessControlType
                                'Rights' = $Setting.ClusterRights
                            }
                            $OutObj += [pscustomobject]$inobj
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    $TableParams = @{
                        Name = "Access Permission - $($Cluster)"
                        List = $false
                        ColumnWidths = 60, 20, 20
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