function Get-AbrWinFOCluster {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft FailOver Cluster configuration
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
        Write-PScriboMessage "Collecting Host FailOver Cluster Server information."
    }

    process {
        try {
            $Settings = Invoke-Command -Session $TempPssSession { Get-Cluster | Select-Object -Property * }
            if ($Settings) {
                $OutObj = @()
                try {
                    $inObj = [ordered] @{
                        'Name' = $Settings.Name
                        'Domain' = $Settings.Domain
                        'Shared Volumes Root' = $Settings.SharedVolumesRoot
                        'Administrative Access Point' = $Settings.AdministrativeAccessPoint
                        'Description' = $Settings.Description
                    }
                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                } catch {
                    Write-PScriboMessage -IsWarning $_.Exception.Message
                }

                $TableParams = @{
                    Name = "FailOver Cluster Servers Settings - $($System.split(".")[0])"
                    List = $true
                    ColumnWidths = 40, 60
                }
                if ($Report.ShowTableCaptions) {
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                }
                $OutObj | Table @TableParams
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}