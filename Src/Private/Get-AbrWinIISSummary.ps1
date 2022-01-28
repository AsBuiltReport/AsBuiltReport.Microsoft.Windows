function Get-AbrWinIISSummary {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server IIS Summary information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
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
        Write-PScriboMessage "IIS InfoLevel set at $($InfoLevel.IIS)."
        Write-PscriboMessage "Collecting IIS Summary information."
    }

    process {
        if ($InfoLevel.IIS -ge 1) {
            try {
                $IISApplicationDefaults = Invoke-Command -Session $TempPssSession { (Get-IISServerManager).ApplicationDefaults }
                $IISSiteDefaults = Invoke-Command -Session $TempPssSession { (Get-IISServerManager).SiteDefaults | Select-Object ServerAutoStart,@{name='Directory'; Expression={$_.Logfile.Directory}} }
                if ($IISApplicationDefaults -and $IISSiteDefaults) {
                    $IISServerManagerReport = [PSCustomObject]@{
                        'Default Application Pool' = ($IISApplicationDefaults).ApplicationPoolName
                        'Enabled Protocols' = (($IISApplicationDefaults).EnabledProtocols).toUpper()
                        'Logfile Path' = ($IISSiteDefaults).Directory
                        'Server Auto Start' = ConvertTo-TextYN ($IISSiteDefaults).ServerAutoStart
                    }
                    $TableParams = @{
                        Name = "IIS Host Settings"
                        List = $false
                        ColumnWidths = 25, 25, 25, 25
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $IISServerManagerReport | Table @TableParams
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}