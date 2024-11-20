function Get-AbrWinIISSummary {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server IIS Summary information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.6
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
        Write-PScriboMessage "Collecting IIS Summary information."
    }

    process {
        if ($InfoLevel.IIS -ge 1) {
            try {
                $OutObj = @()
                $IISApplicationDefaults = Invoke-Command -Session $TempPssSession { (Get-IISServerManager).ApplicationDefaults }
                $IISSiteDefaults = Invoke-Command -Session $TempPssSession { (Get-IISServerManager).SiteDefaults | Select-Object ServerAutoStart, @{name = 'Directory'; Expression = { $_.Logfile.Directory } } }
                if ($IISApplicationDefaults -and $IISSiteDefaults) {
                    try {
                        $inObj = [ordered] @{
                            'Default Application Pool' = ($IISApplicationDefaults).ApplicationPoolName
                            'Enabled Protocols' = (($IISApplicationDefaults).EnabledProtocols).toUpper()
                            'Logfile Path' = ($IISSiteDefaults).Directory
                            'Server Auto Start' = ($IISSiteDefaults).ServerAutoStart
                        }
                        $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)

                        $TableParams = @{
                            Name = "IIS Host Settings"
                            List = $false
                            ColumnWidths = 25, 25, 25, 25
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Table @TableParams
                    } catch {
                        Write-PScriboMessage -IsWarning $_.Exception.Message
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}