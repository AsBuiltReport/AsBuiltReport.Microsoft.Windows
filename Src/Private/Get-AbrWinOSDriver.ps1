function Get-AbrWinOSDriver {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Operating System Drivers information.
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
        Write-PScriboMessage "Operating System InfoLevel set at $($InfoLevel.OperatingSystem)."
        Write-PScriboMessage "Collecting Operating System Drivers information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            try {
                $HostDriversList = Invoke-Command -Session $TempPssSession { Get-WindowsDriver -Online }
                if ($HostDriversList) {
                    Section -Style Heading3 'Drivers' {
                        Invoke-Command -Session $TempPssSession { Import-Module DISM }
                        $OutObj = @()
                        ForEach ($HostDriver in $HostDriversList) {
                            try {
                                $inObj = [ordered] @{
                                    'Class Description' = $HostDriver.ClassDescription
                                    'Provider Name' = $HostDriver.ProviderName
                                    'Driver Version' = $HostDriver.Version
                                    'Version Date' = $HostDriver.Date.ToShortDateString()
                                }
                                $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Hardware Drivers"
                            List = $false
                            ColumnWidths = 30, 30, 20, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Sort-Object -Property 'Class Description' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}