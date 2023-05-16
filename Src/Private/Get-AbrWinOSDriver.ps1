function Get-AbrWinOSDriver {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Operating System Drivers information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
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
        Write-PscriboMessage "Collecting Operating System Drivers information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            try {
                $HostDriversList = Invoke-Command -Session $TempPssSession { Get-WindowsDriver -Online }
                if ($HostDriversList) {
                    Section -Style Heading3 'Drivers' {
                        Invoke-Command -Session $TempPssSession { Import-Module DISM }
                        $HostDriverReport = @()
                        ForEach ($HostDriver in $HostDriversList) {
                            try {
                                $TempDriver = [PSCustomObject] @{
                                    'Class Description' = $HostDriver.ClassDescription
                                    'Provider Name' = $HostDriver.ProviderName
                                    'Driver Version' = $HostDriver.Version
                                    'Version Date' = $HostDriver.Date.ToShortDateString()
                                }
                                $HostDriverReport += $TempDriver
                            } catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
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
                        $HostDriverReport | Sort-Object -Property 'Class Description' | Table @TableParams
                    }
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}