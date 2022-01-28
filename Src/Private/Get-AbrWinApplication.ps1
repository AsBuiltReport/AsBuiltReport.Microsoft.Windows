function Get-AbrWinApplication {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Application Inventory information.
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
        Write-PscriboMessage "Collecting Application Inventory information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            try {
                [array]$AddRemove = @()
                $AddRemove += Invoke-Command -Session $TempPssSession { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* }
                $AddRemove += Invoke-Command -Session $TempPssSession { Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* }
                if ($AddRemove) {
                    Section -Style Heading3 'Installed Applications' {
                        Paragraph 'The following settings details applications listed in Add/Remove Programs'
                        Blankline
                        [array]$AddRemoveReport = @()
                        ForEach ($App in $AddRemove) {
                                try {
                                $TempAddRemoveReport = [PSCustomObject]@{
                                    'Application Name' = $App.DisplayName
                                    'Publisher' = $App.Publisher
                                    'Version' = $App.Version
                                    'Install Date' = Switch (($App.InstallDate).count) {
                                        0 {"-"}
                                        default {$App.InstallDate}
                                    }
                                }
                                $AddRemoveReport += $TempAddRemoveReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Installed Applications"
                            List = $false
                            ColumnWidths = 30, 30, 20, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $AddRemoveReport | Where-Object { $_.'Application Name' -notlike $null } | Sort-Object -Property 'Application Name' | Table @TableParams
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