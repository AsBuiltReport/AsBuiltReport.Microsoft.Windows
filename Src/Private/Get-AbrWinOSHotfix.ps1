function Get-AbrWinOSHotfix {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Operating System HotFix information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
        Author:         Andrew Ramsay
        Editor:         Jonathan Colon
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
        Write-PScriboMessage "Operating System InfoLevel set at $($InfoLevel.OperatingSystem)."
        Write-PscriboMessage "Collecting Operating System HotFix information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            try {
                $HotFixes = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-HotFix }
                if ($HotFixes) {
                    Section -Style Heading3 'Installed Hotfixes' {
                        Paragraph 'The following table details the OS Hotfixes installed'
                        Blankline
                        $HotFixReport = @()
                        Foreach ($HotFix in $HotFixes) {
                            try {
                                $TempHotFix = [PSCustomObject] @{
                                    'Hotfix ID' = $HotFix.HotFixID
                                    'Description' = $HotFix.Description
                                    'Installation Date' = $HotFix.InstalledOn.ToShortDateString()
                                }
                                $HotfixReport += $TempHotFix
                            } catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Installed Hotfixes"
                            List = $false
                            ColumnWidths = 34, 33, 33
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $HotFixReport | Sort-Object -Property 'Hotfix ID' | Table @TableParams
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