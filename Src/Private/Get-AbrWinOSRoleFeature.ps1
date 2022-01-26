function Get-AbrWinOSRoleFeature {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Role & Features information.
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
        Write-PscriboMessage "Collecting Role & Features information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            try {
                $HostRolesAndFeatures = Invoke-Command -Session $TempPssSession -ScriptBlock{Get-WindowsFeature | Where-Object { $_.Installed -eq $True }}
                if ($HostRolesAndFeatures) {
                    Section -Style Heading3 'Roles and Features' {
                        Paragraph 'The following settings details host roles and features installed'
                        Blankline
                        [array]$HostRolesAndFeaturesReport = @()
                        ForEach ($HostRoleAndFeature in $HostRolesAndFeatures) {
                            try {
                                $TempHostRolesAndFeaturesReport = [PSCustomObject] @{
                                    'Feature Name' = $HostRoleAndFeature.DisplayName
                                    'Feature Type' = $HostRoleAndFeature.FeatureType
                                    'Description' = $HostRoleAndFeature.Description
                                }
                                $HostRolesAndFeaturesReport += $TempHostRolesAndFeaturesReport
                            } catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Roles and Features"
                            List = $false
                            ColumnWidths = 20, 10, 70
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $HostRolesAndFeaturesReport | Sort-Object -Property 'Feature Name' | Table @TableParams
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