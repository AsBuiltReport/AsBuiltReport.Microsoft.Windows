function Get-AbrWinOSRoleFeature {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Role & Features information.
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
        Write-PScriboMessage "Collecting Role & Features information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1 -and $OSType.Value -ne 'WorkStation') {
            try {
                $HostRolesAndFeatures = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-WindowsFeature | Where-Object { $_.Installed -eq $True } }
                if ($HostRolesAndFeatures) {
                    Section -Style Heading3 'Roles' {
                        $OutObj = @()
                        ForEach ($HostRoleAndFeature in $HostRolesAndFeatures) {
                            if ( $HostRoleAndFeature.FeatureType -eq 'Role') {
                                try {
                                    $inObj = [ordered] @{
                                        'Name' = $HostRoleAndFeature.DisplayName
                                        'Type' = $HostRoleAndFeature.FeatureType
                                        'Description' = $HostRoleAndFeature.Description
                                    }
                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                } catch {
                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                }
                            }
                        }
                        $TableParams = @{
                            Name = "Roles"
                            List = $false
                            ColumnWidths = 20, 10, 70
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Sort-Object -Property 'Name' | Table @TableParams
                        if ($InfoLevel.OperatingSystem -ge 2) {
                            try {
                                if ($HostRolesAndFeatures) {
                                    Section -Style Heading3 'Features and Role Services' {
                                        $OutObj = @()
                                        ForEach ($HostRoleAndFeature in $HostRolesAndFeatures) {
                                            if ( $HostRoleAndFeature.FeatureType -eq 'Role Service' -or $HostRoleAndFeature.FeatureType -eq 'Feature') {
                                                try {
                                                    $inObj = [ordered] @{
                                                        'Name' = $HostRoleAndFeature.DisplayName
                                                        'Type' = $HostRoleAndFeature.FeatureType
                                                        'Description' = $HostRoleAndFeature.Description
                                                    }
                                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                                } catch {
                                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                                }
                                            }
                                        }
                                        $TableParams = @{
                                            Name = "Feature & Role Services"
                                            List = $false
                                            ColumnWidths = 20, 10, 70
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $OutObj | Sort-Object -Property 'Name' | Table @TableParams
                                    }
                                }
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}