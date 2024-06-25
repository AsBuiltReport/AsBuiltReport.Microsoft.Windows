function Get-AbrWinIISWebSite {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server IIS Sites information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.5
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
        Write-PScriboMessage "Collecting IIS Sites information."
    }

    process {
        if ($InfoLevel.IIS -ge 1) {
            try {
                $IISWebSites = Invoke-Command -Session $TempPssSession { Get-Website }
                if ($IISWebSites) {
                    Section -Style Heading3 'Sites Summary' {
                        Paragraph 'The following table provide a summary of IIS Web Sites'
                        BlankLine
                        $IISWebSitesrReport = @()
                        foreach ($IISWebSite in $IISWebSites) {
                            try {
                                $TempIISWebSitesrReport = [PSCustomObject]@{
                                    'Name' = $IISWebSite.Name
                                    'Status' = $IISWebSite.state
                                    'Binding' = $IISWebSite.bindings.Collection
                                    'Application Pool' = $IISWebSite.applicationPool
                                }
                                $IISWebSitesrReport += $TempIISWebSitesrReport
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
                            }
                        }

                        $TableParams = @{
                            Name = "IIS Web Sites"
                            List = $false
                            ColumnWidths = 25, 25, 25, 25
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $IISWebSitesrReport | Table @TableParams
                        try {
                            $IISWebSites = Invoke-Command -Session $TempPssSession { Get-Website }
                            if ($IISWebSites) {
                                Section -Style Heading4 'Sites Configuration' {
                                    Paragraph 'The following section details IIS Web Sites configuration'
                                    BlankLine
                                    $IISWebSitesrReport = @()
                                    foreach ($IISWebSite in $IISWebSites) {
                                        try {
                                            Section -Style Heading5 "$($IISWebSite.Name)" {
                                                Paragraph "The following table details $($IISWebSite.Name) settings"
                                                BlankLine
                                                $SiteURL = Invoke-Command -Session $TempPssSession { Get-WebURL -PSPath "IIS:\Sites\$(($using:IISWebSite).Name)" }
                                                $TempIISWebSitesrReport = [PSCustomObject]@{
                                                    'Name' = $IISWebSite.Name
                                                    'Auto Start' = ConvertTo-TextYN $IISWebSite.serverAutoStart
                                                    'Enabled Protocols ' = $IISWebSite.enabledProtocols
                                                    'URL' = Switch (($SiteURL.ResponseUri).count) {
                                                        0 { "--" }
                                                        default { $SiteURL.ResponseUri }
                                                    }
                                                    'Path ' = $IISWebSite.physicalPath
                                                    'Log Path' = $IISWebSite.logFile.directory

                                                }
                                                $IISWebSitesrReport = $TempIISWebSitesrReport

                                                $TableParams = @{
                                                    Name = "IIS Web Sites - $($IISWebSite.Name)"
                                                    List = $true
                                                    ColumnWidths = 40, 60
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $IISWebSitesrReport | Table @TableParams
                                                try {
                                                    $IISWebApps = Invoke-Command -Session $TempPssSession { Get-WebApplication -Site $(($using:IISWebSite).Name) }
                                                    if ($IISWebApps) {
                                                        Section -Style Heading5 "Web Applications" {
                                                            Paragraph "The following table details $($IISWebSite.Name) Web Application"
                                                            BlankLine
                                                            $IISWebAppsReport = @()
                                                            foreach ($IISWebApp in $IISWebApps) {
                                                                try {
                                                                    $TempIISWebAppsReport = [PSCustomObject]@{
                                                                        'Name' = $IISWebApp.Path
                                                                        'Application pool' = $IISWebApp.Applicationpool
                                                                        'Physical Path ' = $IISWebApp.PhysicalPath
                                                                    }
                                                                    $IISWebAppsReport += $TempIISWebAppsReport
                                                                } catch {
                                                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                                                }
                                                            }

                                                            $TableParams = @{
                                                                Name = "Web Applications - $($IISWebSite.Name)"
                                                                List = $false
                                                                ColumnWidths = 35, 20, 45
                                                            }
                                                            if ($Report.ShowTableCaptions) {
                                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                                            }
                                                            $IISWebAppsReport | Table @TableParams
                                                        }
                                                    }
                                                } catch {
                                                    Write-PScriboMessage -IsWarning $_.Exception.Message
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
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}