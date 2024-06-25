function Get-AbrWinOSConfig {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Operating System Configuration information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.2
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
        Write-PScriboMessage "Collecting Oprating System Configuration information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            Section -Style Heading3 'OS Configuration' {
                Paragraph 'The following section details host OS configuration'
                BlankLine
                $HostOSReport = [PSCustomObject] @{
                    'Windows Product Name' = $HostInfo.WindowsProductName
                    'Windows Version' = $HostInfo.WindowsCurrentVersion
                    'Windows Build Number' = $HostInfo.OsVersion
                    'Windows Install Type' = $HostInfo.WindowsInstallationType
                    'AD Domain' = $HostInfo.CsDomain
                    'Windows Installation Date' = switch (($HostInfo.OsInstallDate).count) {
                        0 { "--" }
                        default { $HostInfo.OsInstallDate.ToShortDateString() }
                    }
                    'Time Zone' = $HostInfo.TimeZone
                    'License Type' = Switch ([string]::IsNullOrEmpty($HostLicense.ProductKeyChannel)) {
                        $true { "--" }
                        $false { $HostLicense.ProductKeyChannel }
                        default { "Unknown" }
                    }
                    'Partial Product Key' = Switch ([string]::IsNullOrEmpty($HostLicense.PartialProductKey)) {
                        $true { "--" }
                        $false { $HostLicense.PartialProductKey }
                        default { "Unknown" }
                    }
                }
                $TableParams = @{
                    Name = "OS Settings"
                    List = $true
                    ColumnWidths = 50, 50
                }
                if ($Report.ShowTableCaptions) {
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                }
                $HostOSReport | Table @TableParams
            }
        }
    }
    end {}
}
