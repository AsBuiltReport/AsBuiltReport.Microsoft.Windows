function Get-AbrWinSQLBuild {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows SQL Server build information.
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
        Write-PScriboMessage "SQL Server InfoLevel set at $($InfoLevel.SQLServer)."
        Write-PscriboMessage "Collecting SQL Server build information."
    }

    process {
        if ($InfoLevel.SQLServer -ge 1) {
            try {
                $Build = Get-DbaBuild -SqlInstance $SQLServer
                if ($Build) {
                    Section -Style Heading3 'Build' {
                        Paragraph 'The following table details sql server build information'
                        Blankline
                        [array]$SQLServerObjt = @()
                        $TempSQLServerObjt = [PSCustomObject]@{
                            'Instance Name' = $Build.SqlInstance
                            'Build' = $Build.Build
                            'Level' = $Build.NameLevel
                            'Service Pack' = $Build.SPLevel
                            'Comulative Update' = ConvertTo-EmptyToFiller $Build.CULevel
                            'KB Level' = $Build.KBLevel
                            'Supported Until' = $Build.SupportedUntil.ToShortDateString()
                            'Warning' = ConvertTo-EmptyToFiller $Build.Warning
                        }
                        $SQLServerObjt += $TempSQLServerObjt

                        $TableParams = @{
                            Name = "SQL Server Build"
                            List = $True
                            ColumnWidths = 40, 60
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $SQLServerObjt | Table @TableParams
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