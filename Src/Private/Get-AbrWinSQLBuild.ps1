function Get-AbrWinSQLBuild {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows SQL Server Properties information.
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
        Write-PScriboMessage "Collecting SQL Server Properties information."
    }

    process {
        if ($InfoLevel.SQLServer -ge 1) {
            try {
                $Properties = Get-DbaInstanceProperty -SqlInstance $SQLServer | ForEach-Object { @{$_.Name = $_.Value } }
                $Build = Get-DbaBuild -SqlInstance $SQLServer -WarningAction SilentlyContinue
                if ($Properties) {
                    Section -Style Heading3 'General Information' {
                        Paragraph 'The following table details sql server Properties information'
                        BlankLine
                        [array]$SQLServerObjt = @()
                        $TempSQLServerObjt = [PSCustomObject]@{
                            'Instance Name' = $Build.SqlInstance
                            'Fully Qualified Net Name' = $Properties.FullyQualifiedNetName
                            'Supported Until' = $Build.SupportedUntil.ToShortDateString()
                            'Edition' = $Properties.Edition
                            'Level' = "Microsoft SQL Server $($Build.NameLevel)"
                            'Build' = $Properties.VersionString
                            'Service Pack' = $Properties.ProductLevel
                            'Comulative Update' = ConvertTo-EmptyToFiller $Build.CULevel
                            'KB Level' = $Build.KBLevel
                            'Case Sensitive' = ConvertTo-TextYN $Properties.IsCaseSensitive
                            'Full Text Installed' = ConvertTo-TextYN $Properties.IsFullTextInstalled
                            'XTP Supported' = ConvertTo-TextYN $Properties.IsXTPSupported
                            'Clustered' = ConvertTo-TextYN $Properties.IsClustered
                            'Single User' = ConvertTo-TextYN $Properties.IsSingleUser
                            'Language' = $Properties.Language
                            'Collation' = $Properties.Collation
                            'Sql CharSet Name' = $Properties.SqlCharSetName
                            'Root Directory' = $Properties.RootDirectory
                            'Master DB Path' = $Properties.MasterDBPath
                            'Master DB Log Path' = $Properties.MasterDBLogPath
                            'Backup Directory' = $Properties.BackupDirectory
                            'Default File' = $Properties.DefaultFile
                            'Default Log' = $Properties.DefaultLog
                            'Login Mode' = $Properties.LoginMode
                            'Mail Profile' = ConvertTo-EmptyToFiller $Properties.MailProfile
                            'Warning' = ConvertTo-EmptyToFiller $Build.Warning
                        }
                        $SQLServerObjt += $TempSQLServerObjt

                        $TableParams = @{
                            Name = "General Information"
                            List = $True
                            ColumnWidths = 40, 60
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $SQLServerObjt | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}