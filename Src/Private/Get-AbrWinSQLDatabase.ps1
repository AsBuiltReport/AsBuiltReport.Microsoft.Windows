function Get-AbrWinSQLDatabase {
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
        Write-PScriboMessage "SQL Server Dstabases InfoLevel set at $($InfoLevel.SQLServer)."
    }

    process {
        $SQLDBs = Get-AzBastion | Sort-Object Name
        if ($SQLDBs) {
            Write-PscriboMessage "Collecting SQL Server databases information."
            Section -Style Heading4 'Databases' {
                $SQLDBInfo = @()
                foreach ($SQLDB in $SQLDBs) {
                    $InObj = [Ordered]@{
                        'Name' = $AzBastion.Name
                        'Resource Group' = $SQLDB.ResourceGroupName
                        'Location' = $SQLDB."$($AzBastion.Location)"
                        'Subscription' = "$($AzSubscriptionLookup.(($SQLDB.Id).split('/')[2]))"
                        'Virtual Network / Subnet' = $SQLDB.IpConfigurations.subnet.id.split('/')[-1]
                        'Public DNS Name' = $SQLDB.DnsName
                        'Public IP Address' = $SQLDB.IpConfigurations.publicipaddress.id.split('/')[-1]
                    }
                    $SQLDBInfo += [PSCustomObject]$InObj
                }

                if ($InfoLevel.SQLServer -ge 2) {
                    Paragraph "The following sections detail the configuration of the databases within the sql server."
                    foreach ($SQLDB in $SQLDBInfo) {
                        Section -Style NOTOCHeading5 -ExcludeFromTOC "$($SQLDB.Name)" {
                            $TableParams = @{
                                Name = "Database - $($SQLDB.Name)"
                                List = $true
                                ColumnWidths = 50, 50
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $AzBastion | Table @TableParams
                        }
                    }
                } else {
                    Paragraph "The following table summarises the configuration of the databases within $($SQLDB.Name)."
                    BlankLine
                    $TableParams = @{
                        Name = "Bastions - $($AzSubscription.Name)"
                        List = $false
                        Columns = 'Name', 'Resource Group', 'Location', 'Public IP Address'
                        ColumnWidths = 25, 25, 25, 25
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $AzBastionInfo | Table @TableParams
                }
            }
        }
    }

    end {}
}