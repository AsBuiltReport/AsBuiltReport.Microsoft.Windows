function Get-AbrWinHyperVSummary {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Hyper-V Summary information.
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
        Write-PScriboMessage "Hyper-V InfoLevel set at $($InfoLevel.HyperV)."
        Write-PScriboMessage "Collecting Hyper-V Summary information."
    }

    process {
        if ($InfoLevel.HyperV -ge 1) {
            try {
                $script:VmHost = Invoke-Command -Session $TempPssSession { Get-VMHost }
                if ($VmHost) {
                    $VmHostReport = [PSCustomObject]@{
                        'Logical Processor Count' = $VmHost.LogicalProcessorCount
                        'Memory Capacity' = "$([Math]::Round($VmHost.MemoryCapacity / 1gb)) GB"
                        'VM Default Path' = $VmHost.VirtualMachinePath
                        'VM Disk Default Path' = $VmHost.VirtualHardDiskPath
                        'Supported VM Versions' = $VmHost.SupportedVmVersions -Join ","
                        'Numa Spannning Enabled' = ConvertTo-TextYN $VmHost.NumaSpanningEnabled
                        'Iov Support' = ConvertTo-TextYN $VmHost.IovSupport
                        'VM Migrations Enabled' = ConvertTo-TextYN $VmHost.VirtualMachineMigrationEnabled
                        'Allow any network for Migrations' = ConvertTo-TextYN $VmHost.UseAnyNetworkForMigration
                        'VM Migration Authentication Type' = $VmHost.VirtualMachineMigrationAuthenticationType
                        'Max Concurrent Storage Migrations' = $VmHost.MaximumStorageMigrations
                        'Max Concurrent VM Migrations' = $VmHost.MaximumStorageMigrations
                    }
                    $TableParams = @{
                        Name = "Hyper-V Host Settings"
                        List = $true
                        ColumnWidths = 50, 50
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $VmHostReport | Table @TableParams
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}