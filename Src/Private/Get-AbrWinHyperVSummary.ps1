function Get-AbrWinHyperVSummary {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Hyper-V Summary information.
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
        Write-PScriboMessage "Hyper-V InfoLevel set at $($InfoLevel.HyperV)."
        Write-PScriboMessage "Collecting Hyper-V Summary information."
    }

    process {
        if ($InfoLevel.HyperV -ge 1) {
            try {
                $script:VmHost = Invoke-Command -Session $TempPssSession { Get-VMHost }
                if ($VmHost) {
                    $OutObj = @()
                    $inObj = [ordered] @{
                        'Logical Processor Count' = $VmHost.LogicalProcessorCount
                        'Memory Capacity' = "$([Math]::Round($VmHost.MemoryCapacity / 1gb)) GB"
                        'VM Default Path' = $VmHost.VirtualMachinePath
                        'VM Disk Default Path' = $VmHost.VirtualHardDiskPath
                        'Supported VM Versions' = $VmHost.SupportedVmVersions -Join ","
                        'Numa Spannning Enabled' = $VmHost.NumaSpanningEnabled
                        'Iov Support' = $VmHost.IovSupport
                        'VM Migrations Enabled' = $VmHost.VirtualMachineMigrationEnabled
                        'Allow any network for Migrations' = $VmHost.UseAnyNetworkForMigration
                        'VM Migration Authentication Type' = $VmHost.VirtualMachineMigrationAuthenticationType
                        'Max Concurrent Storage Migrations' = $VmHost.MaximumStorageMigrations
                        'Max Concurrent VM Migrations' = $VmHost.MaximumStorageMigrations
                    }
                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)

                    $TableParams = @{
                        Name = "Hyper-V Host Settings"
                        List = $true
                        ColumnWidths = 50, 50
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Table @TableParams
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}