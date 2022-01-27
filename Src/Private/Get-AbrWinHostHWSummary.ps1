function Get-AbrWinHostHWSummary {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Hardware Inventory information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
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
        Write-PScriboMessage "Hardware InfoLevel set at $($InfoLevel.Hardware)."
        Write-PscriboMessage "Collecting Host Inventory information."
    }

    process {
        if ($InfoLevel.Hardware -ge 1) {
            try {
                Section -Style Heading2 'Host Hardware Settings' {
                    Paragraph 'The following section details hardware settings for the host'
                    $HostHardware = [PSCustomObject] @{
                        'Manufacturer' = $HostComputer.Manufacturer
                        'Model' = $HostComputer.Model
                        'Product ID' = $HostComputer.SystemSKUNumbe
                        'Serial Number' = $HostBIOS.SerialNumber
                        'BIOS Version' = $HostBIOS.Version
                        'Processor Manufacturer' = $HostCPU[0].Manufacturer
                        'Processor Model' = $HostCPU[0].Name
                        'Number of Processors' = $HostCPU.Length
                        'Number of CPU Cores' = $HostCPU[0].NumberOfCores
                        'Number of Logical Cores' = $HostCPU[0].NumberOfLogicalProcessors
                        'Physical Memory (GB)' = [Math]::Round($HostComputer.TotalPhysicalMemory / 1Gb)
                    }
                    $TableParams = @{
                        Name = "Host Hardware Specifications"
                        List = $true
                        ColumnWidths = 50, 50
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $HostHardware | Table @TableParams
                }
            }
            catch {
                Write-PscriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}