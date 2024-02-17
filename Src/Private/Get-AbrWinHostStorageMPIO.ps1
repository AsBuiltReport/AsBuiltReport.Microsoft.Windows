function Get-AbrWinHostStorageMPIO {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Host Storage MPIO information.
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
        Write-PScriboMessage "Storage InfoLevel set at $($InfoLevel.Storage)."
        Write-PScriboMessage "Collecting Host Storage MPIO information."
    }

    process {
        if ($InfoLevel.Storage -ge 1 -and $OSType.Value -ne 'WorkStation') {
            try {
                $MPIOInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "Multipath*" } }
                if ($MPIOInstalledCheck.InstallState -eq "Installed") {
                    try {
                        Section -Style Heading3 'Host MPIO Settings' {
                            Paragraph 'The following section details host MPIO Settings'
                            BlankLine
                            [string]$MpioLoadBalance = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-MSDSMGlobalDefaultLoadBalancePolicy }
                            Paragraph "The default load balancing policy is: $MpioLoadBalance"
                            BlankLine
                            $MpioAutoClaim = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-MSDSMAutomaticClaimSettings | Select-Object -ExpandProperty Keys }
                            if ($MpioAutoClaim) {
                                Section -Style Heading4 'Multipath I/O AutoClaim' {
                                    Paragraph 'The following table details the BUS types MPIO will automatically claim for'
                                    BlankLine
                                    $MpioAutoClaimReport = @()
                                    foreach ($key in $MpioAutoClaim) {
                                        try {
                                            $Temp = "" | Select-Object BusType, State
                                            $Temp.BusType = $key
                                            $Temp.State = 'Enabled'
                                            $MpioAutoClaimReport += $Temp
                                        } catch {
                                            Write-PScriboMessage -IsWarning $_.Exception.Message
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "Multipath I/O Auto Claim Settings"
                                        List = $false
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $MpioAutoClaimReport | Sort-Object -Property 'BusType' | Table @TableParams
                                }
                            }
                            try {
                                $MpioAvailableHws = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-MPIOAvailableHw }
                                if ($MpioAvailableHws) {
                                    Section -Style Heading4 'MPIO Detected Hardware' {
                                        Paragraph 'The following table details the hardware detected and claimed by MPIO'
                                        BlankLine
                                        $MpioAvailableHwReport = @()
                                        foreach ($MpioAvailableHw in $MpioAvailableHws) {
                                            try {
                                                $TempMpioAvailableHwReport = [PSCustomObject]@{
                                                    'Vendor' = $MpioAvailableHw.VendorId
                                                    'Product' = $MpioAvailableHw.ProductId
                                                    'BusType' = $MpioAvailableHw.BusType
                                                    'Multipathed' = ConvertTo-TextYN $MpioAvailableHw.IsMultipathed
                                                }
                                                $MpioAvailableHwReport += $TempMpioAvailableHwReport
                                            } catch {
                                                Write-PScriboMessage -IsWarning $_.Exception.Message
                                            }
                                        }
                                        $TableParams = @{
                                            Name = "MPIO Available Hardware"
                                            List = $false
                                            ColumnWidths = 25, 25, 25, 25
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $MpioAvailableHwReport | Table @TableParams
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
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}