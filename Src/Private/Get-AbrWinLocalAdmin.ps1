function Get-AbrWinLocalAdmin {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Local Administrator information.
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
        Write-PScriboMessage "Account InfoLevel set at $($InfoLevel.Account)."
        Write-PScriboMessage "Collecting Local Administrator information."
    }

    process {
        if ($InfoLevel.Account -ge 1) {
            try {
                if ($LocalAdmins) {
                    Section -Style Heading3 'Local Administrators' {
                        $LocalAdminsReport = @()
                        ForEach ($LocalAdmin in $LocalAdmins) {
                            try {
                                $TempLocalAdminsReport = [PSCustomObject]@{
                                    'Account Name' = $LocalAdmin.Name
                                    'Account Type' = $LocalAdmin.ObjectClass
                                    'Account Source' = ConvertTo-EmptyToFiller $LocalAdmin.PrincipalSource
                                }
                                $LocalAdminsReport += $TempLocalAdminsReport
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Local Administrators"
                            List = $false
                            ColumnWidths = 40, 30, 30
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $LocalAdminsReport | Sort-Object -Property 'Account Name' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}