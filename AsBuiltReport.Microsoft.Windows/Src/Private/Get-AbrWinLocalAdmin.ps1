function Get-AbrWinLocalAdmin {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Local Administrator information.
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
        Write-PScriboMessage "Account InfoLevel set at $($InfoLevel.Account)."
        Write-PScriboMessage "Collecting Local Administrator information."
    }

    process {
        if ($InfoLevel.Account -ge 1) {
            try {
                if ($LocalAdmins) {
                    Section -Style Heading3 'Local Administrators' {
                        $OutObj = @()
                        ForEach ($LocalAdmin in $LocalAdmins) {
                            try {
                                $inObj = [ordered] @{
                                    'Account Name' = $LocalAdmin.Name
                                    'Account Type' = $LocalAdmin.ObjectClass
                                    'Account Source' = $LocalAdmin.PrincipalSource
                                }
                                $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
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
                        $OutObj | Sort-Object -Property 'Account Name' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}