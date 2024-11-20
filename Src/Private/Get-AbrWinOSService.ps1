function Get-AbrWinOSService {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Operating System Services information.
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
        Write-PScriboMessage "Operating System InfoLevel set at $($InfoLevel.OperatingSystem)."
        Write-PScriboMessage "Collecting Operating System Service information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            try {
                $Available = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service "W32Time" | Select-Object DisplayName, Name, Status }
                if ($Available) {
                    Section -Style Heading3 'Services' {
                        Paragraph 'The following table details status of important services'
                        BlankLine
                        $Services = @('DNS', 'DFS Replication', 'Intersite Messaging', 'Kerberos Key Distribution Center', 'Active Directory Domain Services', 'W32Time', 'ADWS', 'DHCPServer', 'Dnscache', 'gpsvc', 'HvHost', 'vmcompute', 'vmms', 'iphlpsvc', 'MSiSCSI', 'Netlogon', 'RasMan', 'SessionEnv', 'TermService', 'RpcSs', 'RpcEptMapper', 'SamSs', 'LanmanServer', 'Schedule', 'lmhosts', 'UsoSvc', 'mpssvc', 'W3SVC', 'MSSQLSERVER', 'ClusSvc')
                        $OutObj = @()
                        Foreach ($Service in $Services) {
                            try {
                                $Status = Invoke-Command -Session $TempPssSession -ScriptBlock { Get-Service $using:Service -ErrorAction SilentlyContinue | Select-Object DisplayName, Name, Status }
                                if ($Status) {
                                    $inObj = [ordered] @{
                                        'Display Name' = $Status.DisplayName
                                        'Short Name' = $Status.Name
                                        'Status' = $Status.Status
                                    }
                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                }
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
                            }
                        }

                        if ($HealthCheck.OperatingSystem.Services) {
                            $OutObj | Where-Object { $_.'Status' -notlike 'Running' } | Set-Style -Style Warning -Property 'Status'
                        }

                        $TableParams = @{
                            Name = "Services Status"
                            List = $false
                            ColumnWidths = 50, 25, 25
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Sort-Object -Property 'Display Name' | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}