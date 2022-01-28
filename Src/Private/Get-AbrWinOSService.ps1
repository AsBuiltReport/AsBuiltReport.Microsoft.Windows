function Get-AbrWinOSService {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Operating System Services information.
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
        Write-PScriboMessage "Operating System InfoLevel set at $($InfoLevel.OperatingSystem)."
        Write-PscriboMessage "Collecting Operating System Service information."
    }

    process {
        if ($InfoLevel.OperatingSystem -ge 1) {
            try {
                $Available = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-Service "W32Time" | Select-Object DisplayName, Name, Status}
                if ($Available) {
                    Section -Style Heading3 'Services' {
                        Paragraph 'The following table details status of important services'
                        Blankline
                        $Services = @('DNS','DFS Replication','Intersite Messaging','Kerberos Key Distribution Center','Active Directory Domain Services','W32Time','ADWS''Dhcp','Dnscache','gpsvc','HvHost','vmcompute','vmms','iphlpsvc','MSiSCSI','Netlogon','RasMan','SessionEnv','TermService','RpcSs','RpcEptMapper','SamSs','LanmanServer','Schedule','lmhosts','UsoSvc','mpssvc','W3SVC','MSSQLSERVER')
                        $ServicesReport = @()
                        Foreach ($Service in $Services) {
                            try {
                                $Status = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-Service $using:Service -ErrorAction SilentlyContinue | Select-Object DisplayName, Name, Status}
                                if ($Status) {
                                    $TempServicesReport = [PSCustomObject] @{
                                        'Display Name' = $Status.DisplayName
                                        'Short Name' = $Status.Name
                                        'Status' = $Status.Status
                                    }
                                    $ServicesReport += $TempServicesReport
                                }
                            } catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }

                        if ($HealthCheck.OperatingSystem.Services) {
                            $ServicesReport | Where-Object { $_.'Status' -notlike 'Running'} | Set-Style -Style Warning -Property 'Status'
                        }

                        $TableParams = @{
                            Name = "Services Status"
                            List = $false
                            ColumnWidths = 50, 25, 25
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $ServicesReport | Sort-Object -Property 'Display Name' | Table @TableParams
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