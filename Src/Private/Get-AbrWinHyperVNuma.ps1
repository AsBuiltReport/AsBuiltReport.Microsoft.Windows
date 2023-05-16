function Get-AbrWinHyperVNuma {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server Hyper-V Numa information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
        Author:         Andrew Ramsay
        Editor:         Jonathan Colon
        Twitter:        @jcolonfzenpr
        Github:         rebelinux
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows
    #>
    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "Hyper-V InfoLevel set at $($InfoLevel.HyperV)."
        Write-PscriboMessage "Collecting Hyper-V Numa information."
    }

    process {
        if ($InfoLevel.HyperV -ge 1) {
            try {
                $VmHostNumaNodes = Invoke-Command -Session $TempPssSession { Get-VMHostNumaNode }
                if ($VmHostNumaNodes) {
                    Section -Style Heading3 "Hyper-V NUMA Boundaries" {
                        [array]$VmHostNumaReport = @()
                        foreach ($Node in $VmHostNumaNodes) {
                            try {
                                $TempVmHostNumaReport = [PSCustomObject]@{
                                    'Numa Node Id' = $Node.NodeId
                                    'Memory Available' = "$([math]::Round(($Node.MemoryAvailable)/1024,0)) GB"
                                    'Memory Total' =  "$([math]::Round(($Node.MemoryTotal)/1024,0)) GB"
                                }
                                $VmHostNumaReport += $TempVmHostNumaReport
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
                        $TableParams = @{
                            Name = "Host NUMA Nodes"
                            List = $false
                            ColumnWidths = 34, 33, 33
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $VmHostNumaReport | Sort-Object -Property 'Numa Node Id' | Table @TableParams
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