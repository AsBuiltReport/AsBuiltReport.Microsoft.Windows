function Get-AbrWinSMBNetworkInterface {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server File Server NIC information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.6
        Author:         Jonathan Colon
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
        Write-PScriboMessage "SMB InfoLevel set at $($InfoLevel.SMB)."
        Write-PScriboMessage "Collecting File Server Network Interface information."
    }

    process {
        if ($InfoLevel.SMB -ge 1) {
            try {
                $SMBNICs = Invoke-Command -Session $TempPssSession { Get-SmbServerNetworkInterface }
                if ($SMBNICs) {
                    Section -Style Heading3 "SMB Network Interface" {
                        Paragraph "The following table provide a summary of the SMB protocol network interface information"
                        BlankLine
                        $OutObj = @()
                        foreach ($SMBNIC in $SMBNICs) {
                            try {
                                $inObj = [ordered] @{
                                    'Name' = Switch (($SMBNIC.InterfaceIndex).count) {
                                        0 { "Unknown" }
                                        default { Invoke-Command -Session $TempPssSession { (Get-NetAdapter -InterfaceIndex ($using:SMBNIC).InterfaceIndex).Name } }
                                    }
                                    'RSS Capable' = $SMBNIC.RssCapable
                                    'RDMA Capable' = $SMBNIC.RdmaCapable
                                    'IP Address' = $SMBNIC.IpAddress
                                }
                                $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                            } catch {
                                Write-PScriboMessage -IsWarning $_.Exception.Message
                            }
                        }

                        $TableParams = @{
                            Name = "SMB Network Interfaces"
                            List = $false
                            ColumnWidths = 34, 16, 16, 34
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Table @TableParams
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}