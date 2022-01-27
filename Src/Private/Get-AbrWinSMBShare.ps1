function Get-AbrWinSMBShare {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server File Server Shares information.
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.2.0
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
        Write-PscriboMessage "Collecting File Server Shares information."
    }

    process {
        if ($InfoLevel.SMB -ge 1) {
            try {
                if ($SMBShares) {
                    Section -Style Heading3 'Network Shares' {
                        Paragraph 'The following table details network shares'
                        Blankline
                        $SMBSharesReport = @()
                        foreach ($SMBShare in $SMBShares) {
                            try {
                                $ShareAccess = Invoke-Command -Session $TempPssSession { Get-SmbShareAccess -Name ($using:SMBShare).Name }
                                $TempSMBSharesReport = [PSCustomObject]@{
                                    'Name' = $SMBShare.Name
                                    'Scope Name' = $SMBShare.ScopeName
                                    'Path' = $SMBShare.Path
                                    'Description' =  $SMBShare.Description
                                    'Access Based Enumeration Mode' = $SMBShare.FolderEnumerationMode
                                    'Caching Mode' = $SMBShare.CachingMode
                                    'Encrypt Data' = $SMBShare.EncryptData
                                    'State' = $SMBShare.ShareState
                                    'Share Access' = $ShareAccess.AccountName
                                }
                                $SMBSharesReport = $TempSMBSharesReport

                                $TableParams = @{
                                    Name = "File Server Share - $($SMBShare.Name)"
                                    List = $true
                                    ColumnWidths = 40, 60
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $SMBSharesReport | Table @TableParams
                            }
                            catch {
                                Write-PscriboMessage -IsWarning $_.Exception.Message
                            }
                        }
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