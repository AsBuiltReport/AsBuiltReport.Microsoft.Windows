function Get-AbrWinSMBShare {
    <#
    .SYNOPSIS
    Used by As Built Report to retrieve Windows Server File Server Shares information.
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
        Write-PScriboMessage "Collecting File Server Shares information."
    }

    process {
        if ($InfoLevel.SMB -ge 1) {
            try {
                if ($SMBShares) {
                    Section -Style Heading3 'File Shares' {
                        Paragraph 'The following section details network shares'
                        BlankLine
                        $OutObj = @()
                        foreach ($SMBShare in $SMBShares) {
                            Section -Style Heading4 "$($SMBShare.Name) Share" {
                                Paragraph "The following table details shares configuration"
                                BlankLine
                                try {
                                    $ShareAccess = Invoke-Command -Session $TempPssSession { Get-SmbShareAccess -Name ($using:SMBShare).Name }
                                    $inObj = [ordered] @{
                                        'Name' = $SMBShare.Name
                                        'Scope Name' = $SMBShare.ScopeName
                                        'Path' = $SMBShare.Path
                                        'Description' = $SMBShare.Description
                                        'Access Based Enumeration Mode' = $SMBShare.FolderEnumerationMode
                                        'Caching Mode' = $SMBShare.CachingMode
                                        'Encrypt Data' = $SMBShare.EncryptData
                                        'State' = $SMBShare.ShareState
                                    }
                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                                    $TableParams = @{
                                        Name = "File Server Shares - $($SMBShare.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $OutObj | Table @TableParams
                                } catch {
                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                }
                                try {
                                    $ShareAccess = Invoke-Command -Session $TempPssSession { Get-SmbShareAccess -Name ($using:SMBShare).Name }
                                    if ($ShareAccess) {
                                        Section -Style Heading5 'Permissions' {
                                            Paragraph "The following table details $($SMBShare.Name) shares permissions"
                                            BlankLine
                                            $OutObj = @()
                                            foreach ($SMBACL in $ShareAccess) {
                                                try {
                                                    $inObj = [ordered] @{
                                                        'Scope Name' = $SMBACL.ScopeName
                                                        'Account Name' = $SMBACL.AccountName
                                                        'Access Control Type' = $SMBACL.AccessControlType
                                                        'Access Right' = $SMBACL.AccessRight
                                                    }
                                                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)

                                                } catch {
                                                    Write-PScriboMessage -IsWarning $_.Exception.Message
                                                }
                                            }

                                            $TableParams = @{
                                                Name = "Share Permissions - $($SMBShare.Name)"
                                                List = $false
                                                ColumnWidths = 25, 25, 25, 25
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
                    }
                }
            } catch {
                Write-PScriboMessage -IsWarning $_.Exception.Message
            }
        }
    }
    end {}
}