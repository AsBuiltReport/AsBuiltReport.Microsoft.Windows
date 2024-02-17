function ConvertTo-TextYN {
    <#
    .SYNOPSIS
    Used by As Built Report to convert true or false automatically to Yes or No.
    .DESCRIPTION

    .NOTES
        Version:        0.4.0
        Author:         LEE DAILEY

    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [AllowEmptyString()]
        [string]
        $TEXT
    )

    switch ($TEXT) {
        "" { "--"; break }
        $Null { "--"; break }
        "True" { "Yes"; break }
        "False" { "No"; break }
        default { $TEXT }
    }
} # end

function ConvertTo-FileSizeString {
    <#
    .SYNOPSIS
    Used by As Built Report to convert bytes automatically to GB or TB based on size.
    .DESCRIPTION

    .NOTES
        Version:        0.4.0
        Author:         LEE DAILEY

    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [int64]
        $Size
    )

    switch ($Size) {
        { $_ -gt 1TB }
        { [string]::Format("{0:0.00} TB", $Size / 1TB); break }
        { $_ -gt 1GB }
        { [string]::Format("{0:0.00} GB", $Size / 1GB); break }
        { $_ -gt 1MB }
        { [string]::Format("{0:0.00} MB", $Size / 1MB); break }
        { $_ -gt 1KB }
        { [string]::Format("{0:0.00} KB", $Size / 1KB); break }
        { $_ -gt 0 }
        { [string]::Format("{0} B", $Size); break }
        { $_ -eq 0 }
        { "0 KB"; break }
        default
        { "0 KB" }
    }
} # end >> function Format-FileSize

function ConvertTo-EmptyToFiller {
    <#
    .SYNOPSIS
    Used by As Built Report to convert empty culumns to "--".
    .DESCRIPTION
    .NOTES
        Version:        0.5.0
        Author:         Jonathan Colon
    .EXAMPLE
    .LINK
    #>
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [AllowEmptyString()]
        [string]$TEXT
    )
    
    switch ([string]::IsNullOrEmpty($TEXT)) {
        $true { "--"; break }
        default { $TEXT }
    }
}

function Convert-IpAddressToMaskLength {
    <#
    .SYNOPSIS
    Used by As Built Report to convert subnet mask to dotted notation.
    .DESCRIPTION

    .NOTES
        Version:        0.4.0
        Author:         Ronald Rink

    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [string]
        $SubnetMask
    )

    [IPAddress] $MASK = $SubnetMask
    $octets = $MASK.IPAddressToString.Split('.')
    foreach ($octet in $octets) {
        while (0 -ne $octet) {
            $octet = ($octet -shl 1) -band [byte]::MaxValue
            $result++;
        }
    }
    return $result;
}

function ConvertTo-ADObjectName {
    <#
    .SYNOPSIS
    Used by As Built Report to translate Active Directory DN to Name.
    .DESCRIPTION

    .NOTES
        Version:        0.4.0
        Author:         Jonathan Colon

    .EXAMPLE

    .LINK

    #>
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $DN,
        $Session
    )
    $ADObject = @()
    foreach ($Object in $DN) {
        $ADObject += Invoke-Command -Session $Session { Get-ADObject $using:Object | Select-Object -ExpandProperty Name }
    }
    return $ADObject;
}# end
