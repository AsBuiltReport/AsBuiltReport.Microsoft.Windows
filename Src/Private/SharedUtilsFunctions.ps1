function ConvertTo-TextYN {
    <#
    .SYNOPSIS
        Used by As Built Report to convert true or false automatically to Yes or No.
    .DESCRIPTION

    .NOTES
        Version:        0.3.0
        Author:         LEE DAILEY

    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    [OutputType([String])]
    param (
        [Parameter (
            Position = 0,
            Mandatory)]
        [AllowEmptyString()]
        [string] $TEXT
    )

    switch ($TEXT) {
        "" { "--"; break }
        " " { "--"; break }
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
        Version:        0.1.0
        Author:         Jonathan Colon
    .EXAMPLE
    .LINK
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [int64]
        $Size
    )

    $Unit = switch ($Size) {
        { $Size -gt 1PB } { 'PB' ; break }
        { $Size -gt 1TB } { 'TB' ; break }
        { $Size -gt 1GB } { 'GB' ; break }
        { $Size -gt 1Mb } { 'MB' ; break }
        Default { 'KB' }
    }
    return "$([math]::Round(($Size / $("1" + $Unit)), 0)) $Unit"
} # end

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
    param
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
    param
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

function Get-LocalGroupMembership {
    <#
    .SYNOPSIS
        Recursively list all members of a specified Local group.

    .DESCRIPTION
        Recursively list all members of a specified Local group. This can be run against a local or
        remote system or systems. Recursion is unlimited unless specified by the -Depth parameter.

        Alias: glgm

    .NOTES
        Version:        0.5.4
        Author:         Boe Prox (Updated by Graham Flynn (26/07/2024)
        Changes:        Updated to ouput PrincipalSource and ObjectClass so output is similar to Microsoft's Get-LocalGroupMember for compatibility

    .PARAMETER Computername
        Local or remote computer/s to perform the query against.
        Default value is the local system.

    .PARAMETER Group
        Name of the group to query on a system for all members.
        Default value is 'Administrators'

    .PARAMETER Depth
        Limit the recursive depth of a query.
        Default value is 2147483647.

    .PARAMETER Throttle
        Number of concurrently running jobs to run at a time
        Default value is 10

    .EXAMPLE
        Get-LocalGroupMembership

        Name              ParentGroup       isGroup     ObjectClass PrincipalSource   Computername Depth
        ----              -----------       -------     ----        ------------      -----
        Administrator     Administrators      False     User        Local             DC1              1
        boe               Administrators      False     User        Domain            DC1              1
        testuser          Administrators      False     User        Local             DC1              1
        bob               Administrators      False     User        Domain            DC1              1
        proxb             Administrators      False     User        Domain            DC1              1
        Enterprise Admins Administrators      True      Group       Domain            DC1              1
        Sysops Admins     Enterprise Admins   True      Group       Domain            DC1              2
        Domain Admins     Enterprise Admins   True      Group       Domain            DC1              2
        Administrator     Enterprise Admins   False     User        Domain            DC1              2
        Domain Admins     Administrators      True      Group       Domain            DC1              1
        proxb             Domain Admins       False     User        Domain            DC1              2
        Administrator     Domain Admins       False     User        Domain            DC1              2
        Sysops Admins     Administrators      True      Group       Domain            DC1              1
        Org Admins        Sysops Admins       True      Group       Domain            DC1              2
        Enterprise Admins Sysops Admins       True      Group       Domain            DC1              2

        Description
        -----------
        Gets all of the members of the 'Administrators' group on the local system.

    .EXAMPLE
        Get-LocalGroupMembership -Group 'Administrators' -Depth 1

        Name              ParentGroup    isGroup    ObjectClass PrincipalSource   Computername Depth
        ----              -----------    -------    ----        ------------      -----
        Administrator     Administrators   False    User        Local             DC1              1
        boe               Administrators   False    User        Domain            DC1              1
        testuser          Administrators   False    User        Local             DC1              1
        bob               Administrators   False    User        Domain            DC1              1
        proxb             Administrators   False    User        Domain            DC1              1
        Enterprise Admins Administrators   True     Group       Domain            DC1              1
        Domain Admins     Administrators   True     Group       Domain            DC1              1
        Sysops Admins     Administrators   True     Group       Domain            DC1              1

        Description
        -----------
        Gets the members of 'Administrators' with only 1 level of recursion.

    .LINK
        Original Script: https://github.com/proxb/PowerShell_Scripts/blob/master/Get-LocalGroupMembership.ps1

    #>
    [cmdletbinding()]
    param (
        [parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [Alias('CN', '__Server', 'Computer', 'IPAddress')]
        [string[]]$Computername = $env:COMPUTERNAME,
        [parameter()]
        [string]$Group = "Administrators",
        [parameter()]
        [int]$Depth = ([int]::MaxValue),
        [parameter()]
        [Alias("MaxJobs")]
        [int]$Throttle = 10
    )
    begin {
        $PSBoundParameters.GetEnumerator() | ForEach-Object {
            Write-Verbose $_
        }
        # region Extra Configurations
        Write-Verbose ("Depth: {0}" -f $Depth)
        # endregion Extra Configurations
        # Define hash table for Get-RunspaceData function
        $runspacehash = @{}
        # Function to perform runspace job cleanup
        function Get-RunspaceData {
            [cmdletbinding()]
            param(
                [switch]$Wait
            )
            do {
                $more = $false
                foreach ($runspace in $runspaces) {
                    if ($runspace.Runspace.isCompleted) {
                        $runspace.powershell.EndInvoke($runspace.Runspace)
                        $runspace.powershell.dispose()
                        $runspace.Runspace = $null
                        $runspace.powershell = $null
                    } elseif ($runspace.Runspace -ne $null) {
                        $more = $true
                    }
                }
                if ($more -and $PSBoundParameters['Wait']) {
                    Start-Sleep -Milliseconds 100
                }
                # Clean out unused runspace jobs
                $temphash = $runspaces.clone()
                $temphash | Where-Object {
                    $_.runspace -eq $Null
                } | ForEach-Object {
                    Write-Verbose ("Removing {0}" -f $_.computer)
                    $Runspaces.remove($_)
                }
            } while ($more -and $PSBoundParameters['Wait'])
        }

        # region ScriptBlock
        $scriptBlock = {
            param ($Computer, $Group, $Depth, $NetBIOSDomain, $ObjNT, $Translate)
            $Script:Depth = $Depth
            $Script:ObjNT = $ObjNT
            $Script:Translate = $Translate
            $Script:NetBIOSDomain = $NetBIOSDomain
            function Get-LocalGroupMemberObj {
                [cmdletbinding()]
                param (
                    [parameter()]
                    [System.DirectoryServices.DirectoryEntry]$LocalGroup
                )
                # Invoke the Members method and convert to an array of member objects.
                $Members = @($LocalGroup.psbase.Invoke("Members")) | ForEach-Object { ([System.DirectoryServices.DirectoryEntry]$_) }
                $Counter++
                foreach ($Member in $Members) {
                    try {

                        $Name = $Member.InvokeGet("Name")
                        $Path = $Member.InvokeGet("AdsPath")

                        # Check if this member is a group.
                        $isGroup = ($Member.InvokeGet("Class") -eq "group")

                        # Remove the domain from the computername to fix the type comparison when supplied with FQDN
                        if ($Computer.Contains('.')) {
                            $Computer = $computer.Substring(0, $computer.IndexOf('.'))
                        }

                        if (($Path -like "*/$Computer/*")) {
                            $Type = 'Local'
                        } else { $Type = 'Domain' }

                        # Add Objectclass to match Get-LocalGroupMember output
                        if ($isGroup) {
                            $ObjectClass = 'Group'
                        } elseif ($isGroup -eq $false) {
                            $ObjectClass = 'User'
                        } else {
                            'Unknown'
                        }

                        New-Object PSObject -Property @{
                            Computername = $Computer
                            Name = $Name
                            PrincipalSource = $Type
                            ParentGroup = $LocalGroup.Name[0]
                            isGroup = $isGroup
                            ObjectClass = $ObjectClass
                            Depth = $Counter
                            Group = $Group
                        }
                        if ($isGroup) {
                            # Check if this group is local or domain.
                            # $host.ui.WriteVerboseLine("(RS)Checking if Counter: {0} is less than Depth: {1}" -f $Counter, $Depth)
                            if ($Counter -lt $Depth) {
                                if ($Type -eq 'Local') {
                                    if ($Groups[$Name] -notcontains 'Local') {
                                        $host.ui.WriteVerboseLine(("{0}: Getting local group members" -f $Name))
                                        $Groups[$Name] += , 'Local'
                                        # Enumerate members of local group.
                                        Get-LocalGroupMemberObj $Member
                                    }
                                } else {
                                    if ($Groups[$Name] -notcontains 'Domain') {
                                        $host.ui.WriteVerboseLine(("{0}: Getting domain group members" -f $Name))
                                        $Groups[$Name] += , 'Domain'
                                        # Enumerate members of domain group.
                                        Get-DomainGroupMember $Member $Name $True
                                    }
                                }
                            }
                        }
                    } catch {
                        $host.ui.WriteWarningLine(("GLGM{0}" -f $_.Exception.Message))
                    }
                }
            }

            function Get-DomainGroupMember {
                [cmdletbinding()]
                param (
                    [parameter()]
                    $DomainGroup,
                    [parameter()]
                    [string]$NTName,
                    [parameter()]
                    [string]$blnNT
                )
                try {
                    if ($blnNT -eq $True) {
                        # Convert NetBIOS domain name of group to Distinguished Name.
                        $objNT.InvokeMember("Set", "InvokeMethod", $Null, $Translate, (3, ("{0}{1}" -f $NetBIOSDomain.Trim(), $NTName)))
                        $DN = $objNT.InvokeMember("Get", "InvokeMethod", $Null, $Translate, 1)
                        $ADGroup = [ADSI]"LDAP://$DN"
                    } else {
                        $DN = $DomainGroup.distinguishedName
                        $ADGroup = $DomainGroup
                    }
                    $Counter++
                    foreach ($MemberDN in $ADGroup.Member) {
                        $MemberGroup = [ADSI]("LDAP://{0}" -f ($MemberDN -replace '/', '\/'))

                        # Add Objectclass to match Get-LocalGroupMember output
                        if ($MemberGroup.Class -eq "group") {
                            $ObjectClass = 'Group'
                        } else {
                            $ObjectClass = 'User'
                        }

                        New-Object PSObject -Property @{
                            Computername = $Computer
                            Name = $MemberGroup.name[0]
                            PrincipalSource = 'Domain'
                            ParentGroup = $NTName
                            isGroup = ($MemberGroup.Class -eq "group")
                            ObjectClass = $ObjectClass
                            Depth = $Counter
                            Group = $Group
                        }
                        # Check if this member is a group.
                        if ($MemberGroup.Class -eq "group") {
                            if ($Counter -lt $Depth) {
                                if ($Groups[$MemberGroup.name[0]] -notcontains 'Domain') {
                                    Write-Verbose ("{0}: Getting domain group members" -f $MemberGroup.name[0])
                                    $Groups[$MemberGroup.name[0]] += , 'Domain'
                                    # Enumerate members of domain group.
                                    Get-DomainGroupMember $MemberGroup $MemberGroup.Name[0] $False
                                }
                            }
                        }
                    }
                } catch {
                    $host.ui.WriteWarningLine(("GDGM{0}" -f $_.Exception.Message))
                }
            }
            # region Get Local Group Members
            $Script:Groups = @{}
            $Script:Counter = 0
            # Bind to the group object with the WinNT provider.
            $ADSIGroup = [ADSI]"WinNT://$Computer/$Group,group"
            Write-Verbose ("Checking {0} membership for {1}" -f $Group, $Computer)
            $Groups[$Group] += , 'Local'
            Get-LocalGroupMemberObj -LocalGroup $ADSIGroup
            # endregion Get Local Group Members
        }
        # endregion ScriptBlock
        Write-Verbose ("Checking to see if connected to a domain")
        try {
            $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $Root = $Domain.GetDirectoryEntry()
            $Base = ($Root.distinguishedName)

            # Use the NameTranslate object.
            $Script:Translate = New-Object -ComObject "NameTranslate"
            $Script:objNT = $Translate.GetType()

            # Initialize NameTranslate by locating the Global Catalog.
            $objNT.InvokeMember("Init", "InvokeMethod", $Null, $Translate, (3, $Null))

            # Retrieve NetBIOS name of the current domain.
            $objNT.InvokeMember("Set", "InvokeMethod", $Null, $Translate, (1, "$Base"))
            [string]$Script:NetBIOSDomain = $objNT.InvokeMember("Get", "InvokeMethod", $Null, $Translate, 3)
        } catch {
            Out-Null
        }

        # region Runspace Creation
        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
        $runspacepool.Open()

        Write-Verbose ("Creating empty collection to hold runspace jobs")
        $Script:runspaces = New-Object System.Collections.ArrayList
        # endregion Runspace Creation
    }

    process {
        foreach ($Computer in $Computername) {
            # Create the powershell instance and supply the scriptblock with the other parameters
            $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($computer).AddArgument($Group).AddArgument($Depth).AddArgument($NetBIOSDomain).AddArgument($ObjNT).AddArgument($Translate)

            # Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspacepool

            # Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell, Runspace, Computer
            $Temp.Computer = $Computer
            $temp.PowerShell = $powershell

            # Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()
            Write-Verbose ("Adding {0} collection" -f $temp.Computer)
            $runspaces.Add($temp) | Out-Null

            Write-Verbose ("Checking status of runspace jobs")
            Get-RunspaceData @runspacehash
        }
    }
    end {
        Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@(($runspaces | Where-Object { $_.Runspace -ne $Null }).Count)))
        $runspacehash.Wait = $true
        Get-RunspaceData @runspacehash

        # region Cleanup Runspace
        Write-Verbose ("Closing the runspace pool")
        $runspacepool.close()
        $runspacepool.Dispose()
        # endregion Cleanup Runspace
    }
}


function ConvertTo-HashToYN {
    <#
    .SYNOPSIS
        Used by As Built Report to convert array content true or false automatically to Yes or No.
    .DESCRIPTION
        Used by As Built Report to convert array content true or false automatically to Yes or No.
        Now also strips non-printable ASCII characters from string values while creating the array hash.
		This is required for Word Document Output as PSCribo cannot create Word documents with non-ASCII characters
    .NOTES
        Version:        0.1.1
        Author:         Jonathan Colon
        Changes:        0.1.1 - Updated to include non-unicode character string cleaning. Graham Flynn - 30/07/2025

    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param (
        [Parameter (Position = 0, Mandatory)]
        [AllowEmptyString()]
        [Hashtable] $TEXT
    )

    $result = [ordered] @{}

    foreach ($i in $inObj.GetEnumerator()) {
        try {
            $valueToProcess = $i.Value

            # Check if the value is a string before attempting to clean it
            if ($valueToProcess -is [string]) {
                $valueToProcess = $valueToProcess | Remove-NonPrintableAscii
            }

            $convertedValue = ConvertTo-TextYN $valueToProcess

            $result.add($i.Key, $convertedValue)
        } catch {
            # If ConvertTo-TextYN fails, still try to clean the original value if it's a string
            $originalValue = $i.Value
            if ($originalValue -is [string]) {
                $originalValue = $originalValue | Remove-NonPrintableAscii
            }
            $result.add($i.Key, ($originalValue)) # Add the (potentially cleaned) original value
        }
    }
    if ($result) {
        return $result
    } else {
        # If $TEXT was empty or processing failed to produce results, return the original (empty) $TEXT
        # Note: If $inObj was the source, and $TEXT is not used, this 'else' block might need review
        # based on how $TEXT is intended to be used when $inObj is empty.
        return $TEXT
    }
} # end



function Remove-NonPrintableAscii {
    <#
    .SYNOPSIS
        Removes non-printable ASCII characters from a string.
    .DESCRIPTION
        This function takes a string as input and returns a new string
        where all characters outside the printable ASCII range (ASCII 32-126)
        have been removed. If the input is null or empty, it returns an empty string.
    .PARAMETER InputString
        The string from which to remove non-printable ASCII characters.
    .EXAMPLE
        Remove-NonPrintableAscii -InputString "Hello`nWorld`t!"
        # Output: "HelloWorld!"

    .EXAMPLE
        "This string has a null character: `0" | Remove-NonPrintableAscii
        # Output: "This string has a null character: "

    .EXAMPLE
        $null | Remove-NonPrintableAscii
        # Output: "" (empty string)

    .EXAMPLE
        "" | Remove-NonPrintableAscii
        # Output: "" (empty string)
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([String])]

    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$InputString
    )

    process {
        if ($PSCmdlet.ShouldProcess($InputString, "Remove non-printable ASCII characters")) {
            # Check if the input string is null or empty.
            # If it is, return an empty string immediately to avoid errors.
            if ([string]::IsNullOrEmpty($InputString)) {
                return ""
            }

            # Regular expression to match any character that is NOT a printable ASCII character.
            # [^\x20-\x7E] matches any character that is not in the range of ASCII 32 (space) to 126 (tilde).
            $cleanedString = $InputString -replace '[^\x20-\x7E]', ''
            return $cleanedString
        }
    }
}