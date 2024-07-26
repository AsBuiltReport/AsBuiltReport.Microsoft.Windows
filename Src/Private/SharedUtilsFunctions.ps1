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

Function Get-LocalGroupMembership {
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
    Param (
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
    Begin {
        $PSBoundParameters.GetEnumerator() | ForEach {
            Write-Verbose $_
        }
        # region Extra Configurations
        Write-Verbose ("Depth: {0}" -f $Depth)
        # endregion Extra Configurations
        # Define hash table for Get-RunspaceData function
        $runspacehash = @{}
        # Function to perform runspace job cleanup
        Function Get-RunspaceData {
            [cmdletbinding()]
            param(
                [switch]$Wait
            )
            Do {
                $more = $false         
                Foreach ($runspace in $runspaces) {
                    If ($runspace.Runspace.isCompleted) {
                        $runspace.powershell.EndInvoke($runspace.Runspace)
                        $runspace.powershell.dispose()
                        $runspace.Runspace = $null
                        $runspace.powershell = $null                 
                    }
                    ElseIf ($runspace.Runspace -ne $null) {
                        $more = $true
                    }
                }
                If ($more -AND $PSBoundParameters['Wait']) {
                    Start-Sleep -Milliseconds 100
                }   
                # Clean out unused runspace jobs
                $temphash = $runspaces.clone()
                $temphash | Where {
                    $_.runspace -eq $Null
                } | ForEach {
                    Write-Verbose ("Removing {0}" -f $_.computer)
                    $Runspaces.remove($_)
                }             
            } while ($more -AND $PSBoundParameters['Wait'])
        }

        # region ScriptBlock
        $scriptBlock = {
            Param ($Computer, $Group, $Depth, $NetBIOSDomain, $ObjNT, $Translate)            
            $Script:Depth = $Depth
            $Script:ObjNT = $ObjNT
            $Script:Translate = $Translate
            $Script:NetBIOSDomain = $NetBIOSDomain
            Function Get-LocalGroupMember {
                [cmdletbinding()]
                Param (
                    [parameter()]
                    [System.DirectoryServices.DirectoryEntry]$LocalGroup
                )
                # Invoke the Members method and convert to an array of member objects.
                $Members = @($LocalGroup.psbase.Invoke("Members")) | foreach { ([System.DirectoryServices.DirectoryEntry]$_) }
                $Counter++
                ForEach ($Member In $Members) {                
                    Try {
						
                        $Name = $Member.InvokeGet("Name")
                        $Path = $Member.InvokeGet("AdsPath")

                        # Check if this member is a group.
                        $isGroup = ($Member.InvokeGet("Class") -eq "group")

                        # Remove the domain from the computername to fix the type comparison when supplied with FQDN
                        IF ($Computer.Contains('.')) {
                            $Computer = $computer.Substring(0, $computer.IndexOf('.'))
                        }

                        If (($Path -like "*/$Computer/*")) {
                            $Type = 'Local'
                        }
                        Else { $Type = 'Domain' }
						
                        # Add Objectclass to match Get-LocalGroupMember output
                        if ($isGroup) {
                            $ObjectClass = 'Group'
                        }
                        elseif ($isGroup -eq $false) {
                            $ObjectClass = 'User'
                        }
                        else {
                            'Unknown'
                        }

                        New-Object PSObject -Property @{
                            Computername    = $Computer
                            Name            = $Name
                            PrincipalSource = $Type
                            ParentGroup     = $LocalGroup.Name[0]
                            isGroup         = $isGroup
                            ObjectClass     = $ObjectClass
                            Depth           = $Counter
                            Group           = $Group
                        }
                        If ($isGroup) {
                            # Check if this group is local or domain.
                            # $host.ui.WriteVerboseLine("(RS)Checking if Counter: {0} is less than Depth: {1}" -f $Counter, $Depth)
                            If ($Counter -lt $Depth) {
                                If ($Type -eq 'Local') {
                                    If ($Groups[$Name] -notcontains 'Local') {
                                        $host.ui.WriteVerboseLine(("{0}: Getting local group members" -f $Name))
                                        $Groups[$Name] += , 'Local'
                                        # Enumerate members of local group.
                                        Get-LocalGroupMember $Member
                                    }
                                }
                                Else {
                                    If ($Groups[$Name] -notcontains 'Domain') {
                                        $host.ui.WriteVerboseLine(("{0}: Getting domain group members" -f $Name))
                                        $Groups[$Name] += , 'Domain'
                                        # Enumerate members of domain group.
                                        Get-DomainGroupMember $Member $Name $True
                                    }
                                }
                            }
                        }
                    }
                    Catch {
                        $host.ui.WriteWarningLine(("GLGM{0}" -f $_.Exception.Message))
                    }
                }
            }

            Function Get-DomainGroupMember {
                [cmdletbinding()]
                Param (
                    [parameter()]
                    $DomainGroup, 
                    [parameter()]
                    [string]$NTName, 
                    [parameter()]
                    [string]$blnNT
                )
                Try {
                    If ($blnNT -eq $True) {
                        # Convert NetBIOS domain name of group to Distinguished Name.
                        $objNT.InvokeMember("Set", "InvokeMethod", $Null, $Translate, (3, ("{0}{1}" -f $NetBIOSDomain.Trim(), $NTName)))
                        $DN = $objNT.InvokeMember("Get", "InvokeMethod", $Null, $Translate, 1)
                        $ADGroup = [ADSI]"LDAP://$DN"
                    }
                    Else {
                        $DN = $DomainGroup.distinguishedName
                        $ADGroup = $DomainGroup
                    }         
                    $Counter++   
                    ForEach ($MemberDN In $ADGroup.Member) {
                        $MemberGroup = [ADSI]("LDAP://{0}" -f ($MemberDN -replace '/', '\/'))
						
                        # Add Objectclass to match Get-LocalGroupMember output
                        if ($MemberGroup.Class -eq "group") {
                            $ObjectClass = 'Group'
                        }
                        else {
                            $ObjectClass = 'User'
                        }
						
                        New-Object PSObject -Property @{
                            Computername    = $Computer
                            Name            = $MemberGroup.name[0]
                            PrincipalSource = 'Domain'
                            ParentGroup     = $NTName
                            isGroup         = ($MemberGroup.Class -eq "group")
                            ObjectClass     = $ObjectClass
                            Depth           = $Counter
                            Group           = $Group
                        }
                        # Check if this member is a group.
                        If ($MemberGroup.Class -eq "group") {              
                            If ($Counter -lt $Depth) {
                                If ($Groups[$MemberGroup.name[0]] -notcontains 'Domain') {
                                    Write-Verbose ("{0}: Getting domain group members" -f $MemberGroup.name[0])
                                    $Groups[$MemberGroup.name[0]] += , 'Domain'
                                    # Enumerate members of domain group.
                                    Get-DomainGroupMember $MemberGroup $MemberGroup.Name[0] $False
                                }                                                
                            }
                        }
                    }
                }
                Catch {
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
            Get-LocalGroupMember -LocalGroup $ADSIGroup
            # endregion Get Local Group Members
        }
        # endregion ScriptBlock
        Write-Verbose ("Checking to see if connected to a domain")
        Try {
            $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $Root = $Domain.GetDirectoryEntry()
            $Base = ($Root.distinguishedName)

            # Use the NameTranslate object.
            $Script:Translate = New-Object -comObject "NameTranslate"
            $Script:objNT = $Translate.GetType()

            # Initialize NameTranslate by locating the Global Catalog.
            $objNT.InvokeMember("Init", "InvokeMethod", $Null, $Translate, (3, $Null))

            # Retrieve NetBIOS name of the current domain.
            $objNT.InvokeMember("Set", "InvokeMethod", $Null, $Translate, (1, "$Base"))
            [string]$Script:NetBIOSDomain = $objNT.InvokeMember("Get", "InvokeMethod", $Null, $Translate, 3)  
        }
        Catch { # Write-Warning ("{0}" -f $_.Exception.Message) 
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

    Process {
        ForEach ($Computer in $Computername) {
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
    End {
        Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@(($runspaces | Where { $_.Runspace -ne $Null }).Count)))
        $runspacehash.Wait = $true
        Get-RunspaceData @runspacehash
    
        # region Cleanup Runspace
        Write-Verbose ("Closing the runspace pool")
        $runspacepool.close()  
        $runspacepool.Dispose() 
        # endregion Cleanup Runspace    
    }
}
