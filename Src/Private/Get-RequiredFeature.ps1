function Get-RequiredFeature {
    <#
    .SYNOPSIS
        Function to check if the required version of windows feature is installed
    .DESCRIPTION
        Function to check if the required version of windows feature is installed
    .NOTES
        Version:        0.5.2
        Author:         Jonathan Colon
        Twitter:        @jcolonfzenpr
        Github:         rebelinux
    .PARAMETER Name
        The name of the required windows feature
    .PARAMETER Version
        The version of the required windows feature
    #>

    Param
    (
        [CmdletBinding()]
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $OSType,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [Switch]
        $Feature = $False,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [Switch]
        $Status = $False,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Service
    )

    process {
        # Check if the required version of Module is installed
        if ($OSType -eq 'WorkStation') {
            if ($Feature) {
                $RequiredFeature = Invoke-Command -Session $TempPssSession { Get-WindowsOptionalFeature -FeatureName $Using:Name -Online }
                if ($Status) {
                    if ($RequiredFeature.State -eq "Enabled") {
                        return $true
                    } else {
                        return $false
                    }
                }
                if (-Not $Status) {
                    if ($RequiredFeature.State -ne "Enabled") {
                        Write-PScriboMessage -IsWarning "$Name module is required to be installed on $System to be able to document $Service service. Run 'Enable-WindowsOptionalFeature -Online -FeatureName '$($Name)'' to install the required modules."
                    }
                }
            } else {
                $RequiredFeature = Invoke-Command -Session $TempPssSession { Get-WindowsCapability -Online -Name $Using:Name }
                if ($Status) {
                    if ($RequiredFeature.State -eq "Installed") {
                        return $true
                    } else {
                        return $false
                    }
                }
                if (-Not $Status) {
                    if ($RequiredFeature.State -ne "Installed") {
                        Write-PScriboMessage -IsWarning "$Name module is required to be installed on $System to be able to document $Service service. Run 'Add-WindowsCapability -online -Name '$($Name)'' to install the required modules."
                    }
                }
            }
        } elseif ($OSType -eq 'Server' -or $OSType -eq 'DomainController') {
            $RequiredFeature = Invoke-Command -Session $TempPssSession { Get-WindowsFeature -Name $Using:Name }
            if ($Status) {
                if ($RequiredFeature.InstallState -eq 'Installed') {
                    return $true
                } else {
                    return $false
                }
            }
            if (-Not $Status) {
                if ($RequiredFeature.InstallState -ne 'Installed') {
                    Write-PScriboMessage -IsWarning "$Name module is required to be installed on $System to be able to document $Service service. Run 'Install-WindowsFeature -Name '$($Name)'' to install the required modules."
                }
            }
        }
        else {
            throw "Unable to validate if $Name is installed. https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows"
        }
    }
    end {}
}