function Get-RequiredFeature {
    <#
    .SYNOPSIS
        Function to check if the required version of windows feature is installed
    .DESCRIPTION
        Function to check if the required version of windows feature is installed
    .NOTES
        Version:        0.1.0
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
        [Bool]
        $Feature = $False
    )

    process {
        # Check if the required version of Module is installed
        if ($OSType -eq 'WorkStation') {
            if ($Feature) {
                $RequiredFeature = Invoke-Command -Session $TempPssSession { Get-WindowsOptionalFeature -FeatureName $Using:Name -Online }
                if ($RequiredFeature.State -ne "Enabled")  {
                    throw "$Name is required to run the Microsoft AD As Built Report. Run 'Enable-WindowsOptionalFeature -Online -FeatureName '$($Name)'' to install the required modules. https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.AD"
                }
            } else {
                $RequiredFeature = Invoke-Command -Session $TempPssSession { Get-WindowsCapability -online -Name $Using:Name }
                if ($RequiredFeature.State -ne "Installed")  {
                    throw "$Name is required to run the Microsoft AD As Built Report. Run 'Add-WindowsCapability -online -Name '$($Name)'' to install the required modules. https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.AD"
                }
            }
        }
        elseif ($OSType -eq 'Server' -or $OSType -eq 'DomainController') {
            $RequiredFeature = Invoke-Command -Session $TempPssSession { Get-WindowsFeature -Name $Using:Name }
            if ($RequiredFeature.InstallState -ne 'Installed')  {
                throw "$Name is required to run the Microsoft AD As Built Report. Run 'Install-WindowsFeature -Name '$($Name)'' to install the required modules. https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.AD"
            }
        }
        else {
            throw "Unable to validate if $Name is installed. https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.AD"
        }
    }
    end {}
}