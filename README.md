<p align="center">
    <a href="https://www.asbuiltreport.com/" alt="AsBuiltReport"></a>
            <img src='https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport/master/AsBuiltReport.png' width="8%" height="8%" /></a>
</p>
<p align="center">
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.Windows/" alt="PowerShell Gallery Version">
        <img src="https://img.shields.io/powershellgallery/v/AsBuiltReport.Microsoft.Windows.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.Windows/" alt="PS Gallery Downloads">
        <img src="https://img.shields.io/powershellgallery/dt/AsBuiltReport.Microsoft.Windows.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.Windows/" alt="PS Platform">
        <img src="https://img.shields.io/powershellgallery/p/AsBuiltReport.Microsoft.Windows.svg" /></a>
</p>
<p align="center">
    <a href="https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows/graphs/commit-activity" alt="GitHub Last Commit">
        <img src="https://img.shields.io/github/last-commit/AsBuiltReport/AsBuiltReport.Microsoft.Windows/master.svg" /></a>
    <a href="https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows/master/LICENSE" alt="GitHub License">
        <img src="https://img.shields.io/github/license/AsBuiltReport/AsBuiltReport.Microsoft.Windows.svg" /></a>
    <a href="https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows/graphs/contributors" alt="GitHub Contributors">
        <img src="https://img.shields.io/github/contributors/AsBuiltReport/AsBuiltReport.Microsoft.Windows.svg"/></a>
</p>
<p align="center">
    <a href="https://twitter.com/AsBuiltReport" alt="Twitter">
            <img src="https://img.shields.io/twitter/follow/AsBuiltReport.svg?style=social"/></a>
</p>
<p align="center">
    <a href='https://ko-fi.com/B0B7DDGZ7' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://cdn.ko-fi.com/cdn/kofi1.png?v=3' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>
</p>

# Microsoft Windows As Built Report

Microsoft Windows As Built Report is a PowerShell module which works in conjunction with [AsBuiltReport.Core](https://github.com/AsBuiltReport/AsBuiltReport.Core).

[AsBuiltReport](https://github.com/AsBuiltReport/AsBuiltReport) is an open-sourced community project which utilises PowerShell to produce as-built documentation in multiple document formats for multiple vendors and technologies.

Please refer to the AsBuiltReport [website](https://www.asbuiltreport.com) for more detailed information about this project.

# :books: Sample Reports

## Sample Report - Custom Style 1

Sample Microsoft Windows As Built report HTML file: [Sample Microsoft Windows As-Built Report.html](https://htmlpreview.github.io/?https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows/dev/Samples/Sample%20Microsoft%20Windows%20As%20Built%20Report.html "Sample Microsoft Windows As-Built Report")

Sample Microsoft Windows As Built report PDF file: [Sample Microsoft Windows As Built Report.pdf](https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows/dev/Samples/Sample%20Microsoft%20Windows%20As%20Built%20Report.pdf)

# :beginner: Getting Started

Below are the instructions on how to install, configure and generate a Microsoft Windows As Built report.

## :floppy_disk: Supported Versions
<!-- ********** Update supported Windows versions ********** -->
The Microsoft Windows As Built Report supports the following Windows Server versions;

- 2012, 2016, 2019 & 2022

### PowerShell

This report is compatible with the following PowerShell versions;

<!-- ********** Update supported PowerShell versions ********** -->
| Windows PowerShell 5.1 |     PowerShell 7    |
|:----------------------:|:--------------------:|
|   :white_check_mark:   | :x: |

## :wrench: System Requirements
<!-- ********** Update system requirements ********** -->
PowerShell 5.1 and the following PowerShell modules are required for generating a Microsoft Windows As Built report.

- [AsBuiltReport.Microsoft.Windows Module](https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.Windows/)
- [IISAdministration Module](https://docs.microsoft.com/en-us/powershell/module/iisadministration/?view=windowsserver2022-ps)
- [SmbShare Module](https://docs.microsoft.com/en-us/powershell/module/smbshare/?view=windowsserver2022-ps)
- [Hyper-V Module](https://docs.microsoft.com/en-us/powershell/module/hyper-v/?view=windowsserver2022-ps)
- [DhcpServer Module](https://docs.microsoft.com/en-us/powershell/module/dhcpserver/?view=windowsserver2019-ps)
- [DnsServer Module](https://docs.microsoft.com/en-us/powershell/module/dnsserver/?view=windowsserver2019-ps)

### Linux & macOS

This report does not support Linux or Mac due to the fact that the Windows modules are dependent on the .NET Framework. Until Microsoft migrates these modules to native PowerShell Core, only PowerShell >= 5.1.x will be supported on Windows.

### :closed_lock_with_key: Required Privileges

A Microsoft Windows As Built Report can be generated with Administrator level privileges. Since this report relies extensively on the WinRM component, you should make sure that it is enabled and configured

## :package: Module Installation

The installation of the modules will depend on the roles that are being served on the server to be documented.

### PowerShell v5.x running on a Windows server (Target)

```powershell
Install-Module AsBuiltReport.Microsoft.Windows

# DNS/DHCP Server powershell modules
Install-WindowsFeature -Name RSAT-DNS-Server
Install-WindowsFeature -Name RSAT-DHCP

# Hyper-V Server powershell modules
Install-WindowsFeature -Name Hyper-V-PowerShell

#IIS Server powershell modules
Install-WindowsFeature -Name web-mgmt-console
Install-WindowsFeature -Name Web-Scripting-Tools

#FailOver Cluster powershell modules
Install-WindowsFeature -Name RSAT-Clustering-PowerShell

```

### PowerShell v5.x running on Windows client computer
<!-- ********** Add installation for any additional PowerShell module(s) ********** -->
```powershell
Install-Module AsBuiltReport.Microsoft.Windows

# DNS/DHCP Server powershell Modules
Add-WindowsCapability –online –Name 'Rsat.Dns.Tools~~~~0.0.1.0'
Add-WindowsCapability -Online -Name 'Rsat.DHCP.Tools~~~~0.0.1.0'

#FailOver Cluster powershell modules
Add-WindowsCapability -Online -Name 'Rsat.FailoverCluster.Management.Tools~~~~0.0.1.0'

# Hyper-V Server powershell modules
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell

#IIS Server powershell modules
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerManagementTools
Enable-WindowsOptionalFeature -Online -FeatureName IIS-ManagementScriptingTools

```

### GitHub

If you are unable to use the PowerShell Gallery, you can still install the module manually. Ensure you repeat the following steps for the [system requirements](https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows#wrench-system-requirements) also.

1. Download the code package / [latest release](https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows/releases/latest) zip from GitHub
2. Extract the zip file
3. Copy the folder `AsBuiltReport.Microsoft.Windows` to a path that is set in `$env:PSModulePath`.
4. Open a PowerShell terminal window and unblock the downloaded files with

    ```powershell
    $path = (Get-Module -Name AsBuiltReport.Microsoft.Windows -ListAvailable).ModuleBase; Unblock-File -Path $path\*.psd1; Unblock-File -Path $path\Src\Public\*.ps1; Unblock-File -Path $path\Src\Private\*.ps1
    ```

5. Close and reopen the PowerShell terminal window.

_Note: You are not limited to installing the module to those example paths, you can add a new entry to the environment variable PSModulePath if you want to use another path._

## :pencil2: Configuration

The Microsoft Windows As Built Report utilises a JSON file to allow configuration of report information, options, detail and healthchecks.

A Microsoft Windows report configuration file can be generated by executing the following command;

```powershell
New-AsBuiltReportConfig -Report Microsoft.Windows -FolderPath <User specified folder> -Filename <Optional>
```

Executing this command will copy the default Microsoft Windows report JSON configuration to a user specified folder.

All report settings can then be configured via the JSON file.

The following provides information of how to configure each schema within the report's JSON file.

### Report

The **Report** schema provides configuration of the Microsoft Windows report information.

| Sub-Schema          | Setting      | Default                        | Description                                                  |
|---------------------|--------------|--------------------------------|--------------------------------------------------------------|
| Name                | User defined | Microsoft Windows As Built Report | The name of the As Built Report                              |
| Version             | User defined | 1.0                            | The report version                                           |
| Status              | User defined | Released                       | The report release status                                    |
| ShowCoverPageImage  | true / false | true                           | Toggle to enable/disable the display of the cover page image |
| ShowTableOfContents | true / false | true                           | Toggle to enable/disable table of contents                   |
| ShowHeaderFooter    | true / false | true                           | Toggle to enable/disable document headers & footers          |
| ShowTableCaptions   | true / false | true                           | Toggle to enable/disable table captions/numbering            |

### Options

The **Options** schema allows certain options within the report to be toggled on or off.

### InfoLevel

The **InfoLevel** schema allows configuration of each section of the report at a granular level. The following sections can be set.

There are 3 levels (0-2) of detail granularity for each section as follows;

| Setting | InfoLevel         | Description                                                                                                                                |
|:-------:|-------------------|--------------------------------------------------------------------------------------------------------------------------------------------|
|    0    | Disabled          | Does not collect or display any information                                                                                                |
|    1    | Enabled           | Provides summarised information for a collection of objects                                                                                |
|    2    | Adv Summary       | Provides condensed, detailed information for a collection of objects                                                                       |

The table below outlines the default and maximum **InfoLevel** settings for each section.

| Sub-Schema   | Default Setting | Maximum Setting |
|--------------|:---------------:|:---------------:|
| Hardware       |        1        |        1        |
| OperatingSystem       |        1        |        2        |
| Storage          |        1        |        1        |
| Networking         |        1        |        1        |
| IIS           |        1        |        1        |
| HyperV           |        1        |        1        |
| DHCP           |        1        |        2        |
| DNS           |        1        |        2        |

### Healthcheck

The **Healthcheck** schema is used to toggle health checks on or off.

## :computer: Examples

There are a few examples listed below on running the AsBuiltReport script against a Microsoft Windows server target. Refer to the `README.md` file in the main AsBuiltReport project repository for more examples.

```powershell

# Generate a Microsoft Windows As Built Report for Server 'win-server-01v.contoso.local' using specified credentials. Export report to HTML & DOCX formats. Use default report style. Append timestamp to report filename. Save reports to 'C:\Users\Jon\Documents'
PS C:\> New-AsBuiltReport -Report Microsoft.Windows -Target 'win-server-01v.contoso.local' -Username 'administrator@contoso.local' -Password 'P@ssw0rd' -Format Html,Word -OutputFolderPath 'C:\Users\Jon\Documents' -Timestamp

# Generate a Microsoft Windows As Built Report for Server 'win-server-01v.contoso.local' using specified credentials and report configuration file. Export report to Text, HTML & DOCX formats. Use default report style. Save reports to 'C:\Users\Jon\Documents'. Display verbose messages to the console.
PS C:\> New-AsBuiltReport -Report Microsoft.Windows -Target 'win-server-01v.contoso.local' -Username 'administrator@contoso.local' -Password 'P@ssw0rd' -Format Text,Html,Word -OutputFolderPath 'C:\Users\Jon\Documents' -ReportConfigFilePath 'C:\Users\Jon\AsBuiltReport\AsBuiltReport.Microsoft.Windows.json' -Verbose

# Generate a Microsoft Windows As Built Report for Server 'win-server-01v.contoso.local' using stored credentials. Export report to HTML & Text formats. Use default report style. Highlight environment issues within the report. Save reports to 'C:\Users\Jon\Documents'.
PS C:\> $Creds = Get-Credential
PS C:\> New-AsBuiltReport -Report Microsoft.Windows -Target 'win-server-01v.contoso.local' -Credential $Creds -Format Html,Text -OutputFolderPath 'C:\Users\Jon\Documents' -EnableHealthCheck

# Generate a Microsoft Windows As Built Report for Server 'win-server-01v.contoso.local' using specified credentials. Export report to HTML & DOCX formats. Use default report style. Reports are saved to the user profile folder by default. Attach and send reports via e-mail.
PS C:\> New-AsBuiltReport -Report Microsoft.Windows -Target 'win-server-01v.contoso.local' -Username 'administrator@contoso.local' -Password 'P@ssw0rd' -Format Html,Word -OutputFolderPath 'C:\Users\Jon\Documents' -SendEmail
```

## :x: Known Issues

- Issues with WinRM when using the IP address instead of the "Fully Qualified Domain Name".
- The report provides the ability to extract the configuration of the DNS/DHCP/Hyper-V/IIS/FailOver-Cluster services. In order to obtain this information it is required that the servers running these services have the corresponding powershell modules installed.
