param (
    # Provide module configuration
    [Parameter(Mandatory = $false)]
    [psobject] $ModuleConfiguration
)

## Set Strict Mode for Module. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode
Set-StrictMode -Version 3.0

#Write-Warning 'The MSAL.PS PowerShell module wraps MSAL.NET functionality into PowerShell-friendly cmdlets and is not supported by Microsoft. Microsoft support does not extend beyond the underlying MSAL.NET library. For any inquiries regarding the PowerShell module itself, you may contact the author on GitHub or PowerShell Gallery.'

$script:ModuleConfigDefault = Import-Config -Path (Join-Path $PSScriptRoot 'config.json')
$script:ModuleConfig = $script:ModuleConfigDefault.psobject.Copy()
Import-Config | Set-Config
Set-Config -ResolveFromEnvironmentVariables
if ($PSBoundParameters.ContainsKey('ModuleConfiguration')) { Set-Config $ModuleConfiguration }
#Export-Config

$script:ModuleFeatureSupport = [ordered]@{
    WebView1Support   = $PSVersionTable.PSEdition -eq 'Desktop'
    WebView2Support   = [System.Environment]::OSVersion.Platform -eq 'Win32NT' -and [System.Environment]::Is64BitProcess -and ($PSVersionTable.PSVersion -lt [version]'6.0' -or $PSVersionTable.PSVersion -ge [version]'7.0' -and (Get-Item -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}' -ErrorAction SilentlyContinue))
    #EmbeddedWebViewSupport = $WebView1Support -or $WebView2Support
    DeviceCodeSupport = $true
}

## PowerShell Desktop 5.1 does not dot-source ScriptsToProcess when a specific version is specified on import. This is a bug.
# if ($PSEdition -eq 'Desktop') {
#     $ModuleManifest = Import-PowershellDataFile -LiteralPath (Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name.Replace('.psm1','.psd1'))
#     if ($ModuleManifest.ContainsKey('ScriptsToProcess')) {
#         foreach ($Path in $ModuleManifest.ScriptsToProcess) {
#             . (Join-Path $PSScriptRoot $Path)
#         }
#     }
# }

## Azure Automation module import fails when ScriptsToProcess is specified in manifest. Referencing import script directly.
. (Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name.Replace('.psm1', '.ps1'))

## Global Variables
[System.Collections.Generic.List[Microsoft.Identity.Client.IPublicClientApplication]] $PublicClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IPublicClientApplication]'
[System.Collections.Generic.List[Microsoft.Identity.Client.IConfidentialClientApplication]] $ConfidentialClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IConfidentialClientApplication]'
$script:ModuleState = @{
    DeviceRegistrationStatus = $null
    UseWebView2              = $true
}
