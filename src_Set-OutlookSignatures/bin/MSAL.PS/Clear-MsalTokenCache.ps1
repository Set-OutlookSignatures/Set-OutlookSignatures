<#
.SYNOPSIS
    Clear all client applications from local session cache.
.DESCRIPTION
    This cmdlet clear all client application objects from the local session cache.
.EXAMPLE
    PS C:\>Clear-MsalTokenCache
    Clear all client applications from local session cache.
.EXAMPLE
    PS C:\>Clear-MsalTokenCache -FromDisk
    Clear all client applications from persistent cache on disk.
#>
function Clear-MsalTokenCache {
    [CmdletBinding()]
    param(
        # Clear the token cache from disk.
        [Parameter(Mandatory = $false)]
        [switch] $FromDisk
    )

    if ($FromDisk) {
        $TokenCachePath = $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Set-OutlookSignatures\MSAL.PS\MSAL.PS.msalcache.bin3')
        if (Test-Path -LiteralPath $TokenCachePath) { Remove-Item -LiteralPath $TokenCachePath -Force }
    } else {
        $script:PublicClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IPublicClientApplication]'
        $script:ConfidentialClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IConfidentialClientApplication]'
    }
}
