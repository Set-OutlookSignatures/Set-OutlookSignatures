<#
.SYNOPSIS
    Enable client application to use persistent token cache on disk.
.DESCRIPTION
    This cmdlet will enable a client application object to use persistent token cache on disk.
.EXAMPLE
    PS C:\>Enable-MsalTokenCacheOnDisk $ClientApplication
    Enable client application to use persistent token cache on disk.
.EXAMPLE
    PS C:\>Enable-MsalTokenCacheOnDisk $ClientApplication -PassThru
    Enable client application to use persistent token cache on disk and return the object.
#>
function Enable-MsalTokenCacheOnDisk {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.PublicClientApplication], [Microsoft.Identity.Client.ConfidentialClientApplication])]
    param
    (
        # Public client application
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IPublicClientApplication] $PublicClientApplication,
        # Confidential client application
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication,
        # Returns client application
        [Parameter(Mandatory = $false)]
        [switch] $PassThru
    )

    switch ($PSCmdlet.ParameterSetName) {
        'PublicClient' {
            $ClientApplication = $PublicClientApplication
            break
        }
        'ConfidentialClient' {
            $ClientApplication = $ConfidentialClientApplication
            break
        }
    }

    if ([System.Environment]::OSVersion.Platform -eq 'Win32NT' -and ($PSVersionTable.PSVersion -lt [version]'6.0' -or $PSVersionTable.PSVersion -ge [version]'7.0')) {
        if ($ClientApplication -is [Microsoft.Identity.Client.IConfidentialClientApplication]) {
            [TokenCacheHelper]::EnableSerialization($ClientApplication.AppTokenCache)
        }
        [TokenCacheHelper]::EnableSerialization($ClientApplication.UserTokenCache)
    } else {
        Write-Warning 'Using TokenCache On Disk only works on Windows platform using Windows PowerShell or PowerShell 7+. The token cache will be stored in memory and not persisted on disk.'
    }

    if ($PassThru) {
        Write-Output $ClientApplication
    }
}
