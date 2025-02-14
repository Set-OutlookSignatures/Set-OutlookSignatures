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


    if ([System.Environment]::OSVersion.Platform -eq 'Win32NT' -and $PSVersionTable.PSVersion -lt [version]'6.0') {
        if ($ClientApplication -is [Microsoft.Identity.Client.IConfidentialClientApplication]) {
            [TokenCacheHelper]::EnableSerialization($ClientApplication.AppTokenCache)
        }
        [TokenCacheHelper]::EnableSerialization($ClientApplication.UserTokenCache)

        $ClientApplication | Add-Member -MemberType NoteProperty -Name 'cacheInfo' -Value "Encrypted file '$([TokenCacheHelper]::CacheFilePath)'"
    } else {
        $cacheFilePath = $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Set-OutlookSignatures\MSAL.PS\MSAL.PS.msalcache.bin3')
        $cacheFileName = [System.IO.Path]::GetFileName($cacheFilePath)
        $cacheDir = [System.IO.Path]::GetDirectoryName($cacheFilePath)

        if ($IsWindows) {
            $storageProperties = [Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder]::new(
                $cacheFileName,
                $cacheDir
            ).
            WithCacheChangedEvent(
                $ClientApplication.AppConfig.ClientId,
                $ClientApplication.AppConfig.Authority.AuthorityInfo.CanonicalAuthority.AbsoluteUri
            ).
            CustomizeLockRetry(1000, 3).
            Build()
        } elseif ( $IsLinux ) {
            $storageProperties = [Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder]::new(
                $cacheFileName,
                $cacheDir
            ).
            WithCacheChangedEvent(
                $ClientApplication.AppConfig.ClientId,
                $ClientApplication.AppConfig.Authority.AuthorityInfo.CanonicalAuthority.AbsoluteUri
            ).
            WithLinuxKeyring(
                'at.explicitconsulting.setoutlooksignatures.tokencache',
                [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::LinuxKeyRingDefaultCollection,
                'Set-OutlookSignatures Microsoft Graph token via MSAL.Net',
                (New-Object 'System.Collections.Generic.KeyValuePair[String, String]' -ArgumentList 'Version', '1'),
                (New-Object 'System.Collections.Generic.KeyValuePair[String, String]' -ArgumentList 'Product', 'Set-OutlookSignatures')
            ).
            CustomizeLockRetry(1000, 3).
            Build()
        } elseif ($IsMacOS) {
            $storageProperties = [Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder]::new(
                $cacheFileName,
                $cacheDir
            ).
            WithCacheChangedEvent(
                $ClientApplication.AppConfig.ClientId,
                $ClientApplication.AppConfig.Authority.AuthorityInfo.CanonicalAuthority.AbsoluteUri
            ).
            WithMacKeyChain(
                'Set-OutlookSignatures Microsoft Graph token via MSAL.Net',
                'Set-OutlookSignatures Microsoft Graph token via MSAL.Net'
            ).
            CustomizeLockRetry(1000, 3).
            Build()
        }

        $cacheHelper = [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::CreateAsync(
            $storageProperties
        ).
        GetAwaiter().
        GetResult()

        try {
            $cacheHelper.VerifyPersistence()

            $ClientApplication | Add-Member -MemberType NoteProperty -Name 'cacheInfo' -Value $(
                if ($IsWindows) {
                    "Encrypted file '$($cacheFilePath)'"
                } elseif ($IsLinux) {
                    "Encrypted default keyring entry 'Set-OutlookSignatures Microsoft Graph token via MSAL.Net'"
                } elseif ($IsMacOS) {
                    "Encrypted default keychain entry 'Set-OutlookSignatures Microsoft Graph token via MSAL.Net'"
                }
            )
        } catch {
            if ($IsWindows -or $IsMacOS) {
                $storageProperties = [Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder]::new(
                    $cacheFileName,
                    $cacheDir
                ).
                WithCacheChangedEvent(
                    $ClientApplication.AppConfig.ClientId,
                    $ClientApplication.AppConfig.Authority.AuthorityInfo.CanonicalAuthority.AbsoluteUri
                ).
                WithUnprotectedFile().
                CustomizeLockRetry(1000, 3).
                Build()
            } elseif ($IsLinux) {
                $storageProperties = [Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder]::new(
                    $cacheFileName,
                    $cacheDir
                ).
                WithCacheChangedEvent(
                    $ClientApplication.AppConfig.ClientId,
                    $ClientApplication.AppConfig.Authority.AuthorityInfo.CanonicalAuthority.AbsoluteUri
                ).
                WithLinuxUnprotectedFile().
                CustomizeLockRetry(1000, 3).
                Build()
            }

            $cacheHelper = [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::CreateAsync(
                $storageProperties
            ).
            GetAwaiter().
            GetResult()

            $ClientApplication | Add-Member -MemberType NoteProperty -Name 'cacheInfo' -Value $(
                "Unencrypted file '$($cacheFilePath)'"
            )
        }

        if ($ClientApplication -is [Microsoft.Identity.Client.IConfidentialClientApplication]) {
            $cacheHelper.RegisterCache($ClientApplication.AppTokenCache)
        }

        $cacheHelper.RegisterCache($ClientApplication.UserTokenCache)
    }

    if ($PassThru) {
        Write-Output $ClientApplication
    }
}
