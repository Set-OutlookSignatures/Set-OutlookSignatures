<#
.SYNOPSIS
    Create new client application.
.DESCRIPTION
    This cmdlet will return a new client application object which can be used with the Get-MsalToken cmdlet.
.EXAMPLE
    PS C:\>New-MsalClientApplication -ClientId '00000000-0000-0000-0000-000000000000'
    Get public client application using default settings.
.EXAMPLE
    PS C:\>$PublicClientOptions = New-Object Microsoft.Identity.Client.PublicClientApplicationOptions -Property @{ ClientId = '00000000-0000-0000-0000-000000000000' }
    PS C:\>$PublicClientOptions | New-MsalClientApplication -TenantId '00000000-0000-0000-0000-000000000000'
    Pipe in public client options object to get a public client application and target a specific tenant.
.EXAMPLE
    PS C:\>$ClientCertificate = Get-Item Cert:\CurrentUser\My\0000000000000000000000000000000000000000
    PS C:\>$ConfidentialClientOptions = New-Object Microsoft.Identity.Client.ConfidentialClientApplicationOptions -Property @{ ClientId = '00000000-0000-0000-0000-000000000000'; TenantId = '00000000-0000-0000-0000-000000000000' }
    PS C:\>$ConfidentialClientOptions | New-MsalClientApplication -ClientCertificate $ClientCertificate
    Pipe in confidential client options object to get a confidential client application using a client certificate and target a specific tenant.
#>
function New-MsalClientApplication {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.PublicClientApplication], [Microsoft.Identity.Client.ConfidentialClientApplication])]
    param
    (
        # Identifier of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', Position = 1, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientSecret', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientClaims', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientAssertion', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', Position = 1, ValueFromPipelineByPropertyName = $true)]
        [string] $ClientId,
        # Secure secret of the client requesting the token.
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClientSecret', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        #[AllowNull()]
        [securestring] $ClientSecret,
        # Client assertion certificate of the client requesting the token.
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientCertificate', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientClaims', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $ClientCertificate,
        # Set the specific client claims to sign. ClientCertificate must also be specified.
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientClaims', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [hashtable] $ClientClaims,
        # Set client assertion used to prove the identity of the application to Azure AD. This is a Base-64 encoded JWT.
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClientAssertion', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'ConfidentialClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [string] $ClientAssertion,
        # Address to return to upon receiving a response from the authority.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [uri] $RedirectUri,
        # Instance of Azure Cloud
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.AzureCloudInstance] $AzureCloudInstance,
        # Tenant identifier of the authority to issue token.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string] $TenantId,
        # Address of the authority to issue token.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [uri] $Authority,
        # Use Platform Authentication Broker
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient', ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $false, ParameterSetName = 'PublicClient-InputObject', ValueFromPipelineByPropertyName = $true)]
        [switch] $AuthenticationBroker,
        # Sets Extra Query Parameters for the query string in the HTTP authentication request.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [hashtable] $ExtraQueryParameters,
        # Allows usage of experimental features and APIs.
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [switch] $EnableExperimentalFeatures,
        # Add Application and TokenCache to list for this PowerShell session.
        #[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
        #[switch] $AddToSessionCache,
        # Read and save encrypted TokenCache to disk for persistance across PowerShell sessions.
        #[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
        #[switch] $UseTokenCacheOnDisk,
        # Public client application options
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient-InputObject', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.PublicClientApplicationOptions] $PublicClientOptions,
        # Confidential client application options
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClient-InputObject', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Microsoft.Identity.Client.ConfidentialClientApplicationOptions] $ConfidentialClientOptions
    )

    switch -Wildcard ($PSCmdlet.ParameterSetName) {
        "PublicClient*" {
            if ($PublicClientOptions) {
                $ClientApplicationBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::CreateWithApplicationOptions($PublicClientOptions)
            }
            else {
                $ClientApplicationBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId)
            }

            ## Check Device Registration Status
            if (!$script:ModuleState.DeviceRegistrationStatus) {
                $script:ModuleState.DeviceRegistrationStatus = Get-DeviceRegistrationStatus
                $script:ModuleState.UseWebView2 = $script:ModuleFeatureSupport.WebView2Support -and ($script:ModuleState.DeviceRegistrationStatus['AzureAdPrt'] -eq 'NO' -or !$script:ModuleFeatureSupport.WebView1Support)
            }

            if ($PSBoundParameters.ContainsKey('EnableExperimentalFeatures')) { [void] $ClientApplicationBuilder.WithExperimentalFeatures($EnableExperimentalFeatures) }  # Must be called before other experimental features
            if ($script:ModuleState.UseWebView2) { [void] [Microsoft.Identity.Client.Desktop.DesktopExtensions]::WithDesktopFeatures($ClientApplicationBuilder) }
            if ($RedirectUri) { [void] $ClientApplicationBuilder.WithRedirectUri($RedirectUri.AbsoluteUri) }
            elseif (!$PublicClientOptions -or !$PublicClientOptions.RedirectUri) {
                if ($script:ModuleState.UseWebView2) { [void] $ClientApplicationBuilder.WithRedirectUri('https://login.microsoftonline.com/common/oauth2/nativeclient') }
                else { [void] $ClientApplicationBuilder.WithDefaultRedirectUri() }
            }
            if ($PSBoundParameters.ContainsKey('AuthenticationBroker')) {
                if ([System.Environment]::OSVersion.Platform -eq 'Win32NT') { [void] [Microsoft.Identity.Client.Desktop.WamExtension]::WithWindowsBroker($ClientApplicationBuilder, $AuthenticationBroker) }
                else { [void] $ClientApplicationBuilder.WithBroker($AuthenticationBroker) }
            }

            $ClientOptions = $PublicClientOptions
        }
        "ConfidentialClient*" {
            if ($ConfidentialClientOptions) {
                $ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::CreateWithApplicationOptions($ConfidentialClientOptions)
            }
            else {
                $ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ClientId)
            }

            if ($PSBoundParameters.ContainsKey('EnableExperimentalFeatures')) { [void] $ClientApplicationBuilder.WithExperimentalFeatures($EnableExperimentalFeatures) }  # Must be called before other experimental features
            if ($ClientSecret) { [void] $ClientApplicationBuilder.WithClientSecret((ConvertFrom-SecureStringAsPlainText $ClientSecret -Force)) }
            if ($ClientAssertion) { [void] $ClientApplicationBuilder.WithClientAssertion($ClientAssertion) }
            if ($ClientClaims) { [void] $ClientApplicationBuilder.WithClientClaims($ClientCertificate, (ConvertTo-Dictionary $ClientClaims -KeyType ([string]) -ValueType ([string]))) }
            elseif ($ClientCertificate) { [void] $ClientApplicationBuilder.WithCertificate($ClientCertificate) }
            if ($RedirectUri) { [void] $ClientApplicationBuilder.WithRedirectUri($RedirectUri.AbsoluteUri) }

            $ClientOptions = $ConfidentialClientOptions
        }
        "*" {
            if ($ClientId) { [void] $ClientApplicationBuilder.WithClientId($ClientId) }
            if ($AzureCloudInstance -and $TenantId) { [void] $ClientApplicationBuilder.WithAuthority($AzureCloudInstance, $TenantId) }
            elseif ($TenantId) { [void] $ClientApplicationBuilder.WithTenantId($TenantId) }
            if ($Authority) { [void] $ClientApplicationBuilder.WithAuthority($Authority) }
            if (!$ClientOptions -or !($ClientOptions.ClientName -or $ClientOptions.ClientVersion)) {
                [void] $ClientApplicationBuilder.WithClientName("PowerShell $($PSVersionTable.PSEdition)")
                [void] $ClientApplicationBuilder.WithClientVersion($PSVersionTable.PSVersion)
            }
            if ($ExtraQueryParameters) { [void] $ClientApplicationBuilder.WithExtraQueryParameters((ConvertTo-Dictionary $ExtraQueryParameters -KeyType ([string]) -ValueType ([string]))) }
            #[void] $ClientApplicationBuilder.WithLogging($null, [Microsoft.Identity.Client.LogLevel]::Verbose, $false, $true)

            $ClientApplication = $ClientApplicationBuilder.Build()
            break
        }
    }

    ## Add to local PowerShell session cache.
    # if ($AddToSessionCache) {
    #     Add-MsalClientApplication $ClientApplication
    # }

    ## Enable custom serialization of TokenCache to disk
    # if ($UseTokenCacheOnDisk) {
    #     Enable-MsalTokenCacheOnDisk $ClientApplication
    # }

    return $ClientApplication
}

# SIG # Begin signature block
# MIIjggYJKoZIhvcNAQcCoIIjczCCI28CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCekiPmjD+mO57Z
# TKT6+L6otfOpwVJmrsDsK0W/fHdO3qCCDYEwggX/MIID56ADAgECAhMzAAACUosz
# qviV8znbAAAAAAJSMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjEwOTAyMTgzMjU5WhcNMjIwOTAxMTgzMjU5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDQ5M+Ps/X7BNuv5B/0I6uoDwj0NJOo1KrVQqO7ggRXccklyTrWL4xMShjIou2I
# sbYnF67wXzVAq5Om4oe+LfzSDOzjcb6ms00gBo0OQaqwQ1BijyJ7NvDf80I1fW9O
# L76Kt0Wpc2zrGhzcHdb7upPrvxvSNNUvxK3sgw7YTt31410vpEp8yfBEl/hd8ZzA
# v47DCgJ5j1zm295s1RVZHNp6MoiQFVOECm4AwK2l28i+YER1JO4IplTH44uvzX9o
# RnJHaMvWzZEpozPy4jNO2DDqbcNs4zh7AWMhE1PWFVA+CHI/En5nASvCvLmuR/t8
# q4bc8XR8QIZJQSp+2U6m2ldNAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUNZJaEUGL2Guwt7ZOAu4efEYXedEw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDY3NTk3MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAFkk3
# uSxkTEBh1NtAl7BivIEsAWdgX1qZ+EdZMYbQKasY6IhSLXRMxF1B3OKdR9K/kccp
# kvNcGl8D7YyYS4mhCUMBR+VLrg3f8PUj38A9V5aiY2/Jok7WZFOAmjPRNNGnyeg7
# l0lTiThFqE+2aOs6+heegqAdelGgNJKRHLWRuhGKuLIw5lkgx9Ky+QvZrn/Ddi8u
# TIgWKp+MGG8xY6PBvvjgt9jQShlnPrZ3UY8Bvwy6rynhXBaV0V0TTL0gEx7eh/K1
# o8Miaru6s/7FyqOLeUS4vTHh9TgBL5DtxCYurXbSBVtL1Fj44+Od/6cmC9mmvrti
# yG709Y3Rd3YdJj2f3GJq7Y7KdWq0QYhatKhBeg4fxjhg0yut2g6aM1mxjNPrE48z
# 6HWCNGu9gMK5ZudldRw4a45Z06Aoktof0CqOyTErvq0YjoE4Xpa0+87T/PVUXNqf
# 7Y+qSU7+9LtLQuMYR4w3cSPjuNusvLf9gBnch5RqM7kaDtYWDgLyB42EfsxeMqwK
# WwA+TVi0HrWRqfSx2olbE56hJcEkMjOSKz3sRuupFCX3UroyYf52L+2iVTrda8XW
# esPG62Mnn3T8AuLfzeJFuAbfOSERx7IFZO92UPoXE1uEjL5skl1yTZB3MubgOA4F
# 8KoRNhviFAEST+nG8c8uIsbZeb08SeYQMqjVEmkwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVVzCCFVMCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAlKLM6r4lfM52wAAAAACUjAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgJxKP//pI
# 575O/vio5H5qNNh3MRfYl3otcegPIqjIE4QwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQBjyfBuFXaR87Q2MBNGNLyxZag70cWQCoEUqLq2m5tS
# 8uLdINGs/SYqJbS/esn5WnwxSZhisgRakr2Rk6i7UHEbqDpXVwvON0X5NjyEXJNd
# mtLV/KMEsu5FddQWfrbp5MrtCaTqXBWP6y151URwztB6FbW6dKuWJTkVWf6jrI+6
# FWk1BacoXvJ7bHSY7ca6XvbrTrifXfzkifbFPAy3iIfP63y3sP6ghf5IdE1psRfg
# SiYoUJRbXSItncO+PwT1G2RywNEUHiNdoTxdIovIF2LYqHUzAn2zQJD+Qt0/RVbG
# IMTleZ2iFpcphCBSnUdM2wPyhxWGqo0XJPL+8MnFaAlDoYIS4TCCEt0GCisGAQQB
# gjcDAwExghLNMIISyQYJKoZIhvcNAQcCoIISujCCErYCAQMxDzANBglghkgBZQME
# AgEFADCCAVAGCyqGSIb3DQEJEAEEoIIBPwSCATswggE3AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIAakCWGuIOww2w5YIAIvhyZ1ahbtQ4iYXUVnpkPU
# gl6KAgZhkuE5IO4YEjIwMjExMTE5MDIzNzQ3LjI4WjAEgAIB9KCB0KSBzTCByjEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWlj
# cm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEmMCQGA1UECxMdVGhhbGVzIFRTUyBF
# U046MTJCQy1FM0FFLTc0RUIxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2Wggg45MIIE8TCCA9mgAwIBAgITMwAAAVPSgnJFbFfjiwAAAAABUzAN
# BgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
# MDExMTIxODI2MDVaFw0yMjAyMTExODI2MDVaMIHKMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBP
# cGVyYXRpb25zMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjoxMkJDLUUzQUUtNzRF
# QjElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBALXrtnQHsxv5s0IC5cLoQYmwW+BUHqLf
# T+OXiCa1TChAbnPdanFAc0mP63Yjh+kRNBAujGKgdKJH6dskVGdPdYKNzpTBsYEy
# dAmDODNVeh9U8cgKKaJg0SIfLzo+8ISlZPy9vqN8Vxo5Wgx77jA3Y2puU5YDijrC
# YRBQWatukpkH5xFUkXtYTUvo0N9GI8T1dF9GT7PNl34wmGzd5ZGvuNV0bXS9USVX
# eGrRgXN+GjuC4/cvRszGKRHLek97hbqDtDB/kxzEOkzQBq3P0I2SxwA/KHiNk4XR
# /t2IvUEYRP9zi+nyJBOV8qoSEJu3cWL8ernhzBeMTVpTHNC9rlv95u0CAwEAAaOC
# ARswggEXMB0GA1UdDgQWBBTGootzSCqpszAZQLfgcG2sMVZEdTAfBgNVHSMEGDAW
# gBTVYzpcijGQ80N7fEYbxTNoWoVtVTBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8v
# Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNUaW1TdGFQQ0Ff
# MjAxMC0wNy0wMS5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1RpbVN0YVBDQV8yMDEw
# LTA3LTAxLmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0G
# CSqGSIb3DQEBCwUAA4IBAQCUaL15fPogRD7HAXNJolHcnEexqJ6DKb0OxFdyl93t
# +uVNRn/mJWTIWrxhfIqgQj7ttU7MmVN/c8XxR1gLLjKxZp7qzMj5uzPb1uhihGXB
# JrF1A6D2HCVfUyR90yeEn4iLDfkRVinusxWN032LppkeV0o/88NA2767MQehso0u
# lmzF74sj+B3G4iIUkM7c715O6fTvFYd+GcjEnbBLvLk8Qel3FlNBfae9ZHRuN2a3
# mvqHLlfSw0PVdmJpTZqmXVZv0drZuc7gDyWt7vBmNL6LxcFVUloTgteZxMV5CWVm
# V2+rzIGa8OjjC+MU4VVySrIFKkIwwmlGrOSh0YPEnTo0MIIGcTCCBFmgAwIBAgIK
# YQmBKgAAAAAAAjANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMTAwNzAxMjEzNjU1WhcNMjUwNzAxMjE0
# NjU1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAKkdDbx3EYo6IOz8E5f1+n9plGt0VBDVpQoAgoX7
# 7XxoSyxfxcPlYcJ2tz5mK1vwFVMnBDEfQRsalR3OCROOfGEwWbEwRA/xYIiEVEMM
# 1024OAizQt2TrNZzMFcmgqNFDdDq9UeBzb8kYDJYYEbyWEeGMoQedGFnkV+BVLHP
# k0ySwcSmXdFhE24oxhr5hoC732H8RsEnHSRnEnIaIYqvS2SJUGKxXf13Hz3wV3Ws
# vYpCTUBR0Q+cBj5nf/VmwAOWRH7v0Ev9buWayrGo8noqCjHw2k4GkbaICDXoeByw
# 6ZnNPOcvRLqn9NxkvaQBwSAJk3jN/LzAyURdXhacAQVPIk0CAwEAAaOCAeYwggHi
# MBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBTVYzpcijGQ80N7fEYbxTNoWoVt
# VTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0T
# AQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNV
# HR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEE
# TjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDCBoAYDVR0gAQH/BIGVMIGS
# MIGPBgkrBgEEAYI3LgMwgYEwPQYIKwYBBQUHAgEWMWh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9QS0kvZG9jcy9DUFMvZGVmYXVsdC5odG0wQAYIKwYBBQUHAgIwNB4y
# IB0ATABlAGcAYQBsAF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0AGUAbQBlAG4AdAAu
# IB0wDQYJKoZIhvcNAQELBQADggIBAAfmiFEN4sbgmD+BcQM9naOhIW+z66bM9TG+
# zwXiqf76V20ZMLPCxWbJat/15/B4vceoniXj+bzta1RXCCtRgkQS+7lTjMz0YBKK
# dsxAQEGb3FwX/1z5Xhc1mCRWS3TvQhDIr79/xn/yN31aPxzymXlKkVIArzgPF/Uv
# eYFl2am1a+THzvbKegBvSzBEJCI8z+0DpZaPWSm8tv0E4XCfMkon/VWvL/625Y4z
# u2JfmttXQOnxzplmkIz/amJ/3cVKC5Em4jnsGUpxY517IW3DnKOiPPp/fZZqkHim
# bdLhnPkd/DjYlPTGpQqWhqS9nhquBEKDuLWAmyI4ILUl5WTs9/S/fmNZJQ96LjlX
# dqJxqgaKD4kWumGnEcua2A5HmoDF0M2n0O99g/DhO3EJ3110mCIIYdqwUB5vvfHh
# AN/nMQekkzr3ZUd46PioSKv33nJ+YWtvd6mBy6cJrDm77MbL2IK0cs0d9LiFAR6A
# +xuJKlQ5slvayA1VmXqHczsI5pgt6o3gMy4SKfXAL1QnIffIrE7aKLixqduWsqdC
# osnPGUFN4Ib5KpqjEWYw07t0MkvfY3v1mYovG8chr1m1rtxEPJdQcdeh0sVV42ne
# V8HR3jDA/czmTfsNv11P6Z0eGTgvvM9YBS7vDaBQNdrvCScc1bN+NR4Iuto229Nf
# j950iEkSoYICyzCCAjQCAQEwgfihgdCkgc0wgcoxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9w
# ZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjEyQkMtRTNBRS03NEVC
# MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYF
# Kw4DAhoDFQCKSk3txw7WT08oIYK9pBYrInnRm6CBgzCBgKR+MHwxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA5UFUHjAiGA8yMDIx
# MTExOTA2MzcxOFoYDzIwMjExMTIwMDYzNzE4WjB0MDoGCisGAQQBhFkKBAExLDAq
# MAoCBQDlQVQeAgEAMAcCAQACAg7NMAcCAQACAhGXMAoCBQDlQqWeAgEAMDYGCisG
# AQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMB
# hqAwDQYJKoZIhvcNAQEFBQADgYEAG2i4uDXRN94X4kYuk0sQB4jtiWdXxeD6sjxT
# oKHyfq3VF0iwHR0Hw0e/ASE1dBN6jjTv1qyJ9wQ7SNFRMDx4G9cin8Wi+esynvLx
# edTVqC59HAXqH2eX/V7fiKYgNiOgBddSeP3ykOp91zhwfrCnKWDEQD6gxWmUxNVd
# ouccIqExggMNMIIDCQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MAITMwAAAVPSgnJFbFfjiwAAAAABUzANBglghkgBZQMEAgEFAKCCAUowGgYJKoZI
# hvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCBXJoayw4MrY2cu
# Z2ACTagS4gHqEL5B0x9yr6N9/wrGJjCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQw
# gb0EIFDBCo85tCAICfyXoZBDppodLIMcb2wOH2rEBWiNtY8AMIGYMIGApH4wfDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAFT0oJyRWxX44sAAAAAAVMw
# IgQgWfBxFcn9g02heXSYl9qd04fGetKE4cuAIsgomAnxTvcwDQYJKoZIhvcNAQEL
# BQAEggEAem5QJel9mY9gBDEWi8sgmzrljJ6yFfaJLc83aWYDE1oMExt0VaJGXKh6
# z+sZ+5EWGpTOTA27WwX2ktr3ysnWcGpq9eSrQy6JLfUJhph4DrPQEc9XVh8J+82W
# 9JEqc7YG4ID79S/E1albeVrBuhVe5y8O0PAsSmH+1bU1/rr/OphbkDWkoWjAe9GG
# 3UG8+0OqSEhwM+NGKyKYgG0KroGS8t8CLjE79HqrKCTO7Ty5BC6BUGCK5PIQvjv6
# XZixPzMdUL+1JePY8M7cxHH6IY3GqQCnJmzVpJQyXB3MVtIFlsqj/VHjFFU5AKIj
# l65QsqOJu37duDW+hQZNVMqt9UjrCA==
# SIG # End signature block
