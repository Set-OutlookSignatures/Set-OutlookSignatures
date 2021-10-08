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
# MIIZrAYJKoZIhvcNAQcCoIIZnTCCGZkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB0fxaxHKbkKmfD
# 3cGiQyPsFDFpBk4Jaw02pBjv9FWz6aCCFJUwggT+MIID5qADAgECAhANQkrgvjqI
# /2BAIc4UAPDdMA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNV
# BAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBUaW1lc3RhbXBpbmcgQ0EwHhcN
# MjEwMTAxMDAwMDAwWhcNMzEwMTA2MDAwMDAwWjBIMQswCQYDVQQGEwJVUzEXMBUG
# A1UEChMORGlnaUNlcnQsIEluYy4xIDAeBgNVBAMTF0RpZ2lDZXJ0IFRpbWVzdGFt
# cCAyMDIxMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAwuZhhGfFivUN
# CKRFymNrUdc6EUK9CnV1TZS0DFC1JhD+HchvkWsMlucaXEjvROW/m2HNFZFiWrj/
# ZwucY/02aoH6KfjdK3CF3gIY83htvH35x20JPb5qdofpir34hF0edsnkxnZ2OlPR
# 0dNaNo/Go+EvGzq3YdZz7E5tM4p8XUUtS7FQ5kE6N1aG3JMjjfdQJehk5t3Tjy9X
# tYcg6w6OLNUj2vRNeEbjA4MxKUpcDDGKSoyIxfcwWvkUrxVfbENJCf0mI1P2jWPo
# GqtbsR0wwptpgrTb/FZUvB+hh6u+elsKIC9LCcmVp42y+tZji06lchzun3oBc/gZ
# 1v4NSYS9AQIDAQABo4IBuDCCAbQwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQC
# MAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQQYDVR0gBDowODA2BglghkgBhv1s
# BwEwKTAnBggrBgEFBQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20vQ1BTMB8G
# A1UdIwQYMBaAFPS24SAd/imu0uRhpbKiJbLIFzVuMB0GA1UdDgQWBBQ2RIaOpLqw
# Zr68KC0dRDbd42p6vDBxBgNVHR8EajBoMDKgMKAuhixodHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLXRzLmNybDAyoDCgLoYsaHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC10cy5jcmwwgYUGCCsGAQUFBwEBBHkw
# dzAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME8GCCsGAQUF
# BzAChkNodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNz
# dXJlZElEVGltZXN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUAA4IBAQBIHNy1
# 6ZojvOca5yAOjmdG/UJyUXQKI0ejq5LSJcRwWb4UoOUngaVNFBUZB3nw0QTDhtk7
# vf5EAmZN7WmkD/a4cM9i6PVRSnh5Nnont/PnUp+Tp+1DnnvntN1BIon7h6JGA078
# 9P63ZHdjXyNSaYOC+hpT7ZDMjaEXcw3082U5cEvznNZ6e9oMvD0y0BvL9WH8dQgA
# dryBDvjA4VzPxBFy5xtkSdgimnUVQvUtMjiB2vRgorq0Uvtc4GEkJU+y38kpqHND
# Udq9Y9YfW5v3LhtPEx33Sg1xfpe39D+E68Hjo0mh+s6nv1bPull2YYlffqe0jmd4
# +TaY4cso2luHpoovMIIFJjCCBA6gAwIBAgIQCm8Gpkn9Nk686mPMJKDEczANBgkq
# hkiG9w0BAQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBT
# SEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDMzMTAwMDAwMFoX
# DTIzMDQwNTEyMDAwMFowYzELMAkGA1UEBhMCVVMxDTALBgNVBAgTBE9oaW8xEzAR
# BgNVBAcTCkNpbmNpbm5hdGkxFzAVBgNVBAoTDkphc29uIFRob21wc29uMRcwFQYD
# VQQDEw5KYXNvbiBUaG9tcHNvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAMVnygZO0wvpQ3NjGpEREqp0r/FN5C0X0Gn6HxrhPXAyGZaBlEjV0eO6bz8N
# BVFwyHsQ0BFxT7CrGvCCwvekm7bqIZaIJe9kFYAvOVBDK+S042dGaT8cUSxU6QIk
# gXL2IZKZu8R8H0+26rehGpadj+onbqzFshaS8C18/1oFv27W/3FeOwAkXbE8Mbpu
# c9ntR/6PUV4biw3AYUITVps0PmfTB1f06DmrbWa3orHVDO1yEL/E1hoe0jpXPAHz
# vtNlLMtZg5LeRrGdkfasq8V94XicNWU8XFy6D5cFlIg0RPcSzMJRJb78nfpQInrp
# DAagviDCUVR5ZwLsvDk096h8kCUCAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrE
# uXsqCqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBS+5845JPvDWenjXahLo4XUCcTn
# MjAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAw
# bjA1oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1j
# cy1nMS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtY3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYB
# BQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGE
# BggrBgEFBQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBOBggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQC
# MAAwDQYJKoZIhvcNAQELBQADggEBAER9rMHu+w+qJrQmh6at6GrAPYuHi2zuU04n
# dRRzTSmHUKvzS1DvEYxLp6cO//3gHEqBV1S0YV58Rn5idMii7fmANSfO1Og4x77/
# CmmnpwB8aoSCpbRxqcIBE+pUm7r7JBT4xNEKT3FkgcpVymE4VuIscBgnekEmmaVf
# Doh1Xm4cQ+hvtyZ8+3+bNQ/Oe008RSk5zmiWiS++eGeB1D5v6yLs2bHAHldKKCp8
# Mg322VqRB2C9bFlQSxS97FB/s4J4jGxjSSl6MmcYLzkw+Copc5/9c1QEzBe+9rZM
# aAPwb6e977tkFtFOCfiekESAjku2NPqjj83EtLOOllrv3r81oWcwggUwMIIEGKAD
# AgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0x
# MzEwMjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0Ew
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZ
# z9D7RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnko
# On7p0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXU
# LaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q
# 8grkV7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5S
# lsHyDxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPR
# AgMBAAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIB
# hjATBgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBG
# MDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt
# 9mV1DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcN
# AQELBQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L
# /e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xG
# Tlz/kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGiv
# m6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZ
# Hen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdi
# bqFT+hKUGIUukpHqaGxEMrJmoecYpJpkUe8wggUxMIIEGaADAgECAhAKoSXW1jIb
# fkHkBdo2l8IVMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xNjAxMDcxMjAwMDBa
# Fw0zMTAxMDcxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBUaW1lc3RhbXBpbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQC90DLuS82Pf92puoKZxTlUKFe2I0rEDgdFM1EQ
# fdD5fU1ofue2oPSNs4jkl79jIZCYvxO8V9PD4X4I1moUADj3Lh477sym9jJZ/l9l
# P+Cb6+NGRwYaVX4LJ37AovWg4N4iPw7/fpX786O6Ij4YrBHk8JkDbTuFfAnT7l3I
# mgtU46gJcWvgzyIQD3XPcXJOCq3fQDpct1HhoXkUxk0kIzBdvOw8YGqsLwfM/fDq
# R9mIUF79Zm5WYScpiYRR5oLnRlD9lCosp+R1PrqYD4R/nzEU1q3V8mTLex4F0IQZ
# chfxFwbvPc3WTe8GQv2iUypPhR3EHTyvz9qsEPXdrKzpVv+TAgMBAAGjggHOMIIB
# yjAdBgNVHQ4EFgQU9LbhIB3+Ka7S5GGlsqIlssgXNW4wHwYDVR0jBBgwFoAUReui
# r/SSy4IxLVGLp6chnfNtyA8wEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8E
# BAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwUAYDVR0g
# BEkwRzA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4IBAQBx
# lRLpUYdWac3v3dp8qmN6s3jPBjdAhO9LhL/KzwMC/cWnww4gQiyvd/MrHwwhWiq3
# BTQdaq6Z+CeiZr8JqmDfdqQ6kw/4stHYfBli6F6CJR7Euhx7LCHi1lssFDVDBGiy
# 23UC4HLHmNY8ZOUfSBAYX4k4YU1iRiSHY4yRUiyvKYnleB/WCxSlgNcSR3CzddWT
# hZN+tpJn+1Nhiaj1a5bA9FhpDXzIAbG5KHW3mWOFIoxhynmUfln8jA/jb7UBJrZs
# pe6HUSHkWGCbugwtK22ixH67xCUrRwIIfEmuE7bhfEJCKMYYVs9BNLZmXbZ0e/VW
# MyIvIjayS6JKldj1po5SMYIEbTCCBGkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QQIQCm8Gpkn9Nk686mPMJKDEczANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD+YGci2d3W
# Qc5w1STmSqokKJbb7o0aFyRR16yJd1Ll5zANBgkqhkiG9w0BAQEFAASCAQBnvhRM
# XkjWrd4JYrnCPiTfCDYfT1jCFHCPcrGNYBtPUL2VkloXP62q3CoLM8BRCLDopB7J
# SI6TBK4GIhpF8WrpIXZ1nIesCHrocIuQiDf7MStmGdwmfvYgoeEbGIHrWFJ2tmos
# 14NqZxaHhH7KCPR7VgiOUtM6kR+qyPDDCzXLz5YhoQs0nGh8E3SID/1hb++uiMuY
# jn4b5EEMEantWN1iS+F+VSOGblk7tl105AFU80enRedleuT12qOxb9OfzkovVrgW
# MHhDCEJnGpliXNOy8N+VO3cLi0V5ccp0AB+9LKHnjuPy+Ix7LUk2k9AS2pMWbnIE
# xArZTGdq4KSqYLiVoYICMDCCAiwGCSqGSIb3DQEJBjGCAh0wggIZAgEBMIGGMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBUaW1lc3RhbXBpbmcgQ0ECEA1CSuC+Ooj/YEAhzhQA8N0wDQYJYIZIAWUDBAIB
# BQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0y
# MTA5MjExMzI3NTNaMC8GCSqGSIb3DQEJBDEiBCBFQQshl8cUJtUxx984G1zfluuH
# Gt4eaQYbTiw+7RXu3jANBgkqhkiG9w0BAQEFAASCAQBXhRRs3OPVW0Dy9Ctv6t0y
# 1enkBFLalceqjZrj9AkQfxao/nhB2ChRIEzhyCyWDJTWVRBXzSnEXAtJuPlpumCm
# Hio9b40FWpZ++0E9XUrLn3lhRgPIA+8cq1FCrbbpQEbf8ouZz+hyi3tgXClf4Dd6
# l87hcuAt/6/jxUGPvM1MYQUvGkNMyenNo2I/eqzflH5xEU+nlcfVcXZr1ogu4jgF
# 5JGoX/ckv3wRqCZleK2FnfsEbpyl4eZSexHkg4BJTJ/VcgfwfNwL/fRXtmvPcwrV
# 2/7t9YqnrMATTmi62OwZz2Sey6naficyfRqtMx+shzIXQiPwkPYUfCo8lsr7xwBa
# SIG # End signature block
