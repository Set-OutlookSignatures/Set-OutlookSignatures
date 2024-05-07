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
        'PublicClient*' {
            if ($PublicClientOptions) {
                $ClientApplicationBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::CreateWithApplicationOptions($PublicClientOptions)
            } else {
                $ClientApplicationBuilder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId)
            }

            ## Check Device Registration Status
            if (!$script:ModuleState.DeviceRegistrationStatus) {
                $script:ModuleState.DeviceRegistrationStatus = Get-DeviceRegistrationStatus
                $script:ModuleState.UseWebView2 = $script:ModuleFeatureSupport.WebView2Support -and ($script:ModuleState.DeviceRegistrationStatus['AzureAdPrt'] -eq 'NO' -or !$script:ModuleFeatureSupport.WebView1Support)
            }

            if ($PSBoundParameters.ContainsKey('EnableExperimentalFeatures')) { [void] $ClientApplicationBuilder.WithExperimentalFeatures($EnableExperimentalFeatures) }  # Must be called before other experimental features
            #if ($script:ModuleState.UseWebView2) { [void] [Microsoft.Identity.Client.Desktop.DesktopExtensions]::WithDesktopFeatures($ClientApplicationBuilder) }
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
        'ConfidentialClient*' {
            if ($ConfidentialClientOptions) {
                $ClientApplicationBuilder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::CreateWithApplicationOptions($ConfidentialClientOptions)
            } else {
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
        '*' {
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
